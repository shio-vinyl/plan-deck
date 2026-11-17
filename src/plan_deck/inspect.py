"""读一份 pptx：逐页列形状（名字、位置 cm、文字、字号），并做几项检查。

组内的形状也逐个查，位置换算到页面上。
每个形状标 placement（inside 版面内、crossing 跨边、outside 版面外、unknown 读不到位置）；版面外的内容用 plan-deck offcanvas 看。
检查：页面绝不出现的用语（免责和口径说明、待定和核查、解说和自我澄清，见 words.BANNED；不查隐藏页和版面外）、
比正文小的字（给了 style 时）、同页重名的 pd: 形状、画布外的对象（只列出，不动）、动画引用了不存在的形状、图片的实际 dpi。
版面警告（按字宽估算，不查隐藏页，看渲染图确认）：字放不下、文字或图片部分出界、文字或图被后放的不透明形状盖住。
警告（不影响退出码，看过后改或报告）：整段或整栏标红（给了 style 时）、页面上的“第 N 页”引用。
线索类警告（不查隐藏页，要逐条判断，见 docs/visual-language.md“不当讲解员”“框和卡片”）：讲解员口吻和讨论口语、
偏长的标题和解释性的副标题、阴影、同排等尺寸的模块、字多却没有强调色的页、多页相同的备注。
"""
import re

from pptx import Presentation
from pptx.enum.dml import MSO_COLOR_TYPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.shapes.group import GroupShape
from pptx.util import Emu

from .audit import SHORT_AT, _label_rows, _page, _pictures, _Theme, content_bottom, content_page, mid_layer
from .clean import duplicate_notes
from .measure import LINE, PT_CM, text_lines, text_units
from .offcanvas import placements
from .words import SPOKEN, banned

CM = 360000
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
# 整段标红：一段至少这么多字、强调色占这么多；整栏标红：一个文本框至少这么多段、强调色占这么多（docs/visual-language.md“加粗与标红”）
RED_PARA = (40, 0.8)
RED_BLOCK = (3, 0.5)
HEADER_SLOTS = ('title', 'parent-title', 'subtitle', 'conclusion')      # 页眉和红字结论可以整句用强调色
PAGE_REF = re.compile(r'第\s*\d+\s*页')
# 讲解员口吻（docs/visual-language.md“不当讲解员”）。词表只是线索，讲解员能换说法，同样的词也常是正当内容，只作警告
PAGE_WORDS = [
    (re.compile(r'不是[^，。；\n]{0,16}[，,]\s*(?:而)?是'), '讲解员口吻：替读者纠正理解'),
    (re.compile(r'本页|本表|这一页|上表|下表|怎么查|如何查'), '讲解员口吻：教读者怎么读'),
    (re.compile(r'要解决什么|要解决的问题|从案例看|怎么[判写算做]|这个案例的'), '讲解员口吻：带读者往下读'),
    (re.compile(r'靠的是|要做的[，,]?是|关键在于|回答[^，。\n]{0,8}问题|解决[“"]'), '讲解员口吻：替读者点重点'),
    (re.compile(r'向[^，。、\n]{1,8}要[^，。、\n]{0,4}$|让[^，。\n]{1,8}(?:算清|看懂|明白)'), '讲解员口吻：替读者总结意义'),
    (SPOKEN, '讨论时的口语'),
]
# 标题和副标题（docs/visual-language.md“标题”）：带事实的结果句可以有逗号，只提醒明显偏长的；真实汇报的标题中位 11—33 字
TITLE_MAX, LEAD_MAX = 36, 18        # 汉字宽
JUDGING_PAGES = ('question', 'conclusion')
NO_RED_CHARS = 200      # 正文这么多字以上才查有没有强调色，封面、目录一般到不了
LEAD_NAMES = ('subtitle', 'header:lead')      # 其他页型的 :lead 是正文里的说明条，不是副标题
EDGE = 0.05 * CM        # 出界的容差
COVERED_TEXT, COVERED_PICTURE = 0.5, 0.9      # 被后放的形状盖住多少才算：文本框常比字宽，图片压一部分常是有意的
LABEL_ROW_SHARE = 0.1   # 一册里用左色块＋右底框标签行的页超过这个比例就提醒（docs/patterns.md“标签行”）
LINK_INSET = 0.12 * CM  # 连线离文字框边这么远以内不算穿过（端点粘在框边上）


def _cm(v):
    return round(Emu(v or 0) / CM, 2)


def _sizes(shape):
    out = set()
    frames = []
    if shape.has_text_frame:
        frames.append(shape.text_frame)
    if getattr(shape, 'has_table', False) and shape.has_table:
        frames += [c.text_frame for row in shape.table.rows for c in row.cells]
    for tf in frames:
        for p in tf.paragraphs:
            for r in p.runs:
                if r.font.size is not None and r.text.strip():
                    out.add(r.font.size.pt)
    return out


def _text(shape):
    if shape.has_text_frame:
        return shape.text_frame.text
    if getattr(shape, 'has_table', False) and shape.has_table:
        return ' | '.join(c.text_frame.text for c in shape.table.rows[0].cells)
    return ''


def _red(shape, emphasis, body):
    """正文字号里强调色的字数：[(段落字数, 强调色字数)]。大号数字（比正文大的字）不计。"""
    out = []
    for p in shape.text_frame.paragraphs:
        n = red = 0
        for r in p.runs:
            k = len(''.join(r.text.split()))
            if not k or (r.font.size is not None and r.font.size.pt > body + 0.01):
                continue
            n += k
            c = r.font.color
            if c.type == MSO_COLOR_TYPE.RGB and str(c.rgb) == emphasis:
                red += k
        if n:
            out.append((n, red))
    return out


def _red_warnings(shape, emphasis, body):
    paras = _red(shape, emphasis, body)
    total, red = sum(n for n, _ in paras), sum(r for _, r in paras)
    if len(paras) >= RED_BLOCK[0] and red >= RED_BLOCK[1] * total:
        return [f'整栏标红：{len(paras)} 段里强调色占 {red}/{total} 字']
    return [f'整段标红：第 {i} 段强调色占 {r}/{n} 字' for i, (n, r) in enumerate(paras, 1)
            if n >= RED_PARA[0] and r >= RED_PARA[1] * n]


def _all_text(shape):
    if getattr(shape, 'has_table', False) and shape.has_table:
        return '\n'.join(c.text_frame.text for row in shape.table.rows for c in row.cells)
    return shape.text_frame.text if shape.has_text_frame else ''


def _word_warnings(shape):
    text = _all_text(shape)
    return [f'{why}“{m.group()}”：页面只呈现内容，解释和说明放讲稿或备注' for rx, why in PAGE_WORDS for m in [rx.search(text)] if m]


def _word_problems(shape):
    return [f'{kind}“{word}”：页面绝不出现，删掉或移到备注（改已有稿时原稿里的也一样）'
            for kind, word, _ in banned(_all_text(shape))]


def _title_warnings(page_type, slot, text):
    text = text.replace('\n', ' ').strip()
    out = []
    if slot == 'title' and page_type not in JUDGING_PAGES and text_units(text) > TITLE_MAX:
        out.append(f'标题约 {text_units(text):.0f} 字，偏长：是否带了解释')
    if (slot == 'subtitle' or f'{page_type}:{slot}' in LEAD_NAMES) and text:
        if re.search(r'[，,；;。]', text) or text_units(text) > LEAD_MAX:
            out.append('副标题像一句解释：默认不写，只用于区分同题页、声明页面性质、交代口径')
    return out


def _theme_shadows(master):
    """主题三档效果样式各自有没有阴影：[False, True, True]。读不到主题时按都有算。"""
    from lxml import etree
    for rel in master.part.rels.values():
        if rel.reltype.endswith('/theme'):
            lst = etree.fromstring(rel.target_part.blob).find('.//' + A + 'effectStyleLst')
            if lst is not None:
                return [es.find('.//' + A + 'outerShdw') is not None for es in lst]
    return [True, True, True]


def _has_shadow(shape, shadows=(True, True, True)):
    el = shape._element
    sp = el.find(P + 'spPr')
    if sp is None:
        return False
    effects = sp.find(A + 'effectLst')
    if effects is not None:
        return effects.find(A + 'outerShdw') is not None
    style = el.find(P + 'style')
    ref = style.find(A + 'effectRef') if style is not None else None
    idx = int(ref.get('idx', '0')) if ref is not None else 0
    return idx > 0 and (shadows[idx - 1] if idx <= len(shadows) else True)


def _styled(shape):
    sp = shape._element.find(P + 'spPr')
    if sp is None or 'PICTURE' in str(shape.shape_type):
        return False
    line = sp.find(A + 'ln')
    return (sp.find(A + 'solidFill') is not None or sp.find(A + 'gradFill') is not None
            or (line is not None and line.find(A + 'noFill') is None and len(line) > 0))


def _reddish(hexv):
    r, g, b = (int(hexv[i:i + 2], 16) for i in (0, 2, 4))
    return r > 140 and r - g >= 80 and r - b >= 80      # 含橙红：手工稿的强调色不全是正红


def _no_emphasis_warning(slide, emphasis):
    """正文字多却完全没有强调色。给了 style 时比强调色，没给时认红、橙红的字；主题色认不出；含组内形状和表格，页眉不算。"""
    total = 0
    for el in slide._element.iter(P + 'sp', P + 'graphicFrame'):
        name = el.find('.//' + P + 'cNvPr')
        if name is not None and name.get('name', '').startswith('pd:header'):
            continue
        for r in el.iter(A + 'r'):
            text = (r.find(A + 't').text or '').strip()
            total += len(text)
            clr = r.find(A + 'rPr/' + A + 'solidFill/' + A + 'srgbClr')
            if text and clr is not None:
                val = clr.get('val', '').upper()
                if val == emphasis or (not emphasis and len(val) == 6 and _reddish(val)):
                    return None
    return f'正文约 {total} 字，没有强调色：引领词、关键短语按条目标' if total >= NO_RED_CHARS else None


def _grid_warning(slide, W, H):
    """同排至少 3 个近似同尺寸、里面有字的带框形状。文字和底框常是两个形状，按位置判断。"""
    conts, texts = [], []
    for sh in slide.shapes:
        x, y, w, h = sh.left or 0, sh.top or 0, sh.width or 0, sh.height or 0
        if sh.has_text_frame and len(sh.text_frame.text.strip()) >= 15:
            texts.append((x + w / 2, y + h / 2))
        if _styled(sh) and W * H * .006 < w * h < W * H * .85 and w > W * .03 and h > H * .03:
            conts.append((x, y, w, h))
    full = [c for c in conts if any(c[0] <= a <= c[0] + c[2] and c[1] <= b <= c[1] + c[3] for a, b in texts)]
    best = 0
    for c in full:
        same = {round(o[0] / W, 2) for o in full
                if abs(o[1] - c[1]) < H * .015 and abs(o[2] - c[2]) < W * .015 and abs(o[3] - c[3]) < H * .015}
        best = max(best, len(same))
    return f'同排 {best} 个等尺寸的模块：形状有没有表示关系，还是只把文字分组' if best >= 3 else None


def _walk(shapes, sx=1, sy=1, ox=0, oy=0, group=None):
    """按叠放顺序给出每个形状（含组内）：(形状, x, y, w, h, 所在组的名字)。组内坐标换算到页面上。"""
    for sh in shapes:
        x, y, w, h = ox + (sh.left or 0) * sx, oy + (sh.top or 0) * sy, (sh.width or 0) * sx, (sh.height or 0) * sy
        yield sh, x, y, w, h, group
        if isinstance(sh, GroupShape):
            xf = sh._element.grpSpPr.find(A + 'xfrm')
            che = None if xf is None else xf.find(A + 'chExt')
            if che is None:     # 组没写变换：子形状的坐标就是外层的
                yield from _walk(sh.shapes, sx, sy, ox, oy, sh.name)
                continue
            cho = xf.find(A + 'chOff')
            kx, ky = w / (int(che.get('cx')) or 1), h / (int(che.get('cy')) or 1)
            yield from _walk(sh.shapes, kx, ky, x - int(cho.get('x')) * kx, y - int(cho.get('y')) * ky, sh.name)


def _text_need(sh, w):
    """(字排出来约多高 cm, 最大字号 pt)。不换行、自动缩字、有没写字号的字时估不了，给 None。一段里字号不同时，换行按平均字宽算。"""
    tf = sh.text_frame
    if tf.word_wrap is False or tf.auto_size == MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE:
        return None
    side, need, top = (tf.margin_left + tf.margin_right) / CM, (tf.margin_top + tf.margin_bottom) / CM, 0
    for p in tf.paragraphs:
        runs = [r for r in p.runs if r.text.strip()]
        if not runs:
            continue
        if any(r.font.size is None for r in runs):
            return None
        pt = max(r.font.size.pt for r in runs)
        text = ''.join(r.text for r in runs)
        mean = sum(text_units(r.text) * r.font.size.pt for r in runs) / (text_units(text) or 1)
        spacing = p.line_spacing if isinstance(p.line_spacing, float) else 1.0
        need += text_lines(text, w / CM - side + 0.15, mean) * pt * PT_CM * LINE * spacing
        top = max(top, pt)
    return (need, top) if top else None


def _opaque(sh):
    """盖在上面会挡住下层的形状：图片、表格、不透明的实色填充。"""
    if 'PICTURE' in str(sh.shape_type) or getattr(sh, 'has_table', False):
        return True
    solid = sh._element.find(P + 'spPr/' + A + 'solidFill')
    return solid is not None and solid.find('.//' + A + 'alpha') is None


def _overlap(a, b):
    w = min(a[0] + a[2], b[0] + b[2]) - max(a[0], b[0])
    h = min(a[1] + a[3], b[1] + b[3]) - max(a[1], b[1])
    return w * h if w > 0 and h > 0 else 0


def _layout_warnings(placed, W, H, place):
    """一页的版面线索。placed：_walk 给的 (形状, x, y, w, h, 组)，已去掉画布外的和组本身；place：placements() 的结果。"""
    out = []
    for k, (sh, x, y, w, h, _) in enumerate(placed):
        words = _all_text(sh).strip()
        if not words and 'PICTURE' not in str(sh.shape_type):
            continue
        need = _text_need(sh, w) if words and sh.has_text_frame else None
        bottom = max(y + h, y + need[0] * CM) if need else y + h
        if need and len(words) > 1 and need[0] > h / CM + need[1] * PT_CM * LINE / 2:      # 单个字符是运算符、记号，压不到别的
            out.append((sh, f'字约需 {need[0]:.1f} cm 高，框只有 {_cm(h)} cm，会压到下面：拆分或换版式，不要缩字号'))
        if words and (x < -EDGE or y < -EDGE or x + w > W + EDGE or bottom > H + EDGE):
            out.append((sh, '文字部分出了页面'))
        elif not words and place.get(sh._element) == 'crossing':
            out.append((sh, '图片部分在版面外，渲染图上看不全：plan-deck offcanvas 导出了原图'))
        limit = COVERED_TEXT if words else COVERED_PICTURE
        top = next((o[0] for o in placed[k + 1:] if _opaque(o[0]) and _overlap((x, y, w, h), o[1:5]) >= limit * w * h > 0), None)
        if top is not None:
            out.append((sh, f'{"文字" if words else "图片"}被后放的 {top.name} 盖住：要在上面就后画或调叠放顺序，透明图忽略'))
    return out


def _link_path(sh, x, y, w, h):
    """连接线在页面上的折点：直线两点；bentConnector3 横—竖—横（中段位置取 adj1）；bentConnector2 横—竖。"""
    xf = sh._element.find(P + 'spPr/' + A + 'xfrm')
    fh = xf is not None and xf.get('flipH') == '1'
    fv = xf is not None and xf.get('flipV') == '1'
    rot = int(xf.get('rot', '0')) if xf is not None else 0
    x0, y0 = (x + w, x)[not fh], (y + h, y)[not fv]
    x1, y1 = (x, x + w)[not fh], (y, y + h)[not fv]
    geom = sh._element.find('.//' + A + 'prstGeom')
    kind = geom.get('prst') if geom is not None else 'line'
    if rot or kind not in ('bentConnector2', 'bentConnector3'):
        return [(x0, y0), (x1, y1)]
    if kind == 'bentConnector2':
        return [(x0, y0), (x1, y0), (x1, y1)]
    gd = geom.find(A + 'avLst/' + A + 'gd')
    adj = int(gd.get('fmla').split()[-1]) / 100000 if gd is not None and gd.get('fmla', '').startswith('val') else 0.5
    xm = x0 + (x1 - x0) * adj
    return [(x0, y0), (xm, y0), (xm, y1), (x1, y1)]


def _has_arrow(sh):
    ln = sh._element.find(P + 'spPr/' + A + 'ln')
    return ln is not None and any(e is not None and e.get('type', 'none') != 'none'
                                  for e in (ln.find(A + 'headEnd'), ln.find(A + 'tailEnd')))


def _text_band(sh, x, y, w, h):
    """框里文字实际占的那一条（页面坐标，EMU）：横向取整个框宽，纵向按字的高度和锚点。估不了时用整个框。"""
    need = _text_need(sh, w)
    if not need:
        return x, y, x + w, y + h
    th = min(need[0] * CM, h)
    tf = sh.text_frame
    anchor = tf.vertical_anchor
    top = y + (h - th) / 2 if anchor == MSO_ANCHOR.MIDDLE else y + h - th if anchor == MSO_ANCHOR.BOTTOM else y
    left, right = x, x + w
    paras = [p for p in tf.paragraphs if p.text.strip()]
    if len(paras) == 1:     # 一行字：横向只算字的宽度，按对齐放
        tw = text_units(paras[0].text) * need[1] * PT_CM * CM + (tf.margin_left + tf.margin_right)
        if tw < w:
            al = paras[0].alignment
            left = x + (w - tw) / 2 if al == PP_ALIGN.CENTER else x + w - tw if al == PP_ALIGN.RIGHT else x
            right = left + tw
    return left, top, right, top + th


def _crossing_warnings(placed):
    """带箭头的连接线从字上压过：线画在字的上面，或字所在的框是透明的（不算端点所在的框边）。"""
    out = []
    boxes = []
    for k, (sh, x, y, w, h, _) in enumerate(placed):
        if sh.has_text_frame and len(sh.text_frame.text.strip()) >= 2 and w > 2 * LINK_INSET and h > 2 * LINK_INSET:
            l, t, r, b = _text_band(sh, x, y, w, h)
            boxes.append((k, sh, l + LINK_INSET, t + LINK_INSET, r - LINK_INSET, b - LINK_INSET))
    for j, (sh, x, y, w, h, _) in enumerate(placed):
        if sh._element.tag != P + 'cxnSp' or not _has_arrow(sh):
            continue
        pts = _link_path(sh, x, y, w, h)
        samples = [(a[0] + (b[0] - a[0]) * t / 24, a[1] + (b[1] - a[1]) * t / 24) for a, b in zip(pts, pts[1:]) for t in range(25)]
        ends = (pts[0], pts[-1])        # 端点附近 0.3 cm 不算：箭头尖贴着标签是常见的引线
        samples = [q for q in samples if min(abs(q[0] - e[0]) + abs(q[1] - e[1]) for e in ends) > 0.3 * CM]
        for k, other, l, t, r, b in boxes:
            if (j > k or not _opaque(other)) and any(l < px < r and t < py < b for px, py in samples):
                out.append((sh, f'箭头从 {other.name} 的字上压过：改走线或挪节点'))
                break
    return out


def _dangling(slide):
    """动画（p:timing）引用了页上不存在的形状：形状被删了，时间轴还留着。"""
    timing = slide._element.find(P + 'timing')
    ids = {e.get('id') for e in slide._element.iter(P + 'cNvPr')}
    return sorted({e.get('spid') for e in timing.iter() if e.get('spid') and e.get('spid') not in ids}) if timing is not None else []


def inspect(path, style=None):
    prs = Presentation(path)
    W, H = prs.slide_width, prs.slide_height
    pages, problems, warnings = [], [], []
    k = 19.05 / (H / CM)
    themes, label_pages, checked = {}, [], 0
    body = style.size('body') if style else None
    emphasis = style.color('emphasis').lstrip('#').upper() if style else None
    for n, slide in enumerate(prs.slides, 1):
        shapes, names, placed = [], {}, []
        hidden = slide._element.get('show') == '0'
        place = placements(slide.shapes, W, H)
        shadows = _theme_shadows(slide.slide_layout.slide_master)
        for sh, x, y, w, h, group in _walk(slide.shapes):
            item = {'name': sh.name, 'kind': sh.shape_type and str(sh.shape_type).split('.')[-1].split(' ')[0].lower(),
                    'cm': [_cm(x), _cm(y), _cm(w), _cm(h)], 'placement': place.get(sh._element, 'unknown')}
            if group:
                item['group'] = group
            text = _text(sh)
            if text:
                item['text'] = text[:60]
            sizes = _sizes(sh)
            if sizes:
                item['pt'] = sorted(sizes)
            if sh.shape_type is not None and 'PICTURE' in str(sh.shape_type):
                try:
                    px = sh.image.size[0]
                except AttributeError:      # svg、emf 等 python-pptx 读不了的图
                    px = None
                item['dpi'] = round(px / (Emu(w) / CM / 2.54)) if w and px else None
            off = item['placement'] == 'outside'
            if off:
                item['offCanvas'] = True
                problems.append({'page': n, 'shape': sh.name, 'problem': '画布外的对象（只列出，不动；内容用 plan-deck offcanvas 看）'})
            if body and sizes and min(sizes) < body - 0.01 and not off:
                problems.append({'page': n, 'shape': sh.name, 'problem': f'有 {min(sizes):g} pt 的字，小于正文 {body:g} pt'})
            if sh.name.startswith('pd:'):
                names[sh.name] = names.get(sh.name, 0) + 1
            if not off and not isinstance(sh, GroupShape):
                placed.append((sh, x, y, w, h, group))
            if not off and not hidden:
                problems += [{'page': n, 'shape': sh.name, 'problem': w} for w in _word_problems(sh)]
                warnings += [{'page': n, 'shape': sh.name, 'warning': w} for w in _word_warnings(sh)]
                if _has_shadow(sh, shadows):
                    sp = sh._element.find(P + 'spPr')
                    on_text = sp is not None and sp.find(A + 'solidFill') is None and sh.has_text_frame and sh.text_frame.text.strip()
                    warnings.append({'page': n, 'shape': sh.name, 'warning': (
                        '有阴影，框没有填充，阴影落在字上像重影' if on_text else '有阴影：规划汇报的形状不加阴影')
                        + '；python-pptx 直接画的形状带着主题效果（p:style 的 effectRef），用运行时的 shape()，'
                          '或在 spPr 里加空的 a:effectLst'})
                parts = sh.name.split(':')
                if sh.name.startswith('pd:') and len(parts) == 3 and sh.has_text_frame:
                    warnings += [{'page': n, 'shape': sh.name, 'warning': w}
                                 for w in _title_warnings(parts[1], parts[2], sh.text_frame.text)]
            if text and not off:
                ref = PAGE_REF.search(text)
                if ref:
                    warnings.append({'page': n, 'shape': sh.name, 'warning': f'页面上有页码引用“{ref.group()}”，页序一调就对不上'})
                slot = sh.name.split(':')[2] if sh.name.startswith('pd:') and sh.name.count(':') >= 2 else None
                if emphasis and sh.has_text_frame and slot not in HEADER_SLOTS:
                    warnings += [{'page': n, 'shape': sh.name, 'warning': w} for w in _red_warnings(sh, emphasis, body)]
            shapes.append(item)
        problems += [{'page': n, 'shape': f'id {spid}', 'problem': '动画引用的形状不存在：删掉这条动画或补回形状'} for spid in _dangling(slide)]
        if not hidden:
            warnings += [{'page': n, 'shape': sh.name, 'warning': w} for sh, w in _layout_warnings(placed, W, H, place)]
            warnings += [{'page': n, 'shape': sh.name, 'warning': w} for sh, w in _crossing_warnings(placed)]
            master = slide.slide_layout.slide_master
            items, runs = _page(slide, themes.setdefault(id(master._element), _Theme(master)), W, H, k)
            if content_page(items, runs):
                checked += 1
                bottom = content_bottom(items)
                if bottom is not None and bottom < SHORT_AT:
                    warnings.append({'page': n, 'shape': 'slide', 'warning': f'正文止于页高 {bottom:.0%}，下面空着：把内容铺到版心底'
                                     f'（组件给 bottom），加图，或问用户是否和相邻节点合页'})
                n_body, n_mid = mid_layer(runs)
                pic_area = sum(min(1, i['w'] * i['h']) for i in _pictures(items))
                if n_body >= 200 and not n_mid and pic_area < .15:
                    warnings.append({'page': n, 'shape': 'slide', 'warning': f'正文 {n_body} 字没有 14 磅以上的字，缺中间层：'
                                     f'把本页最重要的两三件事放大（栏头、标签、通栏、节点链，docs/patterns.md）'})
                if _label_rows(items, with_text=True) >= 2:
                    label_pages.append(n)
            grid = _grid_warning(slide, W, H)
            if grid:
                warnings.append({'page': n, 'shape': 'slide', 'warning': grid})
            plain = _no_emphasis_warning(slide, emphasis)
            if plain:
                warnings.append({'page': n, 'shape': 'slide', 'warning': plain})
        for nm, c in names.items():
            if c > 1:
                problems.append({'page': n, 'shape': nm, 'problem': f'同页有 {c} 个同名形状'})
        notes = slide.notes_slide.notes_text_frame.text if slide.has_notes_slide else ''
        pd = sum(1 for s in shapes if s['name'].startswith('pd:'))
        pages.append({'page': n, 'shapes': len(shapes), 'pd': pd, 'other': len(shapes) - pd,
                      'title': next((s.get('text') for s in shapes if s['name'].endswith(':title')), None),
                      'notes': notes[:80], 'items': shapes})
    warnings += [{'page': d['pages'][0], 'shape': 'notes', 'warning': f"第 {'、'.join(map(str, d['pages']))} 页备注相同：“{d['text']}”"}
                 for d in duplicate_notes(prs)]
    if len(label_pages) >= 2 and len(label_pages) > LABEL_ROW_SHARE * checked:
        warnings.append({'page': label_pages[0], 'shape': 'deck', 'warning': f'第 {"、".join(map(str, label_pages))} 页都是左色块＋右底框的'
                         f'标签行（{len(label_pages)}/{checked} 页）：换成栏头并列、节点链、表或正文，标签行不当默认做法'})
    return {'file': str(path), 'pageCm': [_cm(W), _cm(H)], 'pages': pages, 'problems': problems, 'warnings': warnings}
