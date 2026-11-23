"""改稿审核对比：两版逐页配对，改过的页出“旧 | 新”并排图，新增页、删除页单独出图，合成一份 PDF 给人审。

配对：先按全文相同对齐，剩下的按文字相似度配对，挪了位置的页也认得出来。
分类：文字和形状位置都没变的不出图；只有数字变了（节号、页码顺延）的只在清单里列页码。
新图上的红框是和旧页像素不同的区域，只作提示；图下列出文字改动。

渲染图用 render() 的输出（只含可见页，按顺序编号）：不给渲染目录时这里现渲染。
"""
import difflib
import glob
import os
import re

from PIL import Image, ImageChops, ImageDraw, ImageFont
from pptx import Presentation

from .measure import _font_dirs
from .render import render

PAGE_W = 2240          # 审核页宽（px），两张页面图并排
MARGIN = 40
MATCH_MIN = 0.45       # 文字相似度低于这个不算同一页
FONT_FILES = (('msyh.ttc', 'msyh.ttf', 'PingFang.ttc', 'Hiragino Sans GB.ttc', 'NotoSansCJK-Regular.ttc', 'simhei.ttf'),
              ('msyhbd.ttc', 'msyhbd.ttf', 'PingFang.ttc', 'Hiragino Sans GB.ttc', 'NotoSansCJK-Bold.ttc', 'simhei.ttf'))
_FONT_PATHS = {}


class CompareError(Exception):
    pass


def _font_path(bold):
    if bold not in _FONT_PATHS:
        _FONT_PATHS[bold] = None
        for want in FONT_FILES[1 if bold else 0]:
            for d in _font_dirs():
                for root, _, files in os.walk(d):
                    if want in files:
                        _FONT_PATHS[bold] = os.path.join(root, want)
                        break
                if _FONT_PATHS[bold]:
                    break
            if _FONT_PATHS[bold]:
                break
    return _FONT_PATHS[bold]


def _font(size, bold=False):
    p = _font_path(bold)
    return ImageFont.truetype(p, size) if p else ImageFont.load_default()


def _walk(shapes):
    for x in shapes:
        if x.shape_type == 6:
            yield from _walk(x.shapes)
        else:
            yield x


def _clean(t):
    return re.sub(r'\s+', ' ', t).strip()


def slides(path):
    """每页：页码、是否隐藏、逐段文字、形状位置、标题（优先取带三级编号的小标题）。"""
    out = []
    for i, s in enumerate(Presentation(path).slides, 1):
        lines, geo, head = [], [], ''
        for x in _walk(s.shapes):
            geo.append((x.left, x.top, x.width, x.height))
            if x.has_text_frame:
                lines += [t for t in (_clean(''.join(r.text for r in p.runs)) for p in x.text_frame.paragraphs) if t]
                if not head and x.name.endswith('header:title'):
                    head = _clean(x.text_frame.text)
            if getattr(x, 'has_table', False) and x.has_table:
                lines += [t for t in (_clean(c.text) for r in x.table.rows for c in r.cells) if t]
        sub = next((l for l in lines if re.match(r'^\d+(\.\d+){2,}\s', l)), '')
        out.append({'no': i, 'hidden': s._element.get('show') == '0', 'lines': lines, 'text': '\n'.join(lines),
                    'geo': sorted(geo), 'title': sub or head or (lines[0] if lines else '')})
    return out


def _ratio(a, b):
    return difflib.SequenceMatcher(None, a, b, autojunk=False).ratio()


def pair(old, new):
    """返回 (配对 [(旧页, 新页)…] 按新页序, 新增页, 删除页)。"""
    sm = difflib.SequenceMatcher(None, [s['text'] for s in old], [s['text'] for s in new], autojunk=False)
    pairs, o_left, n_left = [], [], []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == 'equal':
            pairs += [(old[i1 + k], new[j1 + k]) for k in range(i2 - i1)]
        else:
            o_left += old[i1:i2]
            n_left += new[j1:j2]
    cands = sorted(((_ratio(o['text'], n['text']), o['no'], n['no']) for o in o_left for n in n_left), reverse=True)
    om, nm = {s['no']: s for s in old}, {s['no']: s for s in new}
    used_o, used_n = set(), set()
    for r, oi, ni in cands:
        if r < MATCH_MIN or oi in used_o or ni in used_n:
            continue
        used_o.add(oi)
        used_n.add(ni)
        pairs.append((om[oi], nm[ni]))
    return (sorted(pairs, key=lambda p: p[1]['no']), [n for n in n_left if n['no'] not in used_n],
            [o for o in o_left if o['no'] not in used_o])


def kind(o, n):
    """same 没变；renumber 只有数字变了；changed 改过。"""
    if o['geo'] != n['geo']:
        return 'changed'
    if o['text'] == n['text']:
        return 'same'
    if re.sub(r'\d+', '#', o['text']) == re.sub(r'\d+', '#', n['text']):
        return 'renumber'
    return 'changed'


def text_changes(o, n):
    """[('-', 旧版有新版没有的段), ('+', 新版新加的段)]"""
    return [('-', l) for l in o['lines'] if l not in n['lines']] + [('+', l) for l in n['lines'] if l not in o['lines']]


def diff_boxes(a, b, cell=12, thresh=40):
    """两张同尺寸图的差异区域（原图像素坐标），按网格合并成框。"""
    if a.size != b.size:
        return []
    d = ImageChops.difference(a.convert('L'), b.convert('L')).point(lambda v: 255 if v > thresh else 0)
    gw, gh = (d.width + cell - 1) // cell, (d.height + cell - 1) // cell
    px = d.resize((gw, gh), Image.BOX).load()
    mark = {(x, y) for y in range(gh) for x in range(gw) if px[x, y] > 8}
    boxes = []
    while mark:
        stack = [mark.pop()]
        xs, ys = [], []
        while stack:
            cx, cy = stack.pop()
            xs.append(cx)
            ys.append(cy)
            for dx in (-2, -1, 0, 1, 2):
                for dy in (-2, -1, 0, 1, 2):
                    if (cx + dx, cy + dy) in mark:
                        mark.discard((cx + dx, cy + dy))
                        stack.append((cx + dx, cy + dy))
        boxes.append((min(xs) * cell, min(ys) * cell, (max(xs) + 1) * cell, (max(ys) + 1) * cell))
    return sorted(boxes, key=lambda b: (b[1], b[0]))


def _wrap(draw, text, f, width):
    lines, cur = [], ''
    for ch in text:
        if cur and draw.textlength(cur + ch, font=f) > width:
            lines.append(cur)
            cur = ch
        else:
            cur += ch
    return lines + ([cur] if cur else [])


def _sheet(header, left, right, labels, changes=(), boxes=(), limit=16):
    imgs = [im for im in (left, right) if im is not None]
    iw = (PAGE_W - MARGIN * 3) // 2 if len(imgs) == 2 else PAGE_W - MARGIN * 2
    scaled = [im.resize((iw, round(im.height * iw / im.width))) for im in imgs]
    ih = max(im.height for im in scaled)
    f_h, f_l, f_t = _font(34, True), _font(26, True), _font(22)
    probe = ImageDraw.Draw(Image.new('RGB', (8, 8)))
    rows = []
    for sign, t in list(changes)[:limit]:
        rows += [(sign if k == 0 else ' ', seg) for k, seg in enumerate(_wrap(probe, t, f_t, PAGE_W - MARGIN * 2 - 40))]
    if len(changes) > limit:
        rows.append((' ', f'……另有 {len(changes) - limit} 处文字改动'))
    page = Image.new('RGB', (PAGE_W, 130 + ih + (70 + len(rows) * 32 if rows else 40)), 'white')
    d = ImageDraw.Draw(page)
    d.text((MARGIN, 26), header, font=f_h, fill='#1F3864')
    y0 = 130
    xs = [MARGIN, MARGIN * 2 + iw] if len(scaled) == 2 else [MARGIN]
    for x, im, lab, col in zip(xs, scaled, labels, ('#7F7F7F', '#C00000') if len(scaled) == 2 else ('#C00000',)):
        d.text((x, 88), lab, font=f_l, fill=col)
        page.paste(im, (x, y0))
        d.rectangle([x - 1, y0 - 1, x + iw, y0 + im.height], outline='#BFBFBF', width=2)
    if boxes and len(scaled) == 2:
        s = iw / right.width
        for a, b, c, e in boxes:
            d.rectangle([xs[1] + a * s - 3, y0 + b * s - 3, xs[1] + c * s + 3, y0 + e * s + 3], outline='#FF0000', width=3)
    y = y0 + ih + 30
    if rows:
        d.text((MARGIN, y), '文字改动（- 旧版有、新版没有；+ 新版新加）', font=f_l, fill='#404040')
        y += 40
        for sign, seg in rows:
            col = {'-': '#7F7F7F', '+': '#C00000'}.get(sign, '#404040')
            d.text((MARGIN, y), sign, font=f_t, fill=col)
            d.text((MARGIN + 34, y), seg, font=f_t, fill=col)
            y += 32
    return page


def _cover(title, rows):
    f_h, f_t = _font(44, True), _font(26)
    probe = ImageDraw.Draw(Image.new('RGB', (8, 8)))
    lines = []
    for r in rows:
        lines += _wrap(probe, r, f_t, PAGE_W - MARGIN * 2) or ['']
    page = Image.new('RGB', (PAGE_W, max(1260, 170 + len(lines) * 40)), 'white')
    d = ImageDraw.Draw(page)
    d.text((MARGIN, 50), title, font=f_h, fill='#1F3864')
    for k, l in enumerate(lines):
        d.text((MARGIN, 140 + k * 40), l, font=f_t, fill='#262626')
    return page


def _pngs(folder):
    files = glob.glob(os.path.join(folder, '*.png'))
    num = lambda p: int(re.search(r'(\d+)\.png$', p).group(1))
    return sorted((p for p in files if re.search(r'\d+\.png$', p)), key=num)


def _render_map(info, pngs, which):
    visible = [s['no'] for s in info if not s['hidden']]
    if len(pngs) != len(visible):
        raise CompareError(f'{which}渲染图 {len(pngs)} 张，可见页 {len(visible)} 页，对不上；渲染目录里只能放这一版的图')
    return dict(zip(visible, pngs))


def compare(old, new, out_pdf, old_renders=None, new_renders=None, engine='powerpoint', dpi=80, labels=('旧版', '新版')):
    """出审核 PDF，返回 {pdf, sheets, changed, added, deleted, renumbered, renders, notes}。"""
    for p in (old, new):
        if not os.path.isfile(p):
            raise CompareError(f'找不到 {p}')
    lo, ln = labels
    so, sn = slides(old), slides(new)
    base = os.path.splitext(out_pdf)[0] + '.renders'
    dirs = {}
    for key, path, given in (('old', old, old_renders), ('new', new, new_renders)):
        if given:
            dirs[key] = _pngs(given)
        else:
            res = render(path, os.path.join(base, key), engine, dpi)
            if not res['pages']:
                raise CompareError('渲染只出了 pdf，没有逐页图：' + '；'.join(res['notes']))
            dirs[key] = res['pages']
    ro, rn = _render_map(so, dirs['old'], lo), _render_map(sn, dirs['new'], ln)
    pairs, added, deleted = pair(so, sn)
    changed = [(o, n) for o, n in pairs if kind(o, n) == 'changed']
    renum = [(o, n) for o, n in pairs if kind(o, n) == 'renumber']

    rows = [f'{lo}：{len(so)} 页　{ln}：{len(sn)} 页', '',
            f'改动 {len(changed)} 页，新增 {len(added)} 页，删除 {len(deleted)} 页；只改了编号的 {len(renum)} 页不出图。', '',
            f'改动页（{lo} → {ln}）：']
    for o, n in changed:
        flag = [w for w, c in (('挪了位置', abs(o['no'] - n['no']) > 5), ('隐藏页', o['hidden'] or n['hidden'])) if c]
        rows.append(f'　第 {o["no"]} 页 → 第 {n["no"]} 页　{n["title"][:40]}' + (f'（{"、".join(flag)}）' if flag else ''))
    if added:
        rows += ['', '新增页：'] + [f'　第 {n["no"]} 页　{n["title"][:40]}' for n in added]
    if deleted:
        rows += ['', '删除页：'] + [f'　{lo} 第 {o["no"]} 页　{o["title"][:40]}' for o in deleted]
    if renum:
        rows += ['', '只改了编号：' + '、'.join(f'{o["no"]}→{n["no"]}' for o, n in renum)]
    rows += ['', '并排图左为旧页、右为新页；右图红框是和旧页像素不同的区域，只作提示。隐藏页没有渲染图，用灰底代替，只看文字改动。']
    sheets = [_cover(f'改稿审核：{lo} → {ln}', rows)]

    blank = lambda ref: Image.new('RGB', ref.size if ref else (1067, 600), '#F2F2F2')
    seq = sorted([(n['no'], o, n) for o, n in changed] + [(n['no'], None, n) for n in added], key=lambda t: t[0])
    for _, o, n in seq:
        ni = Image.open(rn[n['no']]).convert('RGB') if n['no'] in rn else None
        if o is None:
            sheets.append(_sheet(f'新增　{ln} 第 {n["no"]} 页　{n["title"]}', None, ni or blank(None),
                                 (f'{ln} 第 {n["no"]} 页（新增）' + ('（隐藏页）' if ni is None else ''),)))
            continue
        oi = Image.open(ro[o['no']]).convert('RGB') if o['no'] in ro else None
        head = f'{lo} 第 {o["no"]} 页 → {ln} 第 {n["no"]} 页　{n["title"]}'
        boxes = diff_boxes(oi, ni) if oi is not None and ni is not None else ()
        sheets.append(_sheet(head + ('（隐藏页）' if oi is None or ni is None else ''), oi or blank(ni), ni or blank(oi),
                             (f'{lo} 第 {o["no"]} 页', f'{ln} 第 {n["no"]} 页'), text_changes(o, n), boxes))
    for o in deleted:
        oi = Image.open(ro[o['no']]).convert('RGB') if o['no'] in ro else blank(None)
        sheets.append(_sheet(f'删除　{lo} 第 {o["no"]} 页　{o["title"]}', oi, None, (f'{lo} 第 {o["no"]} 页（删除）',)))

    os.makedirs(os.path.dirname(os.path.abspath(out_pdf)), exist_ok=True)
    sheets[0].save(out_pdf, save_all=True, append_images=sheets[1:], resolution=150)
    notes = [] if _font_path(False) else ['没找到中文字体，审核页上的字可能显示成方框']
    return {'pdf': out_pdf, 'sheets': len(sheets),
            'changed': [{'old': o['no'], 'new': n['no'], 'title': n['title']} for o, n in changed],
            'added': [{'new': n['no'], 'title': n['title']} for n in added],
            'deleted': [{'old': o['no'], 'title': o['title']} for o in deleted],
            'renumbered': [[o['no'], n['no']] for o, n in renum],
            'renders': {'old': os.path.dirname(dirs['old'][0]) if dirs['old'] else None,
                        'new': os.path.dirname(dirs['new'][0]) if dirs['new'] else None},
            'notes': notes}
