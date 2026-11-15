"""只读体检：统计全册版面和文字的“过于整齐”线索，逐页列出问题，不改文件。

页面绝不出现的用语（words.BANNED）进 problems，影响退出码；其余文字线索进 language，只作线索。

数字只作线索，不是及格线（docs/humanize.md）。参考范围来自五份人做的规划汇报，样本出自同一团队，配色有传承；
12 磅以下的小字是那些稿子的毛病，不学。
给 against 时比对两版文字，确认去 AI 味前后文字没变，或找出工作稿改过哪些页。
"""
import colorsys
import re
from collections import Counter
from statistics import median

from pptx import Presentation      # 先导入它：缺依赖时入口脚本按 pptx 报，不是它带的 lxml
from lxml import etree  # noqa: E402

from .words import SPOKEN, TERSE_VERB, banned

A ='{http://schemas.openxmlformats.org/drawingml/2006/main}'
P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
COLORS = ('srgbClr', 'schemeClr', 'sysClr', 'prstClr')
PRESET = {'black': '000000', 'white': 'FFFFFF', 'red': 'FF0000', 'green': '008000', 'blue': '0000FF', 'yellow': 'FFFF00', 'gray': '808080', 'grey': '808080'}
# Office 默认主题色和标准色，含常用深浅
OFFICE = set('4472C4 ED7D31 A5A5A5 FFC000 5B9BD5 70AD47 8FAADC B4C7E7 DAE3F3 2F5597 1F3864 F4B183 F8CBAD FBE5D6 C55A11 843C0C '
             'DBDBDB EDEDED C9C9C9 7B7B7B 525252 FFD966 FFE699 FFF2CC BF9000 7F6000 9DC3E6 BDD7EE DEEBF7 2E75B6 1F4E79 A9D18E C5E0B4 '
             'E2F0D9 548235 385723 FF0000 C00000 FFFF00 00B050 0070C0 002060 7030A0 00B0F0 92D050'.split())
GRID = (32, 18)
SIMILAR = .78       # 两页版面网格重合这么多算高度相似；自定算法，看量级
EQUAL_PAGE = .7     # 一页七成以上的框等尺寸
# 五份人做的规划汇报的范围（低, 高），只作线索
REFERENCE = {
    'adjacentColorOverlap': (.25, .33), 'officeColorShare': (.36, .56), 'fillSaturation': (.53, .79), 'tintedPaleShare': (.01, .24),
    'blackLineShare': (.14, .60), 'equalSizePages': (.12, .46), 'labelRowPages': (0, .06), 'similarAdjacentPages': (.06, .26),
    'commonestLayoutShare': (.12, .18), 'bleedPerPage': (.20, .73), 'largeImagePages': (.27, .70), 'noImagePages': (.29, .63),
    'charsPerPage': (196, 387), 'bigRedPages': (.08, .33), 'smallTextShare': (0, 0),
}
LABELS = {
    'adjacentColorOverlap': '相邻两页的彩色重合度', 'officeColorShare': '彩色填充里 Office 默认色的比例', 'fillSaturation': '填充饱和度（按面积）',
    'tintedPaleShare': '浅底里偏蓝灰的面积比例', 'blackLineShare': '线条中纯黑的比例', 'equalSizePages': '七成以上框等尺寸的页',
    'labelRowPages': '左色块＋右底框成行的页', 'similarAdjacentPages': '相邻页版面高度相似', 'commonestLayoutShare': '最常见的一种版面占全册',
    'bleedPerPage': '每页伸出页边的形状', 'largeImagePages': '图占页面超 25% 的页', 'noImagePages': '完全没有图的页',
    'charsPerPage': '每页字数（中位）', 'bigRedPages': '标题下大号红字的页', 'smallTextShare': '12 磅以下字的比例',
}
CHARS = re.compile(r'\S')


def _mods(hexv, el):
    r, g, b = (int(hexv[i:i + 2], 16) / 255 for i in (0, 2, 4))
    for m in el:
        name, val = etree.QName(m).localname, m.get('val')
        val = int(val) / 100000 if val and val.lstrip('-').isdigit() else None
        if val is None:
            continue
        if name in ('lumMod', 'lumOff'):
            h, l, s = colorsys.rgb_to_hls(r, g, b)
            r, g, b = colorsys.hls_to_rgb(h, min(1, max(0, l * val if name == 'lumMod' else l + val)), s)
        elif name == 'shade':
            r, g, b = r * val, g * val, b * val
        elif name == 'tint':
            r, g, b = 1 - (1 - r) * val, 1 - (1 - g) * val, 1 - (1 - b) * val
    return '%02X%02X%02X' % tuple(max(0, min(255, round(x * 255))) for x in (r, g, b))


class _Theme:
    def __init__(self, master):
        self.map = dict(master._element.find(P + 'clrMap').attrib)
        self.colors = {}
        for rel in master.part.rels.values():
            if rel.reltype.endswith('/theme'):
                scheme = etree.fromstring(rel.target_part.blob).find('.//' + A + 'clrScheme')
                for c in scheme:
                    self.colors[etree.QName(c).localname] = c[0].get('val') if etree.QName(c[0]).localname == 'srgbClr' else c[0].get('lastClr', '000000')

    def color(self, el):
        """el 或它的子元素是颜色元素时给六位色值，否则 None。"""
        if el is None:
            return None
        for c in ([el] if etree.QName(el).localname in COLORS else el):
            name = etree.QName(c).localname
            if name == 'srgbClr':
                return _mods(c.get('val').upper(), c)
            if name == 'schemeClr':
                v = c.get('val')
                return _mods(self.colors.get(self.map.get(v, v), self.colors.get(v, '000000')).upper(), c)
            if name == 'sysClr':
                return _mods((c.get('lastClr') or '000000').upper(), c)
            if name == 'prstClr':
                return _mods(PRESET.get(c.get('val'), '808080'), c)
        return None


def _family(h):
    hue, l, s = colorsys.rgb_to_hls(*(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4)))
    if l >= .97:
        return '白'
    if l <= .08:
        return '黑'
    if s < .13 or (l > .9 and s < .25):
        return '灰'
    d = hue * 360
    for limit, name in ((14, '红'), (45, '橙'), (70, '黄'), (165, '绿'), (195, '青'), (255, '蓝'), (345, '紫')):
        if d < limit:
            return name
    return '红'


def _walk(el, tf=(0, 0, 1, 1)):
    for c in el:
        name = etree.QName(c).localname if isinstance(c.tag, str) else ''
        if name == 'grpSp':
            xf = c.find(P + 'grpSpPr/' + A + 'xfrm')
            parts = None if xf is None else [xf.find(A + k) for k in ('off', 'ext', 'chOff', 'chExt')]
            if not parts or None in parts:
                yield from _walk(c, tf)
                continue
            o, e, co, ce = parts
            sx, sy = int(e.get('cx')) / (int(ce.get('cx')) or 1), int(e.get('cy')) / (int(ce.get('cy')) or 1)
            yield from _walk(c, (tf[0] + tf[2] * (int(o.get('x')) - int(co.get('x')) * sx), tf[1] + tf[3] * (int(o.get('y')) - int(co.get('y')) * sy), tf[2] * sx, tf[3] * sy))
        elif name in ('sp', 'pic', 'cxnSp', 'graphicFrame'):
            yield c, name, tf
        elif name in ('AlternateContent', 'Choice'):
            yield from _walk(c, tf)


def _box(c, kind, tf, W, H):
    xf = c.find(P + 'xfrm') if kind == 'graphicFrame' else c.find(P + 'spPr/' + A + 'xfrm')
    if xf is None or xf.find(A + 'off') is None or xf.find(A + 'ext') is None:
        return None
    o, e = xf.find(A + 'off'), xf.find(A + 'ext')
    return ((tf[0] + tf[2] * int(o.get('x'))) / W, (tf[1] + tf[3] * int(o.get('y'))) / H, tf[2] * int(e.get('cx')) / W, tf[3] * int(e.get('cy')) / H)


def _page(slide, theme, W, H, k):
    """一页的形状和文字。k：把字号折算到 19.05 cm 页高。"""
    items, runs = [], []

    def add_text(tb):
        n = 0
        for r in tb.iter(A + 'r'):
            t = r.find(A + 't')
            chars = len(CHARS.findall(t.text or '')) if t is not None else 0
            if not chars:
                continue
            rp = r.find(A + 'rPr')
            sz = int(rp.get('sz')) / 100 * k if rp is not None and rp.get('sz') else None
            fill = rp.find(A + 'solidFill') if rp is not None else None
            runs.append({'chars': chars, 'pt': sz, 'color': theme.color(fill) if fill is not None else None})
            n += chars
        return n

    for c, kind, tf in _walk(slide._element.find(P + 'cSld/' + P + 'spTree')):
        box = _box(c, kind, tf, W, H)
        if box is None or box[0] >= 1 or box[1] >= 1 or box[0] + box[2] <= 0 or box[1] + box[3] <= 0:
            continue
        it = {'kind': kind, 'x': box[0], 'y': box[1], 'w': box[2], 'h': box[3], 'fill': None, 'line': None, 'chars': 0}
        sp, style = c.find(P + 'spPr'), c.find(P + 'style')
        if kind == 'sp' or kind == 'cxnSp':
            solid = sp.find(A + 'solidFill')
            if solid is not None:
                it['fill'] = theme.color(solid)
            elif sp.find(A + 'noFill') is None and style is not None and style.find(A + 'fillRef') is not None and int(style.find(A + 'fillRef').get('idx', '0')) > 0:
                it['fill'] = theme.color(style.find(A + 'fillRef'))
            ln = sp.find(A + 'ln')
            if ln is not None and ln.find(A + 'solidFill') is not None:
                it['line'] = theme.color(ln.find(A + 'solidFill'))
            elif (ln is None or ln.find(A + 'noFill') is None) and style is not None and style.find(A + 'lnRef') is not None and int(style.find(A + 'lnRef').get('idx', '0')) > 0:
                it['line'] = theme.color(style.find(A + 'lnRef'))
        tb = c.find(P + 'txBody')
        if tb is not None:
            it['chars'] = add_text(tb)
        if kind == 'graphicFrame' and c.find('.//' + A + 'tbl') is not None:
            it['kind'] = 'table'
            for tc in c.iter(A + 'tc'):
                pr = tc.find(A + 'tcPr')
                if pr is not None and pr.find(A + 'solidFill') is not None:
                    color = theme.color(pr.find(A + 'solidFill'))
                    if color:
                        it.setdefault('cells', []).append(color)
                if tc.find(A + 'txBody') is not None:
                    it['chars'] += add_text(tc.find(A + 'txBody'))
                for side in ('lnL', 'lnR', 'lnT', 'lnB'):
                    ln = pr.find(A + side) if pr is not None else None
                    if ln is not None and ln.find(A + 'solidFill') is not None:
                        it.setdefault('lines', []).append(theme.color(ln.find(A + 'solidFill')))
        items.append(it)
    return items, runs


def _colored(c):
    return bool(c) and _family(c) not in ('白', '黑', '灰')


def _hues(items, runs):
    out = set()
    for i in items:
        out |= {c for c in [i['fill'], i['line'], *i.get('cells', []), *i.get('lines', [])] if _colored(c)}
    return out | {r['color'] for r in runs if _colored(r['color'])}


def _boxes(items):
    return [i for i in items if not (i['w'] >= .95 and i['h'] >= .9) and i['kind'] != 'cxnSp' and i['w'] >= .025 and i['h'] >= .02 and i['y'] + i['h'] >= .12
            and (i['kind'] in ('pic', 'table', 'graphicFrame') or i['fill'] or i['line'])]


def _equal_share(boxes):
    return sum(any(o is not a and abs(o['w'] - a['w']) < .006 and abs(o['h'] - a['h']) < .008 for o in boxes) for a in boxes) / len(boxes)


def _label_rows(items):
    """窄实色块右侧紧贴一块底框的行数。"""
    rows = 0
    for a in items:
        if not a['fill'] or not a['chars'] or a['y'] <= .12 or a['w'] >= .2 or a['h'] <= .05 or colorsys.rgb_to_hls(*(int(a['fill'][i:i + 2], 16) / 255 for i in (0, 2, 4)))[1] >= .6:
            continue
        rows += any(b is not a and b['fill'] and abs(b['x'] - a['x'] - a['w']) < .012 and abs(b['y'] - a['y']) < .012 and abs(b['h'] - a['h']) < .02 and b['w'] > a['w'] for b in items)
    return rows


def _fingerprint(items):
    nx, ny = GRID
    g = [[0] * nx for _ in range(ny)]
    for i in items:
        if i['w'] > .95 and i['h'] > .9:
            continue
        v = 3 if i['kind'] == 'pic' else 2 if i['kind'] in ('table', 'graphicFrame') or i['fill'] else 1 if i['chars'] else 0
        if not v:
            continue
        for y in range(int(max(0, i['y']) * ny), min(ny, int(min(1, i['y'] + i['h']) * ny + .999))):
            for x in range(int(max(0, i['x']) * nx), min(nx, int(min(1, i['x'] + i['w']) * nx + .999))):
                g[y][x] = max(g[y][x], v)
    return g


def _same(a, b):
    return sum(x == y for r, q in zip(a, b) for x, y in zip(r, q)) / (GRID[0] * GRID[1])


def _pictures(items):
    return [i for i in items if i['kind'] == 'pic' and i['w'] > .04 and i['h'] > .04 and not (i['w'] > .95 and i['y'] + i['h'] < .15)]


def _off(c, kind, tf, W, H):
    box = _box(c, kind, tf, W, H)
    return box is not None and (box[0] >= 1 or box[1] >= 1 or box[0] + box[2] <= 0 or box[1] + box[3] <= 0)


def _slide_text(slide, W, H):
    """页面上每个文本框的文字，排序后比较，不管叠放顺序和位置。版面外的文字也比，但标出来，不和正文混。"""
    out = []
    for c, kind, tf in _walk(slide._element.find(P + 'cSld/' + P + 'spTree')):
        t = ''.join(x.text or '' for x in c.iter(A + 't'))
        if t.strip():
            out.append(f'[版面外] {t}' if _off(c, kind, tf, W, H) else t)
    return sorted(out)


def compare_text(path, against):
    """两版逐页比文字。页数不同时只比前面相同页数的部分，并报出来。"""
    pa, pb = Presentation(path), Presentation(against)
    a, b = pa.slides, pb.slides
    changed = []
    for n, (x, y) in enumerate(zip(a, b), 1):
        tx, ty = Counter(_slide_text(x, pa.slide_width, pa.slide_height)), Counter(_slide_text(y, pb.slide_width, pb.slide_height))
        if tx != ty:
            changed.append({'page': n, 'onlyHere': [t[:60] for t in (tx - ty)], 'onlyThere': [t[:60] for t in (ty - tx)]})
    return {'pages': [len(a), len(b)], 'changedPages': [c['page'] for c in changed], 'changes': changed}


def _language(n, slide, W, H):
    """页面文字：(页面绝不出现的用语, 线索)。版面外的草稿不算页面文字，备注不查。"""
    problems, clues = [], []
    for c, kind, tf in _walk(slide._element.find(P + 'cSld/' + P + 'spTree')):
        if _off(c, kind, tf, W, H):
            continue
        for para in c.iter(A + 'p'):
            t = ''.join('\n' if etree.QName(x).localname == 'br' else x.text or '' for x in para.iter(A + 't', A + 'br'))
            problems += [{'page': n, 'kind': why, 'word': word, 'text': line[:50]} for why, word, line in banned(t)]
            for rx, why in ((SPOKEN, '讨论口语'), (TERSE_VERB, '单字动词接术语')):
                m = rx.search(t)
                if m:
                    clues.append({'page': n, 'kind': why, 'word': m.group(), 'text': t[:50]})
    return problems, clues


def audit(path, against=None):
    prs = Presentation(path)
    W, H = prs.slide_width, prs.slide_height
    k = 19.05 / (H / 360000)
    themes, pages, problems, language = {}, [], [], []
    for n, slide in enumerate(prs.slides, 1):
        if slide._element.get('show') == '0':
            continue
        master = slide.slide_layout.slide_master
        theme = themes.setdefault(id(master._element), _Theme(master))
        items, runs = _page(slide, theme, W, H, k)
        pages.append({'n': n, 'items': items, 'runs': runs, 'hues': _hues(items, runs), 'fp': _fingerprint(items)})
        found, clues = _language(n, slide, W, H)
        problems += found
        language += clues

    area, pale_tint, pale_grey, lines, black = Counter(), 0, 0, 0, 0
    fills = Counter()
    for p in pages:
        for i in p['items']:
            for c in [i['line'], *i.get('lines', [])]:
                if c:
                    lines += 1
                    black += c == '000000'
            for c in i.get('cells', []):
                if _colored(c):
                    fills[c] += 1
            if not i['fill'] or (i['w'] > .95 and i['h'] > .9):
                continue
            if _colored(i['fill']):
                fills[i['fill']] += 1
            if i['y'] > .12:
                a = i['w'] * i['h']
                area[i['fill']] += a
                _, l, s = colorsys.rgb_to_hls(*(int(i['fill'][j:j + 2], 16) / 255 for j in (0, 2, 4)))
                if .8 < l < .985:
                    if s < .04:
                        pale_grey += a
                    elif s < .45:
                        pale_tint += a
    colored_area = {c: v for c, v in area.items() if _colored(c)}
    saturation = sum(colorsys.rgb_to_hls(*(int(c[j:j + 2], 16) / 255 for j in (0, 2, 4)))[2] * v for c, v in colored_area.items()) / (sum(colored_area.values()) or 1)

    n = len(pages) or 1
    overlaps, similar = [], []
    clusters = []
    for a, b in zip(pages, pages[1:]):
        overlaps.append(len(a['hues'] & b['hues']) / (len(a['hues'] | b['hues']) or 1))
        similar.append(_same(a['fp'], b['fp']))
    for p in pages:
        for c in clusters:
            if _same(c[0], p['fp']) >= SIMILAR:
                c[1] += 1
                break
        else:
            clusters.append([p['fp'], 1])

    equal_pages, sized, label_pages, bleed, large, none, chars, big_red = 0, 0, 0, 0, 0, 0, [], 0
    pt_total, pt_small = 0, 0
    rows = []
    for idx, p in enumerate(pages):
        boxes, pics = _boxes(p['items']), _pictures(p['items'])
        share = _equal_share(boxes) if len(boxes) >= 3 else None
        pic_area = min(1, sum(min(1, i['w'] * i['h']) for i in pics))
        labels = _label_rows(p['items'])
        total = sum(r['chars'] for r in p['runs'])
        red = any(r['pt'] and r['pt'] >= 19 and r['color'] and _family(r['color']) == '红' for r in p['runs'])
        sized += share is not None
        equal_pages += share is not None and share >= EQUAL_PAGE
        label_pages += labels >= 2
        bleed += sum(i['x'] < -.002 or i['y'] < -.002 or i['x'] + i['w'] > 1.002 or i['y'] + i['h'] > 1.002 for i in p['items'])
        large += pic_area > .25
        none += not pics
        chars.append(total)
        big_red += red
        for r in p['runs']:
            if r['pt']:
                pt_total += r['chars']
                pt_small += r['chars'] if r['pt'] < 11.75 else 0
        flags = []
        if share is not None and share >= EQUAL_PAGE:
            flags.append(f'七成以上的框等尺寸（{share:.0%}）')
        if labels >= 2:
            flags.append(f'{labels} 行左色块＋右底框')
        if idx and similar[idx - 1] >= SIMILAR:
            flags.append(f'和上一页版面相似（第 {pages[idx - 1]["n"]} 页）')
        if not pics:
            flags.append('没有图')
        rows.append({'page': p['n'], 'chars': total, 'imageArea': round(pic_area, 2), 'bigRed': red, 'flags': flags})

    colored_fills = sum(fills.values())
    metrics = {
        'adjacentColorOverlap': round(median(overlaps), 2) if overlaps else 0,
        'officeColorShare': round(sum(v for c, v in fills.items() if c in OFFICE) / (colored_fills or 1), 2),
        'fillSaturation': round(saturation, 2),
        'tintedPaleShare': round(pale_tint / ((pale_tint + pale_grey) or 1), 2),
        'blackLineShare': round(black / (lines or 1), 2),
        'equalSizePages': round(equal_pages / (sized or 1), 2),
        'labelRowPages': round(label_pages / n, 2),
        'similarAdjacentPages': round(sum(s >= SIMILAR for s in similar) / (len(similar) or 1), 2),
        'commonestLayoutShare': round(max((c[1] for c in clusters), default=0) / n, 2),
        'bleedPerPage': round(bleed / n, 2),
        'largeImagePages': round(large / n, 2),
        'noImagePages': round(none / n, 2),
        'charsPerPage': round(median(chars)) if chars else 0,
        'bigRedPages': round(big_red / n, 2),
        'smallTextShare': round(pt_small / (pt_total or 1), 2),
    }
    table = [{'metric': key, 'label': LABELS[key], 'value': v, 'reference': list(REFERENCE[key]),
              'outside': not REFERENCE[key][0] <= v <= REFERENCE[key][1]} for key, v in metrics.items()]
    res = {'visiblePages': len(pages), 'totalPages': len(prs.slides), 'metrics': table, 'pages': rows, 'problems': problems, 'language': language,
           'note': '参考范围只作线索，不是及格线；看每页标记时对照内容判断，不要为了凑数字改页面'}
    if against:
        res['textCompare'] = compare_text(path, against)
    return res
