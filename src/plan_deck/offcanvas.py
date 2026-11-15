"""版面外审计：把渲染图上看不到的内容逐页列出来，原 pptx 不改。

同事常把资料、批注用的表和素材放在幻灯片版面外。渲染图看不到它们，不分位置的文字提取又把它们和正文混在一起。
每个对象按页面上的绝对坐标（含组的缩放、平移、旋转）标位置：inside 版面内、crossing 跨边、outside 版面外、unknown 读不到位置。

输出目录里：
  offcanvas.json   全部对象（文字全文、表格、图表数据、位置、隐藏页、窗格隐藏、裁剪）
  版面外清单.md     人看的清单，只列渲染图上看不到的对象
  全文提取.md       每页全部文字，每个对象标明在版面内还是版面外
  assets/          版面外、跨边、窗格隐藏、裁剪较多的图片（原图，含裁掉的部分），图表的内嵌工作簿，嵌入对象
  renders/         --render 时连版面外一起渲染的图，红框是版面边界
读不了的文件（flat XML、加密、旧版 .ppt……）明确报告没审计，不跳过。
同一页版面内外有结构相同、内容对不上的表格时列为线索（pages[].tableVersions）：默认以版面内为准，版面外的照样列出，什么都不改。
"""
import hashlib
import json
import math
import os
import re
import shutil
import tempfile
import zipfile

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.shapes.group import GroupShape
from pptx.util import Emu, Pt

from .render import RenderError, render

CM = 360000
INCH = 914400
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
C = '{http://schemas.openxmlformats.org/drawingml/2006/chart}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
MC = '{http://schemas.openxmlformats.org/markup-compatibility/2006}'
DGM_REL = '{http://schemas.openxmlformats.org/drawingml/2006/diagram}relIds'
SHAPE_TAGS = tuple(P + t for t in ('sp', 'pic', 'graphicFrame', 'cxnSp', 'grpSp'))
LABEL = {'inside': '版面内', 'crossing': '跨边', 'outside': '版面外', 'unknown': '位置不明'}
EDGE = round(0.25 * CM)     # 出界不到这么多不算跨边：满版图常略出一点
CROP_MIN = 0.2              # 裁掉的面积占原图这么多才列为线索
MAX_SLIDE = 56 * INCH       # PowerPoint 的页面尺寸上限
PAD = round(0.2 * INCH)
TITLE_MAX = 40
VERSION_LABELS = 0.8        # 两张表算同结构：行标题这么多相同，首行至少一半相同
VERSION_SAME = 0.5          # 数据格这么多相同才算两版；或者数据格够多（VERSION_CELLS），小表数字全不同更可能是另一组数据（如不同街道）
VERSION_CELLS = 6
NOTICE = ('版面外常是任务说明、素材、草稿或其他项目的模板：引用要注明来源（第几页版面外、对象名），判断后再用，不当正文。'
          '同一内容版面内外各有一版、对不上时，默认以版面内为准，版面外视为旧版或暂存；版面外独有的资料照样判断。')
_ID = (1, 0, 0, 1, 0, 0)


# 位置

def _times(m, n):
    a, b, c, d, e, f = m
    g, h, i, j, k, l = n
    return (a * g + c * h, b * g + d * h, a * i + c * j, b * i + d * j, a * k + c * l + e, b * k + d * l + f)


def _rotate(deg, cx, cy):
    t = math.radians(deg)
    cos, sin = math.cos(t), math.sin(t)
    return (cos, sin, -sin, cos, cx - cos * cx + sin * cy, cy - sin * cx - cos * cy)


def _group_matrix(m, xf):
    """组的变换：子坐标 → 按 chOff/chExt 到 off/ext 缩放 → 绕组中心翻转、旋转。组没写完整变换时子形状的坐标就是外层的。"""
    if xf is None:
        return m
    off, ext, cho, che = (xf.find(A + k) for k in ('off', 'ext', 'chOff', 'chExt'))
    if off is None or ext is None or cho is None or che is None:
        return m
    x, y, w, h = int(off.get('x')), int(off.get('y')), int(ext.get('cx')), int(ext.get('cy'))
    sx, sy = w / (int(che.get('cx')) or 1), h / (int(che.get('cy')) or 1)
    cx, cy = x + w / 2, y + h / 2
    local = (sx, 0, 0, sy, x - int(cho.get('x')) * sx, y - int(cho.get('y')) * sy)
    fx, fy = xf.get('flipH') in ('1', 'true'), xf.get('flipV') in ('1', 'true')
    if fx or fy:
        local = _times((-1 if fx else 1, 0, 0, -1 if fy else 1, 2 * cx if fx else 0, 2 * cy if fy else 0), local)
    rot = int(xf.get('rot', '0')) / 60000
    return _times(m, _times(_rotate(rot, cx, cy), local) if rot else local)


def _xfrm(el):
    return next((x for x in (el.find(P + 'spPr/' + A + 'xfrm'), el.find(P + 'xfrm'), el.find(P + 'grpSpPr/' + A + 'xfrm')) if x is not None), None)


def _box(sh, m):
    """形状在页面上的外接矩形 (x, y, w, h)，读不到位置时 None。"""
    if None in (sh.left, sh.top, sh.width, sh.height):
        return None
    xf = _xfrm(sh._element)
    x, y, w, h = sh.left, sh.top, sh.width, sh.height
    n = _times(m, _rotate(int(xf.get('rot', '0')) / 60000, x + w / 2, y + h / 2)) if xf is not None and xf.get('rot') else m
    pts = [(n[0] * px + n[2] * py + n[4], n[1] * px + n[3] * py + n[5]) for px, py in ((x, y), (x + w, y), (x, y + h), (x + w, y + h))]
    xs, ys = [p[0] for p in pts], [p[1] for p in pts]
    return min(xs), min(ys), max(xs) - min(xs), max(ys) - min(ys)


def placement(box, W, H):
    if box is None:
        return 'unknown'
    x, y, w, h = box
    if x + w <= 0 or y + h <= 0 or x >= W or y >= H:
        return 'outside'
    return 'crossing' if x < -EDGE or y < -EDGE or x + w > W + EDGE or y + h > H + EDGE else 'inside'


def _members(shapes):
    """python-pptx 不进 mc:AlternateContent，里面的形状（新式图表、公式、墨迹）要自己取，优先 Fallback。"""
    yield from shapes
    for ac in shapes._spTree.findall(MC + 'AlternateContent'):
        hosts = [ac.find(MC + 'Fallback'), ac.find(MC + 'Choice')]
        el = next((c for h in hosts if h is not None for c in h if c.tag in SHAPE_TAGS), None)
        if el is not None:
            yield shapes._shape_factory(el)


def walk(shapes, m=_ID, path=(), hidden=False):
    """按叠放顺序给出每个形状（含组和组内）：(形状, 外接矩形或 None, 所在组的名字路径, 是否窗格隐藏, 读取错误或 None)。"""
    for sh in _members(shapes):
        try:
            box, name = _box(sh, m), sh.name
        except Exception as e:      # 墨迹（contentPart）一类 python-pptx 读不了的对象，报出来，不跳过
            yield sh, None, path, hidden, f'{etree.QName(sh._element).localname}：{e}'
            continue
        nv = sh._element.find('*/' + P + 'cNvPr')
        hide = hidden or (nv is not None and nv.get('hidden') in ('1', 'true'))
        yield sh, box, path, hide, None
        if isinstance(sh, GroupShape):
            yield from walk(sh.shapes, _group_matrix(m, _xfrm(sh._element)), path + (name,), hide)


def placements(shapes, W, H):
    """{形状元素: 位置}。inspect、clean 用它区分版面内外。"""
    return {sh._element: placement(box, W, H) for sh, box, *_ in walk(shapes)}


# 内容

def _chart(sh):
    """图表数据从 XML 读：所有图表类型、组合图都一样，python-pptx 的 plots 取不到散点的 x。"""
    space = sh.chart._chartSpace
    title = space.find(C + 'chart/' + C + 'title')
    out = {'title': ''.join(t.text or '' for t in title.iter(A + 't')) if title is not None else None, 'series': []}

    def points(ser, *names):
        el = next((e for e in map(ser.find, [C + n for n in names]) if e is not None), None)
        vals = {}
        for pt in el.iter(C + 'pt') if el is not None else ():
            v = pt.find(C + 'v')
            i = int(pt.get('idx'))
            vals[i] = f"{vals[i]} {v.text}" if i in vals else (v.text if v is not None else None)     # 多级类别按级拼起来
        return [vals.get(i) for i in range(max(vals) + 1)] if vals else []

    for ser in space.iter(C + 'ser'):
        tx = ser.find(C + 'tx')
        out['series'].append({'name': ''.join(v.text or '' for v in tx.iter(C + 'v')) if tx is not None else None,
                              'type': etree.QName(ser.getparent()).localname, 'categories': points(ser, 'cat', 'xVal'),
                              'values': [_number(v) for v in points(ser, 'val', 'yVal')]})
    return out


def _crop(blip):
    rect = blip.getparent().find(A + 'srcRect')
    if rect is None:
        return None
    side = {k: max(0, int(rect.get(k, '0'))) / 100000 for k in 'ltrb'}
    share = 1 - max(0, 1 - side['l'] - side['r']) * max(0, 1 - side['t'] - side['b'])
    return {'left': round(side['l'], 3), 'top': round(side['t'], 3), 'right': round(side['r'], 3), 'bottom': round(side['b'], 3), 'share': round(share, 3)} if share > 0 else None


def _outside_share(box, W, H):
    x, y, w, h = box
    if w <= 0 or h <= 0:
        return 0 if placement(box, W, H) == 'inside' else 1
    keep = max(0, min(x + w, W) - max(x, 0)) * max(0, min(y + h, H) - max(y, 0))
    return round(1 - keep / (w * h), 3)


class _Assets:
    """导出到 assets/，同一份内容只存一个文件。"""

    def __init__(self, root):
        self.root, self.seen = root, {}

    def _same(self, name, blob):
        with open(os.path.join(self.root, name), 'rb') as f:
            return f.read() == blob

    def save(self, blob, stem, ext):
        key = hashlib.sha1(blob).hexdigest()
        if key not in self.seen:
            os.makedirs(self.root, exist_ok=True)
            base = re.sub(r'[^0-9A-Za-z\u4e00-\u9fff_-]+', '_', stem).strip('_')
            name, k = f'{base}.{ext}', 1
            while os.path.exists(os.path.join(self.root, name)) and not self._same(name, blob):      # 同一目录重跑时覆盖同样的内容，不另存
                k += 1
                name = f'{base}_{k}.{ext}'
            with open(os.path.join(self.root, name), 'wb') as f:
                f.write(blob)
            self.seen[key] = name
        return 'assets/' + self.seen[key]


def _number(text):
    if text is None:
        return None
    try:
        v = float(text)
    except ValueError:
        return text
    return int(v) if v.is_integer() else v


def _part_ext(part, fallback='bin'):
    return str(part.partname).rsplit('.', 1)[-1] if '.' in str(part.partname) else fallback


def _describe(sh, box, path, hide, ctx, W, H, crop_min):
    """一个形状的内容。图片和嵌入文件只有看不到的对象才导出，版面内的渲染图上有。"""
    place = placement(box, W, H)
    el = sh._element
    nv = el.find('*/' + P + 'cNvPr')
    item = {'shape': sh.name, 'id': sh.shape_id, 'kind': 'shape', 'placement': place,
            'cm': [round(v / CM, 2) for v in box] if box else None}
    if path:
        item['group'] = ' > '.join(path)
    if place == 'crossing':
        item['outsideShare'] = _outside_share(box, W, H)
    if hide:
        item['hiddenInPane'] = True
    if nv is not None and (nv.get('descr') or nv.get('title')):
        item['alt'] = nv.get('descr') or nv.get('title')
    unseen = place != 'inside' or hide
    stem = f"p{ctx['page']:03d}_{LABEL[place]}_{sh.name}"

    if sh.has_text_frame and sh.text_frame.text.strip():
        item.update(kind='text', text=sh.text_frame.text.replace('\v', '\n').strip())
    if getattr(sh, 'has_table', False) and sh.has_table:
        item.update(kind='table', rows=[['' if c.is_spanned else c.text_frame.text.replace('\v', '\n').strip() for c in r.cells] for r in sh.table.rows])
    if getattr(sh, 'has_chart', False) and sh.has_chart:
        item.update(kind='chart', chart=_chart(sh))
        ext_data = sh.chart._chartSpace.find(C + 'externalData')
        rel = sh.chart.part.rels.get(ext_data.get(R + 'id')) if ext_data is not None else None
        if rel is not None and rel.is_external:      # 数据链到外部文件：路径常指向别的项目
            item['workbookLink'] = rel.target_ref
        elif rel is not None and unseen:
            item['workbook'] = ctx['assets'].save(rel.target_part.blob, stem, _part_ext(rel.target_part, 'xlsx'))
    rel_ids = el.find('.//' + DGM_REL)
    if rel_ids is not None:     # SmartArt 的文字在单独的数据部件里
        data = etree.fromstring(sh.part.related_part(rel_ids.get(R + 'dm')).blob)
        item.update(kind='smartart', text='\n'.join(t.text for t in data.iter(A + 't') if t.text and t.text.strip()))
    ole = next(el.iter(P + 'oleObj'), None)
    if ole is not None:
        item.update(kind='ole', progId=ole.get('progId'))
        if ole.find(P + 'embed') is None and ole.find(P + 'link') is not None:
            item['linked'] = True
        elif unseen and ole.get(R + 'id'):
            part = sh.part.related_part(ole.get(R + 'id'))
            item['embedded'] = ctx['assets'].save(part.blob, stem, _part_ext(part))
    pics = []
    for k, blip in enumerate(el.iter(A + 'blip')):
        crop = _crop(blip)
        pic = {'crop': crop} if crop else {}
        if blip.get(R + 'link') and not blip.get(R + 'embed'):
            pic['linked'] = sh.part.rels[blip.get(R + 'link')].target_ref
        elif blip.get(R + 'embed'):
            big_crop = crop is not None and crop['share'] >= crop_min
            if unseen or big_crop:
                part = sh.part.related_part(blip.get(R + 'embed'))
                pic['file'] = ctx['assets'].save(part.blob, stem if not k else f'{stem}_{k + 1}', _part_ext(part))
            if big_crop:
                pic['cropClue'] = True
        pics.append(pic)
    if pics:
        item['pictures'] = pics
        if item['kind'] == 'shape':
            item['kind'] = 'picture'
    if el.find('.//' + A + 'videoFile') is not None or el.find('.//' + A + 'audioFile') is not None:
        item['media'] = True
    item['hasContent'] = any(k in item for k in ('text', 'rows', 'chart', 'pictures', 'progId')) or item['kind'] == 'smartart'
    return item


def _objects(shapes, ctx, W, H, crop_min, errors, where):
    out = []
    for sh, box, path, hide, err in walk(shapes):
        if err:
            errors.append({**where, 'error': f'读不了这个对象（{err}）'})
            out.append({'shape': '（读不了）', 'kind': 'unreadable', 'placement': 'unknown', 'cm': None, 'hasContent': True, 'error': err})
            continue
        if isinstance(sh, GroupShape):
            continue
        try:
            out.append(_describe(sh, box, path, hide, ctx, W, H, crop_min))
        except Exception as e:
            errors.append({**where, 'shape': sh.name, 'error': f'{type(e).__name__}: {e}'})
            out.append({'shape': sh.name, 'kind': 'unreadable', 'placement': placement(box, W, H), 'cm': [round(v / CM, 2) for v in box] if box else None,
                        'hasContent': True, 'error': str(e)})
    return out


def _norm(text):
    return re.sub(r'(?<=\d),(?=\d)', '', ''.join(str(text).split()))


def _row_keys(rows):
    """{(首列文字, 第几次出现): 行号}。"""
    seen, out = {}, {}
    for i, row in enumerate(rows):
        label = _norm(row[0])
        seen[label] = seen.get(label, 0) + 1
        out[(label, seen[label])] = i
    return out


def _column_names(rows):
    """每列的名字：表头行（第一行，和紧接着首列为空的行）按列拼起来，合并单元格向右填。"""
    head = 1
    while head < len(rows) and not rows[head][0].strip():
        head += 1
    names = [[] for _ in rows[0]]
    for row in rows[:head]:
        last = ''
        for k, cell in enumerate(row):
            last = cell.strip() or last
            if last and last not in names[k]:
                names[k].append(last)
    return [' / '.join(n) for n in names], head


def _compare_tables(inside, outside):
    """两张表是不是同一内容的两版：列数相同、行标题和首行基本相同、数据大半相同（或者数据格够多）。不是时给 None。"""
    a, b = inside['rows'], outside['rows']
    if not a or not b or len(a[0]) != len(b[0]) or len(a[0]) < 2:
        return None
    ka, kb = _row_keys(a), _row_keys(b)
    common = [k for k in ka if k in kb]
    if len(common) < VERSION_LABELS * max(len(ka), len(kb)) or sum(_norm(x) == _norm(y) for x, y in zip(a[0], b[0])) < len(a[0]) / 2:
        return None
    columns, head = _column_names(a)
    cells, total = [], 0
    for key in common:
        i, j = ka[key], kb[key]
        if i < head:
            continue
        for k in range(1, len(a[0])):
            total += 1
            if _norm(a[i][k]) != _norm(b[j][k]):
                cells.append({'row': a[i][0].strip(), 'column': columns[k], 'inside': a[i][k], 'outside': b[j][k]})
    if not total:
        return None
    same = 1 - len(cells) / total
    if same < VERSION_SAME and total < VERSION_CELLS:
        return None
    return {'inside': inside['shape'], 'outside': outside['shape'], 'sameShare': round(same, 3), 'cells': cells,
            'rowsOnlyInside': [a[ka[k]][0].strip() for k in ka if k not in kb], 'rowsOnlyOutside': [b[kb[k]][0].strip() for k in kb if k not in ka]}


def _table_versions(objs):
    """版面外的表和同页版面内最像的那张比。一模一样的不算线索。给对得上的版面外对象标 olderVersionOf，内容不动。"""
    tables = [o for o in objs if o['kind'] == 'table' and o.get('rows')]
    shown = [o for o in tables if o['placement'] == 'inside' and not o.get('hiddenInPane')]
    out = []
    for off in (o for o in tables if o['placement'] == 'outside'):
        found = [c for c in (_compare_tables(i, off) for i in shown) if c]
        best = max(found, key=lambda c: c['sameShare'], default=None)
        if best and (best['cells'] or best['rowsOnlyInside'] or best['rowsOnlyOutside']):
            off['olderVersionOf'] = best['inside']
            out.append(best)
    return out


def _md_versions(versions):
    if not versions:
        return []
    out = ['### 版面内外两版对不上的表格', '', '默认以版面内为准，版面外视为旧版或暂存；这里只列线索，没有改动任何内容。', '']
    for v in versions:
        out += [f"**版面内 {v['inside']}**（采用）和**版面外 {v['outside']}**：{v['sameShare']:.0%} 的数据格相同，{len(v['cells'])} 格不同", '']
        if v['cells']:
            out += _md_table([['行', '列', '版面内', '版面外']] + [[c['row'], c['column'], c['inside'], c['outside']] for c in v['cells']]) + ['']
        if v['rowsOnlyInside']:
            out += [f"只有版面内有的行：{'、'.join(v['rowsOnlyInside'])}", '']
        if v['rowsOnlyOutside']:
            out += [f"只有版面外有的行：{'、'.join(v['rowsOnlyOutside'])}", '']
    return out


def _notes(slide):
    frame = slide.notes_slide.notes_text_frame if slide.has_notes_slide else None
    return frame.text.strip() if frame is not None else ''


def _comments(slide):
    out = []
    for rel in slide.part.rels.values():
        if not rel.is_external and rel.reltype.lower().endswith(('comments', 'comment')):
            out += [t for t in (''.join(cm.itertext()).strip() for cm in etree.fromstring(rel.target_part.blob)) if t]
    return out


def _title(slide, objs, H):
    """页标题，只用来在清单里认页：标题占位符；手工稿没有时取页面上部最靠上的短文字。"""
    t = slide.shapes.title
    if t is not None and t.has_text_frame and t.text_frame.text.strip():
        return ' '.join(t.text_frame.text.split())
    top = [o for o in objs if o.get('text') and o['placement'] == 'inside' and o['cm'] and o['cm'][1] * CM < H * .25 and len(o['text']) <= TITLE_MAX]
    return ' '.join(min(top, key=lambda o: o['cm'][1])['text'].split()) if top else None


def _cropped(obj):
    return any(p.get('cropClue') for p in obj.get('pictures', ()))


# 没审计的文件

def _not_audited(path, err):
    item = {'file': str(path), 'audited': False}
    if os.path.isdir(path):
        return {**item, 'format': 'directory', 'reason': '这是目录，要逐个文件传入', 'advice': '用 shell 通配符传入目录里的 pptx'}
    if not os.path.isfile(path):
        return {**item, 'format': 'missing', 'reason': '找不到文件', 'advice': '核对路径'}
    with open(path, 'rb') as f:
        head = f.read(512)
    lead = head.lstrip(b'\xef\xbb\xbf \t\r\n')
    item['head'] = head[:16].hex()
    if lead.startswith((b'<?xml', b'<pkg:package')) or b'pkg:package' in head:
        return {**item, 'format': 'flat-xml', 'reason': 'Flat OPC XML：整份演示文稿是一个 XML 文件，不是 zip，python-pptx 读不了',
                'advice': '在 PowerPoint 里另存为标准 pptx 后再跑'}
    if head[:8] == bytes.fromhex('d0cf11e0a1b11ae1'):
        return {**item, 'format': 'ole-compound', 'reason': '旧版 .ppt，或带密码的 pptx（OLE 复合文件）',
                'advice': '在 PowerPoint 里解密并另存为 pptx 后再跑'}
    if bytes(b ^ 0x5c for b in head[:7]) == b'DG_FILE':
        return {**item, 'format': 'encrypted-wrapper', 'reason': '文件被加密外壳包住（文件头异或 0x5C 后是 DG_FILE），不是 pptx 包',
                'advice': '在能解密的环境里另存，或请对方提供未加密的版本'}
    if zipfile.is_zipfile(path):
        return {**item, 'format': 'broken-package', 'reason': f'是 zip，但不是能读的 pptx：{type(err).__name__}: {err}',
                'advice': '在 PowerPoint 里打开，修复后另存'}
    return {**item, 'format': 'unknown', 'reason': f'不是 zip，认不出格式：{type(err).__name__}: {err}', 'advice': '在 PowerPoint 里打开，另存为 pptx 后再跑'}


# 输出

def _cell(t):
    return str(t).replace('|', '\\|').replace('\n', '<br>')


def _md_table(rows):
    if not rows:
        return []
    width = max(map(len, rows))
    rows = [r + [''] * (width - len(r)) for r in rows]
    return ['| ' + ' | '.join(map(_cell, rows[0])) + ' |', '|' + ' --- |' * width] + ['| ' + ' | '.join(map(_cell, r)) + ' |' for r in rows[1:]]


def _md_chart(chart):
    blocks = [f"图表标题：{chart['title']}"] if chart['title'] else []
    series = chart['series']
    if not series:
        return blocks + ['（没读到系列）']
    cats = next((s['categories'] for s in series if s['categories']), [])
    n = max(len(cats), *(len(s['values']) for s in series))
    head = [''] + [s['name'] or f"系列{k}" for k, s in enumerate(series, 1)]
    rows = [[cats[i] if i < len(cats) else '', *[s['values'][i] if i < len(s['values']) else '' for s in series]] for i in range(n)]
    return blocks + ['\n'.join(_md_table([head] + rows))]


def _md_body(obj):
    """对象的内容：一组 markdown 段落，段落之间空一行。"""
    out = []
    if obj.get('text'):
        out += [obj['text'].replace('\n', '  \n')]
    if obj.get('rows'):
        out.append('\n'.join(_md_table(obj['rows'])))
    if obj.get('chart'):
        out += _md_chart(obj['chart'])
    if obj.get('workbook'):
        out.append(f"内嵌工作簿：[{obj['workbook']}]({obj['workbook']})")
    if obj.get('workbookLink'):
        out.append(f"图表数据链接到外部文件：{obj['workbookLink']}")
    if obj.get('progId'):
        out.append(f"嵌入对象 {obj['progId']}" + (f"：[{obj['embedded']}]({obj['embedded']})" if obj.get('embedded') else '（链接到外部文件）' if obj.get('linked') else ''))
    for p in obj.get('pictures', ()):
        if p.get('crop') and p.get('cropClue'):
            c = p['crop']
            out.append(f"裁掉 {c['share']:.0%}（左 {c['left']:.0%}、上 {c['top']:.0%}、右 {c['right']:.0%}、下 {c['bottom']:.0%}），assets 里是原图")
        if p.get('file'):
            out.append(f"![{obj['shape']}]({p['file']})")
        if p.get('linked'):
            out.append(f"图片链接到外部文件：{p['linked']}，没有内容可导出")
    if obj.get('alt'):
        out.append(f"替代文字：{obj['alt']}")
    if obj.get('error'):
        out.append(f"读取失败：{obj['error']}")
    return out


def _md_head(obj):
    cm = '位置读不到' if not obj.get('cm') else f"x {obj['cm'][0]}、y {obj['cm'][1]}、宽 {obj['cm'][2]}、高 {obj['cm'][3]} cm"
    bits = [{'text': '文字', 'table': '表格', 'chart': '图表', 'picture': '图片', 'ole': '嵌入对象', 'smartart': 'SmartArt', 'unreadable': '读不了'}.get(obj['kind'], '形状'), cm]
    if obj.get('group'):
        bits.append(f"组内：{obj['group']}")
    if 'outsideShare' in obj:
        bits.append(f"{obj['outsideShare']:.0%} 在版面外")
    if obj.get('olderVersionOf'):
        bits.append(f"疑似旧版，版面内是 {obj['olderVersionOf']}")
    if obj.get('hiddenInPane'):
        bits.append('选择窗格里隐藏')
    if obj.get('media'):
        bits.append('音视频')
    return f"**{obj['shape']}**（{'，'.join(bits)}）"


def _md_section(title, objs):
    if not objs:
        return []
    out = [f'### {title}', '']
    for obj in objs:
        out += [_md_head(obj), ''] + [x for block in _md_body(obj) for x in (block, '')]
    return out


def _write_list(data, path):
    s = data['summary']
    hide = '、'.join(map(str, s['hiddenPages'])) or '无'
    lines = [f"# 版面外清单：{os.path.basename(data['file'])}", '',
             f"共 {s['pages']} 页，隐藏页 {len(s['hiddenPages'])} 页（{hide}）。有内容的对象：版面外 {s['outside']} 个，跨边 {s['crossing']} 个，"
             f"选择窗格隐藏 {s['hiddenInPane']} 个，位置不明 {s['unknown']} 个；裁掉 {s['cropMin']:.0%} 以上的图片 {s['cropped']} 张；"
             f"版面内外两版对不上的表格 {s['tableVersions']} 组。", '',
             '这些内容都不在渲染图上；隐藏页的版面内内容见 `全文提取.md`。' + NOTICE, '']
    if data['errors']:
        lines += ['## 没读全的对象', ''] + [f"- {'第 ' + str(e['page']) + ' 页 ' if 'page' in e else ''}{e.get('shape', '')} {e['error']}" for e in data['errors']] + ['']
    rows = [['页', '隐藏页', '版面外', '跨边', '窗格隐藏', '裁剪线索', '内外差异', '标题']]
    for pg in data['pages']:
        c = pg['counts']
        if c['outside'] or c['crossing'] or c['hiddenInPane'] or c['cropped'] or c['unknown']:
            rows.append([pg['page'], '是' if pg['hidden'] else '', c['outside'] or '', c['crossing'] or '', c['hiddenInPane'] or '', c['cropped'] or '', len(pg['tableVersions']) or '', pg['title'] or ''])
    lines += ['## 索引', ''] + (_md_table(rows) if len(rows) > 1 else ['没有版面外、跨边、窗格隐藏或裁剪较多的对象。']) + ['']
    for pg in data['pages']:
        objs = pg['objects']
        content = [o for o in objs if o['hasContent']]
        out, cross = [o for o in content if o['placement'] == 'outside'], [o for o in content if o['placement'] == 'crossing']
        where = [o for o in content if o['placement'] == 'unknown']
        hid = [o for o in content if o['placement'] == 'inside' and o.get('hiddenInPane')]
        crop = [o for o in objs if _cropped(o) and o['placement'] == 'inside' and not o.get('hiddenInPane')]
        rest = [o['shape'] for o in objs if not o['hasContent'] and o['placement'] != 'inside']
        if not (out or cross or where or hid or crop or rest):
            continue
        lines += [f"## 第 {pg['page']} 页{'（隐藏页）' if pg['hidden'] else ''}" + (f"：{pg['title']}" if pg['title'] else ''), '']
        lines += _md_versions(pg['tableVersions']) + _md_section('版面外', out) + _md_section('跨边', cross) + _md_section('位置不明', where) + _md_section('版面内但选择窗格隐藏', hid)
        if crop:
            lines += ['### 裁剪较多的图片（裁掉的部分可能有内容，assets 里是原图）', '']
            for o in crop:
                lines += [_md_head(o), ''] + [x for block in _md_body(o) for x in (block, '')]
        if rest:
            lines += [f"另有 {len(rest)} 个没有内容的形状在版面外或跨边：{'、'.join(rest[:12])}{'……' if len(rest) > 12 else ''}", '']
    if data['layouts']:
        lines += ['## 母版和版式', '']
        for lay in data['layouts']:
            lines += _md_section(lay['layout'], lay['objects'])
    with open(path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines).rstrip() + '\n')


def _write_text(data, path):
    lines = [f"# 全文提取：{os.path.basename(data['file'])}", '', '每个对象标明位置；[版面外]、[跨边] 的内容不在渲染图上。' + NOTICE, '']
    for pg in data['pages']:
        lines += [f"## 第 {pg['page']} 页{'（隐藏页）' if pg['hidden'] else ''}" + (f"：{pg['title']}" if pg['title'] else ''), '']
        for obj in pg['objects']:
            if obj['hasContent']:
                tag = f"[{LABEL[obj['placement']]}]" + ('[窗格隐藏]' if obj.get('hiddenInPane') else '') + (f"[疑似旧版，版面内是 {obj['olderVersionOf']}]" if obj.get('olderVersionOf') else '')
                if obj['kind'] == 'picture' and not any(p.get('file') for p in obj.get('pictures', ())):
                    lines += [f"{tag} {obj['shape']}（图片）", '']
                    continue
                lines += [f"{tag} {obj['shape']}", ''] + [x for block in _md_body(obj) for x in (block, '')]
        if pg['notes']:
            lines += ['[备注]', '', pg['notes'].replace('\n', '  \n'), '']
        lines += [x for c in pg['comments'] for x in (f'[批注] {c}', '')]
    with open(path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines).rstrip() + '\n')


# 连版面外一起渲染

def _shift(tree, dx, dy):
    for el in tree:
        if el.tag == MC + 'AlternateContent':
            for host in el:
                _shift(host, dx, dy)
        elif el.tag in SHAPE_TAGS:
            xf = _xfrm(el)
            off = xf.find(A + 'off') if xf is not None else None
            if off is not None:
                off.set('x', str(int(off.get('x')) + dx))
                off.set('y', str(int(off.get('y')) + dy))


def _trim(lo, hi, size):
    """超过 PowerPoint 页面上限时按比例收窄两侧，版面始终留在窗口里。"""
    if hi - lo <= MAX_SLIDE - 2 * PAD:
        return lo, hi, False
    k = (MAX_SLIDE - 2 * PAD - size) / ((-lo) + (hi - size))
    return lo * k, size + (hi - size) * k, True


def _extent(slide, W, H):
    boxes = [b for _, b, *_ in walk(slide.shapes) if b]
    return (min([0] + [b[0] for b in boxes]), min([0] + [b[1] for b in boxes]),
            max([W] + [b[0] + b[2] for b in boxes]), max([H] + [b[1] + b[3] for b in boxes]))


def render_with_offcanvas(path, out_dir, pages, engine='libreoffice', dpi=100):
    """指定页放大画布连版面外一起渲染，红框是版面边界。在临时副本上做，原文件不动。
    所有页用同一个画布（页面尺寸全册一致），渲染后裁成各页自己的范围。"""
    prs = Presentation(path)
    W, H = prs.slide_width, prs.slide_height
    slides = list(prs.slides)
    res = {'engine': engine, 'dpi': dpi, 'pages': [], 'warnings': [], 'notes': []}
    bad = [n for n in pages if not 1 <= n <= len(slides)]
    if bad:
        res['warnings'].append(f"没有这些页：{'、'.join(map(str, bad))}（共 {len(slides)} 页）")
    pages = sorted({n for n in pages if n not in bad})
    if not pages:
        return res
    ext = {n: _extent(slides[n - 1], W, H) for n in pages}
    x0, x1, cut_x = _trim(min(e[0] for e in ext.values()), max(e[2] for e in ext.values()), W)
    y0, y1, cut_y = _trim(min(e[1] for e in ext.values()), max(e[3] for e in ext.values()), H)
    if cut_x or cut_y:
        res['warnings'].append('对象离版面太远，超过 PowerPoint 的页面上限，渲染图只画到上限范围；内容以 offcanvas.json 为准')
    dx, dy = round(PAD - x0), round(PAD - y0)
    for master in prs.slide_masters:
        _shift(master.shapes._spTree, dx, dy)
        for layout in master.slide_layouts:
            _shift(layout.shapes._spTree, dx, dy)
    keep = {slides[n - 1].slide_id for n in pages}
    for n in pages:
        sl = slides[n - 1]
        _shift(sl.shapes._spTree, dx, dy)
        sl._element.attrib.pop('show', None)
        frame = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(dx), Emu(dy), W, H)
        frame.name = 'plan-deck:版面边界'
        frame.fill.background()
        frame.line.color.rgb = RGBColor(255, 0, 0)
        frame.line.width = Pt(3)
        style = frame._element.find(P + 'style')
        if style is not None:       # 主题样式引用会带阴影
            frame._element.remove(style)
    ids = prs.slides._sldIdLst
    for e in list(ids):
        if int(e.get('id')) not in keep:
            prs.part.drop_rel(e.rId)
            ids.remove(e)
    root = prs.part._element
    for el in list(root.iter('{http://schemas.microsoft.com/office/powerpoint/2010/main}sectionLst')) + list(root.iter(P + 'custShowLst')):
        el.getparent().remove(el)       # 节和自定义放映还引用删掉的页，PowerPoint 会报要修复
    prs.slide_width, prs.slide_height = Emu(round(x1 - x0 + 2 * PAD)), Emu(round(y1 - y0 + 2 * PAD))
    root.find(P + 'sldSz').attrib.pop('type', None)
    os.makedirs(out_dir, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='plan-deck-offcanvas-') as tmp:
        work = os.path.join(tmp, 'deck.pptx')
        prs.save(work)
        done = render(work, tmp, engine, dpi)
        pdf = os.path.join(out_dir, f'连版面外.{engine}.pdf')
        shutil.move(done['pdf'], pdf)
        res.update(pdf=pdf, fonts=done['fonts'], warnings=res['warnings'] + done['warnings'], notes=done['notes'])
        if engine == 'libreoffice':
            res['notes'].append('LibreOffice 缺字体时会换字，字形和行距不作数；文字内容以 offcanvas.json 为准')
        if len(done['pages']) != len(pages):
            res['warnings'].append(f"要渲染 {len(pages)} 页，出了 {len(done['pages'])} 张图")
        for n, src in zip(pages, done['pages']):
            e = ext[n]
            with Image.open(src) as im:
                k = im.width / (x1 - x0 + 2 * PAD)
                crop = (max(0, round((e[0] + dx - PAD) * k)), max(0, round((e[1] + dy - PAD) * k)),
                        min(im.width, round((e[2] + dx + PAD) * k)), min(im.height, round((e[3] + dy + PAD) * k)))
                png = os.path.join(out_dir, f'第{n:03d}页_连版面外.png')
                im.crop(crop).save(png)
            res['pages'].append({'page': n, 'png': png, 'canvasCm': [round(W / CM, 2), round(H / CM, 2)],
                                 'extentCm': [round(v / CM, 2) for v in e]})
    return res


# 入口

def scan(path, out, crop_min=CROP_MIN):
    """读一份 pptx，导出素材，返回数据。读不了时抛 NotAudited。"""
    try:
        prs = Presentation(path)
    except Exception as e:
        raise NotAudited(_not_audited(path, e)) from e
    W, H = prs.slide_width, prs.slide_height
    assets = _Assets(os.path.join(out, 'assets'))
    pages, errors = [], []
    for n, slide in enumerate(prs.slides, 1):
        where = {'page': n}
        hidden = slide._element.get('show') == '0'
        try:
            objs = _objects(slide.shapes, {'page': n, 'assets': assets}, W, H, crop_min, errors, where)
            notes, comments, title = _notes(slide), _comments(slide), _title(slide, objs, H)
        except Exception as e:
            errors.append({**where, 'error': f'整页没读全：{type(e).__name__}: {e}'})
            objs, notes, comments, title = [], '', [], None
        content = [o for o in objs if o['hasContent']]
        pages.append({'page': n, 'hidden': hidden, 'title': title, 'notes': notes, 'comments': comments, 'objects': objs, 'tableVersions': _table_versions(objs),
                      'counts': {'outside': sum(o['placement'] == 'outside' for o in content), 'crossing': sum(o['placement'] == 'crossing' for o in content),
                                 'unknown': sum(o['placement'] == 'unknown' for o in content),
                                 'hiddenInPane': sum(bool(o.get('hiddenInPane')) for o in content), 'cropped': sum(_cropped(o) for o in objs)}})
    layouts, seen = [], set()
    for master in prs.slide_masters:
        for host in [master, *master.slide_layouts]:
            if id(host._element) in seen:
                continue
            seen.add(id(host._element))
            objs = [o for o in _objects(host.shapes, {'page': 0, 'assets': assets}, W, H, crop_min, errors, {'layout': host.name}) if o['hasContent'] and (o['placement'] != 'inside' or o.get('hiddenInPane'))]
            if objs:
                layouts.append({'layout': host.name, 'objects': objs})
    total = lambda k: sum(p['counts'][k] for p in pages)
    summary = {'pages': len(pages), 'hiddenPages': [p['page'] for p in pages if p['hidden']], 'outside': total('outside'), 'crossing': total('crossing'),
               'unknown': total('unknown'), 'hiddenInPane': total('hiddenInPane'), 'cropped': total('cropped'), 'cropMin': crop_min,
               'pagesWithContentOutside': [p['page'] for p in pages if p['counts']['outside']],
               'tableVersions': sum(len(p['tableVersions']) for p in pages), 'pagesWithTableVersions': [p['page'] for p in pages if p['tableVersions']]}
    return {'file': str(path), 'audited': True, 'pageCm': [round(W / CM, 2), round(H / CM, 2)], 'summary': summary, 'errors': errors, 'pages': pages, 'layouts': layouts}


class NotAudited(Exception):
    def __init__(self, info):
        super().__init__(info['reason'])
        self.info = info


def page_spec(spec, data):
    """'32,58' '24-31' → 页码列表；'all' = 有版面外或跨边内容的页。"""
    if spec.strip() == 'all':
        return [p['page'] for p in data['pages'] if p['counts']['outside'] or p['counts']['crossing']]
    out = []
    for part in spec.split(','):
        a, _, b = part.strip().partition('-')
        out += range(int(a), int(b or a) + 1)
    return out


def offcanvas(path, out, crop_min=CROP_MIN, render_pages=None, engine='libreoffice', dpi=100):
    """审计一份 pptx，写输出目录，返回摘要。读不了的文件抛 NotAudited。"""
    data = scan(path, out, crop_min)
    os.makedirs(out, exist_ok=True)
    files = {'json': os.path.join(out, 'offcanvas.json'), 'list': os.path.join(out, '版面外清单.md'), 'text': os.path.join(out, '全文提取.md')}
    if render_pages:
        try:
            data['render'] = render_with_offcanvas(path, os.path.join(out, 'renders'), page_spec(render_pages, data), engine, dpi)
        except (RenderError, ValueError) as e:
            data['render'] = {'error': str(e)}
    with open(files['json'], 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    _write_list(data, files['list'])
    _write_text(data, files['text'])
    if os.path.isdir(os.path.join(out, 'assets')):
        files['assets'] = os.path.join(out, 'assets')
    res = {'file': data['file'], 'audited': True, 'summary': data['summary'], 'errors': data['errors'], 'outputs': files}
    if 'render' in data:
        res['render'] = data['render']
    return res


def run(paths, out, crop_min=CROP_MIN, render_pages=None, engine='libreoffice', dpi=100):
    """命令行入口：一份直接写进 out，多份各写一个子目录。没审计的文件列进 notAudited，也写一份 未审计.md。"""
    decks, missed, used = [], [], set()
    for path in paths:
        target = out
        if len(paths) > 1:
            stem, k = os.path.splitext(os.path.basename(path.rstrip('/\\')))[0], 1
            name = stem
            while name in used:
                k += 1
                name = f'{stem}_{k}'
            used.add(name)
            target = os.path.join(out, name)
        try:
            decks.append(offcanvas(path, target, crop_min, render_pages, engine, dpi))
        except NotAudited as e:
            missed.append(e.info)
    if missed:
        os.makedirs(out, exist_ok=True)
        with open(os.path.join(out, '未审计.md'), 'w', encoding='utf-8') as f:
            f.write('# 没有审计的文件\n\n这些文件的版面外内容没有看过，不能当作没有。\n\n')
            f.write(''.join(f"- {m['file']}：{m['reason']}。{m['advice']}。\n" for m in missed))
    render_failed = any('error' in d.get('render', {}) for d in decks)
    return {'ok': not missed and not render_failed and not any(d['errors'] for d in decks), 'notice': NOTICE, 'decks': decks, 'notAudited': missed}
