"""页面框架与组件。接口一律用厘米；颜色、字号只按名字取自项目视觉系统（style.json）。

工具生成的形状统一命名 pd:<页型>:<槽位>[:<序号>]；同一页同一槽位出现多次时自动编号。
"""
import math
import re

from PIL import Image
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.chart.plot import BarPlot, LinePlot, RadarPlot
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.dml import MSO_LINE_DASH_STYLE, MSO_PATTERN
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.util import Cm, Pt

from .blocks import Blocks
from .maps import LineSample, MapAsset, MapError, class_mismatches, legend_rows
from .svgshape import LibrarySymbol, parse as parse_svg
from .measure import LINE, MARK, PT_CM, text_height, text_width
from .style import DataColor, Style, StyleError
from .template import apply_template

ANCHOR = {'top': MSO_ANCHOR.TOP, 'middle': MSO_ANCHOR.MIDDLE, 'bottom': MSO_ANCHOR.BOTTOM}
ALIGN = {'left': PP_ALIGN.LEFT, 'center': PP_ALIGN.CENTER, 'right': PP_ALIGN.RIGHT}
SHAPES = {'rect': MSO_SHAPE.RECTANGLE, 'oval': MSO_SHAPE.OVAL, 'diamond': MSO_SHAPE.DIAMOND}
CHARTS = {'bar': XL_CHART_TYPE.BAR_CLUSTERED, 'bar-stacked': XL_CHART_TYPE.BAR_STACKED,
          'column': XL_CHART_TYPE.COLUMN_CLUSTERED, 'column-stacked': XL_CHART_TYPE.COLUMN_STACKED,
          'line': XL_CHART_TYPE.LINE, 'area': XL_CHART_TYPE.AREA, 'doughnut': XL_CHART_TYPE.DOUGHNUT, 'pie': XL_CHART_TYPE.PIE}
# 文本条目的种类 → (字号级别, 颜色, 加粗)
ITEM_KINDS = {'b': ('body', 'text', False), 'm': ('body', 'muted', False), 'h': ('heading', 'title', True),
              'big': ('figure', 'emphasis', True)}
# 加粗记号 → 颜色；__ 不在这里，用所在段落的颜色
MARK_COLORS = {'**': 'emphasis', '^^': 'title'}
PD_NAME = re.compile(r'pd:([^:]+):([^:]+)(?::\d+)?')


def _pick(table, key, what, enum=None):
    """枚举参数：查表，也直接收 python-pptx 的枚举（enum）；不认识时列出可用值。"""
    if enum is not None and isinstance(key, enum):
        return key
    if key not in table:
        extra = f'，或 {enum.__name__} 的成员' if enum is not None else ''
        raise ValueError(f'{what} {key!r} 不认识；可用：{"、".join(table)}{extra}')
    return table[key]


class Box:
    """组件实际占的位置（cm）。bottom / right 供后面的元素接着排。"""

    def __init__(self, x, y, w, h, shape=None):
        self.x, self.y, self.w, self.h, self.shape = x, y, w, h, shape

    @property
    def bottom(self):
        return self.y + self.h

    @property
    def right(self):
        return self.x + self.w

    def __iter__(self):
        return iter((self.x, self.y, self.w, self.h))

    def __repr__(self):
        return f'Box(x={self.x:.2f}, y={self.y:.2f}, w={self.w:.2f}, h={self.h:.2f})'


def _items(items):
    for it in items:
        kind, text = it if isinstance(it, tuple) else ('b', it)
        if kind not in ITEM_KINDS:
            raise ValueError(f'文本条目种类 {kind!r} 不认识；可用：b 正文、m 次要、h 小标题、big 大号数字')
        yield kind, text


class Deck:
    def __init__(self, style, prs=None):
        self.style = style if isinstance(style, Style) else Style(style) if isinstance(style, dict) else Style.load(style)
        self.prs = Presentation() if prs is None else prs
        if prs is None:
            apply_template(self.prs, self.style)
        self._blank = min(self.prs.slide_layouts, key=lambda lay: len(list(lay.iter_cloneable_placeholders())))
        self.warnings = list(self.style.unknown_keys)      # 写了不生效的键：只跑出页脚本的人也要看到
        self.placements = []
        self.chapter = None         # 页眉章名行（style.header 有 chapter 行时）；page(chapter=) 没给时用它
        self.meanings = {}          # 页码 → {色值: {类别配色 (方案, 类别)}}：查同页、相邻页同色异义
        self.fitted = []            # 矩阵表按版心选的字号：[{page, slot, pt}]

    @classmethod
    def open(cls, path, style):
        """打开已有 pptx 接着改（决策 4）：模板不动，同事的形状不动。deck.at(n) 取已有的页，deck.page() 加新页。"""
        deck = cls(style, Presentation(path))
        got = deck.prs.slide_width / 360000, deck.prs.slide_height / 360000
        if abs(got[0] - deck.style.page_w) > 0.05 or abs(got[1] - deck.style.page_h) > 0.05:
            deck.warn(f'{path} 的页面是 {got[0]:.2f}×{got[1]:.2f} cm，style.json 是 {deck.style.page_w:g}×{deck.style.page_h:g} cm：'
                      f'栏位和字号都按 style.json 算，先确认这份 style.json 对不对')
        return deck

    def at(self, number, page_type=None):
        """已有的第 number 页（1 起算）。页上已有的名字都算占用，新增形状不会重名；页型默认取页上 pd: 名字里最多的那个。"""
        if not 1 <= number <= len(self.prs.slides):
            raise ValueError(f'没有第 {number} 页：这份稿子共 {len(self.prs.slides)} 页')
        slide = self.prs.slides[number - 1]
        types = [m[1] for m in map(PD_NAME.fullmatch, (sh.name for sh in slide.shapes)) if m]
        page_type = page_type or max(set(types), key=types.count, default=None)
        if page_type is None:
            raise ValueError(f'第 {number} 页没有 pd: 形状，认不出页型：写 at({number}, page_type=…)')
        page = Page(self, slide, page_type, number)
        page._used.update(el.get('name') for el in slide._element.iter(qn('p:cNvPr')))
        for sh in slide.shapes:
            m = PD_NAME.fullmatch(sh.name)
            if m and m[1] == page_type:
                page._names.setdefault(f'pd:{m[1]}:{m[2]}', []).append(sh)
        return page

    def warn(self, message, once=False):
        """记一条警告；once=True 时同样的话只记一次（同一份发布版放多张图时）。"""
        if not (once and message in self.warnings):
            self.warnings.append(message)

    def page(self, page_type, title='', parent=None, lead=None, header=True, chapter=None):
        """新增一页并画页眉。title 是本页节点名（与大纲一字不差），parent 是上级节点名，lead 是副标题，默认不写；
        红字结论不走 lead，用 p.conclusion()。chapter 是页眉章名行（style.header 有 chapter 行时画），不给就用 deck.chapter。
        header=False 不画页眉（满版图、章小结页）；要把标题压在图上时，放图后再调 p.header()。"""
        slide = self.prs.slides.add_slide(self._blank)
        for ph in list(slide.placeholders):     # 接手的稿子可能没有空白版式
            ph._element.getparent().remove(ph._element)
        page = Page(self, slide, page_type, len(self.prs.slides))
        page.title, page.parent, page.lead = title, parent, lead
        page.chapter = chapter if chapter is not None else self.chapter
        if header:
            page.header()
        return page

    def save(self, path):
        self._check_meanings()
        self.prs.save(path)
        return path

    def _check_meanings(self):
        """同一颜色在同一页或相邻两页表示两种类别时记警告（一种颜色一类内容，visual-language.md“配色”）。"""
        seen = set()
        pages = sorted(self.meanings)
        for n in pages:
            for m in (n, n + 1):
                if m not in self.meanings:
                    continue
                for hexv in self.meanings[n].keys() & self.meanings[m].keys():
                    refs = sorted(self.meanings[n][hexv] | self.meanings[m][hexv])
                    if len(refs) > 1 and (hexv, tuple(refs)) not in seen:
                        seen.add((hexv, tuple(refs)))
                        where = f'第 {n} 页' if m == n else f'第 {n}、{m} 页'
                        self.warn(f'{where}：{hexv} 同时表示 {"、".join(".".join(r) for r in refs)}，读者会当成同一类；'
                                  f'换一组颜色，确是同一类就用同一个类别')


class Page(Blocks):
    def __init__(self, deck, slide, page_type, number):
        self.deck, self.slide, self.type, self.number = deck, slide, page_type, number
        self.style = deck.style
        self.title = self.parent = self.lead = self.chapter = None
        self._names = {}
        self._used = set()

    # ---------- 命名 ----------
    def name(self, shape, slot, i=None):
        """命名 pd:<页型>:<槽位>[:<序号>]。不给序号时，同一槽位第二次出现起自动编号（第一个改成 :0）；
        给了序号却与本页已有的名字相同时报错。"""
        base = f'pd:{self.type}:{slot}'
        if i is not None:
            full = f'{base}:{i}'
            if full in self._used:
                raise ValueError(f'第 {self.number} 页形状重名：{full}；同一槽位的序号不能重复')
            shape.name = full
            self._used.add(full)
            return shape
        seen = self._names.setdefault(base, [])
        if len(seen) == 1:
            seen[0].name = f'{base}:0'
            self._used.discard(base)
            self._used.add(seen[0].name)
        seen.append(shape)
        shape.name = base if len(seen) == 1 else f'{base}:{len(seen) - 1}'
        if shape.name in self._used:
            raise ValueError(f'第 {self.number} 页形状重名：{shape.name}；同一槽位不要混用自动编号和手写序号，'
                             f'接着编号写 i=p.next_index(槽位)')
        self._used.add(shape.name)
        return shape

    def next_index(self, slot):
        """本页这个槽位下一个没用过的序号（同一页画两份图例、两组节点时接着编）。"""
        prefix = f'pd:{self.type}:{slot}:'
        used = [int(n[len(prefix):]) for n in self._used if n.startswith(prefix) and n[len(prefix):].isdigit()]
        return max(used) + 1 if used else 0

    # ---------- 取值 ----------
    def rgb(self, ref):
        hexv = self.style.color(ref)
        if isinstance(ref, tuple):      # 类别配色：记下这页用它表示哪一类
            self.deck.meanings.setdefault(self.number, {}).setdefault(hexv, set()).add(tuple(ref))
        return RGBColor.from_string(hexv.lstrip('#'))

    @property
    def family(self):
        """量字宽用的字体（页面中文字体）。"""
        return self.style.fonts['cjk']

    def _size(self, role):
        return Pt(self.style.size(role))

    # ---------- 文字 ----------
    def _marked(self, text):
        """把 text 拆成 [(片段, 记号或 None)]。__ 和 ^^ 没配对、中间是空的（如填空横线 ____），或记号嵌套，都报错，
        不把记号留在页面上，也不吞字。** 没配对不报错，后面的字都算强调：旧脚本照旧能跑。"""
        def bad(problem):
            hint = '（要显示下划线本身用全角 ＿）' if mark == '__' else ''
            return ValueError(f'第 {self.number} 页文字标记 {mark} {problem}{hint}：{text}')
        toks = MARK.split(text)
        out, mark = [], None
        for k, tok in enumerate(toks):
            if k % 2 == 0:
                if tok:
                    out.append((tok, mark))
            elif mark is None:
                mark = tok
            elif tok != mark:
                raise bad(f'里不能再套 {tok}')
            elif mark != '**' and not toks[k - 1]:
                raise bad('中间是空的')
            else:
                mark = None
        if mark not in (None, '**'):
            raise bad('没有配对')
        return out

    def runs(self, paragraph, text, size='body', color='text', bold=False):
        """text 里 **…** 是强调色加粗，__…__ 是所在段落的颜色加粗，^^…^^ 是标题色加粗；不嵌套。
        size 是字号级别名；矩阵表按版心选字号时直接给 pt。"""
        fonts = self.style.fonts
        ppr = paragraph._p.get_or_add_pPr()       # 中文断行：标点不落行首
        ppr.set('eaLnBrk', '1')
        ppr.set('hangingPunct', '1')
        for part, mark in self._marked(text):
            r = paragraph.add_run()
            r.text = part
            f = r.font
            f.name, f.size = fonts.get('latin', fonts['cjk']), Pt(size) if isinstance(size, (int, float)) else self._size(size)
            f.bold = bold or mark is not None
            f.color.rgb = self.rgb(MARK_COLORS.get(mark, color))
            rpr = r._r.get_or_add_rPr()
            rpr.set('lang', 'zh-CN')
            ea = rpr.find(qn('a:ea'))
            if ea is None:
                ea = rpr.makeelement(qn('a:ea'), {})
                rpr.append(ea)
            ea.set('typeface', fonts['cjk'])

    def chart_font(self, font, size='body'):
        """图表文字（标题、图例、刻度、数据标签）的字号、字体和颜色（style.chart.text）。西文、东亚字体都要写：
        只写一个时 PowerPoint 的中文落回默认字体（如宋体），LibreOffice 渲染看不出来。"""
        fonts = self.style.fonts
        font.size = self._size(size)
        font.name = fonts.get('latin', fonts['cjk'])
        font.color.rgb = self.rgb(self.style.metrics['chart']['text'])
        rpr = font._rPr
        ea = rpr.find(qn('a:ea'))
        if ea is None:
            ea = rpr.makeelement(qn('a:ea'), {})
            rpr.find(qn('a:latin')).addnext(ea)
        ea.set('typeface', fonts['cjk'])

    def items_height(self, items, w):
        """估算 text() 排这些条目的高度（cm）。"""
        t = self.style.metrics['text']
        h = 0.1
        for kind, text in _items(items):
            size = self.style.size(ITEM_KINDS[kind][0])
            mark = t['headingMark'] if kind == 'h' else ''
            h += text_height(mark + text, w, size, t['lineSpacing'], self.family, kind in ('h', 'big')) + (0.21 if kind == 'h' else 0.07)
        return h

    def _check_fit(self, y, h, est, size_pt, slot):
        """放不下要主动说（决策 11）：框高不够差出半行、或下边出了页面就记警告，不改版面。"""
        where = f'第 {self.number} 页 {self.type}:{slot}'
        if est > h + size_pt * PT_CM * LINE / 2:
            self.deck.warn(f'{where}：字约需 {est:.1f} cm 高，框只有 {h:.1f} cm，会压到下面；拆分节点或换版式，不要缩字号')
        if y + max(h, est) > self.style.page_h + 0.01:
            self.deck.warn(f'{where}：下边到 {y + max(h, est):.1f} cm，超出页高 {self.style.page_h:g} cm；拆分节点或换版式')

    def text(self, x, y, w, items, h=None, slot='text', i=None, anchor='top', align='left'):
        """正文块。items：字符串（正文）或 ('h', 小标题) / ('b', 正文) / ('m', 次要) / ('big', 大号数字)。
        h 不给时按估算高度；返回 Box，bottom 为估算的下边。放不下或出了页面时记警告。"""
        t = self.style.metrics['text']
        est = self.items_height(items, w)
        h = est if h is None else h
        self._check_fit(y, h, est, max(self.style.size(ITEM_KINDS[k][0]) for k, _ in _items(items)), slot)
        tb = self.name(self.slide.shapes.add_textbox(Cm(x), Cm(y), Cm(w), Cm(h)), slot, i)
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Cm(t['marginCm'])
        tf.vertical_anchor = _pick(ANCHOR, anchor, 'anchor')
        for n, (kind, text) in enumerate(_items(items)):
            p = tf.paragraphs[0] if n == 0 else tf.add_paragraph()
            p.alignment = _pick(ALIGN, align, 'align')
            p.line_spacing = t['lineSpacing']
            size, color, bold = ITEM_KINDS[kind]
            if kind == 'h':
                p.space_before = Pt(t['headingSpaceBeforePt'])
                text = t['headingMark'] + text
            elif kind == 'b':
                p.space_before = Pt(t['bodySpaceBeforePt'])
            self.runs(p, text, size, color, bold)
        return Box(x, y, w, h, tb)

    def label(self, x, y, w, h, text, slot, i=None, size='body', color='text', bold=False, anchor='top', align='left',
              shadow=False):
        """单行标签（轴刻度、行名、运算符）：一个文本框一段文字。h 不够时记警告。shadow=True 加柔和投影，只用于照片上的字。"""
        pt = self.style.size(size)
        self._check_fit(y, h, text_height(text, w, pt, self.style.metrics['text']['lineSpacing'], self.family, bold), pt, slot)
        box = self._label(x, y, w, h, text, slot, i, size, color, bold, anchor, align)
        for r in box.shape.text_frame.paragraphs[0].runs if shadow else ():     # rPr 里 effectLst 排在填充之后、字体之前
            r._r.get_or_add_rPr().find(qn('a:solidFill')).addnext(parse_xml(
                f'<a:effectLst {nsdecls("a")}><a:outerShdw blurRad="76200" dist="19050" dir="2700000" algn="tl" rotWithShape="0">'
                '<a:srgbClr val="000000"><a:alpha val="70000"/></a:srgbClr></a:outerShdw></a:effectLst>'))
        return box

    def conclusion(self, text, x=None, y=None, w=None, h=None, color='emphasis'):
        """红字结论：一行带事实的判断（决策 15），不是副标题。字号级别 type.conclusion；默认放在 style.json 的 conclusion 位置，
        去 AI 味版换位置时给 x、y、w、h。放不下时记警告：少说一件事，不缩字号。返回 Box。"""
        default = self.style.doc.get('conclusion')
        if 'conclusion' not in self.style.type or (default is None and None in (x, y, w, h)):
            raise StyleError('红字结论要在 style.json 里写字号级别 type.conclusion（19.05 cm 页高约 20—24 pt，其他尺寸按页高等比），'
                             '和默认位置 conclusion: {xCm, yCm, wCm, hCm}（标题线下、版心上沿之前）；只换一页的位置时给 x、y、w、h')
        if self.lead and y is None:
            self.deck.warn(f'第 {self.number} 页：副标题和红字结论都在标题线下，默认位置可能重叠；去掉副标题或给红字结论另定位置')
        d = default or {}
        x, y = d.get('xCm') if x is None else x, d.get('yCm') if y is None else y
        w, h = d.get('wCm') if w is None else w, d.get('hCm') if h is None else h
        return self.label(x, y, w, h, text, 'conclusion', size='conclusion', color=color, bold=True)

    def _label(self, x, y, w, h, text, slot, i=None, size='body', color='text', bold=False, anchor='top', align='left'):
        tb = self.name(self.slide.shapes.add_textbox(Cm(x), Cm(y), Cm(w), Cm(h)), slot, i)
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Cm(self.style.metrics['text']['marginCm'])
        tf.vertical_anchor = _pick(ANCHOR, anchor, 'anchor')
        p = tf.paragraphs[0]
        p.alignment = _pick(ALIGN, align, 'align')
        p.line_spacing = self.style.metrics['text']['lineSpacing']
        self.runs(p, text, size, color, bold)
        return Box(x, y, w, h, tb)

    # ---------- 页眉 ----------
    def header(self, title=None, parent=None, lead=None):
        """按 style.json 的 header 画：色条，再按 lines 的顺序画本页标题、上级标题、副标题。全册一样，不手写。
        不给就用 page() 时给的；什么字都没有时不画页眉线。"""
        hd = self.style.doc['header']
        title, parent, lead = title or self.title, parent or self.parent, lead or self.lead
        values = {'title': title, 'parent': parent, 'lead': lead, 'chapter': self.chapter}
        roles = {line['role'] for line in hd.get('lines', [])}
        if lead and 'lead' not in roles:
            self.deck.warn(f'第 {self.number} 页：style.json 的页眉没有副标题行，lead 不会画；副标题默认不写，要写就在 header.lines 加 lead 行')
        for bar in hd.get('bars', []):
            if 'role' in bar and not values.get(bar['role']):      # 跟着某一行标题的色条，那行没字就不画
                continue
            w = self.style.page_w if bar['wCm'] == 'page' else bar['wCm']
            self.shape('rect', bar['xCm'], bar['yCm'], w, bar['hCm'], bar['color'], slot=bar.get('slot', 'header-bar'))
        slots = {'title': 'title', 'parent': 'parent-title', 'lead': 'subtitle', 'chapter': 'chapter'}
        for line in hd.get('lines', []):
            text = values[line['role']]
            if not text:
                continue
            self.label(line['xCm'], line['yCm'], line['wCm'], line['hCm'], text, slots[line['role']],
                       size=line['size'], color=line['color'], bold=line.get('bold', False))
        rule = hd.get('rule')
        if rule and any(values.values()):
            self.line(rule['xCm'], rule['yCm'], rule['xCm'] + rule['wCm'], rule['yCm'], rule['color'],
                      rule.get('widthPt', 0.75), slot='header-rule')
        if hd.get('pageNo'):
            self.page_number(**hd['pageNo'])

    def page_number(self, xCm, yCm, wCm, hCm, size='body', color='title', bold=True, align='right'):
        """页码：自动字段，移页后跟着变（style.header.pageNo）。"""
        tb = self._label(xCm, yCm, wCm, hCm, '', 'page-no', size=size, color=color, align=align, anchor='middle')
        p = tb.shape.text_frame.paragraphs[0]
        fonts = self.style.fonts
        fld = parse_xml(
            f'<a:fld {nsdecls("a")} id="{{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}}" type="slidenum">'
            f'<a:rPr lang="zh-CN" sz="{int(self.style.size(size) * 100)}" b="{1 if bold else 0}">'
            f'<a:solidFill><a:srgbClr val="{self.style.color(color).lstrip("#")}"/></a:solidFill>'
            f'<a:latin typeface="{fonts.get("latin", fonts["cjk"])}"/><a:ea typeface="{fonts["cjk"]}"/></a:rPr>'
            f'<a:t>{self.number}</a:t></a:fld>')
        end = p._p.find(qn('a:endParaRPr'))
        if end is not None:
            end.addprevious(fld)
        else:
            p._p.append(fld)
        return tb

    # ---------- 形状 ----------
    def shape(self, kind, x, y, w, h, fill, slot, i=None, line=None, alpha=None):
        """实色填充、无边框的形状（rect / oval / diamond）。line 给颜色名时只画线框、不填充（框只用其一）。"""
        s = self.name(self.slide.shapes.add_shape(_pick(SHAPES, kind, 'shape', MSO_SHAPE), Cm(x), Cm(y), Cm(w), Cm(h)), slot, i)
        # python-pptx 给自动形状挂了主题样式（effectRef idx=2），PowerPoint 会按主题画阴影（LibreOffice 不画）；
        # 线和填充下面都显式写了，去掉样式引用
        style = s._element.find(qn('p:style'))
        if style is not None:
            s._element.remove(style)
        if line:
            s.fill.background()
            s.line.color.rgb = self.rgb(line)
            s.line.width = Pt(self.style.metrics['strokes']['framePt'])
        else:
            s.fill.solid()
            s.fill.fore_color.rgb = self.rgb(fill)
            s.line.fill.background()
            if alpha is not None and alpha < 1:
                clr = s.fill._xPr.find(qn('a:solidFill'))[0]
                clr.append(clr.makeelement(qn('a:alpha'), {'val': str(int(alpha * 100000))}))
        return s

    def band(self, x, y, w, h, color, slot='band', alpha=0.45):
        """照片上的字下面垫的渐隐暗带：上下全透明，中间最深（不透明度 alpha），见 page-types.md“章小结页”。"""
        s = self.shape('rect', x, y, w, h, color, slot)
        hexv = str(self.rgb(color))
        stops = ''.join(f'<a:gs pos="{pos}"><a:srgbClr val="{hexv}"><a:alpha val="{int(a * 100000)}"/></a:srgbClr></a:gs>'
                        for pos, a in ((0, 0), (50000, alpha), (100000, 0)))
        s.fill._xPr.find(qn('a:solidFill')).getparent().replace(s.fill._xPr.find(qn('a:solidFill')), parse_xml(
            f'<a:gradFill {nsdecls("a")} rotWithShape="1"><a:gsLst>{stops}</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>'))
        return s

    def line(self, x0, y0, x1, y1, color, width_pt, slot='line', i=None, arrow=False):
        c = self.name(self.slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(x0), Cm(y0), Cm(x1), Cm(y1)), slot, i)
        c.line.color.rgb = self.rgb(color)
        c.line.width = Pt(width_pt)
        style = c._element.find(qn('p:style'))      # 同 shape：主题样式会带阴影，线的颜色和宽度都显式写了
        if style is not None:
            c._element.remove(style)
        if arrow:
            ln = c.line._get_or_add_ln()
            ln.append(ln.makeelement(qn('a:tailEnd'), {'type': 'triangle', 'w': 'med', 'len': 'med'}))
        return c

    def arrow(self, x0, y0, x1, y1, slot='arrow', i=None):
        """图示里的箭头，全册一个样式（style.strokes）。"""
        st = self.style.metrics['strokes']
        return self.line(x0, y0, x1, y1, st['arrowColor'], st['arrowPt'], slot, i, arrow=True)

    def node(self, x, y, w, h, text, color, slot='node', i=None, text_color=None, size='body'):
        """图示节点：实色或浅色填充的直角框，框里写短语；多行时第一行加粗。"""
        s = self.shape('rect', x, y, w, h, color, slot, i)
        tf = s.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = tf.margin_right = Cm(self.style.metrics['node']['paddingCm'])
        for k, line in enumerate(text.split('\n')):
            p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
            p.alignment = PP_ALIGN.CENTER
            self.runs(p, line, size, text_color or self.style.metrics['node']['text'], bold=(k == 0))
        return s

    def ledger(self, x, y, items, w, h, slot='ledger', fmt='{:.1f}', number='figure', op='figure'):
        """账算式：items 为 (数字, 说明, 颜色) 或运算符字符串；同一个等式里方块同尺寸。返回 Box。"""
        lg = self.style.metrics['ledger']
        cx, k, j = x, self.next_index('ledger-block'), self.next_index('ledger-op')
        for it in items:
            if isinstance(it, str):
                ow = lg['opWidthCm']
                # 运算符是固定小框里的一个字符，字号由 op 定，框高不用核对
                self._label(cx, y + h / 2 - 0.6, ow, 1.2, it, 'ledger-op', i=j, size=op, color=lg['opColor'], bold=True,
                            align='center')
                cx += ow
                j += 1
            else:
                val, label, color = it
                s = self.node(cx, y, w, h, f'{fmt.format(val)}\n{label}', color, 'ledger-block', k)
                s.text_frame.paragraphs[0].runs[0].font.size = self._size(number)
                cx += w
                k += 1
        return Box(x, y, cx - x, h)

    # ---------- 表格 ----------
    def table(self, x, y, w, rows, widths, slot='table', row_h=None, total=False, highlight=(), fmt='{:.1f}'):
        """表头深色填充；数字右对齐；total=True 时最后一行是合计行（加粗、浅底）；highlight 为要提示的行号。
        行高按单元格换行估，返回 Box（h 为估算总高）。"""
        tb = self.style.metrics['table']
        row_h = tb['rowHeightCm'] if row_h is None else row_h
        size = self.style.size('body')
        if len(widths) != len(rows[0]) or any(len(r) != len(widths) for r in rows):
            raise ValueError(f'table：{len(widths)} 个列宽，各行的格数是 {sorted({len(r) for r in rows})}；要一样多')
        cw = [w * wd / sum(widths) for wd in widths]
        show = lambda v: fmt.format(v) if isinstance(v, float) else ('' if v is None else str(v))
        hs = [max(row_h, max(text_height(show(v), cw[j] - 0.2, size, 1.15, self.family, i == 0) + 0.06 for j, v in enumerate(r)))
              for i, r in enumerate(rows)]
        self._check_fit(y, sum(hs), sum(hs), size, slot)
        gf = self.name(self.slide.shapes.add_table(len(rows), len(rows[0]), Cm(x), Cm(y), Cm(w), Cm(sum(hs))), slot)
        t = gf.table
        for j, wd in enumerate(cw):
            t.columns[j].width = Cm(wd)
        last = len(rows) - 1
        for i, r in enumerate(rows):
            t.rows[i].height = Cm(hs[i])
            for j, v in enumerate(r):
                cell = t.cell(i, j)
                cell.margin_left = cell.margin_right = Cm(tb['paddingCm'])
                cell.margin_top = cell.margin_bottom = Cm(0.03)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                cell.text_frame.word_wrap = True
                p = cell.text_frame.paragraphs[0]
                is_total = total and i == last
                self.runs(p, show(v), 'body', tb['headerText'] if i == 0 else tb['text'], bold=(i == 0) or is_total)
                p.alignment = PP_ALIGN.RIGHT if (i > 0 and isinstance(v, (int, float))) else PP_ALIGN.LEFT
                cell.fill.solid()
                fill = tb['headerFill'] if i == 0 else tb['highlightFill'] if i in highlight else tb['totalFill'] if is_total else tb['fill']
                cell.fill.fore_color.rgb = self.rgb(fill)
                self._cell_lines(cell, top=i > 0)
        tbl = t._tbl
        tbl.tblPr.set('bandRow', '0')
        tbl.tblPr.set('firstRow', '1')
        return Box(x, y, w, sum(hs), gf)

    def _cell_lines(self, cell, top):
        """行间细线，不画全框网格（visual-language.md）。不写时表格样式的白框线会露出来。a:ln* 要排在填充前面。"""
        tb = self.style.metrics['table']
        tc = cell._tc.get_or_add_tcPr()
        for k, (side, on) in enumerate((('a:lnL', False), ('a:lnR', False), ('a:lnT', top), ('a:lnB', True))):
            ln = tc.makeelement(qn(side), {'w': str(Pt(tb['linePt']) if on else 0)})
            fill = ln.makeelement(qn('a:solidFill' if on else 'a:noFill'), {})
            if on:
                fill.append(fill.makeelement(qn('a:srgbClr'), {'val': str(self.rgb(tb['lineColor']))}))
            ln.append(fill)
            tc.insert(k, ln)

    # ---------- 图表 ----------
    def chart(self, kind, x, y, w, h, categories, series, colors, slot='chart', legend=True, labels=False,
              number_format='0.0', gap=60, overlap=None):
        """原生图表（可编辑）。kind：bar / bar-stacked / column / column-stacked / line / area / doughnut / pie，
        或 python-pptx 的 XL_CHART_TYPE 成员（散点、气泡等要另一种数据，不支持）。
        colors 写颜色名或类别配色，与 series 一一对应；doughnut / pie 只有一个系列，colors 与 categories 一一对应。"""
        ctype = _pick(CHARTS, kind, 'chart', XL_CHART_TYPE)
        if not ctype.name.startswith(('BAR_', 'COLUMN_', 'LINE', 'AREA', 'RADAR', 'DOUGHNUT', 'PIE')):
            raise ValueError(f'chart {ctype.name} 不支持；可用：{"、".join(CHARTS)}，或这几种同类的 XL_CHART_TYPE 成员')
        round_ = ctype.name.startswith(('DOUGHNUT', 'PIE'))     # 只有一个系列，每个类别一种颜色，没有坐标轴
        want = categories if round_ else series
        if round_ and len(series) != 1:
            raise ValueError(f'chart {ctype.name}：只能有 1 个系列，给了 {len(series)} 个')
        if len(colors) != len(want):
            raise ValueError(f'chart：{len(want)} 个{"类别" if round_ else "系列"}给了 {len(colors)} 个颜色')
        for c in colors:        # 放进页面之前查完，不留半成品
            self.rgb(c)
        data = CategoryChartData()
        data.categories = categories
        for nm, vals in series:
            data.add_series(nm, vals)
        gf = self.name(self.slide.shapes.add_chart(ctype, Cm(x), Cm(y), Cm(w), Cm(h), data), slot)
        ch = gf.chart
        self.chart_font(ch.font)
        plot = ch.plots[0]
        if not isinstance(plot, BarPlot):       # 单系列时 PowerPoint 会拿系列名自动出标题；柱形、条形的默认外观先不动
            ch.has_title = False
        ch.has_legend = legend and len(want) > 1
        if ch.has_legend:
            ch.legend.position = XL_LEGEND_POSITION.TOP
            ch.legend.include_in_layout = False
            self.chart_font(ch.legend.font)
        if isinstance(plot, BarPlot):
            plot.gap_width = gap
            if overlap is not None:
                plot.overlap = overlap
        if round_:
            for k, c in enumerate(colors):
                self.color_point(ch, 0, k, c)
        for s, c in zip(plot.series, colors if not round_ else ()):
            if isinstance(plot, (LinePlot, RadarPlot)):     # 折线只有线色，写 fill 不生效
                s.format.line.color.rgb = self.rgb(c)
                s.format.line.width = Pt(self.style.metrics['chart']['linePt'])
            else:
                s.format.fill.solid()
                s.format.fill.fore_color.rgb = self.rgb(c)
        if labels:
            plot.has_data_labels = True
            dl = plot.data_labels
            dl.show_value = True        # 环形图默认的标签是关着的
            dl.number_format, dl.number_format_is_linked = number_format, False
            self.chart_font(dl.font)
        if round_:
            return ch
        st = self.style.metrics['strokes']
        va = ch.value_axis
        va.has_major_gridlines = True
        va.major_gridlines.format.line.color.rgb = self.rgb(st['gridColor'])
        self.chart_font(va.tick_labels.font)
        va.format.line.fill.background()
        self.chart_font(ch.category_axis.tick_labels.font)
        ch.category_axis.format.line.color.rgb = self.rgb(st['axisColor'])
        return ch

    def color_point(self, chart, series, point, color):
        """给图表的一个数据点单独上色（颜色名或类别配色）。"""
        fmt = chart.plots[0].series[series].points[point].format
        fmt.fill.solid()
        fmt.fill.fore_color.rgb = self.rgb(color)

    def label_series(self, chart, series, color='text', number_format='0.0'):
        dl = chart.plots[0].series[series].data_labels
        dl.show_value = True
        dl.number_format, dl.number_format_is_linked = number_format, False
        self.chart_font(dl.font)
        dl.font.color.rgb = self.rgb(color)

    # ---------- 图片与地图 ----------
    def picture(self, path, x, y, w, h, slot='picture', fit='contain'):
        """照片、截图：按比例撑满槽位、居中（fit='contain'，可能留白）。fit='cover' 铺满槽位，
        多出的部分裁掉（原图还在，PowerPoint 里可调），用于满版图。地图用 map()。"""
        with Image.open(path) as im:
            iw, ih = im.size
        s = _pick({'contain': min, 'cover': max}, fit, 'fit')(w / iw, h / ih)
        pw, ph = iw * s, ih * s
        if fit == 'cover':
            pic = self.name(self.slide.shapes.add_picture(path, Cm(x), Cm(y), Cm(w), Cm(h)), slot)
            pic.crop_left = pic.crop_right = (1 - w / pw) / 2
            pic.crop_top = pic.crop_bottom = (1 - h / ph) / 2
            return Box(x, y, w, h, pic)
        px, py = x + (w - pw) / 2, y + (h - ph) / 2
        pic = self.name(self.slide.shapes.add_picture(path, Cm(px), Cm(py), Cm(pw), Cm(ph)), slot)
        return Box(px, py, pw, ph, pic)

    def map(self, asset, x, y, w, h, slot='map', align='center', shrink=False):
        """把发布版里的一张图放进槽位 (x, y, w, h)。
        - legend.json 有 print 时按原尺寸放（图上的字就是主题定的 pt）；先核对 PNG 像素与 print 一致。
        - 槽位比 print 小：默认报错（改地图清单的 layout.widthCm / heightCm，或换大的槽位）；
          shrink=True 时缩小并记警告，图上的字会小于 style.mapText 的下限。
        - 没有 print：按比例撑满槽位，记警告（图上字号不可控）。
        返回实际位置 Box。"""
        if not isinstance(asset, MapAsset):
            raise MapError('map() 只放发布版里的图：用 MapRelease(<composer 项目>, <版本>).map(<图 id>) 取，照片用 picture()')
        where = f'第 {self.number} 页 {self.type}:{slot}'
        if asset.qa is not None and asset.qa.status != 'ok':
            self.deck.warn(f'{where}：{asset.qa.message(asset.id, asset.release)}')
        theme = self.style.theme_file
        if theme and asset.theme_sha256 and theme[0] in asset.theme_sha256 and asset.theme_sha256[theme[0]] != theme[1]:
            self.deck.warn(f'发布版 {asset.release} 出图时的地图主题 {theme[0]} 与现在的不同：页面颜色已按现在的主题，'
                           f'图面还是出图时的颜色；重新出图并发布新版本', once=True)
        # 地图用字：没写 fonts.map 就是地图用页面字体
        font_key = 'map' if self.style.fonts.get('map') else 'cjk'
        font_msg = asset.font_mismatch(self.style.fonts.get(font_key), f'fonts.{font_key}')
        if font_msg:
            self.deck.warn(f'发布版 {asset.release}：{font_msg}', once=True)
        pr = asset.check_print()
        iw, ih = asset.pixels
        record = {'page': self.number, 'slot': f'pd:{self.type}:{slot}', 'map': asset.id, 'release': asset.release,
                  'slotCm': [round(w, 2), round(h, 2)]}
        if pr:
            pw, ph, dpi = pr
            k = 1.0
            if pw > w + 0.05 or ph > h + 0.05:
                k = min(w / pw, h / ph)
                minor = self.style.map_text.get('minorPt')
                msg = (f'{where}：{asset.id} 按厘米出图 {pw:g}×{ph:g} cm，槽位只有 {w:.2f}×{h:.2f} cm，要缩到 {k:.0%}'
                       + (f'，图上最小注记约 {minor * k:.1f} pt，低于 mapText.minorPt {minor:g} pt' if minor else ''))
                if not shrink:
                    raise MapError(msg + '。改地图清单的 layout.widthCm / heightCm 后重新出图发布，或换大的槽位；确要缩小写 shrink=True')
                self.deck.warn(msg)
                pw, ph = pw * k, ph * k
            record.update(printCm=[pr[0], pr[1]], dpi=pr[2], scale=round(k, 3))
        else:
            s = min(w / iw, h / ih)
            pw, ph = iw * s, ih * s
            dpi = iw / (pw / 2.54)
            self.deck.warn(f'{where}：{asset.id} 没按厘米出图（legend.json 无 print），按比例撑满槽位，'
                           f'约 {dpi:.0f} dpi，图上字号不可控')
            record.update(printCm=None, dpi=round(dpi), scale=None)
        ax = _pick({'left': 0, 'center': 0.5, 'right': 1}, align, 'align')
        px, py = x + (w - pw) * ax, y + (h - ph) / 2
        pic = self.name(self.slide.shapes.add_picture(asset.image, Cm(px), Cm(py), Cm(pw), Cm(ph)), slot)
        record['placedCm'] = [round(px, 2), round(py, 2), round(pw, 2), round(ph, 2)]
        record['effectiveDpi'] = round(iw / (pw / 2.54))
        self.deck.placements.append(record)
        return Box(px, py, pw, ph, pic)

    def swatch(self, x, y, color, slot, i, symbol='fill', alpha=1.0):
        size = self.style.metrics['legend']['swatchCm']
        if isinstance(symbol, LibrarySymbol):
            return self.library_symbol(x, y, symbol, slot, i)
        if symbol == 'frame':
            s = self.shape('rect', x, y, size, size, None, slot, i, line=color)
            s.line.dash_style = 4
            return s
        s = self.shape('oval' if symbol == 'circle' else 'rect', x, y, size, size, color, slot, i, alpha=alpha)
        if alpha < 1:    # 半透明的色块在白底上很浅，加一圈细线看得见边
            s.line.color.rgb = self.rgb('line')
            s.line.width = Pt(0.5)
        return s

    def library_symbol(self, x, y, symbol, slot, i):
        """符号库的符号画成一组可编辑的自由曲线（legend.json 的 svg）：点符号占色块大小，
        线、面样本按原比例放进色块高、标注左边的空当里。返回组合形状。"""
        lg = self.style.metrics['legend']
        (bx, by, bw, bh), parts = parse_svg(symbol.svg)
        k = min((lg['labelOffsetCm'] - 0.08) / bw, lg['swatchCm'] / bh) if bw > bh * 1.2 else lg['swatchCm'] / max(bw, bh)
        emu = Cm(k)                                  # 每个 SVG 单位多少 EMU
        oy = Cm(y + (lg['swatchCm'] - bh * k) / 2)
        group = self.name(self.slide.shapes.add_group_shape(), slot, i)
        for part in parts:
            if not part.contours:
                continue
            (pts, closed), rest = part.contours[0], part.contours[1:]
            fb = group.shapes.build_freeform((pts[0][0] - bx) * emu, (pts[0][1] - by) * emu, scale=1.0)
            fb.add_line_segments([((px - bx) * emu, (py - by) * emu) for px, py in pts[1:]], close=closed)
            for pts, closed in rest:
                fb.move_to((pts[0][0] - bx) * emu, (pts[0][1] - by) * emu)
                fb.add_line_segments([((px - bx) * emu, (py - by) * emu) for px, py in pts[1:]], close=closed)
            s = fb.convert_to_shape(Cm(x), oy)
            style = s._element.find(qn('p:style'))
            if style is not None:
                s._element.remove(style)
            if part.hatch:
                s.fill.patterned()
                s.fill.pattern = MSO_PATTERN.WIDE_DOWNWARD_DIAGONAL if part.hatch['angle'] % 180 > 90 else MSO_PATTERN.WIDE_UPWARD_DIAGONAL
                s.fill.fore_color.rgb = RGBColor.from_string(part.hatch['color'][1:])
                s.fill.back_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            elif part.fill:
                s.fill.solid()
                s.fill.fore_color.rgb = RGBColor.from_string(part.fill[1:])
                _alpha(s.fill._xPr.find(qn('a:solidFill')), part.fill_opacity)
            else:
                s.fill.background()
            if part.stroke:
                s.line.color.rgb = RGBColor.from_string(part.stroke[1:])
                s.line.width = int(part.width * emu)
                ln = s.line._get_or_add_ln()
                _alpha(ln.find(qn('a:solidFill')), part.stroke_opacity)
                if part.dash:     # 虚线按线宽的百分比写，和 SVG 的长度一致
                    ds = ''.join(f'<a:ds d="{int(d / part.width * 100000)}" sp="{int(g / part.width * 100000)}"/>'
                                 for d, g in zip(part.dash[::2], part.dash[1::2] or part.dash[::2]))
                    ln.append(parse_xml(f'<a:custDash {nsdecls("a")}>{ds}</a:custDash>'))
            else:
                s.line.fill.background()
        group._element.recalculate_extents()        # 自由曲线加进组合后组合的外框不会自己更新
        return group

    def _legend_docs(self, source):
        """图例的来源（MapAsset、legend 字典或它们的列表）→ (legend 文档列表, 名称)；
        图例已画进 PNG、页面与地图类别色不一致时记警告。"""
        docs, names = [], []
        for src in (source if isinstance(source, (list, tuple)) else [source]):
            if isinstance(src, MapAsset):
                if src.legend is None:
                    raise MapError(f'{src.id} 没有 legend.json：重新出图时让 gis-composer 写图例文档（map-legend）')
                docs.append(src.legend)
                names.append(src.id)
            else:
                docs.append(src)
                names.append('legend')
        name = '、'.join(names)
        if any(d.get('drawn') for d in docs):
            self.deck.warn(f'第 {self.number} 页：{name} 的图例已画进 PNG（drawn: true），页面再画会重复；'
                           f'地图清单写 legend: {{visible: false}}')
        for scheme, code, page_c, map_c in class_mismatches(docs, self.style.palettes):
            self.deck.warn(f'第 {self.number} 页：{name} 里“{code}”是 {map_c}，style.json 的 classes.{scheme} 是 {page_c}，'
                           f'页面与地图同一类不同色')
        return docs, name

    def legend(self, source, x, y, w, extra=(), short=True, context=True):
        """按 legend.json 画可编辑图例：分类行出标题加各类，普通行（范围线、示意面）排在后面；
        extra 为补充行 (颜色名, 文字, 符号)，只用于地图图例文档里确实没有的行。
        source 是 MapAsset、已读入的 legend 字典，或它们的列表（如放大图面板 + 整页图例）。返回 Box。
        页面类别配色与地图里同一 code 的颜色不一致时记警告：两边要从同一处取色。"""
        docs, name = self._legend_docs(source)
        rows, skipped = legend_rows(docs, short=short, context=context)
        if skipped:
            self.deck.warn(f'第 {self.number} 页：{name} 有图例行没画（不支持的种类）：{"、".join(skipped)}')
        rows += [('row', self.style.color(c), text, symbol, 1.0) for c, text, symbol in extra]
        lg = self.style.metrics['legend']
        off = lg['labelOffsetCm']
        cy, k = y, max(self.next_index('legend-swatch'), self.next_index('legend-label'))
        for r in rows:
            if r[0] == 'title':
                if r[1]:
                    th = text_height(self.style.metrics['text']['headingMark'] + r[1], w, self.style.size('heading'), 1.15,
                                     self.family, True) + 0.1
                    self.text(x, cy, w, [('h', r[1])], h=th, slot='legend-title')
                    cy += th
                continue
            _, color, text, symbol, alpha = r
            rh = max(lg['rowMinCm'], text_height(text, w - off, self.style.size('body'), 1.15, self.family))
            self.swatch(x, cy + 0.08, color if isinstance(symbol, LibrarySymbol) else DataColor(color.upper()), 'legend-swatch', k, symbol, alpha)
            self.text(x + off, cy, w - off, [text], h=rh, slot='legend-label', i=k)
            cy += rh
            k += 1
        return Box(x, y, w, cy - y)

    # ---------- 图上的注记和图例（鲜艳实用汇报，见 docs/map-types.md） ----------
    def _boxed(self, x, y, w, h, slot, i=None, border='#000000', border_pt=0.75):
        """白底、细边框、直角的框（不带阴影）。"""
        s = self.shape('rect', x, y, w, h, DataColor('#FFFFFF'), slot, i)
        s.line.color.rgb = RGBColor.from_string(border.lstrip('#'))
        s.line.width = Pt(border_pt)
        return s

    def map_label(self, map_box, text, at, box_at, r=0.0, slot='map-label', i=None, size=None, bold=True):
        """地图注记方框：白底黑框直角方框里写名称（纯黑，默认加粗），细黑引线从方框最近一边的中点指到点的边缘。
        map_box 是 map() 返回的 Box；at 是点在图上的位置，box_at 是方框左上角，都是相对地图左上角的 cm；
        r 是点的半径（cm），引线停在点的边缘，点就在方框旁边（距离不到 r 加 0.1 cm）时不画引线。
        方框不出地图的边。字号默认 style.mapText.labelPt。返回方框的 Box。"""
        pt = size or self.style.map_text['labelPt']
        pad = 0.09
        w = text_width(text, pt, self.family, bold) + 2 * pad + 0.1
        h = pt * 1.35 * PT_CM + 2 * pad
        bx = min(max(map_box.x + box_at[0], map_box.x), map_box.right - w)
        by = min(max(map_box.y + box_at[1], map_box.y), map_box.bottom - h)
        cx, cy = map_box.x + at[0], map_box.y + at[1]
        if bx <= cx <= bx + w:                       # 点在方框正上方或正下方：引线竖着接上下边
            ex, ey = cx, (by if cy < by else by + h)
        else:
            ex, ey = (bx if cx < bx else bx + w), by + h / 2
        d = math.hypot(ex - cx, ey - cy)
        k = self.next_index(slot) if i is None else i
        if d > r + 0.1:
            sx, sy = cx + (ex - cx) * r / d, cy + (ey - cy) * r / d
            self.line(sx, sy, ex, ey, DataColor('#000000'), 0.75, f'{slot}-leader', k)
        s = self._boxed(bx, by, w, h, slot, k)
        tf = s.text_frame
        tf.word_wrap = False
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = tf.margin_right = Cm(pad)
        tf.margin_top = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        self.runs(p, text, pt, DataColor('#000000'), bold)
        return Box(bx, by, w, h, s)

    def inset_legend(self, source, map_box, corner='bottom-left', title='图例', extra=(), short=True, context=False,
                     sizes=False, margin=0.15, max_rows=6):
        """图内一列图例：放在地图的一个角上（corner：bottom-left / top-left / bottom-right / top-right），
        白底细边框，顶上一行“图例”（title=None 不写），不出分组标题；字号 style.mapText.labelPt，字纯黑。
        - 点类画同色圆点，面（含只作背景的范围）画同色小方块，线画一段线（legend.json 有 dasharray 就是虚线），
          符号库的符号照 svg 画；尺寸行默认不画（sizes=True 画）；context 行默认不画。
        - extra：补充行 (颜色名, 文字, 符号)，同 legend()；符号可写 'line' / 'dash' 画一段实线或虚线。
        - 行数超过 max_rows 时记警告（一张图别叠太多类，能拆就拆）；点类用了黄色记警告（白底上看不清）。
        source 同 legend()。返回 Box。"""
        docs, name = self._legend_docs(source)
        rows, skipped = legend_rows(docs, short=short, context=context, lines=True, sizes=sizes)
        if skipped:
            self.deck.warn(f'第 {self.number} 页：{name} 有图例行没画（不支持的种类）：{"、".join(skipped)}')
        rows = [r for r in rows if r[0] == 'row']
        for c, text, symbol in extra:
            if symbol in ('line', 'dash'):
                symbol = LineSample([3, 1.5] if symbol == 'dash' else None, 1.5)
            rows.append(('row', self.style.color(c), text, symbol, 1.0))
        if len(rows) > max_rows:
            self.deck.warn(f'第 {self.number} 页：{name} 的图内图例有 {len(rows)} 行：一张图叠的类太多，能拆成几张就拆')
        for _, color, text, symbol, _a in rows:
            if symbol == 'circle' and _yellow(color):
                self.deck.warn(f'第 {self.number} 页：{name} 的点类“{text}”是黄色 {color}：白底上看不清，黄色只给面用')
        pt = self.style.map_text['labelPt']
        row_h, pad, sym_w, gap = pt * 1.4 * PT_CM, 0.22, 0.56, 0.15
        texts = [text for _, _, text, _, _ in rows]
        tm = 2 * self.style.metrics['text']['marginCm']       # 文本框左右内边距
        w = pad * 2 + sym_w + gap + tm + max([text_width(t, pt, self.family) for t in texts] +
                                         [text_width(title, pt, self.family, True) - sym_w - gap if title else 0]) + 0.15
        h = pad * 2 + row_h * (len(rows) + (1 if title else 0))
        right, bottom = corner.endswith('right'), corner.startswith('bottom')
        if corner not in ('bottom-left', 'top-left', 'bottom-right', 'top-right'):
            raise ValueError(f'corner 只能是 bottom-left / top-left / bottom-right / top-right，不是 {corner!r}')
        x0 = map_box.right - margin - w if right else map_box.x + margin
        y0 = map_box.bottom - margin - h if bottom else map_box.y + margin
        self._boxed(x0, y0, w, h, 'map-legend-bg', border='#404040', border_pt=0.5)
        black = DataColor('#000000')
        y = y0 + pad
        if title:
            self._label(x0 + pad - tm / 2, y, w - 2 * pad + tm, row_h, title, 'map-legend-title', size=pt, color=black, bold=True,
                        anchor='middle').shape.text_frame.word_wrap = False      # 单行，框宽按字宽定
            y += row_h
        k = self.next_index('map-legend-swatch')
        for _, color, text, symbol, alpha in rows:
            cy, x = y + row_h / 2, x0 + pad
            dc = color if isinstance(symbol, LibrarySymbol) else DataColor(color.upper())
            if isinstance(symbol, LineSample):
                ln = self.line(x, cy, x + sym_w, cy, dc, min(max(symbol.width_pt, 1.0), 4.0), 'map-legend-swatch', k)
                if symbol.dashed:
                    ln.line.dash_style = MSO_LINE_DASH_STYLE.DASH
            elif symbol == 'circle':
                d = pt * 0.8 * PT_CM
                self.shape('oval', x + (sym_w - d) / 2, cy - d / 2, d, d, dc, 'map-legend-swatch', k, alpha=alpha)
            elif symbol == 'fill':
                d = pt * 0.9 * PT_CM
                self.shape('rect', x + (sym_w - d) / 2, cy - d / 2, d, d, dc, 'map-legend-swatch', k, alpha=alpha)
            else:                                    # 虚线框、符号库的符号
                sw = self.style.metrics['legend']['swatchCm']
                self.swatch(x + (sym_w - sw) / 2, cy - sw / 2, dc, 'map-legend-swatch', k, symbol, alpha)
            self._label(x + sym_w + gap - tm / 2, y, w - 2 * pad - sym_w - gap + tm / 2 + 0.05, row_h, text, 'map-legend-label', k,
                        size=pt, color=black, anchor='middle').shape.text_frame.word_wrap = False
            y += row_h
            k += 1
        return Box(x0, y0, w, h)

    def notes(self, text):
        """讲稿备注：限定、说明性的话放这里，不写在页面上。"""
        self.slide.notes_slide.notes_text_frame.text = text


def _yellow(color):
    """色值是不是黄色（点上不用黄）：色相约 45°–70°、饱和、偏亮。"""
    try:
        r, g, b = (int(str(color).lstrip('#')[k:k + 2], 16) / 255 for k in (0, 2, 4))
    except ValueError:
        return False
    import colorsys
    hue, light, sat = colorsys.rgb_to_hls(r, g, b)
    return 45 <= hue * 360 <= 70 and sat > 0.5 and light > 0.4


def _alpha(fill, opacity):
    """给 a:solidFill 里的颜色加不透明度（1 不加）。"""
    if fill is not None and opacity < 1:
        clr = fill[0]
        clr.append(clr.makeelement(qn('a:alpha'), {'val': str(int(opacity * 100000))}))
