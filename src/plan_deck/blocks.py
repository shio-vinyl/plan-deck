"""正文组件：页面中间层（整块实色、大字、粗箭头）和常用图示，做法和使用场合见 docs/patterns.md。

共同的约定：
- 位置、尺寸用厘米；颜色给颜色名、类别配色 (方案, 类别)，或一个类别配色的方案名（逐项轮流取该方案的颜色）；
  字号给 type 里的级别名（key、figure 没写时退到 heading）。
- 高度按内容定。给 bottom（默认版心底）的组件把多余高度均分下去，铺到 bottom；bottom=False 不铺。
  内容放不下时记警告（deck.warnings），不缩字号。
- 条目可以在最后多给一个颜色，覆盖 colors。文字里 \\n 分段，** __ ^^ 照 Page.runs。
- 实底上的字按底色深浅自动用白字或正文色。
"""
from PIL import Image
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Cm, Pt

from .measure import LINE, PT_CM, text_lines, text_width

_ANCHOR = {'top': MSO_ANCHOR.TOP, 'middle': MSO_ANCHOR.MIDDLE, 'bottom': MSO_ANCHOR.BOTTOM}
_ALIGN = {'left': PP_ALIGN.LEFT, 'center': PP_ALIGN.CENTER, 'right': PP_ALIGN.RIGHT}
ARROWS = {'down': MSO_SHAPE.DOWN_ARROW, 'right': MSO_SHAPE.RIGHT_ARROW, 'left': MSO_SHAPE.LEFT_ARROW, 'up': MSO_SHAPE.UP_ARROW}
NODE_KINDS = ('solid', 'plain', 'emphasis', 'light')
# 矩阵表按版心选字号：正文的这几倍里取放得下的最大值
TABLE_STEPS = (1.25, 1.17, 1.08, 1.0)


class Blocks:
    """Page 的正文组件（混入 Page）。"""

    # ---------- 内部 ----------
    @property
    def _bm(self):
        return self.style.metrics['blocks']

    def _lv(self, size, *fallback):
        if isinstance(size, (int, float)) or size in self.style.type:
            return size
        return self.style.level(*fallback)

    def _pt(self, size):
        return size if isinstance(size, (int, float)) else self.style.size(size)

    def _need(self, text, w, size, bold=False, pad=None, gap_pt=0.0):
        """一段（\\n 分段）文字排在宽 w 的框里要多高（cm），含框的上下边距。"""
        pad = self._bm['padCm'] if pad is None else pad
        pt = self._pt(size)
        paras = str(text).split('\n')
        n = sum(text_lines(p, w - 2 * pad, pt, self.family, bold) for p in paras)
        vpad = min(pad, 0.1)
        return n * pt * PT_CM * LINE * self.style.metrics['text']['lineSpacing'] + (len(paras) - 1) * gap_pt * PT_CM + 2 * vpad + 0.05

    def _nth(self, colors, k, default):
        if colors is None:
            return default
        if isinstance(colors, str):
            if colors in self.style.palettes:       # 方案名：逐项轮流取这组类别色
                codes = list(self.style.classes(colors))
                return colors, codes[k % len(codes)]
            return colors
        if isinstance(colors, tuple) and len(colors) == 2 and colors[0] in self.style.palettes and isinstance(colors[1], str):
            return colors
        seq = list(colors)
        return seq[k % len(seq)]

    def _item_color(self, colors, k, item, idx, default):
        if isinstance(item, (tuple, list)) and len(item) > idx and item[idx] is not None:
            return item[idx]
        return self._nth(colors, k, default)

    def _ink(self, fill):
        """实底上的字色：底色亮就用正文色，暗就用 blocks.labelText（白）。"""
        h = self.style.color(fill).lstrip('#')
        r, g, b = (int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))
        lum = 0.2126 * r + 0.7152 * g + 0.0722 * b
        return 'text' if lum > 0.62 else self._bm['labelText']

    def _paras(self, text, size, color, bold=False, align='left', first=None):
        """字符串 → 段落列表 [(文字, 字号, 颜色, 加粗, 对齐)]。first 给第一段另定 (字号, 颜色, 加粗, 对齐)。"""
        out = []
        for k, t in enumerate(str(text).split('\n')):
            out.append((t, *first) if (k == 0 and first) else (t, size, color, bold, align))
        return out

    def _panel(self, x, y, w, h, paras, slot, fill=None, line=None, anchor='middle', pad=None, gap_pt=0.0, shape='rect'):
        """一个框：fill 实底或 line 线框（都没有时是文本框），写入段落。"""
        pad = self._bm['padCm'] if pad is None else pad
        if fill is not None or line is not None:
            s = self.shape(shape, x, y, w, h, fill, slot, line=line)
        else:
            s = self.name(self.slide.shapes.add_textbox(Cm(x), Cm(y), Cm(w), Cm(h)), slot)
        tf = s.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Cm(pad)
        tf.margin_top = tf.margin_bottom = Cm(min(pad, 0.1))
        tf.vertical_anchor = _ANCHOR[anchor]
        for k, (text, size, color, bold, align) in enumerate(paras):
            p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
            p.alignment = _ALIGN[align]
            p.line_spacing = self.style.metrics['text']['lineSpacing']
            if k and gap_pt:
                p.space_before = Pt(gap_pt)
            if text:
                self.runs(p, text, size, color, bold)
        return s

    def _over(self, slot, need, room):
        self.deck.warn(f'第 {self.number} 页 {self.type}:{slot}：内容约需 {need:.1f} cm 高，只有 {room:.1f} cm；'
                       f'拆分节点、少说一件事或换版式，不要缩字号')

    def _bottom(self, bottom):
        return self.style.y1 if bottom is None else bottom

    def _spread(self, hs, y, gaps, bottom, slot):
        """各行按内容的高度 hs；铺到 bottom 时把多余高度均分给各行。"""
        natural = sum(hs) + gaps
        if bottom is False:
            if y + natural > self.style.page_h + 0.01:
                self._over(slot, natural, self.style.page_h - y)
            return hs
        room = self._bottom(bottom) - y
        if natural > room + 0.05:
            self._over(slot, natural, room)
            return hs
        extra = (room - natural) / len(hs)
        return [h + extra for h in hs]

    def _box(self, x, y, w, h):
        from .deck import Box
        return Box(x, y, w, h)

    # ---------- 标签行 ----------
    def rows(self, x, y, w, items, bottom=None, label_w=None, label_size='key', size='body', body='tint', colors=None,
             gap=None, slot='row'):
        """标签行：左边实色标签写短名（案例名、环节名、文件名），右边正文，自上而下 2—5 行。
        items：(标签, 正文[, 颜色])。body='tint' 正文垫浅底（blocks.bodyFill），'plain' 白底、行间细线。
        不当每页的默认做法：一册用它的页超过一成，inspect 会警告（docs/patterns.md“标签行”）。返回 Box。"""
        m = self._bm
        gap = m['gapCm'] if gap is None else gap
        lsize = self._lv(label_size, 'heading')
        lpt = self._pt(lsize)
        if label_w is None:
            widest = max(text_width(line, lpt, self.family, True) for it in items for line in str(it[0]).split('\n'))
            label_w = min(max(widest + 2 * m['padCm'] + 0.15, 2.2), w * 0.35)
        bw = w - label_w
        hs = [max(self._need(it[0], label_w, lsize, True), self._need(it[1], bw, size)) for it in items]
        hs = self._spread(hs, y, gap * (len(items) - 1), bottom, slot)
        cy = y
        for k, (it, h) in enumerate(zip(items, hs)):
            fill = self._item_color(colors, k, it, 2, m['labelFill'])
            self._panel(x, cy, label_w, h, self._paras(it[0], lsize, self._ink(fill), True, 'center'), f'{slot}-label', fill=fill)
            bfill = m['bodyFill'] if body == 'tint' else None
            self._panel(x + label_w, cy, bw, h, self._paras(it[1], size, 'text'), f'{slot}-body', fill=bfill)
            if body != 'tint' and k < len(items) - 1:
                self.line(x + label_w, cy + h + gap / 2, x + w, cy + h + gap / 2, 'line', 0.75, slot=f'{slot}-rule')
            cy += h + gap
        return self._box(x, y, w, cy - gap - y)

    # ---------- 实色栏头并列 ----------
    def columns(self, x, y, w, items, bottom=None, head_size='key', size='body', subs=None, body='tint', colors=None,
                gap=None, head_h=None, slot='col'):
        """实色栏头并列：2—4 栏等宽的同类内容（几个问题、几类做法），栏头整块实色写短名；
        head_size='figure' 时栏头写三五个字的大字。subs 是栏头下的一行副题（标题色加粗，可选）。
        正文从上往下排，底框铺到 bottom（默认版心底）。items：(栏头, 正文[, 颜色])。返回 Box。"""
        m = self._bm
        n = len(items)
        gap = 2.5 * m['gapCm'] if gap is None else gap
        cw = (w - gap * (n - 1)) / n
        hsize = self._lv(head_size, 'key', 'heading')
        if head_h is None:
            head_h = max(self._need(it[0], cw, hsize, True) for it in items) + 0.15
        gap_pt = 3
        paras = []
        for k, it in enumerate(items):
            ps = self._paras(it[1], size, 'text')
            if subs and subs[k]:
                ps = [(subs[k], size, m['subText'], True, 'center')] + ps
            paras.append(ps)
        need = max(self._need(it[1], cw, size, gap_pt=gap_pt)
                   + (self._need(subs[k], cw, size, True) - 0.25 + gap_pt * PT_CM if subs and subs[k] else 0)
                   for k, it in enumerate(items))
        if bottom is False:
            body_h = need
        else:
            body_h = self._bottom(bottom) - y - head_h
            if need > body_h + 0.05:
                self._over(slot, head_h + need, head_h + body_h)
        for k, it in enumerate(items):
            cx = x + k * (cw + gap)
            fill = self._item_color(colors, k, it, 2, m['labelFill'])
            self._panel(cx, y, cw, head_h, self._paras(it[0], hsize, self._ink(fill), True, 'center'), f'{slot}-head', fill=fill)
            bfill = m['bodyFill'] if body == 'tint' else None
            self._panel(cx, y + head_h, cw, body_h, paras[k], f'{slot}-body', fill=bfill, anchor='top', gap_pt=gap_pt)
        return self._box(x, y, w, head_h + body_h)

    # ---------- 通栏与实心箭头 ----------
    def banner(self, x, y, w, text, h=None, color='title', size='key', align='center', bold=True, slot='banner'):
        """结论通栏：整块实色写一两句共同的判断或要求，一页至多一处（visual-language.md“框和卡片”）。
        浅底的约束条、依据条用 color='emphasisTint' 之类的浅色、size='body'、bold=False、align='left'，引领词照常标红。返回 Box。"""
        sz = self._lv(size, 'key', 'heading')
        need = self._need(text, w, sz, bold) + 0.2
        h = need if h is None else h
        if need > h + self._pt(sz) * PT_CM * LINE / 2:
            self._over(slot, need, h)
        self._panel(x, y, w, h, self._paras(text, sz, self._ink(color), bold, align), slot, fill=color)
        return self._box(x, y, w, h)

    def block_arrow(self, x, y, w, h, direction='down', color='title', slot='block-arrow'):
        """实心箭头：把几栏的内容引到下面的通栏，或把左边的条件引到右边的结论。"""
        from .deck import _pick
        s = self.shape(_pick(ARROWS, direction, 'direction'), x, y, w, h, color, slot)
        return self._box(x, y, w, h) if s is not None else None

    # ---------- 箭头链、节点链 ----------
    def chevrons(self, x, y, w, h, items, colors=None, size='key', slot='chevron'):
        """粗箭头链：首块五边形、后面燕尾箭头，写阶段或环节的短名（两三个到六个字）。items：文字或 (文字, 颜色)。返回 Box。"""
        n = len(items)
        ov = min(h * 0.35, w / n * 0.3)
        cw = (w + ov * (n - 1)) / n
        sz = self._lv(size, 'key', 'heading')
        for k, it in enumerate(items):
            text = it if isinstance(it, str) else it[0]
            fill = self._item_color(colors, k, it, 1, self._bm['labelFill'])
            s = self._panel(x + k * (cw - ov), y, cw, h, self._paras(text, sz, self._ink(fill), True, 'center'), slot,
                            fill=fill, shape=MSO_SHAPE.PENTAGON if k == 0 else MSO_SHAPE.CHEVRON, pad=0.05)
            s.text_frame.margin_left = Cm(ov * (0.2 if k == 0 else 0.75))
            s.text_frame.margin_right = Cm(ov * 0.6)
            room = cw - ov * (1.35 if k else 0.8)
            if self._need(text, room, sz, True, pad=0.05) > h + 0.05:
                self._over(slot, self._need(text, room, sz, True, pad=0.05), h)
        return self._box(x, y, w, h)

    def flow(self, x, y, w, steps, h=None, colors=None, notes=None, size='body', note_size='body', gap=None, slot='flow'):
        """实色节点链：节点之间是粗箭头（style.strokes），只用于真正有先后的程序。
        steps：文字或 (文字, 颜色)；颜色可以用 'fill' 之类的浅色表示一般环节，实色表示重点环节。
        notes：每个节点下方的说明，与 steps 一一对应，可以是 None。返回 Box（含说明）。"""
        n = len(steps)
        gap = max(0.55, w * 0.025) if gap is None else gap
        nw = (w - gap * (n - 1)) / n
        texts = [s if isinstance(s, str) else s[0] for s in steps]
        need = max(self._need(t, nw, size, True, pad=0.1) for t in texts) + 0.1
        h = need if h is None else h
        if need > h + 0.05:
            self._over(slot, need, h)
        bottom = y + h
        for k, st in enumerate(steps):
            nx = x + k * (nw + gap)
            fill = self._item_color(colors, k, st, 1, self._bm['labelFill'])
            self._panel(nx, y, nw, h, self._paras(texts[k], size, self._ink(fill), True, 'center'), slot, fill=fill, pad=0.1)
            if k < n - 1:
                self.arrow(nx + nw + 0.05, y + h / 2, nx + nw + gap - 0.05, y + h / 2, slot=f'{slot}-arrow')
            note = notes[k] if notes and k < len(notes) else None
            if note:
                nh = self._need(note, nw + gap * 0.8, note_size, pad=0.05)
                self._panel(nx, y + h + 0.08, nw + gap * 0.8, nh, self._paras(note, note_size, 'text', False, 'left'),
                            f'{slot}-note', anchor='top', pad=0.05)
                bottom = max(bottom, y + h + 0.08 + nh)
        return self._box(x, y, w, bottom - y)

    # ---------- 泳道 ----------
    def lanes(self, x, y, w, h, lanes, stages, nodes, links=(), lane_w=None, stage_h=None, slot='lane'):
        """泳道流程：行是部门或主体（泳道），列是阶段；用于分工和办理程序。
        lanes：(名称, 颜色[, 高度权重])；stages：阶段名列表；
        nodes：(泳道序号, 阶段序号或 (起, 止), 文字[, 样式])，样式 'solid' 泳道色实底（默认）、'light' 浅底、
        'plain' 只有字、'emphasis' 强调色实底（最后的决定、卡点）；
        links：(节点序号, 节点序号)，画带箭头的折线，端点粘在节点上，移动节点时跟着走。返回各节点的 Box 列表。"""
        m = self._bm
        gap = 0.05
        bpt = self.style.size('body')
        if lane_w is None:
            widest = max(text_width(line, bpt, self.family, True) for ln in lanes for line in str(ln[0]).split('\n'))
            lane_w = min(max(widest + 0.5, 2.4), w * 0.16)
        sw = (w - lane_w) / len(stages)
        if stage_h is None:
            stage_h = max(self._need(s, sw - gap, 'body', True, pad=0.08) for s in stages)
        for k, name in enumerate(stages):
            self._panel(x + lane_w + k * sw + gap / 2, y, sw - gap, stage_h,
                        self._paras(name, 'body', self._ink(m['stageFill']), True, 'center'), f'{slot}-stage', fill=m['stageFill'], pad=0.08)
        weights = [ln[2] if len(ln) > 2 else 1 for ln in lanes]
        room = h - stage_h - gap * len(lanes)
        heights = [room * wt / sum(weights) for wt in weights]
        tops, cy = [], y + stage_h + gap
        for k, (ln, lh) in enumerate(zip(lanes, heights)):
            fill = ln[1]
            self._panel(x, cy, lane_w - gap, lh, self._paras(ln[0], 'body', self._ink(fill), True, 'center'), f'{slot}-head',
                        fill=fill, pad=0.08)
            self.shape('rect', x + lane_w, cy, w - lane_w, lh, m['laneFill'], f'{slot}-band')
            tops.append(cy)
            cy += lh + gap
        boxes, shapes = [], []
        pad = 0.12
        for nd in nodes:
            li, st, text = nd[0], nd[1], nd[2]
            kind = nd[3] if len(nd) > 3 else 'solid'
            if kind not in NODE_KINDS:
                raise ValueError(f'lanes 节点样式 {kind!r} 不认识；可用：{"、".join(NODE_KINDS)}')
            s0, s1 = (st, st) if isinstance(st, int) else st
            nx, nw = x + lane_w + s0 * sw + pad, (s1 - s0 + 1) * sw - 2 * pad
            lh = heights[li]
            need = self._need(text, nw, 'body', True, pad=0.08)
            nh = min(lh - 2 * pad, max(need, (lh - 2 * pad) * 0.7))
            if need > lh - 2 * pad + 0.05:
                self._over(f'{slot}-node', need, lh - 2 * pad)
            ny = tops[li] + (lh - nh) / 2
            fill = {'solid': lanes[li][1], 'emphasis': 'emphasis', 'light': 'white', 'plain': None}[kind]
            color = self._ink(fill) if fill else 'text'
            s = self._panel(nx, ny, nw, nh, self._paras(text, 'body', color, kind != 'plain', 'center' if kind != 'plain' else 'left'),
                            f'{slot}-node', fill=fill, pad=0.08)
            boxes.append(self._box(nx, ny, nw, nh))
            shapes.append(s)
        st = self.style.metrics['strokes']
        for a, b in links:
            A, B = boxes[a], boxes[b]
            if abs((A.y + A.h / 2) - (B.y + B.h / 2)) < 0.05:         # 同一泳道：水平
                sa, sb = (3, 1) if B.x >= A.right - 0.01 else (1, 3)
                kind = MSO_CONNECTOR.STRAIGHT
            elif B.x < A.right and B.right > A.x:                    # 列上重叠：在重叠段中间画竖线
                mx = (max(A.x, B.x) + min(A.right, B.right)) / 2
                y0, y1 = (A.bottom, B.y) if B.y > A.y else (A.y, B.bottom)
                c = self.line(mx, y0, mx, y1, st['linkColor'], st['linkPt'], slot=f'{slot}-link', arrow=True)
                continue
            else:
                sa, sb = (3, 1) if B.x >= A.right else (1, 3)
                kind = MSO_CONNECTOR.ELBOW
            c = self.name(self.slide.shapes.add_connector(kind, 0, 0, Cm(1), Cm(1)), f'{slot}-link')
            style = c._element.find(qn('p:style'))
            if style is not None:
                c._element.remove(style)
            c.begin_connect(shapes[a], sa)
            c.end_connect(shapes[b], sb)
            c.line.color.rgb = self.rgb(st['linkColor'])
            c.line.width = Pt(st['linkPt'])
            ln = c.line._get_or_add_ln()
            ln.append(ln.makeelement(qn('a:tailEnd'), {'type': 'triangle', 'w': 'med', 'len': 'med'}))
        return boxes

    # ---------- 文件摘录图 ----------
    def excerpt(self, path, x, y, w, h, crop=None, marks=(), caption=None, slot='excerpt'):
        """文件摘录图：政策文件、规划文本的截图，裁到重点，细线边框，要点用红色虚线框圈出，图注在下（文件全称和文号）。
        crop=(左, 上, 右, 下)、marks 里的每个框都按原图的比例（0—1）给。红虚线框是证据物的一部分，工作稿就加。
        裁剪不改原图，PowerPoint 里还能调。返回 Box（含图注）。"""
        m = self._bm
        with Image.open(path) as im:
            iw, ih = im.size
        l, t, r, b = crop or (0, 0, 1, 1)
        if not (0 <= l < r <= 1 and 0 <= t < b <= 1):
            raise ValueError(f'excerpt crop 要写成原图比例 (左, 上, 右, 下)，0—1 且左<右、上<下：{crop}')
        vw, vh = iw * (r - l), ih * (b - t)
        cap_h = self._need(caption, w, 'body', pad=0.05) if caption else 0
        room = h - cap_h - (0.08 if caption else 0)
        k = min(w / vw, room / vh)
        pw, ph = vw * k, vh * k
        px = x + (w - pw) / 2
        pic = self.name(self.slide.shapes.add_picture(path, Cm(px), Cm(y), Cm(pw), Cm(ph)), slot)
        pic.crop_left, pic.crop_top, pic.crop_right, pic.crop_bottom = l, t, 1 - r, 1 - b
        pic.line.color.rgb = self.rgb(m['frameColor'])
        pic.line.width = Pt(0.5)
        for fl, ft, fr, fb in marks:
            fl, ft, fr, fb = max(fl, l), max(ft, t), min(fr, r), min(fb, b)
            if fr <= fl or fb <= ft:
                self.deck.warn(f'第 {self.number} 页 {self.type}:{slot}：有一个红框在裁剪范围外，没画')
                continue
            s = self.shape('rect', px + (fl - l) / (r - l) * pw, y + (ft - t) / (b - t) * ph, (fr - fl) / (r - l) * pw,
                           (fb - ft) / (b - t) * ph, None, f'{slot}-mark', line=m['markColor'])
            s.line.width = Pt(m['markPt'])
            s.line.dash_style = MSO_LINE_DASH_STYLE.DASH
        bottom = y + ph
        if caption:
            self._panel(x, bottom + 0.08, w, cap_h, self._paras(caption, 'body', m['captionColor'], False, 'center'),
                        f'{slot}-caption', anchor='top', pad=0.05)
            bottom += 0.08 + cap_h
        return self._box(x, y, w, bottom - y)

    # ---------- 矩阵表 ----------
    def matrix(self, x, y, w, rows, widths, header_rows=1, merges=(), row_heads=None, col_heads=None, highlight=(),
               bottom=None, sizes=None, align=None, fmt='{:.1f}', slot='table'):
        """矩阵表：逐行比对的内容（事权、清单、几个案例的同一组要素）。
        表头深色，可两行并合并（header_rows、merges）；row_heads 给表头以下第 0 列整格实色白字（一个颜色、每行一个
        颜色的列表、方案名，或 True 用 blocks.labelFill）；col_heads 给表头各列一个颜色（类别色表头）；
        highlight 是要提示的格 (行, 列) 或整行行号，浅红底强调色加粗。
        bottom（默认版心底）：在正文 1.25、1.17、1.08、1 倍的字号里取放得下的最大值，多余高度均分给各行；
        bottom=False 用正文字号、按内容定高。rows 里被合并的格写 ''。返回 Box。"""
        m, tb = self._bm, self.style.metrics['table']
        nr, nc = len(rows), len(widths)
        if any(len(r) != nc for r in rows):
            raise ValueError(f'matrix：{nc} 个列宽，各行的格数是 {sorted({len(r) for r in rows})}；要一样多')
        cw = [w * wd / sum(widths) for wd in widths]
        show = lambda v: fmt.format(v) if isinstance(v, float) else ('' if v is None else str(v))
        covered, span_w, vspan = set(), {}, {}
        for r1, c1, r2, c2 in merges:
            for rr in range(r1, r2 + 1):
                for cc in range(c1, c2 + 1):
                    if (rr, cc) != (r1, c1):
                        covered.add((rr, cc))
            span_w[(r1, c1)] = sum(cw[c1:c2 + 1])
            if r2 > r1:
                vspan[(r1, c1)] = r2
        cells = {(r, c) for r in range(nr) for c in range(nc)}
        hl = {(r, c) for h in highlight for (r, c) in ([h] if isinstance(h, tuple) else [(h, c) for c in range(nc)])}
        pad = tb['paddingCm']

        def heights(pt):
            hs = []
            for r in range(nr):
                need = pt * PT_CM * LINE * 1.15 + 0.12
                for c in range(nc):
                    if (r, c) in covered or (r, c) in vspan:
                        continue
                    bold = r < header_rows or (row_heads and c == 0) or (r, c) in hl
                    need = max(need, self._need(show(rows[r][c]), span_w.get((r, c), cw[c]), pt, bold, pad=pad))
                hs.append(need)
            for (r, c), r2 in vspan.items():
                need = self._need(show(rows[r][c]), span_w.get((r, c), cw[c]), pt, True, pad=pad)
                have = sum(hs[r:r2 + 1])
                if need > have:
                    hs[r2] += need - have
            return hs

        body = self.style.size('body')
        options = sizes or [round(body * f * 2) / 2 for f in TABLE_STEPS]
        if bottom is False:
            pt, hs = body, heights(body)
            if y + sum(hs) > self.style.page_h + 0.01:
                self._over(slot, sum(hs), self.style.page_h - y)
        else:
            room = self._bottom(bottom) - y
            pt, hs = options[-1], heights(options[-1])
            for cand in options:
                got = heights(cand)
                if sum(got) <= room:
                    pt, hs = cand, got
                    break
            if sum(hs) > room + 0.05:
                self._over(slot, sum(hs), room)
            else:
                cap = 1.4 * self.style.page_h / 19.05
                extra = min((room - sum(hs)) / nr, cap)
                hs = [v + extra for v in hs]
        gf = self.name(self.slide.shapes.add_table(nr, nc, Cm(x), Cm(y), Cm(w), Cm(sum(hs))), slot)
        t = gf.table
        tblPr = t._tbl.tblPr
        for a in list(tblPr.attrib):
            del tblPr.attrib[a]
        for ch in list(tblPr):
            tblPr.remove(ch)
        for c, wd in enumerate(cw):
            t.columns[c].width = Cm(wd)
        for r, hv in enumerate(hs):
            t.rows[r].height = Cm(hv)
        for r, c in sorted(cells):
            cell = t.cell(r, c)
            v = rows[r][c]
            text = '' if (r, c) in covered else show(v)
            if r < header_rows:
                fill = self._nth(col_heads, c, tb['headerFill']) if col_heads is not None else tb['headerFill']
                color, bold, al = self._ink(fill), True, 'center'
            elif row_heads and c == 0:
                fill = m['labelFill'] if row_heads is True else self._nth(row_heads, r - header_rows, m['labelFill'])
                color, bold, al = self._ink(fill), True, 'left'
            elif (r, c) in hl:
                fill, color, bold = tb['highlightFill'], m['highlightText'], True
                al = 'center'
            else:
                fill, color, bold = None, tb['text'], False
                al = 'right' if isinstance(v, (int, float)) and not isinstance(v, bool) else 'left'
            if align and r >= header_rows and not (row_heads and c == 0):
                al = align[c]
            cell.margin_left = cell.margin_right = Cm(pad)
            cell.margin_top = cell.margin_bottom = Cm(0.03)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            for k, para in enumerate(text.split('\n')):
                p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
                p.alignment = _ALIGN[al]
                if para:
                    self.runs(p, para, pt, color, bold)
            self._matrix_lines(cell, fill is not None, top=r > 0)
            if fill is not None:
                cell.fill.solid()
                cell.fill.fore_color.rgb = self.rgb(fill)
            else:
                cell.fill.background()
        for r1, c1, r2, c2 in merges:
            t.cell(r1, c1).merge(t.cell(r2, c2))
        self.deck.fitted.append({'page': self.number, 'slot': f'pd:{self.type}:{slot}', 'pt': pt})
        return self._box(x, y, w, sum(hs))

    def _matrix_lines(self, cell, solid, top):
        """实底格四边白线隔开；白底格只画上下细线，不画竖线（visual-language.md“框和卡片”的表格）。"""
        tb = self.style.metrics['table']
        tc = cell._tc.get_or_add_tcPr()
        for tag in ('a:lnL', 'a:lnR', 'a:lnT', 'a:lnB'):
            for e in tc.findall(qn(tag)):
                tc.remove(e)
        for k, side in enumerate(('a:lnL', 'a:lnR', 'a:lnT', 'a:lnB')):
            if solid:
                on, color, width = True, self.style.color('white'), Pt(1)
            else:
                on = side in ('a:lnT', 'a:lnB') and (side == 'a:lnB' or top)
                color, width = self.style.color(tb['lineColor']), Pt(tb['linePt'])
            ln = tc.makeelement(qn(side), {'w': str(width if on else 0)})
            fill = ln.makeelement(qn('a:solidFill' if on else 'a:noFill'), {})
            if on:
                fill.append(fill.makeelement(qn('a:srgbClr'), {'val': color.lstrip('#')}))
            ln.append(fill)
            tc.insert(k, ln)

    # ---------- 分档台阶 ----------
    def steps(self, x, y, w, items, h=None, rise=None, colors=None, size='key', note_size='body', gap=None, body='tint',
              slot='step'):
        """分档台阶：逐级升高的实色块，块上写档名，块下写这一档怎么办理；用于按程度分档（预警线、分级办理）。
        items：(档名, 说明[, 颜色])，从低到高；颜色常用一组从浅到深的类别色（如 colors='分档'）。
        给了 h 时级差随之伸缩，整组铺到 y+h，body='tint' 的说明框（blocks.bodyFill）一直铺到底；rise 直接给级差。
        body='plain' 说明只有字。返回 Box。"""
        n = len(items)
        gap = self._bm['gapCm'] if gap is None else gap
        bw = (w - gap * (n - 1)) / n
        sz = self._lv(size, 'key', 'heading')
        block_h = max(self._need(it[0], bw, sz, True) for it in items) + 0.1
        note_h = max(self._need(it[1], bw, note_size, pad=0.12) for it in items) + 0.1
        if rise is None:
            rise = block_h * 0.55
            if h is not None and n > 1:     # 给了整组高度：级差随之伸缩，最矮也要看得出台阶
                rise = min(max((h - block_h - note_h) / (n - 1), block_h * 0.3), block_h * 1.6)
        natural = (n - 1) * rise + block_h + note_h
        if h is not None and natural > h + 0.05:
            self._over(slot, natural, h)
        bottom = y + max(h or 0, natural)
        for k, it in enumerate(items):
            bx, by = x + k * (bw + gap), y + (n - 1 - k) * rise
            fill = self._item_color(colors, k, it, 2, self._bm['labelFill'])
            self._panel(bx, by, bw, block_h, self._paras(it[0], sz, self._ink(fill), True, 'center'), f'{slot}-block', fill=fill)
            ny = by + block_h
            nh = (bottom - ny) if (body == 'tint' and h is not None) else note_h
            self._panel(bx, ny, bw, nh, self._paras(it[1], note_size, 'text', False, 'center'), f'{slot}-note',
                        fill=self._bm['bodyFill'] if body == 'tint' else None, anchor='middle' if body == 'tint' else 'top', pad=0.12)
        return self._box(x, y, w, bottom - y)
