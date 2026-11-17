import io
import json
import os
import unittest
from contextlib import redirect_stdout

from helpers import EXAMPLE_STYLE, TempDir, make_release, style_doc
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.util import Cm

from plan_deck import Deck, MapRelease, Style, StyleError, notes_frame, replace_text
from plan_deck.cli import main
from plan_deck.inspect import inspect
from plan_deck.measure import text_height, text_units


def build(path, root):
    """一份两页的虚构汇报，用到全部组件。"""
    deck = Deck(EXAMPLE_STYLE)
    st = deck.style
    x0, y0, x1, y1 = st.body
    rel = MapRelease(root, 'v1')
    p = deck.page('evidence', '2.1 现状用地', parent='二、现状评估与问题识别', lead='北岸老城**更新潜力 98 公顷**，近期 3 处')
    x, w = st.col(0, 4)
    box = p.text(x, y0, w, [('h', '现状'), '老旧小区 **12 个**，城中村 3 个。', ('m', '次要说明'), ('big', '98.4')])
    t = p.table(x, box.bottom + 0.4, w, [['方式', '处数', '面积（公顷）'], ['拆除重建', 3, 21.5], ['合计', 3, 21.5]],
                [2, 1, 1.4], total=True, highlight=(1,))
    mx, mw = st.col(4, 5)
    m = p.map(rel.map('m1-site'), mx, y0, mw, y1 - y0)
    lx, lw = st.col(9, 3)
    p.legend(rel.map('m1-site'), lx, y0, lw, extra=[('emphasis', '近期启动区', 'frame')])
    p.notes('数字取自虚构测算表。')

    q = deck.page('ledger', '2.2 更新规模测算', parent='二、现状评估与问题识别', lead='净增 **36.2** 万㎡')
    x, w = st.col(0, 6)
    q.ledger(x, y0, [(52.0, '新建', ('更新方式', '拆除重建')), '−', (15.8, '拆除', 'muted'), '=', (36.2, '净增', 'emphasis')],
             w=(w - 1.8) / 3, h=3.2)
    for i in range(3):
        q.node(x + i * 5, y0 + 5, 4, 2, f'步骤 {i + 1}\n说明', ('实施时序', '近期'), 'node-step')
        if i:
            q.arrow(x + i * 5 - 0.9, y0 + 6, x + i * 5 - 0.1, y0 + 6)
    q.shape('diamond', x, y0 + 8, 0.4, 0.4, 'title', 'mark', i=0)
    q.line(x, y0 + 9, x + 5, y0 + 9, 'line', 2, slot='bar')
    cx, cw = st.col(6, 6)
    ch = q.chart('bar-stacked', cx, y0, cw, y1 - y0, ['甲', '乙'], [('计入', [3.0, 4.0]), ('不计', [1.0, 2.0])],
                 [('更新方式', '有机更新'), 'line'], gap=40, overlap=100)
    q.color_point(ch, 0, 1, 'emphasis')
    q.label_series(ch, 0)
    deck.save(path)
    return deck, box, t, m


class DeckTest(unittest.TestCase):
    def test_build_and_reopen(self):
        with TempDir() as tmp:
            make_release(tmp)
            out = os.path.join(tmp, 'deck.pptx')
            deck, box, t, m = build(out, tmp)
            prs = Presentation(out)
            self.assertEqual(prs.slide_width, Cm(52.8))
            self.assertEqual(len(prs.slides), 2)
            s1 = prs.slides[0]
            names = [s.name for s in s1.shapes]
            # 全部是工具生成的、带 pd: 前缀的形状，同页不重名
            self.assertTrue(all(n.startswith('pd:evidence:') for n in names), names)
            self.assertEqual(len(names), len(set(names)))
            for n in ('pd:evidence:title', 'pd:evidence:parent-title', 'pd:evidence:subtitle', 'pd:evidence:map',
                      'pd:evidence:table', 'pd:evidence:header-rule', 'pd:evidence:legend-swatch:0'):
                self.assertIn(n, names)
            title = next(s for s in s1.shapes if s.name == 'pd:evidence:title')
            run = title.text_frame.paragraphs[0].runs[0]
            self.assertEqual(run.font.size.pt, 28)
            self.assertEqual(str(run.font.color.rgb), '194461')
            # 正文里的 ** 是强调色加粗
            body = next(s for s in s1.shapes if s.name == 'pd:evidence:text')
            runs = body.text_frame.paragraphs[1].runs
            self.assertEqual([(str(r.font.color.rgb), r.font.bold) for r in runs],
                             [('262626', False), ('C00000', True), ('262626', False)])
            # 页面上没有比正文小的字
            for sh in s1.shapes:
                if sh.has_text_frame:
                    for p in sh.text_frame.paragraphs:
                        for r in p.runs:
                            self.assertGreaterEqual(r.font.size.pt, 18, sh.name)
            self.assertAlmostEqual(m.w, 10.0)
            self.assertGreater(t.y, box.y)
            self.assertIn('数字取自', s1.notes_slide.notes_text_frame.text)
            s2 = prs.slides[1]
            names2 = [s.name for s in s2.shapes]
            self.assertIn('pd:ledger:ledger-block:2', names2)
            self.assertIn('pd:ledger:ledger-op:1', names2)
            self.assertIn('pd:ledger:node-step:2', names2)      # 同槽位自动编号
            self.assertIn('pd:ledger:node-step:0', names2)
            self.assertIn('pd:ledger:chart', names2)
            # 形状不带主题样式引用（PowerPoint 会按主题画阴影）
            block = next(s for s in s2.shapes if s.name == 'pd:ledger:ledger-block:0')
            self.assertNotIn('effectRef', block._element.xml)
            # 图表文字的西文、东亚字体都写了（只写一个时 PowerPoint 的中文会落回宋体）
            xml = next(s for s in s2.shapes if s.name == 'pd:ledger:chart').chart.part.blob.decode('utf-8')
            self.assertGreaterEqual(xml.count('<a:ea typeface="微软雅黑"/>'), 4)
            self.assertNotIn('<a:latin typeface="微软雅黑"/>', xml)

    def test_raw_hex_refused_in_components(self):
        deck = Deck(EXAMPLE_STYLE)
        p = deck.page('diagram', '1.1 框架')
        with self.assertRaises(StyleError):
            p.node(0, 0, 3, 1, '节点', '#ED7D31')
        with self.assertRaises(StyleError):
            p.chart('bar', 0, 0, 5, 5, ['a'], [('s', [1.0])], ['#ED7D31'])

    def test_duplicate_names_refused(self):
        deck = Deck(EXAMPLE_STYLE)
        p = deck.page('schedule', '1.4 时序')
        p.shape('oval', 0, 0, 1, 1, 'title', 'mark', i=0)
        with self.assertRaisesRegex(ValueError, '重名'):
            p.shape('rect', 0, 0, 1, 1, 'title', 'mark', i=0)

    def test_unknown_item_kind(self):
        deck = Deck(EXAMPLE_STYLE)
        p = deck.page('diagram', '1.1 框架')
        with self.assertRaisesRegex(ValueError, '不认识'):
            p.text(0, 0, 5, [('caption', '小字')])

    def test_enum_args_list_choices(self):
        deck = Deck(EXAMPLE_STYLE)
        p = deck.page('diagram', '1.1 框架')
        for call, choices in ((lambda: p.shape('arc', 0, 0, 1, 1, 'title', 'm'), 'rect、oval、diamond'),
                              (lambda: p.text(0, 0, 5, ['字'], anchor='centre'), 'top、middle、bottom'),
                              (lambda: p.label(0, 0, 5, 1, '字', 'l', align='justify'), 'left、center、right')):
            with self.assertRaisesRegex(ValueError, choices):
                call()
        # 其他形状直接用 python-pptx 的枚举
        s = p.shape(MSO_SHAPE.DONUT, 0, 0, 2, 2, 'title', 'ring')
        self.assertEqual(s.auto_shape_type, MSO_SHAPE.DONUT)
        self.assertNotIn('effectRef', s._element.xml)

    def test_table_shape_mismatch(self):
        p = Deck(EXAMPLE_STYLE).page('table', '1.1 表')
        with self.assertRaisesRegex(ValueError, '列宽'):
            p.table(0, 7, 10, [['a', 'b'], ['c']], [1, 1])
        with self.assertRaisesRegex(ValueError, '列宽'):
            p.table(0, 7, 10, [['a', 'b']], [1, 1, 1])

    def test_bold_marks(self):
        """三种记号的颜色和粗细；__ 跟所在段落的颜色走，页面上不留记号。"""
        deck = Deck(EXAMPLE_STYLE)
        st = deck.style
        p = deck.page('evidence', '2.1 现状用地')
        box = p.text(0, 0, 20, ['普通**强调**普通__正文粗__普通^^标题粗^^', ('m', '次要__次要粗__')])
        tag = lambda r: (r.text, str(r.font.color.rgb), bool(r.font.bold))
        hex_ = lambda name: st.color(name).lstrip('#').upper()
        paras = box.shape.text_frame.paragraphs
        self.assertEqual([tag(r) for r in paras[0].runs],
                         [('普通', hex_('text'), False), ('强调', hex_('emphasis'), True), ('普通', hex_('text'), False),
                          ('正文粗', hex_('text'), True), ('普通', hex_('text'), False), ('标题粗', hex_('title'), True)])
        self.assertEqual([tag(r) for r in paras[1].runs],
                         [('次要', hex_('muted'), False), ('次要粗', hex_('muted'), True)])
        self.assertEqual(box.shape.text_frame.text, '普通强调普通正文粗普通标题粗\n次要次要粗')

    def test_bold_marks_not_counted_in_width(self):
        plain = '各地政策在适用情形、核算标准和审批程序等方面规定较细'
        for mark in ('**', '__', '^^'):
            marked = plain[:5] + mark + plain[5:15] + mark + plain[15:]
            self.assertEqual(text_units(marked), text_units(plain), mark)
            self.assertEqual(text_height(marked, 5, 18), text_height(plain, 5, 18), mark)

    def test_bold_marks_refused(self):
        """__ 和 ^^ 没配对、中间是空的、记号嵌套都报错，错误信息带上那段文字。"""
        deck = Deck(EXAMPLE_STYLE)
        p = deck.page('diagram', '1.1 框架')
        for text, why in (('没配对的__正文粗', '__ 没有配对'), ('没配对的^^标题粗', r'\^\^ 没有配对'),
                          ('__一对__又一个__', '__ 没有配对'), ('填空横线 ____ 年', '__ 中间是空的'),
                          ('^^^^', r'\^\^ 中间是空的'), ('**强调__正文粗__**', r'\*\* 里不能再套 __'),
                          ('__正文粗**强调**__', r'__ 里不能再套 \*\*'), ('__交叉^^记号__^^', r'__ 里不能再套 \^\^')):
            with self.assertRaisesRegex(ValueError, why) as ctx:
                p.text(0, 0, 5, ['别的条目', text], slot='col', i=p.next_index('col'))
            self.assertIn(text, str(ctx.exception))
            self.assertIn('第 1 页', str(ctx.exception))
        # 标签、节点、表格走同一个入口
        for call in (lambda: p.label(0, 0, 5, 1, '标签__', 'mark', i=90), lambda: p.node(0, 0, 3, 1, '节点^^', 'title', 'mark', i=91),
                     lambda: p.table(0, 0, 5, [['表头', '__单元格']], [1, 1], slot='mark', row_h=1)):
            with self.assertRaisesRegex(ValueError, '没有配对'):
                call()

    def test_unpaired_emphasis_mark_unchanged(self):
        """** 没配对的旧写法不报错，后面的字都算强调；重跑旧脚本结果不变。"""
        deck = Deck(EXAMPLE_STYLE)
        p = deck.page('diagram', '1.1 框架')
        box = p.text(0, 0, 20, ['前面**后面'])
        runs = box.shape.text_frame.paragraphs[0].runs
        self.assertEqual([(r.text, bool(r.font.bold)) for r in runs], [('前面', False), ('后面', True)])
        self.assertEqual(str(runs[1].font.color.rgb), deck.style.color('emphasis').lstrip('#').upper())

    def test_header_follows_style(self):
        doc = style_doc()
        doc['header'] = {'bars': [{'slot': 'header-bar', 'xCm': 0, 'yCm': 0, 'wCm': 'page', 'hCm': 0.4, 'color': 'title'}],
                         'lines': [{'role': 'title', 'xCm': 2, 'yCm': 1, 'wCm': 40, 'hCm': 1.4, 'size': 'title', 'color': 'title', 'bold': True}]}
        deck = Deck(Style(doc))
        p = deck.page('evidence', '2.1 现状用地', parent='不画', lead='不画')
        names = [s.name for s in p.slide.shapes]
        self.assertEqual(names, ['pd:evidence:header-bar', 'pd:evidence:title'])
        bar = p.slide.shapes[0]
        self.assertEqual(bar.width, Cm(52.8))


class CliTest(unittest.TestCase):
    def run_cli(self, *args):
        buf = io.StringIO()
        with redirect_stdout(buf):
            code = main(list(args))
        return code, json.loads(buf.getvalue())

    def test_style_check(self):
        code, out = self.run_cli('style', 'check', EXAMPLE_STYLE)
        self.assertEqual(code, 0)
        self.assertEqual(out['classes']['实施时序']['近期'], '#C00000')

    def test_style_check_error(self):
        with TempDir() as tmp:
            doc = style_doc()
            doc['type']['caption'] = 9
            p = os.path.join(tmp, 'style.json')
            with open(p, 'w', encoding='utf-8') as f:
                json.dump(doc, f, ensure_ascii=False)
            code, out = self.run_cli('style', 'check', p)
            self.assertEqual(code, 1)
            self.assertEqual(out['error'], 'STYLE_INVALID')

    def test_maps_check_and_inspect(self):
        with TempDir() as tmp:
            make_release(tmp)
            code, out = self.run_cli('maps', 'check', tmp, '--release', 'v1')
            self.assertEqual(code, 0)
            self.assertEqual(out['sourceRun'], '20260101-000000-render-demo')
            code, out = self.run_cli('maps', 'check', tmp, '--release', 'v9')
            self.assertEqual(code, 1)
            deck_path = os.path.join(tmp, 'deck.pptx')
            build(deck_path, tmp)
            code, out = self.run_cli('inspect', deck_path, '--style', EXAMPLE_STYLE, '--brief')
            self.assertEqual(code, 0, out['problems'])
            self.assertEqual(out['pages'][0]['other'], 0)
            self.assertEqual(out['pages'][0]['title'], '2.1 现状用地')
            # 手工加一个小字文本框和一个画布外对象，inspect 报出来
            prs = Presentation(deck_path)
            tb = prs.slides[0].shapes.add_textbox(Cm(60), Cm(1), Cm(5), Cm(1))
            tb.text_frame.text = '页面任务说明'
            tb2 = prs.slides[0].shapes.add_textbox(Cm(1), Cm(27), Cm(5), Cm(1))
            tb2.text_frame.paragraphs[0].add_run().text = '资料来源'
            from pptx.util import Pt
            tb2.text_frame.paragraphs[0].runs[0].font.size = Pt(9)
            prs.save(deck_path)
            code, out = self.run_cli('inspect', deck_path, '--style', EXAMPLE_STYLE, '--brief')
            self.assertEqual(code, 1)
            problems = [p['problem'] for p in out['problems']]
            self.assertTrue(any('画布外' in p for p in problems))
            self.assertTrue(any('9 pt' in p for p in problems))

    def test_inspect_red_and_page_refs(self):
        """整段、整栏标红和“第 N 页”引用只警告，不影响退出码；页眉论断句、大号数字、按条目标的引领词不算。"""
        lead = '各地已从单项试点转向系统推进，实施路径仍不完整，规划许可与土地供应的衔接尚未建立'
        para = '各地已从单项试点转向系统推进，实施路径仍不完整，规划许可与土地供应的衔接尚未建立，监管规则相对薄弱。'
        with TempDir() as tmp:
            deck = Deck(EXAMPLE_STYLE)
            x0, y0, x1, y1 = deck.style.body
            p = deck.page('section-overview', '2.3 存量更新的实施机制', parent='二、政策研究', lead=f'**{lead}**')
            p.text(x0, y0, 15, [('h', '编制阶段'), '**总体判断：**各地已从单项试点转向系统推进，但实施路径仍不完整。',
                               '各地政策在**适用情形、核算标准和审批程序**等方面规定较细。', ('big', '86.5'),
                               '**底线要求：**主导功能的建筑面积比例**不得低于 50%**。'], slot='col', i=0)
            p.text(x0 + 16, y0, 15, [f'**{para}**', '一般条目，没有强调。'], slot='col', i=1)
            p.text(x0 + 32, y0, 15, ['**用途转换：一宗用地兼容两类以上功能**', '**容积率核算：转换部分不计入**',
                                     '**审批程序：**并联办理，详见第 12 页。'], slot='col', i=2)
            path = os.path.join(tmp, 'deck.pptx')
            deck.save(path)
            code, out = self.run_cli('inspect', path, '--style', EXAMPLE_STYLE, '--brief')
        self.assertEqual(code, 0, out['problems'])
        got = {(w['shape'], w['warning'].split('：')[0]) for w in out['warnings'] if w['shape'] != 'slide'}     # 版面铺满另测
        self.assertEqual(got, {('pd:section-overview:col:1', '整段标红'), ('pd:section-overview:col:2', '整栏标红'),
                               ('pd:section-overview:col:2', '页面上有页码引用“第 12 页”，页序一调就对不上'),
                               ('pd:section-overview:subtitle', '副标题像一句解释')})

    def test_inspect_language_titles_and_notes(self):
        """讲解员口吻、偏长的标题、解释性的副标题、多页相同的备注只警告；带事实的结果句、问题页、声明性的短副标题、隐藏页不算。"""
        same = '这段讲稿讲的是另一页的内容，被复制到了很多页上面。'
        with TempDir() as tmp:
            deck = Deck(EXAMPLE_STYLE)
            x0, y0, x1, y1 = deck.style.body
            p = deck.page('evidence', '地价补缴规则：让企业事先算清“经济账”，以便在申请之前就能够把成本和收益都想清楚',
                          lead='按指南推演：编制时两次触发清单，项目实施时不必再修改控规')
            p.text(x0, y0, 15, [('h', '要解决什么'), '特色场景不是清单，是专项通道。', '怎么查：先选目标。', '两个案例靠的是区政府决定。', '清单没写到的走现行程序。'])
            p.notes(same)
            p = deck.page('evidence', '旧产业区更新潜力空间为 12.6km²，主要分布于三个片区', lead='按指南推演')
            p.text(x0, y0, 15, [('h', '存在问题'), '本项目属于一般禁止。'])
            p.notes(same)
            p = deck.page('question', '问题一：土地变更流程复杂，转换程序比例不清晰')
            p.text(x0, y0, 15, ['正文。'])
            p = deck.page('evidence', '备用：这是隐藏页，里面的“旧标题”不查')
            p.text(x0, y0, 15, ['本页是旧稿。'])
            path = os.path.join(tmp, 'deck.pptx')
            deck.save(path)
            prs = Presentation(path)
            prs.slides[3]._element.set('show', '0')
            prs.save(path)
            code, out = self.run_cli('inspect', path, '--style', EXAMPLE_STYLE, '--brief')
        self.assertEqual(code, 0, out['problems'])
        pages = {w['page'] for w in out['warnings']}
        self.assertEqual(pages, {1}, out['warnings'])
        text = ' '.join(w['warning'] for w in out['warnings'])
        for word in ('纠正理解', '教读者怎么读', '带读者往下读', '替读者点重点', '总结意义', '讨论时的口语', '偏长', '副标题像一句解释', '第 1、2 页备注相同'):
            self.assertIn(word, text)

    def test_inspect_layout_clues(self):
        """阴影、同排等尺寸的模块只警告；运行时画的形状和线不带阴影，默认箭头够粗。"""
        with TempDir() as tmp:
            deck = Deck(EXAMPLE_STYLE)
            x0, y0, x1, y1 = deck.style.body
            p = deck.page('evidence', '土地处置方式')
            arrow = p.arrow(x0, y0, x0 + 5, y0)
            self.assertEqual(arrow.line.width.pt, 3)
            p.line(x0, y0 + 1, x0 + 5, y0 + 1, 'line', 1)
            p.shape('rect', x0, y0 + 2, 5, 2, 'emphasis', 'block')
            for i in range(3):
                p.node(x0 + i * 7, y0 + 5, 6, 3, '同样大小的模块里放着一段用来分组的文字', 'title', slot='card', i=i)
            q = deck.page('evidence', '地价补缴规则')
            q.text(x0, y0, 15, ['正文' * 40])
            path = os.path.join(tmp, 'deck.pptx')
            deck.save(path)
            prs = Presentation(path)
            raw = prs.slides[1].shapes.add_shape(1, Cm(x0), Cm(y0 + 12), Cm(5), Cm(2))
            raw.name = 'raw'        # 直接 add_shape 引用主题效果；运行时建的模板里主题效果已清空，不带阴影
            shadowed = prs.slides[1].shapes.add_shape(1, Cm(x0 + 6), Cm(y0 + 12), Cm(5), Cm(2))
            shadowed.name = 'shadowed'
            shadowed.text_frame.text = '无填充的框里的字'
            shadowed.fill.background()
            shadowed._element.spPr.append(parse_xml(
                f'<a:effectLst {nsdecls("a")}><a:outerShdw blurRad="40000" dist="20000" dir="5400000">'
                '<a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst>'))
            prs.save(path)
            code, out = self.run_cli('inspect', path, '--brief')
        by_page = {}
        for w in out['warnings']:
            by_page.setdefault(w['page'], []).append((w['shape'], w['warning']))
        self.assertTrue(any('同排 3 个等尺寸' in w for _, w in by_page[1]), by_page)
        self.assertFalse(any('有阴影' in w for _, w in by_page[1]), by_page)
        self.assertTrue(any(s == 'shadowed' and '阴影落在字上' in w for s, w in by_page[2]), by_page)
        self.assertFalse(any(s == 'raw' and '阴影' in w for s, w in by_page[2]), by_page)

    def test_inspect_no_emphasis(self):
        """字多却没有强调色的页警告，表格也算；有强调色、字少、隐藏页不报；没给 style 时认偏红的字。"""
        long = '土地混合利用和用途转换按协议方式办理用地手续并补缴地价，已出让土地符合条件的可不收回。' * 6
        with TempDir() as tmp:
            deck = Deck(EXAMPLE_STYLE)
            x0, y0, x1, y1 = deck.style.body
            deck.page('evidence', '无强调').text(x0, y0, 30, [long])
            deck.page('evidence', '有强调').text(x0, y0, 30, ['**引领词：**' + long])
            deck.page('evidence', '字少').text(x0, y0, 30, ['正文不多。'])
            deck.page('evidence', '隐藏').text(x0, y0, 30, [long])
            deck.page('evidence', '表格').table(x0, y0, 30, [['项目', '说明'], ['甲', long]], widths=[1, 3])
            path = os.path.join(tmp, 'deck.pptx')
            deck.save(path)
            prs = Presentation(path)
            prs.slides[3]._element.set('show', '0')
            prs.save(path)
            for extra in (('--style', EXAMPLE_STYLE), ()):
                code, out = self.run_cli('inspect', path, *extra, '--brief')
                got = {w['page'] for w in out['warnings'] if '没有强调色' in w['warning']}
                self.assertEqual(got, {1, 5}, out['warnings'])

    def test_replace_text_across_runs(self):
        """跨 run 替换保留没涉及的部分，并标中文；notes_frame 补占位符时不带别页的备注。"""
        with TempDir() as tmp:
            deck = Deck(EXAMPLE_STYLE)
            x0, y0, x1, y1 = deck.style.body
            p = deck.page('evidence', '土地用途转换')
            p.text(x0, y0, 15, ['**清单没写到**的项目，按现行程序办理。'], slot='body')
            p.notes('别页的备注')
            deck.page('evidence', '土地功能混合').text(x0, y0, 15, ['正文'])
            path = os.path.join(tmp, 'deck.pptx')
            deck.save(path)
            prs = Presentation(path)
            sl = prs.slides[0]
            self.assertEqual(replace_text(sl, '没写到**', 'x'), 0)
            self.assertEqual(replace_text(sl, '写到的项目', '列明的项目'), 1)
            shape = next(s for s in sl.shapes if s.name.endswith(':body'))
            par = shape.text_frame.paragraphs[0]
            self.assertEqual(par.text, '清单没列明的项目，按现行程序办理。')
            self.assertEqual(par.runs[0]._r.find(par.runs[0]._r.tag.replace('}r', '}rPr')).get('lang'), 'zh-CN')
            self.assertEqual(replace_text(sl, '现行', '现行现行'), 1)      # 新文字含旧文字时不死循环
            self.assertEqual(par.text, '清单没列明的项目，按现行现行程序办理。')
            other = prs.slides[1]
            frame = notes_frame(prs, other)
            self.assertEqual(frame.text, '')
            frame.text = '本页讲稿'
            prs.save(path)
            again = Presentation(path)
            self.assertEqual(again.slides[0].notes_slide.notes_text_frame.text, '别页的备注')
            self.assertEqual(again.slides[1].notes_slide.notes_text_frame.text, '本页讲稿')

    def test_clean_copy(self):
        """副本清掉过程信息，讲稿、页面、隐藏页不动；工作稿不变；重名和已存在的副本拒绝。"""
        same = '这段讲稿讲的是另一页的内容，被复制到了很多页上面。'
        with TempDir() as tmp:
            deck = Deck(EXAMPLE_STYLE)
            x0, y0, x1, y1 = deck.style.body
            p = deck.page('evidence', '土地用途转换')
            p.text(x0, y0, 15, ['正文。'])
            p.notes('[改稿说明] 原标题：旧\n讲清转换比例。重点回应甲方的意见。约1分钟。\n待确认：比例口径')
            deck.page('evidence', '土地功能混合').text(x0, y0, 15, ['待确认的比例。'])
            p = deck.page('evidence', '备用页')
            p.text(x0, y0, 15, ['讨论稿内容。'])
            p.notes(same)
            p = deck.page('evidence', '另一页')
            p.text(x0, y0, 15, ['正文。'])
            p.notes(same)
            src = os.path.join(tmp, 'deck.pptx')
            deck.save(src)
            prs = Presentation(src)
            prs.slides[2]._element.set('show', '0')
            prs.save(src)
            out = os.path.join(tmp, 'copy.pptx')
            code, res = self.run_cli('clean', src, out)
            self.assertEqual(code, 0, res)
            copied = Presentation(out)
            self.assertEqual(len(copied.slides._sldIdLst), 4)
            self.assertEqual(res['hidden'], 1)
            self.assertEqual(copied.slides[0].notes_slide.notes_text_frame.text, '讲清转换比例。约1分钟。')
            self.assertEqual({(h['page'], h['word']) for h in res['pageHits']}, {(2, '待确认'), (3, '讨论稿')})
            self.assertEqual(Presentation(src).slides[0].notes_slide.notes_text_frame.text.count('[改稿说明]'), 1)
            code, res = self.run_cli('clean', src, out)
            self.assertEqual((code, res['error']), (1, 'CLEAN_REFUSED'))
            code, res = self.run_cli('clean', src, src, '--force')
            self.assertEqual((code, res['error']), (1, 'CLEAN_REFUSED'))
            code, res = self.run_cli('clean', src, out, '--force', '--drop-hidden', '--pattern', '约1分钟')
            self.assertEqual((code, res['slides'], res['hiddenDropped']), (0, 3, 1))
            copied = Presentation(out)
            self.assertEqual(len(copied.slides._sldIdLst), 3)
            self.assertEqual(copied.slides[0].notes_slide.notes_text_frame.text, '讲清转换比例。')

    def test_inspect_other_bold_marks_not_red(self):
        """整段、整栏用 __ 或 ^^ 加粗不算标红；同样的文字用 ** 才算。页面文字里没有记号字符。"""
        para = '各地已从单项试点转向系统推进，实施路径仍不完整，规划许可与土地供应的衔接尚未建立，监管规则相对薄弱。'
        with TempDir() as tmp:
            deck = Deck(EXAMPLE_STYLE)
            x0, y0, x1, y1 = deck.style.body
            p = deck.page('section-overview', '2.3 存量更新的实施机制')
            p.text(x0, y0, 15, [f'__{para}__', f'^^{para}^^'], slot='col', i=0)
            p.text(x0 + 16, y0, 15, ['__用途转换：一宗用地兼容两类以上功能__', '__容积率核算：转换部分不计入__',
                                     '^^审批程序：并联办理，限时办结^^'], slot='col', i=1)
            p.text(x0 + 32, y0, 15, [f'**{para}**'], slot='col', i=2)
            path = os.path.join(tmp, 'deck.pptx')
            deck.save(path)
            code, out = self.run_cli('inspect', path, '--style', EXAMPLE_STYLE, '--brief')
            texts = [sh.text_frame.text for sh in Presentation(path).slides[0].shapes if sh.has_text_frame]
        self.assertEqual(code, 0, out['problems'])
        self.assertEqual({(w['shape'], w['warning'].split('：')[0]) for w in out['warnings'] if w['shape'] != 'slide'},
                         {('pd:section-overview:col:2', '整段标红')})
        for text in texts:
            for mark in ('**', '__', '^^'):
                self.assertNotIn(mark, text)


class ConclusionTest(unittest.TestCase):
    TEXT = '社区公共服务设施覆盖率不到一半，存量空间有挖潜余地'

    def test_default_position_and_no_subtitle_warning(self):
        """文档自己的正例走 conclusion()：有名字、红色加粗，inspect 不按副标题警告。"""
        with TempDir() as tmp:
            deck = Deck(EXAMPLE_STYLE)
            p = deck.page('evidence', '2.1 现状用地', parent='二、现状评估')
            box = p.conclusion(self.TEXT)
            self.assertEqual(deck.warnings, [])
            self.assertEqual(box.shape.name, 'pd:evidence:conclusion')
            self.assertEqual((box.x, box.y), (1.6, 4.3))
            run = box.shape.text_frame.paragraphs[0].runs[0]
            self.assertEqual(run.font.size.pt, deck.style.size('conclusion'))
            self.assertTrue(run.font.bold)
            self.assertEqual('#' + str(run.font.color.rgb), deck.style.color('emphasis'))
            res = inspect(deck.save(os.path.join(tmp, 'd.pptx')), deck.style)
            self.assertEqual([w for w in res['warnings'] if w['shape'] == box.shape.name], [])

    def test_position_can_move_for_humanized_version(self):
        deck = Deck(EXAMPLE_STYLE)
        box = deck.page('evidence', '2.1 现状用地').conclusion(self.TEXT, x=2, y=26, w=40, h=1.6)
        self.assertEqual((box.x, box.y, box.w), (2, 26, 40))

    def test_too_long_warns(self):
        deck = Deck(EXAMPLE_STYLE)
        deck.page('evidence', '2.1 现状用地').conclusion(self.TEXT * 4)
        self.assertTrue(any('字约需' in w for w in deck.warnings), deck.warnings)

    def test_warns_when_subtitle_shares_the_line(self):
        deck = Deck(EXAMPLE_STYLE)
        deck.page('evidence', '2.1 现状用地', lead='假设情形').conclusion(self.TEXT)
        self.assertTrue(any('副标题和红字结论' in w for w in deck.warnings), deck.warnings)

    def test_missing_style_entries_say_what_to_write(self):
        doc = style_doc()
        doc['type'].pop('conclusion')
        with self.assertRaisesRegex(StyleError, 'type.conclusion'):
            Deck(Style(doc)).page('evidence', '2.1').conclusion(self.TEXT)
        doc = style_doc()
        doc.pop('conclusion')
        page = Deck(Style(doc)).page('evidence', '2.1')
        with self.assertRaisesRegex(StyleError, 'xCm'):
            page.conclusion(self.TEXT)
        self.assertEqual(page.conclusion(self.TEXT, x=2, y=5, w=40, h=1.6).y, 5)

    def test_style_checks_size_and_position(self):
        doc = style_doc()
        doc['type']['conclusion'] = 20      # 29.7 cm 高的页面宜 31—37 pt
        self.assertTrue(any('type.conclusion' in w for w in Style(doc).warnings))
        doc = style_doc()
        del doc['conclusion']['hCm']
        with self.assertRaisesRegex(StyleError, 'conclusion 缺少 hCm'):
            Style(doc)


if __name__ == '__main__':
    unittest.main()
