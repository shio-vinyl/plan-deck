"""通栏页眉预设、正文组件（标签行、栏头并列、通栏、箭头链、节点链、泳道、文件摘录、矩阵表、台阶）和测字。"""
import io
import json
import os
import unittest
from contextlib import redirect_stdout

from helpers import EXAMPLE_STYLE, TempDir, style_doc
from PIL import Image
from pptx import Presentation
from pptx.enum.dml import MSO_FILL
from pptx.enum.shapes import MSO_SHAPE

from plan_deck import Deck, Style
from plan_deck.cli import main
from plan_deck.measure import font, text_lines
from plan_deck.preset import build

CM = 360000


def banner_deck():
    deck = Deck(build('banner'))
    deck.chapter = '02  更新策略'
    return deck


def shapes(path, n=0):
    return {s.name: s for s in Presentation(path).slides[n].shapes}


def fill(shape):
    return str(shape.fill.fore_color.rgb)


def run_texts(shape):
    return [(r.text, r.font.size.pt, str(r.font.color.rgb), r.font.bold) for p in shape.text_frame.paragraphs for r in p.runs]


class PresetTest(unittest.TestCase):
    def test_banner_preset_is_clean(self):
        s = Style(build('banner'))
        self.assertEqual(s.warnings, [])
        self.assertEqual(s.info, [])
        self.assertEqual(s.size('key'), 18)
        self.assertEqual(s.size('conclusion'), 22)

    def test_scaled_to_other_page(self):
        doc = build('banner', page=(52.8, 29.7))
        s = Style(doc)
        self.assertEqual(s.warnings, [])
        self.assertEqual((s.page_w, s.page_h), (52.8, 29.7))
        self.assertEqual(s.size('body'), 18.5)                 # 12 pt 按页高等比，取到 0.5
        bar = doc['header']['bars'][0]
        self.assertEqual(bar['wCm'], 'page')
        self.assertAlmostEqual(doc['body']['leftCm'], round(1.57 * 52.8 / 33.867, 2))
        self.assertAlmostEqual(doc['body']['topCm'], round(4.87 * 29.7 / 19.05, 2))

    def test_style_init_cli(self):
        with TempDir() as tmp:
            out = os.path.join(tmp, 'style.json')
            buf = io.StringIO()
            with redirect_stdout(buf):
                code = main(['style', 'init', out, '--preset', 'banner', '--page', '52.8x29.7', '--name', '江屿市示例'])
            self.assertEqual(code, 0, buf.getvalue())
            with open(out, encoding='utf-8') as f:
                self.assertEqual(json.load(f)['name'], '江屿市示例')
            with redirect_stdout(io.StringIO()) as buf2:
                code = main(['style', 'init', out, '--preset', 'banner'])
            self.assertEqual(code, 1)
            self.assertIn('不覆盖', buf2.getvalue())


class HeaderTest(unittest.TestCase):
    def test_chapter_line_accent_and_page_number(self):
        with TempDir() as tmp:
            deck = banner_deck()
            deck.page('rows', '2.1.3 土地整备流程', parent='2.1 实施路径')
            deck.page('rows', '总览', chapter='')        # 没有上级标题：竖色条不画；章名行可以单页关掉
            path = deck.save(os.path.join(tmp, 'd.pptx'))
            got = shapes(path)
            self.assertEqual(got['pd:rows:chapter'].text_frame.text, '02  更新策略')
            self.assertEqual(run_texts(got['pd:rows:chapter'])[0][1:], (12.0, '4E7491', True))
            self.assertEqual(run_texts(got['pd:rows:parent-title'])[0][1], 26.0)
            self.assertEqual(run_texts(got['pd:rows:title'])[0][1], 18.0)
            self.assertIn('pd:rows:header-accent', got)
            fld = got['pd:rows:page-no']._element.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}fld')
            self.assertEqual([f.get('type') for f in fld], ['slidenum'])
            second = shapes(path, 1)
            self.assertNotIn('pd:rows:header-accent', second)
            self.assertNotIn('pd:rows:chapter', second)

    def test_theme_effects_have_no_shadow(self):
        """新建稿子的主题效果样式清空：直接 add_shape 或在 PowerPoint 里新建的形状不带阴影。"""
        with TempDir() as tmp:
            path = banner_deck().save(os.path.join(tmp, 'd.pptx'))
            master = Presentation(path).slide_master
            theme = next(r.target_part.blob for r in master.part.rels.values() if r.reltype.endswith('/theme'))
        self.assertIn(b'effectStyleLst', theme)
        self.assertNotIn(b'outerShdw', theme)

    def test_lead_without_lead_line_warns(self):
        deck = banner_deck()
        deck.page('rows', '2.1.3', parent='2.1', lead='副标题')
        self.assertTrue(any('没有副标题行' in w for w in deck.warnings), deck.warnings)


class BlocksTest(unittest.TestCase):
    def setUp(self):
        self.deck = banner_deck()
        self.st = self.deck.style
        self.p = self.deck.page('blk', '2.1.3 土地整备流程', parent='2.1 实施路径')
        self.deck.warnings.clear()

    def save(self, tmp):
        return self.deck.save(os.path.join(tmp, 'd.pptx'))

    def test_rows_fill_to_bottom_and_colors(self):
        x0, y0, x1, y1 = self.st.body
        box = self.p.rows(x0, y0, 15, [('申请与\n权属核实', '核实权属，出具**测绘报告**。'), ('方案审批', '报市土地管理委员会批准。', ('类别', '三'))])
        self.assertAlmostEqual(box.bottom, y1, places=2)
        self.assertEqual(self.deck.warnings, [])
        with TempDir() as tmp:
            got = shapes(self.save(tmp))
        self.assertEqual(fill(got['pd:blk:row-label:0']), '194461')
        self.assertEqual(fill(got['pd:blk:row-label:1']), 'A8895A')
        self.assertEqual(fill(got['pd:blk:row-body:0']), 'EEF2F5')
        label = run_texts(got['pd:blk:row-label:0'])
        self.assertEqual(label[0][1:], (18.0, 'FFFFFF', True))
        body = run_texts(got['pd:blk:row-body:0'])
        self.assertIn(('测绘报告', 12.0, 'C00000', True), body)
        # 两行等分多出的高度
        h0, h1 = got['pd:blk:row-label:0'].height, got['pd:blk:row-label:1'].height
        self.assertGreater(h0, h1)

    def test_rows_overflow_warns(self):
        x0, y0, *_ = self.st.body
        self.p.rows(x0, y0, 8, [('甲', '很长的正文。' * 120), ('乙', '正文')])
        self.assertTrue(any('row' in w and '不要缩字号' in w for w in self.deck.warnings), self.deck.warnings)

    def test_rows_plain_and_no_fill(self):
        x0, y0, *_ = self.st.body
        box = self.p.rows(x0, y0, 15, [('甲', '正文'), ('乙', '正文')], body='plain', bottom=False)
        self.assertLess(box.bottom, y0 + 3)
        with TempDir() as tmp:
            got = shapes(self.save(tmp))
        self.assertIn('pd:blk:row-rule', got)
        self.assertNotEqual(got['pd:blk:row-body:0'].fill.type, MSO_FILL.SOLID)

    def test_columns_banner_arrow(self):
        x0, y0, x1, y1 = self.st.body
        cb = self.p.columns(x0, y0, x1 - x0, [('文件到期', '**城市更新实施意见**：已到期。'), ('规则偏窄', '面向立项项目。'),
                                             ('逐案决定', '靠控规修改。')], head_size='figure', subs=['基础文件', '覆盖不足', ''],
                            colors='类别', bottom=y1 - 2)
        self.assertAlmostEqual(cb.bottom, y1 - 2, places=2)
        self.p.block_arrow(10, cb.bottom, 2, 0.6)
        bb = self.p.banner(x0, cb.bottom + 0.7, x1 - x0, '需要一套事先明确的规则')
        self.assertEqual(self.deck.warnings, [])
        with TempDir() as tmp:
            got = shapes(self.save(tmp))
        heads = [got[f'pd:blk:col-head:{k}'] for k in range(3)]
        self.assertEqual([fill(h) for h in heads], ['194461', '6C917A', 'A8895A'])     # 方案名：逐项取类别色
        self.assertEqual(run_texts(heads[0])[0][1], 30.0)
        body = got['pd:blk:col-body:0'].text_frame.paragraphs
        self.assertEqual(body[0].runs[0].text, '基础文件')
        self.assertEqual(str(body[0].runs[0].font.color.rgb), '194461')
        self.assertEqual(got['pd:blk:block-arrow'].auto_shape_type, MSO_SHAPE.DOWN_ARROW)
        self.assertEqual(fill(got['pd:blk:banner']), '194461')
        self.assertGreater(bb.h, 0.5)

    def test_chevrons_and_flow(self):
        x0, y0, x1, _ = self.st.body
        self.p.chevrons(x0, y0, x1 - x0, 1.2, ['征求意见', '公示', ('区政府审定', 'emphasis')])
        fb = self.p.flow(x0, y0 + 2, x1 - x0, ['申请', ('签订协议', ('类别', '一')), ('一般环节', 'fill')],
                         notes=[None, '土地处置方式', None])
        self.assertGreater(fb.h, 1.0)
        with TempDir() as tmp:
            got = shapes(self.save(tmp))
        self.assertEqual(got['pd:blk:chevron:0'].auto_shape_type, MSO_SHAPE.PENTAGON)
        self.assertEqual(got['pd:blk:chevron:2'].auto_shape_type, MSO_SHAPE.CHEVRON)
        self.assertEqual(fill(got['pd:blk:chevron:2']), 'C00000')
        self.assertIn('pd:blk:flow-arrow:1', got)
        self.assertEqual(got['pd:blk:flow-arrow:0'].line.width.pt, 3)
        # 浅底上的字自动用正文色
        self.assertEqual(run_texts(got['pd:blk:flow:2'])[0][2], '333333')
        self.assertEqual(got['pd:blk:flow-note'].text_frame.text, '土地处置方式')

    def test_lanes(self):
        x0, y0, x1, y1 = self.st.body
        boxes = self.p.lanes(x0, y0, x1 - x0, y1 - y0, [('属地', ('类别', '一')), ('规资部门', ('类别', '五'), 2), ('区政府', 'emphasis')],
                             ['收集', '受理', '会商', '审定'],
                             [(0, 0, '收集诉求'), (1, 1, '核对主体'), (1, (2, 3), '组织会商'), (2, 3, '审定', 'emphasis'),
                              (0, 2, '配合', 'plain')],
                             links=[(0, 1), (1, 2), (2, 3)])
        self.assertEqual(len(boxes), 5)
        self.assertGreater(boxes[2].w, boxes[1].w * 1.9)
        with TempDir() as tmp:
            path = self.save(tmp)
            prs = Presentation(path)
            got = {s.name: s for s in prs.slides[0].shapes}
        links = [s for n, s in got.items() if n.startswith('pd:blk:lane-link')]
        self.assertEqual(len(links), 3)
        glued = [s for s in links if s._element.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}stCxn') is not None]
        self.assertEqual(len(glued), 2)          # 竖线不粘，另两条端点粘在节点上
        self.assertEqual(fill(got['pd:blk:lane-head:2']), 'C00000')

    def test_excerpt(self):
        with TempDir() as tmp:
            img = os.path.join(tmp, 'doc.png')
            Image.new('RGB', (1000, 2000), 'white').save(img)
            x0, y0, *_ = self.st.body
            box = self.p.excerpt(img, x0, y0, 10, 12, crop=(0, 0.25, 1, 0.75), marks=[(0.1, 0.4, 0.9, 0.5), (0, 0.9, 1, 1)],
                                 caption='《江屿市土地整备管理办法》摘录')
            self.assertTrue(any('裁剪范围外' in w for w in self.deck.warnings))
            got = shapes(self.save(tmp))
        pic = got['pd:blk:excerpt']
        self.assertAlmostEqual(pic.crop_top, 0.25)
        self.assertAlmostEqual(pic.crop_bottom, 0.25)
        self.assertAlmostEqual(pic.width / pic.height, 1.0, places=2)      # 裁成 1000×1000
        mark = got['pd:blk:excerpt-mark']
        self.assertEqual(str(mark.line.color.rgb), 'C00000')
        self.assertAlmostEqual((mark.top - pic.top) / pic.height, 0.3, places=2)
        self.assertLessEqual(box.h, 12.01)

    def test_matrix_fits_and_merges(self):
        x0, y0, x1, y1 = self.st.body
        rows = [['类别', '详细规划', '', '专项规划'], ['', '功能区', '非功能区', ''],
                ['编制部门', '管委会', '规资局', '规资局'], ['审批', '区政府', '区政府', '区政府；重大规划报市政府']]
        box = self.p.matrix(x0, y0, x1 - x0, rows, [1, 1.5, 1.5, 2], header_rows=2,
                            merges=[(0, 0, 1, 0), (0, 1, 0, 2), (0, 3, 1, 3)], row_heads=True, highlight=[3])
        self.assertAlmostEqual(box.bottom, min(y1, box.bottom), places=2)
        self.assertEqual(self.deck.fitted[0]['pt'], 15.0)            # 字少：取最大一档
        with TempDir() as tmp:
            got = shapes(self.save(tmp))
        t = got['pd:blk:table'].table
        self.assertTrue(t.cell(0, 1).is_merge_origin)
        self.assertEqual(str(t.cell(0, 1).fill.fore_color.rgb), '194461')
        self.assertEqual(str(t.cell(2, 0).fill.fore_color.rgb), '194461')        # 行头一列实色
        self.assertEqual(str(t.cell(3, 1).fill.fore_color.rgb), 'F6E4E2')        # 高亮行
        r = t.cell(3, 1).text_frame.paragraphs[0].runs[0]
        self.assertEqual((r.font.size.pt, str(r.font.color.rgb), r.font.bold), (15.0, 'C00000', True))

    def test_matrix_long_table_steps_down(self):
        x0, y0, x1, y1 = self.st.body
        rows = [['名称', '说明']] + [[f'条目 {i}', '一段比较长的说明文字，' * 3] for i in range(14)]
        self.p.matrix(x0, y0, x1 - x0, rows, [1, 4])
        self.assertLess(self.deck.fitted[0]['pt'], 15)

    def test_steps(self):
        x0, y0, x1, _ = self.st.body
        box = self.p.steps(x0, y0, x1 - x0, [('未触及', '按比例办理'), ('达到预警线', '转清单认定'), ('失衡', '**暂停**办理')],
                           colors='分档', h=6)
        self.assertLessEqual(box.h, 6.05)
        with TempDir() as tmp:
            got = shapes(self.save(tmp))
        tops = [got[f'pd:blk:step-block:{k}'].top for k in range(3)]
        self.assertEqual(tops, sorted(tops, reverse=True))        # 逐级升高
        self.assertEqual(fill(got['pd:blk:step-block:2']), '8B2E25')

    def test_same_color_two_meanings_warns(self):
        doc = build('banner')
        doc['classes']['阶段'] = {'近期': '{palette.navy}', '远期': '{palette.olive}'}
        st = Style(doc)
        self.assertTrue(any('几组类别色里都用' in i for i in st.info))
        deck = Deck(st)
        p = deck.page('a', '1')
        p.rows(2, 5, 10, [('甲', '正文', ('类别', '一'))], bottom=False)
        q = deck.page('a', '2')
        q.rows(2, 5, 10, [('乙', '正文', ('阶段', '近期'))], bottom=False)
        deck.page('a', '3')
        r = deck.page('a', '4')
        r.rows(2, 5, 10, [('丙', '正文', ('类别', '一'))], bottom=False)
        with TempDir() as tmp:
            deck.save(os.path.join(tmp, 'd.pptx'))
        hits = [w for w in deck.warnings if '同时表示' in w]
        self.assertEqual(len(hits), 1, deck.warnings)
        self.assertIn('第 1、2 页', hits[0])

    def test_components_on_example_style(self):
        """没有 key、tint 的旧 style.json 也能用：字号退到 heading，浅底用 fill。"""
        deck = Deck(EXAMPLE_STYLE)
        p = deck.page('a', '1')
        p.rows(2, 7, 30, [('甲', '正文')], bottom=False)
        p.columns(2, 12, 30, [('甲', '正文'), ('乙', '正文')], bottom=False)
        self.assertEqual(deck.warnings, [])


class NamesAndRatiosTest(unittest.TestCase):
    def test_hue_named_colors_warn(self):
        doc = style_doc()
        doc['colors']['navy'] = '#194461'
        doc['colors']['lightBlue'] = '#DEEBF7'
        doc['colors']['redLine'] = '#C00000'
        doc['classes']['更新方式']['深蓝'] = '#194461'
        w = ' '.join(Style(doc).warnings)
        self.assertIn('colors.navy', w)
        self.assertIn('colors.lightBlue', w)
        self.assertIn('classes.更新方式.深蓝', w)
        self.assertNotIn('redLine', w)
        self.assertNotIn('colors.white', w)

    def test_ratio_rule_above_heading(self):
        doc = build('banner')
        doc['type'] = {'body': 12, 'heading': 15, 'title': 18, 'conclusion': 22, 'parent': 26}
        st = Style(doc)
        self.assertFalse(any('只差' in w for w in st.warnings), st.warnings)      # 22→26 是 1.18 倍，标题以上放宽到 1.15
        doc['type']['heading'] = 14
        self.assertTrue(any('12 与 14' in w for w in Style(doc).warnings))


class MeasureTest(unittest.TestCase):
    def test_estimate_when_font_missing(self):
        self.assertEqual(text_lines('汉字' * 10, 3, 12, family='没有这种字体'), text_lines('汉字' * 10, 3, 12))

    @unittest.skipUnless(font('微软雅黑'), '本机没有微软雅黑')
    def test_real_width_and_no_punctuation_at_line_start(self):
        # 一行正好放下 10 个字时，第 11 个字是句号：句号不落行首，带上一个字下行，仍是两行
        width = 10 * 12 * 2.54 / 72 / 0.98 + 0.15 + 0.01
        self.assertEqual(text_lines('一二三四五六七八九十', width, 12, '微软雅黑'), 1)
        self.assertEqual(text_lines('一二三四五六七八九十。', width, 12, '微软雅黑'), 2)
        # 数字和字母比估算的 0.6 个字宽窄
        self.assertLessEqual(text_lines('2025' * 30, 5, 12, '微软雅黑'), text_lines('2025' * 30, 5, 12))


if __name__ == '__main__':
    unittest.main()
