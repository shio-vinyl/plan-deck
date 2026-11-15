"""inspect 的版面核对、满版图组件、图表和表格、默认模板、Deck.open。"""
import os
import unittest
import zipfile

from helpers import EXAMPLE_STYLE, TempDir, style_doc
from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.enum.chart import XL_CHART_TYPE
from pptx.oxml.ns import qn
from pptx.util import Cm, Pt

from plan_deck import Deck, Style, StyleError
from plan_deck.inspect import inspect
from plan_deck.render import missing_fonts


def photo(root, size=(2400, 1000)):
    path = os.path.join(root, 'photo.png')
    Image.new('RGB', size, (120, 150, 170)).save(path)
    return path


def warned(res, text):
    return [(w['page'], w['shape']) for w in res['warnings'] if text in w['warning']]


class InspectLayoutTest(unittest.TestCase):
    def test_full_page_image_hides_header(self):
        with TempDir() as root:
            deck = Deck(EXAMPLE_STYLE)
            st = deck.style
            p = deck.page('plan-map', '2.1 总图', parent='二、规划')
            p.picture(photo(root), 0, 0, st.page_w, st.page_h, fit='cover')
            q = deck.page('plan-map', '2.2 总图', header=False)     # 放图后再画页眉：不报
            q.picture(photo(root), 0, 0, st.page_w, st.page_h, fit='cover', slot='base')
            q.header()
            res = inspect(deck.save(os.path.join(root, 'd.pptx')))
            self.assertEqual(warned(res, '盖住'), [(1, 'pd:plan-map:parent-title'), (1, 'pd:plan-map:title')])

    def test_translucent_and_band_do_not_cover(self):
        with TempDir() as root:
            deck = Deck(EXAMPLE_STYLE)
            p = deck.page('evidence', '2.1 现状')
            p.shape('rect', 0, 0, 52.8, 29.7, 'title', 'scrim', alpha=0.4)
            p.band(0, 0, 52.8, 29.7, 'title')
            self.assertEqual(warned(inspect(deck.save(os.path.join(root, 'd.pptx'))), '盖住'), [])

    def test_text_over_filled_shape_is_fine(self):
        with TempDir() as root:
            deck = Deck(EXAMPLE_STYLE)
            p = deck.page('evidence', '2.1 现状')
            p.shape('rect', 2, 7, 20, 5, 'fill', 'bg')
            p.text(2, 7, 20, ['分组底上的字。'])
            p.node(30, 7, 8, 4, '节点\n说明', 'title')
            self.assertEqual(inspect(deck.save(os.path.join(root, 'd.pptx')))['warnings'], [])

    def test_overflow_and_partly_off_page(self):
        with TempDir() as root:
            deck = Deck(EXAMPLE_STYLE)
            p = deck.page('evidence', '2.1 现状')
            p.label(2, 8, 5, 1, '放不下的长标签文字要折成好几行才能排完', 'short')
            p.label(45, 8, 20, 2, '右边出界', 'right')
            p.text(2, 28, 10, ['底下出界的一段正文。' * 5], h=1, slot='low')
            p.ledger(2, 12, [(1.0, '甲', 'title'), '+', (2.0, '乙', 'title')], 5, 3)
            res = inspect(deck.save(os.path.join(root, 'd.pptx')))
            self.assertIn((1, 'pd:evidence:short'), warned(res, '框只有'))
            self.assertEqual({s for _, s in warned(res, '出了页面')}, {'pd:evidence:right', 'pd:evidence:low'})
            self.assertFalse([s for _, s in warned(res, '框只有') if 'ledger' in s])       # 运算符不算

    def test_mixed_sizes_not_overestimated(self):
        with TempDir() as root:
            deck = Deck(EXAMPLE_STYLE)
            p = deck.page('evidence', '2.1 现状')
            box = p.label(2, 8, 9, 2.4, '1.85', 'kpi', size='figure')
            r = box.shape.text_frame.paragraphs[0].add_run()
            r.text, r.font.size = ' 万人/平方公里', Pt(18)
            self.assertEqual(warned(inspect(deck.save(os.path.join(root, 'd.pptx'))), '框只有'), [])

    def test_shapes_inside_groups(self):
        with TempDir() as root:
            deck = Deck(EXAMPLE_STYLE)
            p = deck.page('evidence', '2.1 现状')
            grp = p.slide.shapes.add_group_shape()
            tb = grp.shapes.add_textbox(Cm(2), Cm(8), Cm(10), Cm(1))
            tb.name = '手工组内'
            r = tb.text_frame.paragraphs[0].add_run()
            r.text, r.font.size = '本页说明', Pt(8)
            # 组缩小一半：组内坐标不是页面坐标
            xf = grp._element.grpSpPr.find(qn('a:xfrm'))
            xf.find(qn('a:off')).set('x', str(Cm(1)))
            xf.find(qn('a:off')).set('y', str(Cm(4)))
            xf.find(qn('a:ext')).set('cx', str(Cm(5)))
            xf.find(qn('a:ext')).set('cy', str(Cm(0.5)))
            res = inspect(deck.save(os.path.join(root, 'd.pptx')), deck.style)
            self.assertTrue([q for q in res['problems'] if q['shape'] == '手工组内' and '8 pt' in q['problem']])
            self.assertTrue(warned(res, '讲解员口吻'))
            item = next(s for s in res['pages'][0]['items'] if s['name'] == '手工组内')
            self.assertEqual(item['cm'], [1.0, 4.0, 5.0, 0.5])
            self.assertEqual(item['group'], grp.name)

    def test_group_without_transform(self):
        with TempDir() as root:
            deck = Deck(EXAMPLE_STYLE)
            grp = deck.page('evidence', '2.1 现状').slide.shapes.add_group_shape()
            grp.shapes.add_textbox(Cm(2), Cm(8), Cm(10), Cm(1)).name = '组内'
            grp._element.grpSpPr.remove(grp._element.grpSpPr.find(qn('a:xfrm')))
            res = inspect(deck.save(os.path.join(root, 'd.pptx')))
            self.assertEqual(next(s for s in res['pages'][0]['items'] if s['name'] == '组内')['cm'], [2.0, 8.0, 10.0, 1.0])

    def test_dangling_animation_reference(self):
        timing = ('<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:tnLst><p:par>'
                  '<p:cTn id="1" nodeType="tmRoot"><p:childTnLst><p:set><p:cBhvr><p:cTn id="2"/><p:tgtEl><p:spTgt spid="{}"/></p:tgtEl>'
                  '</p:cBhvr></p:set></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>')
        with TempDir() as root:
            deck = Deck(EXAMPLE_STYLE)
            p = deck.page('evidence', '2.1 现状')
            good = p.label(2, 8, 10, 1.2, '有动画', 'a')
            p.slide._element.append(etree.fromstring(timing.format(good.shape.shape_id)))
            q = deck.page('evidence', '2.2 现状')
            q.slide._element.append(etree.fromstring(timing.format(9999)))
            res = inspect(deck.save(os.path.join(root, 'd.pptx')))
            self.assertEqual([(x['page'], x['shape']) for x in res['problems']], [(2, 'id 9999')])


class FullImageTest(unittest.TestCase):
    def test_page_without_header(self):
        deck = Deck(EXAMPLE_STYLE)
        p = deck.page('chapter-summary', header=False)
        self.assertEqual(len(p.slide.shapes), 0)
        q = deck.page('evidence', '', header=True)      # 空标题：不留一条孤零零的页眉线
        self.assertEqual([s.name for s in q.slide.shapes], [])

    def test_late_header_on_top(self):
        with TempDir() as root:
            deck = Deck(EXAMPLE_STYLE)
            p = deck.page('evidence', '2.1 总图', parent='二、规划', header=False)
            p.picture(photo(root), 0, 0, 52.8, 29.7, fit='cover')
            p.header()
            names = [s.name for s in p.slide.shapes]
            self.assertEqual(names[0], 'pd:evidence:picture')
            self.assertIn('pd:evidence:title', names)
            self.assertIn('pd:evidence:header-rule', names)

    def test_cover_crops_contain_does_not(self):
        with TempDir() as root:
            p = Deck(EXAMPLE_STYLE).page('x', 't')
            box = p.picture(photo(root), 0, 0, 52.8, 29.7, fit='cover')
            self.assertEqual((box.x, box.y, box.w, box.h), (0, 0, 52.8, 29.7))
            self.assertEqual(box.shape.width, Cm(52.8))
            # 2400×1000 铺进 16:9：左右各裁掉一部分，上下不裁
            keep = (52.8 / 29.7) / 2.4
            self.assertAlmostEqual(box.shape.crop_left, (1 - keep) / 2, 4)
            self.assertAlmostEqual(box.shape.crop_right, (1 - keep) / 2, 4)
            self.assertAlmostEqual(box.shape.crop_top, 0, 6)
            plain = p.picture(photo(root), 0, 0, 52.8, 29.7, slot='p2')
            self.assertEqual(plain.shape.crop_left, 0)
            self.assertLess(plain.h, 29.7)
            with self.assertRaisesRegex(ValueError, 'contain、cover'):
                p.picture(photo(root), 0, 0, 5, 5, slot='p3', fit='fill')

    def test_band_and_shadow_xml(self):
        p = Deck(EXAMPLE_STYLE).page('chapter-summary', header=False)
        s = p.band(0, 10, 52.8, 9, 'title')
        xml = s._element.spPr.xml
        self.assertEqual(s._element.spPr.find(qn('a:solidFill')), None)
        self.assertIn('<a:alpha val="45000"/>', xml)
        self.assertEqual(xml.count('<a:alpha val="0"/>'), 2)
        self.assertEqual([c.tag.split('}')[1] for c in s._element.spPr][:3], ['xfrm', 'prstGeom', 'gradFill'])
        box = p.label(5, 12, 40, 3, '结论句', 'conclusion', size='figure', color='white', bold=True, shadow=True)
        rpr = box.shape.text_frame.paragraphs[0].runs[0]._r.rPr
        self.assertEqual([c.tag.split('}')[1] for c in rpr], ['solidFill', 'effectLst', 'latin', 'ea'])
        plain = p.label(5, 16, 40, 1.2, '副句', 'sub', color='white')
        self.assertNotIn('effectLst', plain.shape._element.xml)


class ChartTableTest(unittest.TestCase):
    def setUp(self):
        self.deck = Deck(EXAMPLE_STYLE)
        self.page = self.deck.page('forecast', '3.1 预测')

    def xml(self, ch):
        return ch.part.blob.decode('utf-8')

    def test_chart_text_uses_style_color(self):
        ch = self.page.chart('column', 2, 7, 20, 10, ['甲', '乙'], [('a', [1.0, 2.0]), ('b', [3.0, 4.0])], ['title', 'muted'], labels=True)
        for rpr in etree.fromstring(ch.part.blob).iter(qn('a:defRPr')):
            kids = [c.tag.split('}')[1] for c in rpr]
            self.assertEqual(kids[0], 'solidFill')          # 顺序错了 PowerPoint 会报修复
            self.assertEqual(rpr[0][0].get('val'), '262626')
            self.assertLess(kids.index('latin'), kids.index('ea'))

    def test_old_kinds_keep_their_auto_title(self):
        """柱形、条形的默认外观不变：单系列的自动标题要不要去掉，等用户定。"""
        ch = self.page.chart('column', 2, 7, 20, 10, ['甲'], [('a', [1.0])], ['title'])
        self.assertNotIn('autoTitleDeleted val="1"', self.xml(ch))

    def test_chart_text_color_from_style_json(self):
        doc = style_doc()
        doc['chart'] = {'text': 'white'}
        deck = Deck(Style(doc))
        ch = deck.page('x', 't').chart('bar', 2, 7, 20, 10, ['甲'], [('a', [1.0])], ['title'])
        self.assertIn('FFFFFF', self.xml(ch))

    def test_line_series_color(self):
        ch = self.page.chart('line', 2, 7, 20, 10, ['2021', '2022'], [('a', [1.0, 2.0])], ['emphasis'])
        ser = etree.fromstring(ch.part.blob).find('.//' + qn('c:ser'))
        self.assertEqual(ser.find(qn('c:spPr')).find(qn('a:ln')).find('.//' + qn('a:srgbClr')).get('val'), 'C00000')
        self.assertEqual(ser.find(qn('c:spPr')).find(qn('a:solidFill')), None)
        self.assertEqual(ser.find(qn('c:spPr')).find(qn('a:ln')).get('w'), str(Pt(2.25)))
        self.assertFalse(ch.has_title)

    def test_area_and_round_kinds(self):
        self.page.chart('area', 2, 7, 20, 10, ['2021', '2022'], [('a', [1.0, 2.0])], ['title'])
        n = len(self.page.slide.shapes)
        ch = self.page.chart('doughnut', 24, 7, 10, 10, ['完成', '未完成'], [('进度', [62.0, 38.0])], ['emphasis', 'line'], slot='ring')
        self.assertEqual(len(self.page.slide.shapes), n + 1)
        self.assertTrue(ch.has_legend)
        self.assertEqual(self.xml(ch).count('<c:dPt>'), 2)
        self.assertFalse(ch.has_title)
        labeled = self.page.chart('doughnut', 24, 7, 10, 10, ['甲', '乙'], [('s', [1.0, 2.0])], ['title', 'line'], slot='r2', labels=True)
        self.assertIn('<c:showVal val="1"/>', self.xml(labeled))
        self.page.chart('pie', 36, 7, 10, 10, ['甲', '乙'], [('s', [1.0, 2.0])], ['title', 'muted'], labels=True, slot='pie')

    def test_unsupported_kinds_leave_nothing(self):
        n = len(self.page.slide.shapes)
        for kind, text in (('radar-x', '不认识'), (XL_CHART_TYPE.XY_SCATTER, '不支持')):
            with self.assertRaisesRegex(ValueError, text):
                self.page.chart(kind, 2, 7, 20, 10, ['甲'], [('a', [1.0])], ['title'])
        with self.assertRaisesRegex(ValueError, '只能有 1 个系列'):
            self.page.chart('pie', 2, 7, 20, 10, ['甲'], [('a', [1.0]), ('b', [2.0])], ['title'])
        with self.assertRaisesRegex(ValueError, '2 个类别给了 1 个颜色'):
            self.page.chart('pie', 2, 7, 20, 10, ['甲', '乙'], [('a', [1.0, 2.0])], ['title'])
        with self.assertRaisesRegex(ValueError, '颜色'):
            self.page.chart('bar', 2, 7, 20, 10, ['甲'], [('a', [1.0])], ['title', 'muted'])
        with self.assertRaisesRegex(StyleError, '没有'):
            self.page.chart('line', 2, 7, 20, 10, ['甲'], [('a', [1.0])], ['不存在的颜色'])
        self.assertEqual(len(self.page.slide.shapes), n)    # 不留半成品

    def test_table_rules(self):
        box = self.page.table(2, 7, 20, [['a', 'b'], ['甲', 1.0], ['合计', 1.0]], [1, 1], total=True)
        tb = box.shape.table
        def sides(i, j):
            pr = tb.cell(i, j)._tc.tcPr
            return [(c.tag.split('}')[1], c[0].tag.split('}')[1]) for c in pr][:5]
        self.assertEqual(sides(0, 0), [('lnL', 'noFill'), ('lnR', 'noFill'), ('lnT', 'noFill'), ('lnB', 'solidFill'), ('solidFill', 'srgbClr')])
        self.assertEqual(sides(1, 1)[2:4], [('lnT', 'solidFill'), ('lnB', 'solidFill')])
        lnb = tb.cell(2, 0)._tc.tcPr.find(qn('a:lnB'))
        self.assertEqual(lnb.get('w'), str(Pt(0.75)))
        self.assertEqual(lnb.find('.//' + qn('a:srgbClr')).get('val'), 'BFBFBF')


class TemplateTest(unittest.TestCase):
    def test_theme_and_layouts_follow_style(self):
        with TempDir() as root:
            deck = Deck(EXAMPLE_STYLE)
            deck.page('x', 't')
            path = deck.save(os.path.join(root, 'd.pptx'))
            with zipfile.ZipFile(path) as z:
                theme = etree.fromstring(z.read('ppt/theme/theme1.xml'))
            for kind in ('a:majorFont', 'a:minorFont'):
                font = theme.find(f'.//{qn(kind)}')
                self.assertEqual(font.find(qn('a:latin')).get('typeface'), 'Arial')
                self.assertEqual(font.find(qn('a:ea')).get('typeface'), '微软雅黑')
                hans = next(f for f in font.findall(qn('a:font')) if f.get('script') == 'Hans')
                self.assertEqual(hans.get('typeface'), '微软雅黑')     # PowerPoint 的中文优先取这里
            scheme = theme.find(f'.//{qn("a:clrScheme")}')
            self.assertEqual(scheme.find(qn('a:dk1'))[0].get('val'), '262626')
            self.assertEqual(scheme.find(qn('a:lt1'))[0].get('val'), 'FFFFFF')
            self.assertEqual(scheme.find(qn('a:accent2'))[0].get('val'), 'C00000')
            prs = Presentation(path)
            W, H = prs.slide_width, prs.slide_height
            for part in (prs.slide_master, *prs.slide_layouts):
                for sh in part.placeholders:
                    self.assertLessEqual(sh.left + sh.width, W + 2, sh.name)
                    self.assertLessEqual(sh.top + sh.height, H + 2, sh.name)
            # 版式占位符撑到新页面的宽度，不停在 4:3 的位置
            title = prs.slide_master.placeholders[0]
            self.assertGreater(title.left + title.width, W * 0.9)


class OpenTest(unittest.TestCase):
    def build(self, root):
        deck = Deck(EXAMPLE_STYLE)
        p = deck.page('evidence', '2.1 现状')
        p.text(2, 7, 20, ['一段'], slot='note')
        path = deck.save(os.path.join(root, 'd.pptx'))
        return path

    def test_add_to_existing_page_keeps_names_unique(self):
        with TempDir() as root:
            deck = Deck.open(self.build(root), EXAMPLE_STYLE)
            p = deck.at(1)
            self.assertEqual(p.type, 'evidence')
            p.text(2, 12, 20, ['又一段'], slot='note')             # 已有一个 note：第一个改名 :0，新的是 :1
            names = [s.name for s in p.slide.shapes]
            self.assertIn('pd:evidence:note:0', names)
            self.assertIn('pd:evidence:note:1', names)
            self.assertNotIn('pd:evidence:note', names)
            with self.assertRaisesRegex(ValueError, '重名'):
                p.shape('rect', 0, 0, 1, 1, 'title', 'note', i=1)
            out = deck.save(os.path.join(root, 'd2.pptx'))
            self.assertEqual(inspect(out)['problems'], [])
            self.assertEqual(len(Presentation(out).slides), 1)

    def test_handmade_names_count_and_new_page(self):
        with TempDir() as root:
            path = self.build(root)
            prs = Presentation(path)
            tb = prs.slides[0].shapes.add_textbox(Cm(1), Cm(1), Cm(1), Cm(1))
            tb.name = 'pd:evidence:hand:0'
            prs.slides.add_slide(prs.slide_layouts[0])
            prs.save(path)
            deck = Deck.open(path, EXAMPLE_STYLE)
            with self.assertRaisesRegex(ValueError, '重名'):
                deck.at(1).shape('rect', 0, 0, 1, 1, 'title', 'hand', i=0)
            with self.assertRaisesRegex(ValueError, '认不出页型'):
                deck.at(2)
            self.assertEqual(deck.at(2, 'cover').type, 'cover')
            with self.assertRaisesRegex(ValueError, '共 2 页'):
                deck.at(3)
            q = deck.page('evidence', '2.2 新页')
            self.assertEqual(q.number, 3)
            self.assertFalse(list(q.slide.placeholders))

    def test_open_keeps_template_and_warns_on_size(self):
        with TempDir() as root:
            prs = Presentation()         # 同事的 4:3 稿
            prs.slides.add_slide(prs.slide_layouts[6])
            path = os.path.join(root, 'old.pptx')
            prs.save(path)
            with zipfile.ZipFile(path) as z:
                before = z.read('ppt/theme/theme1.xml')
            deck = Deck.open(path, EXAMPLE_STYLE)
            self.assertEqual(deck.prs.slide_width, prs.slide_width)
            self.assertIn('style.json 是 52.8×29.7 cm', deck.warnings[0])
            out = deck.save(os.path.join(root, 'new.pptx'))
            with zipfile.ZipFile(out) as z:
                self.assertEqual(z.read('ppt/theme/theme1.xml'), before)


class FontNamesTest(unittest.TestCase):
    def test_serif_swap_is_reported(self):
        got = missing_fonts({'cjk': 'Source Han Sans CN', 'latin': 'Arial'}, ['SourceHanSerifCN-Bold', 'ArialMT'])
        self.assertIn('Source Han Sans CN 没出现在 PDF 里', got[0])
        self.assertIn('SourceHanSerifCN-Bold', got[1])
        self.assertEqual(missing_fonts({'cjk': 'Source Han Sans CN', 'latin': 'Arial'}, ['SourceHanSansCN-Medium', 'Arial-BoldMT']), [])
        self.assertEqual(missing_fonts({'cjk': '思源黑体'}, ['NotoSansCJKsc-Regular']), [])
        self.assertIn('可能被替换', missing_fonts({'latin': 'Helvetica Neue'}, ['ArialMT'])[0])
        self.assertIn('阿里巴巴普惠体', missing_fonts({'cjk': '阿里巴巴普惠体'}, ['SimSun'])[0])


if __name__ == '__main__':
    unittest.main()
