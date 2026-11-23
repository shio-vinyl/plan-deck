"""版面外审计：位置（含组和旋转）、内容、导出、裁剪线索、没审计的文件、连版面外渲染。"""
import hashlib
import io
import json
import os
import unittest
from contextlib import redirect_stdout

from helpers import EXAMPLE_STYLE, TempDir
from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.util import Cm

from plan_deck import Deck, NotAudited, audit, clean, compare_text, offcanvas, placements
from plan_deck.cli import main
from plan_deck.inspect import inspect
from plan_deck.offcanvas import EDGE, _not_audited, placement, render_with_offcanvas

W, H = 33.87, 19.05


def text(sl, name, x, y, w, h, words):
    s = sl.shapes.add_textbox(Cm(x), Cm(y), Cm(w), Cm(h))
    s.name = name
    s.text_frame.text = words
    return s


def png(root, name, color=(200, 30, 30), size=(40, 20)):
    path = os.path.join(root, name)
    Image.new('RGB', size, color).save(path)
    return path


def new_deck():
    prs = Presentation()
    prs.slide_width, prs.slide_height = Cm(W), Cm(H)
    return prs


def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def sample(root):
    """第 1 页版面外有字、表、图表、图；第 2 页隐藏；第 3 页有组、跨边的图、裁剪的图、窗格隐藏的字。"""
    prs = new_deck()
    s1 = blank(prs)
    text(s1, '正文', 2, 5, 10, 2, '版面内的正文')
    text(s1, '外字', W + 2, 1, 10, 2, '任务说明第一行\n第二行')
    tb = s1.shapes.add_table(2, 2, Cm(-12), Cm(1), Cm(10), Cm(3))
    tb.name = '外表'
    for (r, c), v in {(0, 0): '分类', (0, 1): '户数', (1, 0): '一人户', (1, 1): '18436'}.items():
        tb.table.cell(r, c).text = v
    data = CategoryChartData()
    data.categories = ['90㎡以下', '90㎡以上']
    data.add_series('套数', (3120, 1460))
    ch = s1.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Cm(2), Cm(H + 1), Cm(10), Cm(6), data)
    ch.name = '外图表'
    pic = s1.shapes.add_picture(png(root, 'a.png'), Cm(2), Cm(-8), Cm(4), Cm(2))
    pic.name = '外图'
    s1.notes_slide.notes_text_frame.text = '讲稿一句'
    s2 = blank(prs)
    s2._element.set('show', '0')
    text(s2, '隐藏页外字', -9, 1, 8, 2, '需求算法')
    text(s2, '隐藏页内字', 2, 2, 8, 2, '隐藏页正文')
    s3 = blank(prs)
    g = s3.shapes.add_group_shape()
    g.name = '组'
    text(g, '组内字', 2, 2, 4, 2, '组内版面外')
    text(g, '组内字2', 8, 2, 4, 2, '组内也在版面外')
    xf = g._element.grpSpPr.find(qn('a:xfrm'))
    xf.find(qn('a:off')).set('x', str(Cm(W + 5)))     # 组移到版面外；子形状自己的坐标还在版面内
    cross = s3.shapes.add_picture(png(root, 'b.png', (30, 30, 200)), Cm(W - 3), Cm(2), Cm(8), Cm(4))
    cross.name = '跨边图'
    crop = s3.shapes.add_picture(png(root, 'c.png', (30, 200, 30)), Cm(2), Cm(8), Cm(8), Cm(4))
    crop.name = '裁剪图'
    crop.crop_left, crop.crop_right = 0.3, 0.2
    small = s3.shapes.add_picture(png(root, 'd.png', (9, 9, 9)), Cm(12), Cm(8), Cm(4), Cm(4))
    small.name = '小裁剪图'
    small.crop_top = 0.05
    hid = text(s3, '窗格隐藏', 2, 13, 8, 2, '选择窗格里藏着')
    hid._element.nvSpPr.cNvPr.set('hidden', '1')
    bar = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(-5), Cm(1), Cm(2), Cm(2))
    bar.name = '外色块'
    path = os.path.join(root, 'deck.pptx')
    prs.save(path)
    return path


def set_group(g, box, child):
    xf = g._element.grpSpPr.find(qn('a:xfrm'))
    for tag, (a, b), (v1, v2) in (('off', 'xy', box[:2]), ('ext', ('cx', 'cy'), box[2:]), ('chOff', 'xy', child[:2]), ('chExt', ('cx', 'cy'), child[2:])):
        xf.find(qn('a:' + tag)).set(a, str(Cm(v1)))
        xf.find(qn('a:' + tag)).set(b, str(Cm(v2)))


def sha(path):
    with open(path, 'rb') as f:
        return hashlib.sha256(f.read()).hexdigest()


def run_cli(*args):
    buf = io.StringIO()
    with redirect_stdout(buf):
        code = main(list(args))
    return code, json.loads(buf.getvalue())


class PlacementTest(unittest.TestCase):
    def test_edges(self):
        Wm, Hm = Cm(W), Cm(H)
        self.assertEqual(placement((0, 0, Wm, Hm), Wm, Hm), 'inside')
        self.assertEqual(placement((Wm, 0, 10, 10), Wm, Hm), 'outside')
        self.assertEqual(placement((-10, 0, 10, 10), Wm, Hm), 'outside')
        self.assertEqual(placement((-EDGE, 0, Wm + 2 * EDGE, Hm), Wm, Hm), 'inside')      # 满版图略出一点不算
        self.assertEqual(placement((-EDGE - 1, 0, Wm, Hm), Wm, Hm), 'crossing')
        self.assertEqual(placement(None, Wm, Hm), 'unknown')

    def test_group_scale(self):
        prs = new_deck()
        sl = blank(prs)
        g = sl.shapes.add_group_shape()
        child = text(g, '子', 1, 1, 4, 2, '字')
        set_group(g, (W + 1, 1, 8, 4), (0, 0, 4, 2))
        self.assertEqual(placements(sl.shapes, Cm(W), Cm(H))[child._element], 'outside')     # 子形状自己的坐标在版面内
        set_group(g, (0, 0, 8, 4), (0, 0, 4, 2))
        self.assertEqual(placements(sl.shapes, Cm(W), Cm(H))[child._element], 'inside')
        set_group(g, (W - 4, 0, 16, 8), (0, 0, 4, 2))       # 放大 4 倍：子形状从 W 开始
        self.assertEqual(placements(sl.shapes, Cm(W), Cm(H))[child._element], 'outside')

    def test_group_rotation_and_flip(self):
        """组中心 (30, 5)，子形状在中心右边 8 cm，默认在版面外。转 90° 落到中心下方，水平翻转落到中心左边，都回到版面内。"""
        prs = new_deck()
        sl = blank(prs)
        g = sl.shapes.add_group_shape()
        child = text(g, '子', 37, 4, 2, 2, '字')
        set_group(g, (21, 4, 18, 2), (21, 4, 18, 2))
        at = lambda: placements(sl.shapes, Cm(W), Cm(H))[child._element]
        self.assertEqual(at(), 'outside')
        xf = g._element.grpSpPr.find(qn('a:xfrm'))
        xf.set('rot', str(90 * 60000))
        self.assertEqual(at(), 'inside')
        xf.attrib.pop('rot')
        xf.set('flipH', '1')
        self.assertEqual(at(), 'inside')

    def test_shape_rotation(self):
        """宽 4、高 18 的竖条转 90° 后是横条：没转时在版面内，转过去两头出界。"""
        prs = new_deck()
        sl = blank(prs)
        bar = text(sl, '竖条', W - 8, 0.5, 4, 18, '转')
        self.assertEqual(placements(sl.shapes, Cm(W), Cm(H))[bar._element], 'inside')
        bar.rotation = 90
        self.assertEqual(placements(sl.shapes, Cm(W), Cm(H))[bar._element], 'crossing')


class AuditTest(unittest.TestCase):
    def setUp(self):
        self.tmp = TempDir()
        self.root = self.tmp.__enter__()
        self.path = sample(self.root)
        self.out = os.path.join(self.root, 'out')

    def tearDown(self):
        self.tmp.__exit__()

    def load(self):
        with open(os.path.join(self.out, 'offcanvas.json'), encoding='utf-8') as f:
            return json.load(f)

    def obj(self, data, page, name):
        return next(o for o in data['pages'][page - 1]['objects'] if o['shape'] == name)

    def test_full_content_of_offcanvas_objects(self):
        before = sha(self.path)
        res = offcanvas(self.path, self.out)
        self.assertEqual(sha(self.path), before)        # 原 pptx 不改
        data = self.load()
        self.assertEqual(res['summary']['hiddenPages'], [2])
        self.assertEqual(self.obj(data, 1, '正文')['placement'], 'inside')
        out = self.obj(data, 1, '外字')
        self.assertEqual((out['placement'], out['text']), ('outside', '任务说明第一行\n第二行'))
        self.assertEqual(self.obj(data, 1, '外表')['rows'], [['分类', '户数'], ['一人户', '18436']])
        series = self.obj(data, 1, '外图表')['chart']['series'][0]
        self.assertEqual((series['name'], series['categories'], series['values']), ('套数', ['90㎡以下', '90㎡以上'], [3120, 1460]))
        self.assertTrue(os.path.isfile(os.path.join(self.out, self.obj(data, 1, '外图表')['workbook'])))
        exported = self.obj(data, 1, '外图')['pictures'][0]['file']
        with Image.open(os.path.join(self.out, exported)) as im:
            self.assertEqual(im.size, (40, 20))
        self.assertTrue(data['pages'][1]['hidden'])
        self.assertEqual(self.obj(data, 2, '隐藏页外字')['text'], '需求算法')

    def test_group_cross_crop_hidden_and_empty(self):
        res = offcanvas(self.path, self.out)
        data = self.load()
        child = self.obj(data, 3, '组内字')
        self.assertEqual((child['placement'], child['group']), ('outside', '组'))     # 按绝对坐标算，不看子形状自己的坐标
        cross = self.obj(data, 3, '跨边图')
        self.assertEqual(cross['placement'], 'crossing')
        self.assertGreater(cross['outsideShare'], 0.5)
        self.assertTrue(cross['pictures'][0]['file'])
        crop = self.obj(data, 3, '裁剪图')['pictures'][0]
        self.assertEqual((crop['crop']['share'], crop['cropClue']), (0.5, True))
        self.assertTrue(crop['file'])
        small = self.obj(data, 3, '小裁剪图')['pictures'][0]
        self.assertNotIn('cropClue', small)     # 裁得少，版面内，不导出
        self.assertNotIn('file', small)
        self.assertTrue(self.obj(data, 3, '窗格隐藏')['hiddenInPane'])
        self.assertFalse(self.obj(data, 3, '外色块')['hasContent'])
        self.assertEqual(res['summary']['cropped'], 1)
        self.assertEqual(res['summary']['hiddenInPane'], 1)

    def test_human_list_and_text_extraction(self):
        offcanvas(self.path, self.out)
        with open(os.path.join(self.out, '版面外清单.md'), encoding='utf-8') as f:
            listing = f.read()
        for words in ('任务说明第一行', '| 一人户 | 18436 |', '3120', '需求算法', '组内版面外', '选择窗格里藏着', '第 2 页（隐藏页）', '不当正文', '外色块'):
            self.assertIn(words, listing)
        self.assertNotIn('版面内的正文', listing)        # 清单只列渲染图上看不到的
        with open(os.path.join(self.out, '全文提取.md'), encoding='utf-8') as f:
            full = f.read()
        self.assertIn('[版面内] 正文', full)
        self.assertIn('[版面外] 外字', full)
        self.assertIn('[跨边] 跨边图', full)
        self.assertIn('[备注]', full)
        self.assertIn('隐藏页正文', full)

    def test_same_picture_exported_once_and_rerun_is_stable(self):
        offcanvas(self.path, self.out)
        first = sorted(os.listdir(os.path.join(self.out, 'assets')))
        offcanvas(self.path, self.out)
        self.assertEqual(sorted(os.listdir(os.path.join(self.out, 'assets'))), first)

    def test_alternate_content_is_not_skipped(self):
        prs = Presentation(self.path)
        sl = prs.slides[0]
        hidden = text(sl, '备选内容里的字', W + 2, 8, 6, 2, '只在 AlternateContent 里')
        el = hidden._element
        tree = el.getparent()
        tree.remove(el)
        mc = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
        ac = etree.SubElement(tree, f'{{{mc}}}AlternateContent', nsmap={'mc': mc})
        etree.SubElement(ac, f'{{{mc}}}Choice', Requires='a14')
        etree.SubElement(ac, f'{{{mc}}}Fallback').append(el)
        prs.save(self.path)
        offcanvas(self.path, self.out)
        self.assertEqual(self.obj(self.load(), 1, '备选内容里的字')['placement'], 'outside')

    def test_cli(self):
        code, res = run_cli('offcanvas', self.path, '--out', self.out)
        self.assertEqual(code, 0)
        self.assertTrue(res['ok'])
        self.assertIn('不当正文', res['notice'])
        self.assertEqual(res['decks'][0]['summary']['pagesWithContentOutside'], [1, 2, 3])
        self.assertTrue(os.path.isfile(res['decks'][0]['outputs']['list']))

    @unittest.skipUnless(os.environ.get('PLAN_DECK_RENDER_ENGINE'), '真渲染：设 PLAN_DECK_RENDER_ENGINE=powerpoint 或 libreoffice')
    def test_render_with_offcanvas(self):
        before = sha(self.path)
        res = render_with_offcanvas(self.path, os.path.join(self.out, 'renders'), [1, 2], os.environ['PLAN_DECK_RENDER_ENGINE'])
        self.assertEqual(sha(self.path), before)
        self.assertEqual([p['page'] for p in res['pages']], [1, 2])     # 隐藏页也渲染
        with Image.open(res['pages'][0]['png']) as im:
            e = res['pages'][0]['extentCm']
            self.assertGreater(im.width / im.height, (W / H) * 0.9)
            self.assertLess(e[0], 0)
            self.assertGreater(e[2], W)
            pixels = im.convert('RGB')
            frame = round((0 - e[0]) / (e[2] - e[0] + 2 * 0.508) * im.width + 0.508 / 2.54 * res['dpi'])
            row = im.height // 2
            self.assertTrue(any(pixels.getpixel((x, row))[0] > 200 and pixels.getpixel((x, row))[1] < 60 for x in range(max(0, frame - 8), min(im.width, frame + 8))))    # 红框


def table(sl, name, x, y, rows):
    t = sl.shapes.add_table(len(rows), len(rows[0]), Cm(x), Cm(y), Cm(10), Cm(4))
    t.name = name
    for i, row in enumerate(rows):
        for k, v in enumerate(row):
            t.table.cell(i, k).text = v
    return t


HOUSEHOLDS = [['责任单位', '公共租赁住房轮候家庭', ''], ['', '老库', '新库'], ['合计', '12480', '1360'], ['北岸区', '5216', '612'], ['东城区', '4870', '498'], ['西港区', '2394', '250']]


class TableVersionsTest(unittest.TestCase):
    def run_page(self, tables):
        """tables：(名字, x, 行)。返回第 1 页的数据。"""
        with TempDir() as root:
            prs = new_deck()
            sl = blank(prs)
            for name, x, rows in tables:
                table(sl, name, x, 2, rows)
            path = os.path.join(root, 'a.pptx')
            prs.save(path)
            before = sha(path)
            out = os.path.join(root, 'out')
            res = offcanvas(path, out)
            self.assertEqual(sha(path), before)
            with open(os.path.join(out, 'offcanvas.json'), encoding='utf-8') as f:
                page = json.load(f)['pages'][0]
            with open(os.path.join(out, '版面外清单.md'), encoding='utf-8') as f:
                listing = f.read()
            with open(os.path.join(out, '全文提取.md'), encoding='utf-8') as f:
                full = f.read()
            return page, res['summary'], listing, full

    def test_older_version_outside_is_a_clue_and_stays_listed(self):
        old = [r[:] for r in HOUSEHOLDS]
        old[4][1] = '4630'
        old[3][1] = '5,180'
        page, summary, listing, full = self.run_page([('表格 10', 2, HOUSEHOLDS), ('表格 9', W + 2, old)])
        v = page['tableVersions'][0]
        self.assertEqual((v['inside'], v['outside']), ('表格 10', '表格 9'))
        self.assertEqual({(c['row'], c['column'], c['inside'], c['outside']) for c in v['cells']},
                         {('东城区', '公共租赁住房轮候家庭 / 老库', '4870', '4630'), ('北岸区', '公共租赁住房轮候家庭 / 老库', '5216', '5,180')})
        self.assertEqual(summary['pagesWithTableVersions'], [1])
        off = next(o for o in page['objects'] if o['shape'] == '表格 9')
        self.assertEqual(off['olderVersionOf'], '表格 10')
        self.assertEqual(off['rows'], old)                 # 旧版的内容照样列出
        self.assertIn('默认以版面内为准', listing)
        self.assertIn('4630', listing)
        self.assertIn('疑似旧版，版面内是 表格 10', full)

    def test_identical_and_thousands_separator_are_not_clues(self):
        same = [r[:] for r in HOUSEHOLDS]
        same[4][1] = '4,870'
        page, summary, *_ = self.run_page([('内', 2, HOUSEHOLDS), ('外', W + 2, same)])
        self.assertEqual(page['tableVersions'], [])
        self.assertEqual(summary['tableVersions'], 0)

    def test_small_tables_with_all_different_numbers_are_other_data(self):
        """同样的面积段小表，分别是不同街道的数据，不是两版。"""
        a = [['面积段', '比例'], ['90㎡以下', '52%'], ['90-120㎡', '23%'], ['120㎡以上', '25%']]
        b = [['面积段', '比例'], ['90㎡以下', '68%'], ['90-120㎡', '16%'], ['120㎡以上', '16%']]
        self.assertEqual(self.run_page([('内', 2, a), ('外', W + 2, b)])[0]['tableVersions'], [])

    def test_different_tables_and_outside_only_material(self):
        other = [['街道', '项目', '面积'], ['临江街道', '甲地块', '30.00'], ['青石街道', '乙地块', '10.00']]
        page, summary, listing, _ = self.run_page([('内', 2, HOUSEHOLDS), ('外', W + 2, other)])
        self.assertEqual(page['tableVersions'], [])
        self.assertIn('甲地块', listing)      # 版面外独有的资料照样列出
        self.assertNotIn('olderVersionOf', next(o for o in page['objects'] if o['shape'] == '外'))

    def test_row_only_on_one_side_and_best_match(self):
        old = [r[:] for r in HOUSEHOLDS] + [['南湖区', '1520', '180']]
        old[4][1] = '4630'
        decoy = [['年份', '人口', '户数'], ['2023', '1', '2'], ['2024', '3', '4']]
        v = self.run_page([('干扰', 2, decoy), ('内', 14, HOUSEHOLDS), ('外', W + 2, old)])[0]['tableVersions'][0]
        self.assertEqual((v['inside'], v['rowsOnlyOutside'], v['rowsOnlyInside']), ('内', ['南湖区'], []))
        self.assertEqual(len(v['cells']), 1)


class NotAuditedTest(unittest.TestCase):
    def test_flat_xml(self):
        with TempDir() as root:
            path = os.path.join(root, 'flat.pptx')
            with open(path, 'w', encoding='utf-8') as f:
                f.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"/>')
            out = os.path.join(root, 'out')
            before = sha(path)
            code, res = run_cli('offcanvas', path, '--out', out)
            self.assertEqual(code, 1)
            self.assertFalse(res['ok'])
            self.assertEqual(res['decks'], [])
            self.assertEqual((res['notAudited'][0]['format'], res['notAudited'][0]['audited']), ('flat-xml', False))
            self.assertEqual(sha(path), before)
            with open(os.path.join(out, '未审计.md'), encoding='utf-8') as f:
                self.assertIn('flat.pptx', f.read())
            with self.assertRaises(NotAudited):
                offcanvas(path, out)

    def test_batch_reports_the_unreadable_and_audits_the_rest(self):
        with TempDir() as root:
            good = sample(root)
            bad = os.path.join(root, 'legacy.ppt')
            with open(bad, 'wb') as f:
                f.write(bytes.fromhex('d0cf11e0a1b11ae1') + b'\0' * 64)
            out = os.path.join(root, 'out')
            code, res = run_cli('offcanvas', good, bad, os.path.join(root, 'nope.pptx'), '--out', out)
            self.assertEqual(code, 1)
            self.assertEqual(len(res['decks']), 1)
            self.assertEqual([m['format'] for m in res['notAudited']], ['ole-compound', 'missing'])
            self.assertTrue(os.path.isfile(os.path.join(out, 'deck', 'offcanvas.json')))

    def test_formats(self):
        with TempDir() as root:
            wrapped = os.path.join(root, 'w.pptx')
            with open(wrapped, 'wb') as f:
                f.write(bytes(b ^ 0x5c for b in b'DG_FILE VER 5.0') + b'\0' * 32)
            self.assertEqual(_not_audited(wrapped, ValueError('x'))['format'], 'encrypted-wrapper')
            self.assertEqual(_not_audited(root, ValueError('x'))['format'], 'directory')
            junk = os.path.join(root, 'j.pptx')
            with open(junk, 'wb') as f:
                f.write(b'hello')
            self.assertEqual(_not_audited(junk, ValueError('x'))['format'], 'unknown')


class OtherCommandsTest(unittest.TestCase):
    def test_inspect_marks_placement_and_crossing_picture(self):
        with TempDir() as root:
            res = inspect(sample(root))
            items = {i['name']: i for p in res['pages'] for i in p['items']}
            self.assertEqual(items['正文']['placement'], 'inside')
            self.assertEqual(items['外字']['placement'], 'outside')
            self.assertEqual(items['组内字']['placement'], 'outside')
            self.assertEqual(items['跨边图']['placement'], 'crossing')
            self.assertTrue(any(w['shape'] == '跨边图' and '版面外' in w['warning'] for w in res['warnings']))
            self.assertTrue(any('plan-deck offcanvas' in p['problem'] for p in res['problems']))

    def test_audit_keeps_offcanvas_out_of_page_text(self):
        with TempDir() as root:
            prs = new_deck()
            sl = blank(prs)
            text(sl, '版面外草稿', W + 1, 1, 8, 2, '待确认：这段口径')
            text(sl, '正文', 2, 5, 12, 2, '待确认：页面上的这段')
            path = os.path.join(root, 'a.pptx')
            prs.save(path)
            found = audit(path)['problems']
            self.assertTrue(any('页面上的' in x['text'] for x in found))
            self.assertFalse(any('这段口径' in x['text'] for x in found))
            sl.shapes[0].text_frame.text = '待确认：改过的口径'
            other = os.path.join(root, 'b.pptx')
            prs.save(other)
            change = compare_text(path, other)['changes'][0]
            self.assertTrue(change['onlyHere'][0].startswith('[版面外]'))

    def test_clean_marks_offcanvas_hits(self):
        with TempDir() as root:
            prs = new_deck()
            sl = blank(prs)
            text(sl, '版面外草稿', W + 1, 1, 8, 2, '负责人张工')
            text(sl, '正文', 2, 5, 12, 2, '讨论稿')
            path = os.path.join(root, 'a.pptx')
            prs.save(path)
            hits = clean(path, os.path.join(root, 'b.pptx'))['pageHits']
            self.assertEqual({h['text']: h.get('offCanvas', False) for h in hits}, {'负责人张工': True, '讨论稿': False})

    def test_deck_built_pages_have_no_offcanvas(self):
        with TempDir() as root:
            deck = Deck(EXAMPLE_STYLE)
            p = deck.page('evidence', '2.1 现状用地', parent='二、现状评估')
            p.text(deck.style.x0, deck.style.y0, 10, ['正文'])
            path = os.path.join(root, 'd.pptx')
            deck.save(path)
            res = offcanvas(path, os.path.join(root, 'out'))
            self.assertEqual((res['summary']['outside'], res['summary']['crossing']), (0, 0))


if __name__ == '__main__':
    unittest.main()
