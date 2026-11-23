import contextlib
import hashlib
import io
import json
import os
import subprocess
import sys
import unittest

from helpers import EXAMPLE_STYLE, TempDir, legend_doc, make_release

from plan_deck import Deck, MapError, MapRelease
from plan_deck.cli import main as cli
from plan_deck.maps import legend_rows


class ReleaseTest(unittest.TestCase):
    def test_reads_release(self):
        with TempDir() as root:
            make_release(root, panels={'m1-site': ('zoom', (6.0, 4.5, 300))})
            rel = MapRelease(root, 'v1')
            self.assertEqual(rel.map_ids(), ['m1-site', 'm2-pixel'])
            a = rel.map('m1-site')
            self.assertEqual(a.check_print(), (10.0, 8.0, 300.0))
            z = rel.map('m1-site', panel='zoom')
            self.assertEqual(z.id, 'm1-site/zoom')
            rows = rel.check()
            self.assertEqual([r['map'] for r in rows], ['m1-site', 'm1-site/zoom', 'm2-pixel'])
            self.assertTrue(all(r['status'] == 'ok' for r in rows))
            self.assertIn('没按厘米出图', rows[2]['note'])

    def test_missing_release_explains(self):
        with TempDir() as root:
            with self.assertRaisesRegex(MapError, 'project release'):
                MapRelease(root, 'v1')
            make_release(root)
            with self.assertRaisesRegex(MapError, '现有：v1'):
                MapRelease(root, 'v2')
            with self.assertRaisesRegex(MapError, '要指定发布版本'):
                MapRelease(root, None)
            with self.assertRaisesRegex(MapError, '有：m1-site'):
                MapRelease(root, 'v1').map('m9')

    def test_changed_file_is_refused(self):
        with TempDir() as root:
            rel = make_release(root)
            with open(os.path.join(rel, 'output', 'm1-site', 'm1-site.png'), 'ab') as f:
                f.write(b'x')
            with self.assertRaisesRegex(MapError, 'sha256'):
                MapRelease(root, 'v1').map('m1-site')

    def test_legend_contract_version(self):
        with TempDir() as root:
            rel = make_release(root, maps={'m1': (5.0, 5.0, 300)})
            p = os.path.join(rel, 'output', 'm1', 'm1.legend.json')
            doc = read_json(p)
            doc['schema_version'] = 2
            write_json(p, doc)
            make_release_hash(rel)
            with self.assertRaisesRegex(MapError, 'schema_version 2'):
                MapRelease(root, 'v1').map('m1')

    def test_pixels_must_match_print(self):
        with TempDir() as root:
            rel = make_release(root, maps={'m1': (5.0, 5.0, 300)})
            p = os.path.join(rel, 'output', 'm1', 'm1.legend.json')
            doc = read_json(p)
            doc['print']['dpi'] = 200
            write_json(p, doc)
            make_release_hash(rel)
            with self.assertRaisesRegex(MapError, '不符'):
                MapRelease(root, 'v1').map('m1').check_print()


def read_json(p):
    with open(p, encoding='utf-8') as f:
        return json.load(f)


def write_json(p, doc):
    with open(p, 'w', encoding='utf-8') as f:
        json.dump(doc, f, ensure_ascii=False)


def make_release_hash(rel):
    """改了发布版里的文件后重算 release.json 的哈希（模拟“就是这样发布的”）。"""
    rec_p = os.path.join(rel, 'release.json')
    rec = read_json(rec_p)
    for e in rec['files']:
        with open(os.path.join(rel, e['path']), 'rb') as f:
            e['sha256'] = hashlib.sha256(f.read()).hexdigest()
    write_json(rec_p, rec)


class PlacementTest(unittest.TestCase):
    def setUp(self):
        self.tmp = TempDir()
        root = self.tmp.__enter__()
        make_release(root)
        self.rel = MapRelease(root, 'v1')
        self.deck = Deck(EXAMPLE_STYLE)
        self.page = self.deck.page('evidence', '2.1 现状用地', parent='二、现状评估', lead='论断')

    def tearDown(self):
        self.tmp.__exit__()

    def test_native_size_centered(self):
        box = self.page.map(self.rel.map('m1-site'), 2, 6, 14, 10)
        self.assertAlmostEqual(box.w, 10.0)
        self.assertAlmostEqual(box.h, 8.0)
        self.assertAlmostEqual(box.x, 4.0)
        rec = self.deck.placements[-1]
        self.assertEqual(rec['effectiveDpi'], 300)
        self.assertEqual(rec['scale'], 1.0)
        self.assertEqual(self.deck.warnings, [])

    def test_slot_too_small_is_an_error(self):
        with self.assertRaisesRegex(MapError, 'layout.widthCm'):
            self.page.map(self.rel.map('m1-site'), 2, 6, 8, 10)

    def test_shrink_warns_with_map_text(self):
        box = self.page.map(self.rel.map('m1-site'), 2, 6, 8, 10, shrink=True)
        self.assertAlmostEqual(box.w, 8.0)
        self.assertIn('低于 mapText.minorPt', self.deck.warnings[-1])

    def test_pixel_map_warns(self):
        self.page.map(self.rel.map('m2-pixel'), 2, 6, 8, 6)
        self.assertIn('没按厘米出图', self.deck.warnings[-1])

    def test_only_release_assets(self):
        with self.assertRaisesRegex(MapError, 'MapRelease'):
            self.page.map('runs/x/output/m1.png', 0, 0, 5, 5)


class ReleaseQaTest(unittest.TestCase):
    """gis-composer 的已知问题发布：发布版可以带着未过的 QA 发布；读取方默认警告、列出未过的检查，严格模式报错。"""

    def test_failed_checks_from_qa_reports(self):
        with TempDir() as root:
            make_release(root, maps={'m1': (5.0, 5.0, 300), 'm2': (5.0, 5.0, 300)},
                         panels={'m1': ('zoom', (5.0, 5.0, 300))},
                         failed={'m1': [('IMPORTANT_LABEL_MISSING', 'CB-021', 'zoom')]})
            rel = MapRelease(root, 'v1')
            self.assertEqual(rel.qa('m2').status, 'ok')
            q = rel.qa('m1')
            self.assertEqual(q.status, 'failed')
            self.assertEqual(q.issues, [{'code': 'IMPORTANT_LABEL_MISSING', 'panel': 'zoom', 'ids': ['CB-021']}])
            deck = Deck(EXAMPLE_STYLE)
            p = deck.page('gallery', '3.3 推荐地块')
            p.map(rel.map('m1'), 2, 6, 6, 6)
            p.map(rel.map('m2'), 10, 6, 6, 6, slot='map2')
            self.assertEqual(len(deck.warnings), 1)
            self.assertIn('IMPORTANT_LABEL_MISSING（zoom 面板、CB-021）', deck.warnings[0])
            self.assertIn('第 1 页 gallery:map', deck.warnings[0])
            self.assertEqual(rel.qa('m1', panel='zoom').status, 'failed')
            self.assertEqual(rel.qa('m1', panel='locator').status, 'ok')     # 问题只在 zoom 面板

    def test_release_record_with_accepted(self):
        with TempDir() as root:
            make_release(root, maps={'m1': (5.0, 5.0, 300), 'm2': (5.0, 5.0, 300)}, record={
                'status': 'known-issues',
                'maps': [{'id': 'm1', 'map': 'maps/m1.map.json', 'qa': 'failed', 'codes': ['IMPORTANT_LABEL_MISSING'],
                          'issues': [{'code': 'IMPORTANT_LABEL_MISSING', 'panelId': 'zoom', 'layerId': 'tiers-labels', 'featureId': 'CB-013'}]},
                         {'id': 'm2', 'map': 'maps/m2.map.json', 'qa': 'ok', 'codes': [], 'issues': []}],
                'accepted': [{'map': 'm1', 'codes': ['IMPORTANT_LABEL_MISSING'], 'reason': '编号在页面表格里有'}]})
            rel = MapRelease(root, 'v1')
            self.assertEqual(rel.status, 'known-issues')
            q = rel.qa('m1')
            self.assertEqual((q.status, q.reason, q.source), ('failed', '编号在页面表格里有', 'release.json'))
            self.assertEqual(q.issues, [{'code': 'IMPORTANT_LABEL_MISSING', 'panel': 'zoom', 'ids': ['CB-013']}])
            self.assertEqual(rel.qa('m2').status, 'ok')
            rel.map('m1')
            self.assertIn('发布原因：编号在页面表格里有', rel.warnings[0])

    def test_no_conclusion_is_not_a_pass(self):
        with TempDir() as root:
            make_release(root, maps={'m1': (5.0, 5.0, 300), 'm2': (5.0, 5.0, 300)},
                         record={'maps': [{'id': 'm2', 'qa': 'ok'}]})
            rel = MapRelease(root, 'v1')
            self.assertEqual(rel.qa('m1').status, 'unknown')                  # maps 里没列
            rel.map('m1')
            self.assertIn('没有 QA 结论', rel.warnings[0])

    def test_strict_mode(self):
        with TempDir() as root:
            make_release(root, maps={'m1': (5.0, 5.0, 300), 'm2': (5.0, 5.0, 300)},
                         failed={'m1': [('IMPORTANT_LABEL_MISSING', 'CB-021', None)]})
            rel = MapRelease(root, 'v1', strict=True)
            rel.map('m2')
            with self.assertRaisesRegex(MapError, '严格模式'):
                rel.map('m1')

    def test_maps_check_cli(self):
        with TempDir() as root:
            make_release(root, maps={'m1': (5.0, 5.0, 300)}, failed={'m1': [('IMPORTANT_LABEL_MISSING', 'CB-021', None)]})
            out = io.StringIO()
            with contextlib.redirect_stdout(out):
                code = cli(['maps', 'check', root, '--release', 'v1'])
            doc = json.loads(out.getvalue())
            self.assertEqual(code, 0)                                         # 不阻塞
            self.assertEqual(doc['maps'][0]['qa'], 'failed')
            self.assertEqual(doc['maps'][0]['qaIssues'][0]['code'], 'IMPORTANT_LABEL_MISSING')
            self.assertIn('IMPORTANT_LABEL_MISSING', doc['warnings'][0])
            out = io.StringIO()
            with contextlib.redirect_stdout(out):
                code = cli(['maps', 'check', root, '--release', 'v1', '--strict'])
            self.assertEqual(code, 1)
            self.assertEqual(json.loads(out.getvalue())['maps'][0]['status'], 'error')

    def test_maps_check_with_nothing_to_check_fails(self):
        with TempDir() as root:
            make_release(root, maps={})
            out = io.StringIO()
            with contextlib.redirect_stdout(out):
                code = cli(['maps', 'check', root, '--release', 'v1'])
            doc = json.loads(out.getvalue())
            self.assertEqual(code, 1)
            self.assertEqual((doc['ok'], doc['error'], doc['maps']), (False, 'NO_MAPS', []))
            with self.assertRaisesRegex(MapError, '有：（无）'):
                MapRelease(root, 'v1').map('m1')


class SingleRunReleaseTest(unittest.TestCase):
    """单图运行发布的版本：output/<图>.png、qa/<图>.qa.json，多面板图的面板在 output/panels/、qa/panels/。"""

    def test_reads_flat_layout(self):
        with TempDir() as root:
            make_release(root, maps={'m1': (5.0, 5.0, 300)}, flat=True)
            rel = MapRelease(root, 'v1')
            self.assertEqual(rel.map_ids(), ['m1'])
            a = rel.map('m1')
            self.assertEqual(a.check_print(), (5.0, 5.0, 300.0))
            self.assertTrue(a.legend)
            self.assertEqual(a.qa.status, 'ok')
            self.assertEqual([r['status'] for r in rel.check()], ['ok'])

    def test_panels_directory_is_not_a_map(self):
        with TempDir() as root:
            make_release(root, maps={'sheet': (10.0, 8.0, 300)}, panels={'sheet': ('zoom', (6.0, 4.5, 300))}, flat=True,
                         failed={'sheet': [('LABEL_MISSING', 'F-1', 'zoom')]},
                         fonts={'sheet/zoom': [('GC Sans', 'a.otf', 0)]})
            rel = MapRelease(root, 'v1')
            self.assertEqual(rel.map_ids(), ['sheet'])
            self.assertEqual([r['map'] for r in rel.check()], ['sheet', 'sheet/zoom'])
            z = rel.map('sheet', panel='zoom')
            self.assertEqual(z.check_print(), (6.0, 4.5, 300.0))
            self.assertEqual(z.qa.issues[0]['code'], 'LABEL_MISSING')    # 读 qa/panels/ 下的报告
            self.assertEqual(z.fonts[0]['family'], 'GC Sans')

    def test_changed_file_is_refused(self):
        with TempDir() as root:
            rel = make_release(root, maps={'m1': (5.0, 5.0, 300)}, flat=True)
            with open(os.path.join(rel, 'output', 'm1.png'), 'ab') as f:
                f.write(b'x')
            with self.assertRaisesRegex(MapError, 'sha256'):
                MapRelease(root, 'v1').map('m1')


class EntryScriptTest(unittest.TestCase):
    def test_missing_dependency_names_the_interpreter(self):
        entry = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'bin', 'plan-deck')
        result = subprocess.run([sys.executable, '-I', '-B', '-S', entry, '--help'], capture_output=True, text=True, encoding='utf-8')
        self.assertEqual(result.returncode, 1)
        self.assertEqual(result.stdout, '')
        self.assertIn(sys.executable, result.stderr)
        self.assertIn('缺依赖', result.stderr)
        self.assertNotIn('Traceback', result.stderr)


class LegendTest(unittest.TestCase):
    def test_rows(self):
        rows, skipped = legend_rows(legend_doc('x.png'))
        self.assertEqual(rows[0], ('title', '实施时序'))                                   # 分类行先出标题
        self.assertEqual(rows[1][2], '近期')                                             # shortLabel
        self.assertEqual(rows[3][2], '圆的大小：更新规模 / 万㎡（1–12）')
        self.assertEqual(rows[4], ('row', '#194461', '更新单元范围', 'frame', 1.0))      # 普通行排在后面，线画成框
        self.assertEqual(rows[5][4], 0.32)                                               # 不透明度照传
        self.assertEqual(skipped, [])
        long_rows, _ = legend_rows(legend_doc('x.png'), short=False, context=False)
        self.assertEqual(long_rows[1][2], '近期（2026–2027 年，3 处）')
        self.assertEqual(len(long_rows), 5)                                              # context 行不画
        sheet = {'items': [{'label': '放大位置', 'symbol': 'ring', 'color': '#C00000'}]}
        merged, _ = legend_rows([legend_doc('x.png', items=[]), sheet])
        self.assertEqual(merged, [('row', '#C00000', '放大位置', 'frame', 1.0)])         # 多份合并，圈画成框

    def test_page_and_map_class_colours_compared(self):
        deck = Deck(EXAMPLE_STYLE)
        p = deck.page('evidence', '2.1 现状用地')
        p.legend(legend_doc('x.png'), 0, 6, 8)                 # 近期 #C00000、远期 #B4C7E7 与 style 一致
        self.assertEqual(deck.warnings, [])
        doc = legend_doc('x.png')
        doc['items'][1]['classes'][1]['color'] = '#7030A0'
        p.legend(doc, 0, 6, 8)
        self.assertIn('同一类不同色', deck.warnings[0])
        self.assertIn('classes.实施时序', deck.warnings[0])

    def test_unknown_model_reported(self):
        rows, skipped = legend_rows(legend_doc('x.png', items=[{'label': 'dem', 'model': 'continuous', 'stops': []}]))
        self.assertEqual(rows, [])
        self.assertEqual(len(skipped), 1)

    def test_page_legend_shapes(self):
        with TempDir() as root:
            make_release(root)
            a = MapRelease(root, 'v1').map('m1-site')
            deck = Deck(EXAMPLE_STYLE)
            p = deck.page('evidence', '2.1 现状用地')
            box = p.legend(a, 30, 6, 8, extra=[('emphasis', '放大位置', 'frame')])
            names = [s.name for s in p.slide.shapes]
            self.assertIn('pd:evidence:legend-title', names)
            self.assertIn('pd:evidence:legend-swatch:5', names)       # 5 行 + 1 行补充 = 6 个色块
            self.assertNotIn('pd:evidence:legend-swatch:6', names)
            self.assertGreater(box.h, 6 * 0.55)
            labels = [s.text_frame.text for s in p.slide.shapes if s.name.startswith('pd:evidence:legend-label')]
            self.assertEqual(labels[-1], '放大位置')

    def test_library_symbol_rows_draw_editable_shapes(self):
        # gis-composer 符号库的行（librarySymbol + svg）：照 svg 画成一组自由曲线，不画色块
        point = ('<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="-12.5 -12.5 25 25">'
                 '<path d="M-7.5 -12.5 L7.5 -12.5 C10 -12.5 12.5 -10 12.5 -7.5 L12.5 7.5 L-12.5 7.5 Z" fill="#FFFFFF" fill-rule="evenodd" '
                 'stroke="#000000" stroke-width="0.75"/><path d="M-3 -3 L3 -3 L3 3 L-3 3 Z M-1 -1 L1 -1 L1 1 L-1 1 Z" fill="#000000"/></svg>')
        line = ('<svg xmlns="http://www.w3.org/2000/svg" width="24" height="12" viewBox="0 0 18 9">'
                '<path d="M0 4.5 L18 4.5" fill="none" stroke="#000000" stroke-width="2.835"/>'
                '<path d="M0 4.5 L18 4.5" fill="none" stroke="#FEFEFE" stroke-width="1.701" stroke-dasharray="19.843 19.843"/></svg>')
        area = ('<svg xmlns="http://www.w3.org/2000/svg" width="18" height="12" viewBox="0 0 13.5 9"><defs><pattern id="gmc-h-1" '
                'patternUnits="userSpaceOnUse" width="4" height="4" patternTransform="rotate(-45)"><path d="M0 2 H4" stroke="#0000FF" '
                'stroke-width="1"/></pattern></defs><rect x="0.5" y="0.5" width="12.5" height="8" fill="url(#gmc-h-1)"/>'
                '<rect x="0.5" y="0.5" width="12.5" height="8" fill="none" stroke="#0000FF" stroke-width="1"/></svg>')
        ref = lambda kind, name: {'library': 'tkgh-2025', 'kind': kind, 'name': name}
        doc = legend_doc('x.png', items=[
            {'label': '地铁站', 'symbol': 'point', 'librarySymbol': ref('point', '轨道交通站点'), 'svg': point, 'color': '#000000'},
            {'label': '铁路', 'type': 'line', 'librarySymbol': ref('line', '铁路'), 'svg': line, 'color': '#000000'},
            {'label': '水源保护区', 'symbol': 'fill', 'librarySymbol': ref('fill', '饮用水水源地一级保护区'), 'svg': area, 'color': '#0000FF'}])
        rows, skipped = legend_rows(doc)
        self.assertEqual(skipped, [])
        self.assertEqual([(r[2], r[3].name) for r in rows], [('地铁站', '轨道交通站点'), ('铁路', '铁路'), ('水源保护区', '饮用水水源地一级保护区')])
        deck = Deck(EXAMPLE_STYLE)
        p = deck.page('evidence', '2.1 现状用地')
        p.legend(doc, 0, 6, 8)
        groups = {s.name: s for s in p.slide.shapes if s.name.startswith('pd:evidence:legend-swatch')}
        self.assertEqual(sorted(groups), ['pd:evidence:legend-swatch:0', 'pd:evidence:legend-swatch:1', 'pd:evidence:legend-swatch:2'])
        station, rail, water = (groups[f'pd:evidence:legend-swatch:{i}'] for i in range(3))
        self.assertEqual(len(station.shapes), 2)
        self.assertEqual(round(station.width / 360000, 2), 0.42)                     # 点符号占一个色块（0.42 cm）
        xml = rail.shapes[1]._element.xml
        self.assertIn('a:custDash', xml)                                            # 白色虚线照 SVG 的长度
        self.assertIn('FEFEFE', xml)
        self.assertLessEqual(rail.width / 360000, 0.6)                               # 线样本不压到标注
        self.assertIn('a:pattFill', water.shapes[0]._element.xml)
        self.assertEqual(deck.warnings, [])

    def test_drawn_legend_warns(self):
        deck = Deck(EXAMPLE_STYLE)
        p = deck.page('evidence', '2.1 现状用地')
        p.legend(legend_doc('x.png', drawn=True), 0, 6, 8)
        self.assertIn('drawn: true', deck.warnings[0])


if __name__ == '__main__':
    unittest.main()


class VividMapMarksTest(unittest.TestCase):
    """鲜艳实用汇报的图上注记方框和图内一列图例（docs/map-types.md）。"""

    ITEMS = [
        {'label': '重点片区', 'type': 'line', 'color': '#FF0000', 'width': 2, 'dasharray': [3, 1.5]},
        {'label': '片区边界', 'type': 'line', 'color': '#7F7F7F', 'width': 1.33, 'dasharray': [10, 0]},
        {'label': 'land', 'type': 'circle', 'model': 'classified', 'title': '居住用地',
         'classes': [{'code': '已列入计划', 'label': '已列入供地计划', 'color': '#FF0000'},
                     {'code': '储备', 'label': '储备地', 'color': '#0064FF'}]},
        {'label': '用地面积', 'type': 'circle', 'color': '#6a7378', 'model': 'size',
         'symbols': [{'label': '12', 'radius': 10}, {'label': '1', 'radius': 3}]},
        {'label': '站点 500 米范围', 'type': 'fill', 'color': '#E3CB9E', 'opacity': 1},
    ]

    def test_line_rows_keep_dash_and_sizes_can_be_left_out(self):
        rows, _ = legend_rows(legend_doc('x.png', items=self.ITEMS), lines=True, sizes=False)
        rows = [r for r in rows if r[0] == 'row']
        self.assertEqual([r[2] for r in rows], ['已列入供地计划', '储备地', '重点片区', '片区边界', '站点 500 米范围'])
        line, solid = rows[2][3], rows[3][3]
        self.assertTrue(line.dashed)
        self.assertEqual(line.width_pt, 1.5)
        self.assertFalse(solid.dashed)                                 # [10, 0] 是实线
        self.assertEqual(rows[4][3], 'fill')
        plain, _ = legend_rows(legend_doc('x.png', items=self.ITEMS))
        self.assertIn(('row', '#FF0000', '重点片区', 'frame', 1.0), plain)  # 默认照旧画虚线框

    def test_inset_legend_one_column_in_the_map_corner(self):
        deck = Deck(EXAMPLE_STYLE)
        p = deck.page('evidence', '3.1 发展策略')
        from plan_deck.deck import Box
        m = Box(20, 4, 10, 13)
        box = p.inset_legend(legend_doc('x.png', items=self.ITEMS), m)
        self.assertAlmostEqual(box.x, 20.15)
        self.assertAlmostEqual(box.bottom, 17 - 0.15)
        names = [s.name for s in p.slide.shapes]
        self.assertIn('pd:evidence:map-legend-bg', names)
        self.assertIn('pd:evidence:map-legend-title', names)
        self.assertIn('pd:evidence:map-legend-swatch:4', names)        # 5 行，尺寸行不画、不出分组标题
        self.assertNotIn('pd:evidence:map-legend-swatch:5', names)
        labels = [s.text_frame.text for s in p.slide.shapes if s.name.startswith('pd:evidence:map-legend-label')]
        self.assertNotIn('居住用地', labels)
        dashed = [s for s in p.slide.shapes if s.name == 'pd:evidence:map-legend-swatch:2'][0]
        self.assertIsNotNone(dashed.line.dash_style)
        self.assertEqual(deck.warnings, [])
        right = p.inset_legend(legend_doc('x.png', items=self.ITEMS), m, corner='top-right', title=None)
        self.assertAlmostEqual(right.right, 30 - 0.15)
        self.assertAlmostEqual(right.y, 4.15)

    def test_inset_legend_warns_about_many_rows_and_yellow_points(self):
        deck = Deck(EXAMPLE_STYLE)
        p = deck.page('evidence', '3.1 发展策略')
        from plan_deck.deck import Box
        classes = [{'code': str(k), 'label': f'类 {k}', 'color': c}
                   for k, c in enumerate(['#FF0000', '#0064FF', '#FF6A00', '#00B050', '#7A00FF', '#FFE100', '#000000'])]
        p.inset_legend(legend_doc('x.png', items=[{'label': 'pts', 'type': 'circle', 'model': 'classified', 'classes': classes}]),
                       Box(0, 0, 10, 10))
        self.assertTrue(any('7 行' in w and '拆' in w for w in deck.warnings))
        self.assertTrue(any('类 5' in w and '黄色' in w for w in deck.warnings))

    def test_map_label_box_and_leader_without_shadow(self):
        from plan_deck.deck import Box
        from plan_deck.inspect import _has_shadow
        deck = Deck(EXAMPLE_STYLE)
        p = deck.page('evidence', '3.1 发展策略')
        m = Box(20, 4, 10, 13)
        b = p.map_label(m, '西岸 12 街坊', at=(3, 3), box_at=(5, 1), r=0.15)
        far = p.map_label(m, '很长的一个项目名称在右边', at=(9, 12), box_at=(9.5, 12.9))
        self.assertLessEqual(far.right, m.right + 1e-9)               # 方框不出地图
        self.assertLessEqual(far.bottom, m.bottom + 1e-9)
        shapes = {s.name: s for s in p.slide.shapes}
        box, leader = shapes['pd:evidence:map-label:0'], shapes['pd:evidence:map-label-leader:0']
        self.assertEqual(box.text_frame.text, '西岸 12 街坊')
        self.assertEqual(str(box.fill.fore_color.rgb), 'FFFFFF')
        self.assertEqual(str(box.line.color.rgb), '000000')
        self.assertTrue(box.text_frame.paragraphs[0].runs[0].font.bold)
        self.assertEqual(str(box.text_frame.paragraphs[0].runs[0].font.color.rgb), '000000')
        # 引线从方框左边中点出发，停在点的边缘
        self.assertAlmostEqual(leader.end_x / 360000, b.x, places=2)
        self.assertAlmostEqual(leader.end_y / 360000, b.y + b.h / 2, places=2)
        for s in p.slide.shapes:
            if 'map-label' in s.name:
                self.assertFalse(_has_shadow(s), s.name)
