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

    def test_drawn_legend_warns(self):
        deck = Deck(EXAMPLE_STYLE)
        p = deck.page('evidence', '2.1 现状用地')
        p.legend(legend_doc('x.png', drawn=True), 0, 6, 8)
        self.assertIn('drawn: true', deck.warnings[0])


if __name__ == '__main__':
    unittest.main()
