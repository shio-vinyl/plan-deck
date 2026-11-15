import json
import os
import sys
import unittest
from unittest import mock

from helpers import TempDir, make_release, resolved_theme, style_doc

from plan_deck import Deck, MapRelease, Style, StyleError
from plan_deck.theme import composer_cmd, split_command


def themed_doc(**over):
    """示例视觉系统改成从地图主题取强调色和“更新方式”配色，页面自有的颜色按含义命名。"""
    doc = style_doc(**over)
    doc['mapTheme'] = {'project': 'map'}
    doc['fonts']['map'] = '思源黑体'
    doc['palette'] = {'teal': '#0D7578', 'sky': '#33A3DA', 'grey': '#A6A6A6'}
    doc['colors']['emphasis'] = '{theme.palette.highlight}'
    doc['colors']['ledgerGross'] = '{palette.sky}'
    doc['colors']['ledgerDeduct'] = '{palette.grey}'
    doc['colors']['tint'] = '#DEEBF7'
    del doc['classes']['更新方式']
    return doc


def fills(slide):
    out = {}
    for s in slide.shapes:
        try:
            if s.fill.type == 1:
                out[s.name] = '#' + str(s.fill.fore_color.rgb)
        except (AttributeError, TypeError):
            pass
    return out


class ThemeColoursTest(unittest.TestCase):
    def test_theme_and_palette_references(self):
        s = Style(themed_doc(), theme=resolved_theme())
        self.assertEqual(s.color('emphasis'), '#C00000')
        self.assertEqual(s.cls('实施时序', '近期'), '#C00000')            # {colors.emphasis} → 主题强调色
        self.assertEqual(s.cls('更新方式', '有机更新'), '#0D7578')         # 主题的类别配色直接可用，转成大写
        self.assertEqual(s.class_sources, {'实施时序': 'style.json', '更新方式': 'theme'})
        self.assertEqual(s.color('ledgerGross'), '#33A3DA')
        self.assertEqual(s.theme_file, ('style/map.theme.json', 'a' * 64))
        summary = s.summary()
        self.assertEqual(summary['mapTheme']['chain'], ['base', 'research', 'demo-map'])
        self.assertEqual(summary['palette']['grey'], '#A6A6A6')

    def test_palette_names_are_not_page_colours(self):
        s = Style(themed_doc(), theme=resolved_theme())
        with self.assertRaisesRegex(StyleError, '按含义起名'):
            s.color('grey')

    def test_one_place_for_class_colours(self):
        doc = themed_doc()
        doc['classes']['更新方式'] = {'拆除重建': '#E46C3C'}
        with self.assertRaisesRegex(StyleError, '只在一处定'):
            Style(doc, theme=resolved_theme())

    def test_bad_theme_references(self):
        doc = themed_doc()
        doc['colors']['a'] = '{theme.palette.accent2}'
        doc['colors']['b'] = '{theme.constants.labelPrimarySize}'
        doc['colors']['c'] = '{theme.classes.更新方式.拆除重建}'
        with self.assertRaises(StyleError) as e:
            Style(doc, theme=resolved_theme())
        text = '；'.join(e.exception.problems)
        self.assertIn('地图主题里没有 palette.accent2', text)
        self.assertIn('不是 #RRGGBB', text)
        self.assertNotIn('colors.c', text)                                # 类别里的单个颜色也能引用

    def test_theme_reference_needs_map_theme(self):
        doc = style_doc()
        doc['colors']['emphasis'] = '{theme.palette.highlight}'
        with self.assertRaisesRegex(StyleError, '没写 mapTheme'):
            Style(doc)

    def test_map_font_declared(self):
        doc = themed_doc()
        del doc['fonts']['map']
        s = Style(doc, theme=resolved_theme())
        self.assertTrue(any('fonts.map' in w for w in s.warnings))

    def test_t1_page_follows_theme_and_style(self):
        """T1：改主题强调色、改主题里一个类别色、改 style.json 里一个页面颜色，页面三处都跟着变。"""
        def build(doc, theme):
            deck = Deck(Style(doc, theme=theme))
            p = deck.page('ledger', '1.1 测算', lead='论断，**重点**')
            p.ledger(2, 7, [(10.0, '新建', 'ledgerGross'), '−', (2.0, '拆除', 'ledgerDeduct'), '=', (8.0, '计入', 'emphasis')], w=4, h=3)
            p.shape('rect', 2, 12, 4, 1, ('更新方式', '有机更新'), 'mark')
            p.node(8, 12, 4, 1, '条件', 'tint', text_color='text')
            return fills(p.slide)

        before = build(themed_doc(), resolved_theme())
        doc = themed_doc()
        doc['colors']['tint'] = '#E2F0D9'
        after = build(doc, resolved_theme(highlight='#7030A0',
                                          classes={'更新方式': {'拆除重建': '#E46C3C', '综合整治': '#33A3DA', '有机更新': '#00B050'}}))
        self.assertEqual(before['pd:ledger:ledger-block:2'], '#C00000')
        self.assertEqual(after['pd:ledger:ledger-block:2'], '#7030A0')    # 主题强调色
        self.assertEqual(after['pd:ledger:mark'], '#00B050')               # 主题类别色
        self.assertEqual(after['pd:ledger:node'], '#E2F0D9')               # style.json 页面颜色
        self.assertEqual(before['pd:ledger:ledger-block:0'], after['pd:ledger:ledger-block:0'])

    def test_release_rendered_with_other_theme_warns(self):
        with TempDir() as root:
            make_release(root, record={'inputs': [{'ref': 'style/map.theme.json', 'kind': 'file', 'sha256': 'f' * 64}]})
            rel = MapRelease(root, 'v1')
            deck = Deck(Style(themed_doc(), theme=resolved_theme()))
            p = deck.page('evidence', '2.1 现状用地')
            p.map(rel.map('m1-site'), 2, 6, 14, 10)
            p.map(rel.map('m1-site'), 20, 6, 14, 10, slot='map2')
            drift = [w for w in deck.warnings if '地图主题' in w]
            self.assertEqual(len(drift), 1)
            deck = Deck(Style(themed_doc(), theme=resolved_theme(theme_sha='f' * 64)))
            deck.page('evidence', '2.1').map(rel.map('m1-site'), 2, 6, 14, 10)
            self.assertEqual(deck.warnings, [])


FAKE = '''import json, sys
with open(sys.argv[1] + '.argv', 'w') as f:
    json.dump(sys.argv[2:], f)
if 'bad' in sys.argv[-1]:
    print(json.dumps({"ok": False, "error": "THEME_NOT_FOUND"}))
    sys.exit(1)
print(open(sys.argv[1]).read())
'''


class ThemeCommandTest(unittest.TestCase):
    def test_style_load_calls_theme_resolve(self):
        with TempDir() as root:
            os.makedirs(os.path.join(root, 'map'))
            theme_out = os.path.join(root, 'resolved.json')
            with open(theme_out, 'w', encoding='utf-8') as f:
                json.dump(resolved_theme(highlight='#7030A0'), f)
            script = os.path.join(root, 'fake_composer.py')
            with open(script, 'w', encoding='utf-8') as f:
                f.write(FAKE)
            style_path = os.path.join(root, 'style.json')
            with open(style_path, 'w', encoding='utf-8') as f:
                json.dump(themed_doc(), f, ensure_ascii=False)
            cmd = f'"{sys.executable}" "{script}" "{theme_out}"'
            with mock.patch.dict(os.environ, {'GIS_COMPOSER_CMD': cmd}):
                s = Style.load(style_path)
                self.assertEqual(s.color('emphasis'), '#7030A0')
                with open(theme_out + '.argv', encoding='utf-8') as f:
                    self.assertEqual(json.load(f), ['--project', os.path.join(root, 'map'), 'theme', 'resolve'])
                doc = themed_doc()
                doc['mapTheme']['theme'] = 'bad.theme.json'
                with open(style_path, 'w', encoding='utf-8') as f:
                    json.dump(doc, f, ensure_ascii=False)
                with self.assertRaisesRegex(StyleError, 'THEME_NOT_FOUND'):
                    Style.load(style_path)

    def test_missing_composer_explains(self):
        with TempDir() as home:
            with mock.patch.dict(os.environ, {'HOME': home, 'USERPROFILE': home, 'GIS_COMPOSER_CMD': ''}):
                os.makedirs(os.path.join(home, 'map'))
                doc = themed_doc()
                doc['mapTheme']['project'] = os.path.join(home, 'map')
                with self.assertRaisesRegex(StyleError, 'GIS_COMPOSER_CMD'):
                    Style(doc)

    def test_engines_env(self):
        with TempDir() as home:
            os.makedirs(os.path.join(home, '.config', 'gis-plugin'))
            with open(os.path.join(home, '.config', 'gis-plugin', 'engines.env'), 'w', encoding='utf-8') as f:
                f.write('# 注释\nGIS_KIT_CMD="python gis.py"\nGIS_COMPOSER_CMD="/opt/node /Users/a/src/agent-entry.js"\n')
            with mock.patch.dict(os.environ, {'HOME': home, 'USERPROFILE': home, 'GIS_COMPOSER_CMD': ''}):
                self.assertEqual(composer_cmd(), ['/opt/node', '/Users/a/src/agent-entry.js'])
            with mock.patch.dict(os.environ, {'HOME': home, 'USERPROFILE': home, 'GIS_COMPOSER_CMD': 'node "/x y/agent-entry.js"'}):
                self.assertEqual(composer_cmd(), ['node', '/x y/agent-entry.js'])

    def test_backslashes_survive_on_windows(self):
        self.assertEqual(split_command(r'node C:\composer\src\agent-entry.js', windows=True), ['node', r'C:\composer\src\agent-entry.js'])
        self.assertEqual(split_command(r'"C:\Program Files\node.exe" "C:\a b\entry.js"', windows=True),
                         [r'C:\Program Files\node.exe', r'C:\a b\entry.js'])
        self.assertEqual(split_command(r'node a\ b.js', windows=False), ['node', 'a b.js'])


if __name__ == '__main__':
    unittest.main()
