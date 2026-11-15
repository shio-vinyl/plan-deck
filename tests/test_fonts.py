"""地图用字：fonts.map 与页面字体、地图主题的字体、发布版实际加载的字体。"""
import contextlib
import io
import json
import unittest

from helpers import TempDir, make_release, resolved_theme, style_doc

from plan_deck import Deck, MapRelease, Style
from plan_deck.cli import main as cli
from plan_deck.fonts import families, name_aliases, same_font, theme_font_uses


def themed(typography, map_font='思源黑体', cjk='微软雅黑'):
    doc = style_doc()
    doc['mapTheme'] = {'project': 'map'}
    doc['fonts'] = {'cjk': cjk, 'latin': 'Arial', **({'map': map_font} if map_font else {})}
    theme = resolved_theme(classes={})
    theme['typography'] = typography
    return Style(doc, theme=theme)


def font_warnings(s):
    return [w for w in s.warnings if '字体' in w]


# 主题按 fonts.map 声明了字体：族名就是字体名，西文字体在前、中文字体在后
SOURCE_HAN = {'mapFontFamily': '思源黑体', 'fontFamily': '"Source Sans 3", "思源黑体", sans-serif',
              'map': {'label': ['Source Sans 3', '思源黑体'], 'labelStrong': ['思源黑体 Medium']},
              'title': {'family': '"Source Han Sans CN Medium", sans-serif'},
              'fonts': [{'family': '思源黑体'}, {'family': '思源黑体 Medium', 'weight': 500}, {'family': 'Source Sans 3'}]}


# gis-composer 读出的 name 表（theme resolve 的 typography.fonts[].names、metrics 的 render.fonts[].names）
HEI_NAMES = {'family': 'Source Han Sans CN', 'fullName': 'Source Han Sans CN Regular', 'postScriptName': 'SourceHanSansCN-Regular',
             'localized': [{'language': 'zh-CN', 'family': '思源黑体 CN', 'fullName': '思源黑体 CN Regular'}]}
SONG_NAMES = {'family': 'Source Han Serif CN', 'fullName': 'Source Han Serif CN Regular', 'postScriptName': 'SourceHanSerifCN-Regular',
              'localized': [{'language': 'zh-CN', 'family': '思源宋体 CN', 'fullName': '思源宋体 CN Regular'}]}


class FontNamesTest(unittest.TestCase):
    def test_same_font_across_names(self):
        self.assertTrue(same_font('思源黑体', 'Source Han Sans CN'))
        self.assertTrue(same_font('思源黑体', 'SourceHanSansCN-Medium'))
        self.assertTrue(same_font('Noto Sans CJK SC', '思源黑体 Bold'))
        self.assertTrue(same_font('微软雅黑', 'Microsoft YaHei UI'))
        self.assertTrue(same_font('Source Sans 3', 'SourceSans3-It'))
        self.assertFalse(same_font('思源黑体', '微软雅黑'))
        self.assertFalse(same_font('思源黑体', '思源宋体'))
        self.assertFalse(same_font('黑体', '思源黑体'))
        self.assertFalse(same_font('思源黑体', None))

    def test_theme_font_uses(self):
        self.assertEqual(families('"GC Sans", GC Sans Latin, sans-serif'), ['GC Sans', 'GC Sans Latin'])
        uses = dict(theme_font_uses(SOURCE_HAN))
        self.assertEqual(uses['mapFontFamily'], ['思源黑体'])
        self.assertEqual(uses['map.label'], ['Source Sans 3', '思源黑体'])
        self.assertEqual(uses['title.family'], ['Source Han Sans CN Medium'])
        # 没写 mapFontFamily 时注记默认用第一个声明的字体
        self.assertEqual(dict(theme_font_uses({'fonts': [{'family': 'A'}]}))['mapFontFamily'], ['A'])


class StyleFontTest(unittest.TestCase):
    def test_theme_font_matches_map_font(self):
        s = themed(SOURCE_HAN)
        self.assertEqual(font_warnings(s), [])
        s = themed(SOURCE_HAN, map_font='Source Han Sans SC')        # 写英文名、别的地区版本也算同一字体
        self.assertEqual(font_warnings(s), [])

    def test_theme_font_differs_from_map_font(self):
        typo = dict(SOURCE_HAN, mapFontFamily='微软雅黑', map={'label': ['Source Sans 3', '思源黑体'], 'region': ['方正兰亭黑']})
        s = themed(typo)
        w = font_warnings(s)
        self.assertEqual(len(w), 1)
        self.assertIn('fonts.map “思源黑体”不一致', w[0])
        self.assertIn('mapFontFamily 用 微软雅黑', w[0])                  # 说明哪里用了什么
        self.assertIn('map.region 用 方正兰亭黑', w[0])
        self.assertNotIn('map.label', w[0])
        # 内置主题的西文栈（Inter 后接系统字体）：中文落到系统默认字体，也算不一致
        s = themed({'fontFamily': 'Inter, ui-sans-serif, system-ui, sans-serif'})
        self.assertIn('fontFamily 用 Inter', font_warnings(s)[0])

    def test_map_font_differs_from_page_font(self):
        s = themed(SOURCE_HAN)
        info = [i for i in s.info if '两种黑体' in i]
        self.assertEqual(len(info), 1)
        self.assertIn('style.md', info[0])
        self.assertEqual(s.summary()['info'], s.info)
        doc = style_doc()
        doc['fonts']['map'] = 'Microsoft YaHei'                             # 与 cjk 同一字体，写法不同
        self.assertFalse(any('两种黑体' in i for i in Style(doc).info))
        doc['fonts']['map'] = '思源黑体'                                    # 没写 mapTheme 也提示
        self.assertTrue(any('两种黑体' in i for i in Style(doc).info))

    def test_theme_alias_cannot_be_confirmed(self):
        """内置主题的字体名是别名（GC Sans），认不出是哪个字体：不报不一致，给 info 说放图时核对。"""
        s = themed({'mapFontFamily': 'GC Sans', 'map': {'label': ['GC Sans', 'GC Sans Latin']},
                    'fonts': [{'family': 'GC Sans'}, {'family': 'GC Sans Latin'}]})
        self.assertEqual(font_warnings(s), [])
        self.assertTrue(any('GC Sans、GC Sans Latin 是主题自带字体' in i and '发布版' in i for i in s.info))

    def test_theme_alias_read_from_name_table(self):
        """theme resolve 给了 name 表：别名 GC Sans 按真名认，是思源黑体就不报，也不再说无法确认。"""
        self.assertEqual(name_aliases(HEI_NAMES)[:2], ['Source Han Sans CN', 'Source Han Sans CN Regular'])
        self.assertIn('思源黑体 CN', name_aliases(HEI_NAMES))
        s = themed({'mapFontFamily': 'GC Sans', 'fonts': [{'family': 'GC Sans', 'names': HEI_NAMES}]})
        self.assertEqual(font_warnings(s), [])
        self.assertFalse(any('主题自带字体' in i for i in s.info))

    def test_theme_alias_with_other_name_table_warns(self):
        s = themed({'mapFontFamily': 'GC Sans', 'fonts': [{'family': 'GC Sans', 'names': SONG_NAMES}]})
        self.assertEqual(len(font_warnings(s)), 1)
        self.assertIn('mapFontFamily 用 GC Sans（即 Source Han Serif CN）', font_warnings(s)[0])

    def test_missing_map_font_still_warns(self):
        s = themed(SOURCE_HAN, map_font=None)
        self.assertTrue(any('没写 fonts.map' in w for w in s.warnings))


class ReleaseFontTest(unittest.TestCase):
    FACES = [('GC Sans', 'SourceHanSansCN-Normal.otf', 0), ('GC Sans Medium', 'SourceHanSansCN-Medium.otf', 0),
             ('GC Sans Latin', 'SourceSans3-Regular.otf', 37)]           # 西文字体缺图上的汉字，不算主要字体

    def deck(self, map_font, cjk='微软雅黑'):
        doc = style_doc()
        doc['fonts'] = {'cjk': cjk, 'latin': 'Arial', **({'map': map_font} if map_font else {})}
        return Deck(Style(doc))

    def test_release_font_matches(self):
        with TempDir() as root:
            make_release(root, maps={'m1': (5.0, 5.0, 300)}, fonts={'m1': self.FACES})
            a = MapRelease(root, 'v1').map('m1')
            self.assertEqual({f['name'] for f in a.main_fonts()}, {'SourceHanSansCN'})
            deck = self.deck('思源黑体')
            deck.page('evidence', '2.1').map(a, 2, 6, 6, 6)
            self.assertEqual(deck.warnings, [])

    def test_release_font_differs(self):
        with TempDir() as root:
            make_release(root, maps={'m1': (5.0, 5.0, 300)}, fonts={'m1': self.FACES})
            rel = MapRelease(root, 'v1')
            deck = self.deck('微软雅黑')
            p = deck.page('evidence', '2.1')
            p.map(rel.map('m1'), 2, 6, 6, 6)
            p.map(rel.map('m1'), 10, 6, 6, 6, slot='map2')
            self.assertEqual(len(deck.warnings), 1)                        # 同一张图放两次只记一次
            w = deck.warnings[0]
            self.assertIn('图面实际用的字体是 SourceHanSansCN（主题里叫 GC Sans、GC Sans Medium', w)
            self.assertIn('fonts.map 是“微软雅黑”', w)
            self.assertNotIn('SourceSans3', w)
            # 没写 fonts.map 时地图应当用页面字体 fonts.cjk
            deck = self.deck(None)
            deck.page('evidence', '2.1').map(rel.map('m1'), 2, 6, 6, 6)
            self.assertIn('fonts.cjk 是“微软雅黑”', deck.warnings[0])

    def test_release_font_read_from_name_table(self):
        """发布版记了 name 表：字体名按它取（路径是可移植的相对路径，本机读不到文件也认得出）。"""
        with TempDir() as root:
            make_release(root, maps={'m1': (5.0, 5.0, 300)}, fonts={'m1': [('GC Sans', 'GCSans-Regular.otf', 0, HEI_NAMES)]})
            a = MapRelease(root, 'v1').map('m1')
            self.assertEqual([f['name'] for f in a.fonts], ['Source Han Sans CN'])
            self.assertIsNone(a.font_mismatch('思源黑体'))
            self.assertIn('Source Han Sans CN', a.font_mismatch('微软雅黑'))

    def test_sheet_reads_panel_fonts(self):
        with TempDir() as root:
            make_release(root, maps={'m1': (5.0, 5.0, 300)}, panels={'m1': ('zoom', (5.0, 5.0, 300))},
                         fonts={'m1/zoom': [('GC Sans', 'NotoSansCJKsc-Regular.otf', 0)]})
            rel = MapRelease(root, 'v1')
            self.assertEqual([f['name'] for f in rel.map('m1').fonts], ['NotoSansCJKsc'])
            self.assertEqual([f['name'] for f in rel.map('m1', 'zoom').fonts], ['NotoSansCJKsc'])
            self.assertIsNone(rel.map('m1').font_mismatch('思源黑体'))

    def test_release_without_font_record(self):
        with TempDir() as root:
            make_release(root, maps={'m1': (5.0, 5.0, 300)})
            a = MapRelease(root, 'v1').map('m1')
            self.assertIsNone(a.fonts)
            deck = self.deck('微软雅黑')
            deck.page('evidence', '2.1').map(a, 2, 6, 6, 6)
            self.assertEqual(deck.warnings, [])                           # 没有记录不核对

    def test_maps_check_lists_fonts(self):
        with TempDir() as root:
            make_release(root, maps={'m1': (5.0, 5.0, 300), 'm2': (5.0, 5.0, 300)}, fonts={'m1': self.FACES})
            out = io.StringIO()
            with contextlib.redirect_stdout(out):
                cli(['maps', 'check', root, '--release', 'v1'])
            rows = {r['map']: r for r in json.loads(out.getvalue())['maps']}
            self.assertEqual(rows['m1']['fonts'], ['SourceHanSansCN'])
            self.assertIsNone(rows['m2']['fonts'])


if __name__ == '__main__':
    unittest.main()
