import unittest

from helpers import EXAMPLE_STYLE, style_doc

from plan_deck import Deck, Style, StyleError, text_height


class StyleTest(unittest.TestCase):
    def test_example_loads(self):
        s = Style.load(EXAMPLE_STYLE)
        self.assertEqual(s.color('emphasis'), '#C00000')
        self.assertEqual(s.cls('实施时序', '近期'), '#C00000')       # {colors.emphasis} 解析成色值
        self.assertEqual(s.size('body'), 18)
        x, w = s.col(0, 12)
        self.assertAlmostEqual(x, 1.6)
        self.assertAlmostEqual(x + w, 52.8 - 1.6)
        x6, w6 = s.col(6, 6)
        self.assertAlmostEqual(x6 + w6, 52.8 - 1.6)

    def test_raw_hex_and_unknown_names_are_refused(self):
        s = Style.load(EXAMPLE_STYLE)
        with self.assertRaisesRegex(StyleError, '不写色值'):
            s.color('#123456')
        with self.assertRaisesRegex(StyleError, '可用'):
            s.color('accent2')
        with self.assertRaisesRegex(StyleError, '没有配色方案'):
            s.cls('判档', '重点')
        with self.assertRaisesRegex(StyleError, '里没有'):
            s.cls('实施时序', '近近期')
        with self.assertRaisesRegex(StyleError, '字号级别'):
            s.size('caption')
        with self.assertRaisesRegex(StyleError, '栏位越界'):
            s.col(10, 3)

    def test_text_smaller_than_body_is_an_error(self):
        doc = style_doc()
        doc['type']['caption'] = 10
        with self.assertRaisesRegex(StyleError, '页面上没有比正文小的字'):
            Style(doc)

    def test_body_too_small_for_page(self):
        doc = style_doc()
        doc['type'] = {'title': 20, 'heading': 14, 'body': 12}    # 29.7 cm 高的页面约 18 pt
        with self.assertRaisesRegex(StyleError, '正文至少'):
            Style(doc)

    def test_map_text_may_be_smaller(self):
        s = Style.load(EXAMPLE_STYLE)
        self.assertLess(s.map_text['minorPt'], s.size('body'))

    def test_bad_references(self):
        doc = style_doc()
        doc['classes']['x'] = {'a': '{colors.nope}', 'b': 'red'}
        doc['colors']['loop'] = '{colors.loop}'
        with self.assertRaises(StyleError) as e:
            Style(doc)
        text = '；'.join(e.exception.problems)
        self.assertIn('不存在的颜色 nope', text)
        self.assertIn("'red'", text)
        self.assertIn('成环', text)

    def test_contract_version(self):
        with self.assertRaisesRegex(StyleError, 'schema_version'):
            Style(style_doc(schema_version=2))

    def test_header_refs_checked(self):
        doc = style_doc()
        doc['header']['lines'][0]['size'] = 'huge'
        doc['header']['lines'][1]['color'] = '#194461'
        with self.assertRaises(StyleError) as e:
            Style(doc)
        self.assertEqual(len(e.exception.problems), 2)

    def test_close_levels_warn(self):
        doc = style_doc()
        doc['type'] = dict(doc['type'], heading=19)
        s = Style(doc)
        self.assertTrue(any('1.25' in w for w in s.warnings))


class KeysTest(unittest.TestCase):
    def test_unknown_keys_warn(self):
        doc = style_doc()
        doc['page']['background'] = '#000000'       # 没有的能力
        doc['fonts']['display'] = '思源宋体'
        doc['header']['rulee'] = {}                 # 拼错
        doc['header']['lines'][0]['colour'] = 'title'
        doc['table'] = {'headerFil': 'title'}
        doc['text'] = {'lineSpacng': 1.0}
        doc['mapText']['labelPT'] = 10
        doc['extra'] = 1
        s = Style(doc)
        got = ' '.join(s.warnings)
        for text in ('page.background 运行时不读', 'fonts.display', 'header.rulee', 'header.lines[0].colour', 'table.headerFil',
                     '是不是 headerFill', 'text.lineSpacng', '是不是 lineSpacing', 'mapText.labelPT', 'extra 运行时不读'):
            self.assertIn(text, got)
        self.assertEqual(len(s.warnings), 8)
        # 只跑出页脚本、不跑 style check 的人也看得到；只有这一类，其余 style 警告仍看 style check
        self.assertEqual(Deck(s).warnings, s.warnings)

    def test_known_keys_silent(self):
        doc = style_doc()
        doc['purpose'] = 'practical'
        doc['name'] = '只作记录'
        doc['table'] = {'headerFill': 'title', 'lineColor': 'line', 'linePt': 0.5}
        doc['chart'] = {'text': 'text'}
        doc['header']['bars'] = [{'xCm': 0, 'yCm': 0, 'wCm': 'page', 'hCm': 1, 'color': 'title', 'slot': 'band'}]
        doc['palette'] = {'随便起的名字': '#000000'}
        self.assertEqual(Style(doc).warnings, [])

    def test_missing_keys_are_style_errors(self):
        for path, key in (('page', 'widthCm'), ('fonts', 'cjk'), ('grid', 'columns'), ('body', 'topCm')):
            doc = style_doc()
            del doc[path][key]
            with self.assertRaisesRegex(StyleError, f'{path} 缺少 {key}'):
                Style(doc)
        doc = style_doc()
        del doc['header']['lines'][1]['hCm']
        with self.assertRaisesRegex(StyleError, r'header.lines\[1\] 缺少 hCm'):
            Style(doc)


class MeasureTest(unittest.TestCase):
    def test_ascii_narrower_than_cjk(self):
        self.assertLess(text_height('a' * 40, 5, 12), text_height('汉' * 40, 5, 12))

    def test_wraps(self):
        one = text_height('汉' * 5, 10, 12)
        self.assertGreater(text_height('汉' * 60, 10, 12), 2.5 * one)
        self.assertEqual(text_height('**强调**', 10, 12), text_height('强调', 10, 12))


if __name__ == '__main__':
    unittest.main()
