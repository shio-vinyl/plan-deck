import os
import unittest

from helpers import EXAMPLE_STYLE, TempDir

from plan_deck import Deck
from plan_deck.render import missing_fonts, parse_pdffonts, render

PDFFONTS = '''name                                 type              encoding         emb sub uni object ID
------------------------------------ ----------------- ---------------- --- --- --- ---------
AAAAAC+MicrosoftYaHei-Bold           TrueType          WinAnsi          yes yes yes      8  0
AAAAAE+Arial-BoldMT                  TrueType          WinAnsi          yes yes yes     10  0
AAAAAK+MicrosoftYaHei                TrueType          WinAnsi          yes yes yes     17  0
AAAAAM+MicrosoftYaHei-Bold           TrueType          WinAnsi          yes yes yes     19  0
'''


class RenderTest(unittest.TestCase):
    def test_pdffonts(self):
        used = parse_pdffonts(PDFFONTS)
        self.assertEqual(used, ['MicrosoftYaHei-Bold', 'Arial-BoldMT', 'MicrosoftYaHei'])
        self.assertEqual(missing_fonts({'cjk': '微软雅黑', 'latin': 'Arial'}, used), [])
        self.assertIn('可能被替换', missing_fonts({'cjk': '微软雅黑', 'latin': 'Arial'}, ['PingFangSC-Regular', 'Arial'])[0])
        self.assertEqual(missing_fonts({'cjk': '方正兰亭黑'}, used), [])          # 表里没有的字体不核对
        extra = missing_fonts({'cjk': '微软雅黑', 'latin': 'Arial'}, used + ['SimSun'])
        self.assertEqual(len(extra), 1)
        self.assertIn('SimSun', extra[0])                                       # 落回的默认字体

    @unittest.skipUnless(os.environ.get('PLAN_DECK_RENDER_ENGINE'), '真渲染：设 PLAN_DECK_RENDER_ENGINE=powerpoint 或 libreoffice')
    def test_render(self):
        with TempDir() as root:
            deck = Deck(EXAMPLE_STYLE)
            deck.page('evidence', '2.1 现状用地', lead='论断').text(2, 7, 20, ['正文 Arial'])
            path = deck.save(os.path.join(root, 'deck.pptx'))
            res = render(path, os.path.join(root, 'renders'), os.environ['PLAN_DECK_RENDER_ENGINE'],
                         fonts=deck.style.fonts)
            self.assertTrue(os.path.isfile(res['pdf']))
            self.assertEqual(len(res['pages']), 1)
            self.assertEqual(res['warnings'], [])


if __name__ == '__main__':
    unittest.main()
