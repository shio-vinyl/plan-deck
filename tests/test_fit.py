"""生成时的版面警告：框高不够、排出页面（docs/design-principles.md 决策 11）。"""
import unittest

from helpers import EXAMPLE_STYLE

from plan_deck import Deck


class FitTest(unittest.TestCase):
    def setUp(self):
        self.deck = Deck(EXAMPLE_STYLE)
        self.page = self.deck.page('evidence', '2.1 现状')
        self.deck.warnings.clear()

    def test_label_too_short(self):
        box = self.page.label(2, 8, 30, 2, '98.4', 'big', size='figure')
        self.assertEqual(self.deck.warnings, [])            # 44 pt 一行约 2.2 cm，差不到半行
        self.assertEqual(box.h, 2)                          # 只警告，不改版面
        self.page.label(2, 8, 5, 2, '一段放不下的长标签文字，要折成好几行才能排完', 'lab')
        self.assertEqual(len(self.deck.warnings), 1)
        self.assertIn('第 1 页 evidence:lab', self.deck.warnings[0])
        self.assertIn('框只有 2.0 cm', self.deck.warnings[0])

    def test_text_given_height(self):
        self.page.text(2, 8, 7, ['窄栏长文。' * 40], h=1)
        self.assertIn('框只有 1.0 cm', self.deck.warnings[0])

    def test_text_runs_off_page(self):
        box = self.page.text(2, 8, 7, ['窄栏长文。' * 200])
        self.assertGreater(box.bottom, 29.7)
        self.assertEqual(len(self.deck.warnings), 1)
        self.assertIn('超出页高 29.7 cm', self.deck.warnings[0])

    def test_table_runs_off_page(self):
        self.page.table(2, 25, 20, [['a', 'b']] + [['甲', '乙']] * 20, [1, 1])
        self.assertIn('evidence:table', self.deck.warnings[0])
        self.assertIn('超出页高', self.deck.warnings[0])

    def test_ledger_operator_box_not_flagged(self):
        self.page.ledger(2, 8, [(1.0, '甲', 'title'), '+', (2.0, '乙', 'title')], 5, 3)
        self.assertEqual(self.deck.warnings, [])

    def test_header_and_normal_page_silent(self):
        deck = Deck(EXAMPLE_STYLE)
        p = deck.page('evidence', '2.1 现状用地', parent='二、现状评估与问题识别', lead='近期 3 处')
        p.text(2, 7, 20, [('h', '现状'), '老旧小区 12 个。', ('m', '说明'), ('big', '98.4')])
        self.assertEqual(deck.warnings, [])

    def test_header_title_wraps(self):
        deck = Deck(EXAMPLE_STYLE)
        deck.page('evidence', '一个长到一行放不下的标题' * 12)
        self.assertIn('evidence:title', deck.warnings[0])


if __name__ == '__main__':
    unittest.main()
