"""版面铺满、中间层、箭头压字、标签行过多：inspect 的警告和 audit 的指标（docs/patterns.md“检查”）。"""
import os
import unittest

from helpers import TempDir
from pptx import Presentation
from pptx.util import Cm, Pt

from plan_deck import Deck, audit
from plan_deck.inspect import inspect
from plan_deck.preset import build

LONG = '各地政策在适用情形、核算标准和审批程序等方面规定较细，监管规则相对薄弱。' * 4


def warned(res, key):
    return sorted({w['page'] for w in res['warnings'] if key in w['warning']})


class InspectLayoutTest(unittest.TestCase):
    def test_short_page_and_missing_middle_layer(self):
        deck = Deck(build('banner'))
        x0, y0, x1, y1 = deck.style.body
        p1 = deck.page('a', '1.1 只有正文', parent='1 现状')          # 正文全是 12 pt，铺到底
        p1.text(x0, y0, x1 - x0, [LONG * 2, LONG * 2, LONG * 2, LONG * 2], h=y1 - y0)
        p2 = deck.page('a', '1.2 有中间层', parent='1 现状')          # 标签行：18 pt 标签，铺到底
        p2.rows(x0, y0, x1 - x0, [('甲', LONG), ('乙', LONG), ('丙', LONG)])
        p3 = deck.page('a', '1.3 没铺满', parent='1 现状')
        p3.rows(x0, y0, x1 - x0, [('甲', '正文' * 60)], bottom=False)
        with TempDir() as tmp:
            res = inspect(deck.save(os.path.join(tmp, 'd.pptx')))
        self.assertEqual(warned(res, '缺中间层'), [1])
        self.assertEqual(warned(res, '正文止于'), [3])

    def test_arrow_over_text(self):
        deck = Deck(build('banner'))
        p = deck.page('a', '1.1', parent='1')
        p.text(5, 8, 10, ['一段压在箭头下面的说明文字，字够长。'], slot='note')
        p.arrow(3, 8.3, 18, 8.3, slot='over')           # 从字上横穿
        p.flow(2, 12, 28, ['申请', '审查', '审定'])       # 节点之间的箭头不压字
        with TempDir() as tmp:
            res = inspect(deck.save(os.path.join(tmp, 'd.pptx')))
        hits = [w for w in res['warnings'] if '压过' in w['warning']]
        self.assertEqual([(w['shape'], w['page']) for w in hits], [('pd:a:over', 1)])

    def test_label_row_share(self):
        deck = Deck(build('banner'))
        x0, y0, x1, y1 = deck.style.body
        for n in range(4):
            p = deck.page('a', f'1.{n}', parent='1')
            if n < 2:
                p.rows(x0, y0, x1 - x0, [('甲', LONG), ('乙', LONG)])
            else:
                p.columns(x0, y0, x1 - x0, [('甲', LONG), ('乙', LONG)])
        with TempDir() as tmp:
            path = deck.save(os.path.join(tmp, 'd.pptx'))
            res = inspect(path)
            metrics = {m['metric']: m['value'] for m in audit(path)['metrics']}
        deckwide = [w for w in res['warnings'] if w['shape'] == 'deck']
        self.assertEqual(len(deckwide), 1)
        self.assertIn('第 1、2 页', deckwide[0]['warning'])
        self.assertEqual(metrics['labelRowPages'], 0.5)
        self.assertGreater(metrics['midLayerShare'], 0)
        self.assertEqual(metrics['shortPages'], 0)


class AuditLayoutTest(unittest.TestCase):
    def test_short_and_mid_on_plain_deck(self):
        prs = Presentation()
        prs.slide_width, prs.slide_height = Cm(33.87), Cm(19.05)
        for size, h in ((12, 6), (18, 13)):
            sl = prs.slides.add_slide(prs.slide_layouts[6])
            tb = sl.shapes.add_textbox(Cm(2), Cm(4.5), Cm(29), Cm(h))
            tb.text_frame.word_wrap = True
            r = tb.text_frame.paragraphs[0].add_run()
            r.text = LONG * 2
            r.font.size = Pt(size)
        with TempDir() as tmp:
            path = os.path.join(tmp, 'p.pptx')
            prs.save(path)
            res = audit(path)
        metrics = {m['metric']: m['value'] for m in res['metrics']}
        self.assertEqual(metrics['shortPages'], 0.5)
        self.assertEqual(metrics['midLayerShare'], 0.5)
        flags = {p['page']: p['flags'] for p in res['pages']}
        self.assertTrue(any('正文止于页高' in f for f in flags[1]))
        self.assertTrue(any('缺中间层' in f for f in flags[1]))
        self.assertFalse(any('缺中间层' in f for f in flags[2]))


if __name__ == '__main__':
    unittest.main()
