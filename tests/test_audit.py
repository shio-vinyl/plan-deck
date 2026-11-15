"""audit 的版面指标、文字线索和两版文字比对。"""
import io
import json
import os
import unittest
from contextlib import redirect_stdout

from helpers import EXAMPLE_STYLE, TempDir
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Cm

from plan_deck import Deck, audit, compare_text
from plan_deck.cli import main
from plan_deck.inspect import inspect
from plan_deck.words import banned


def deck_of(root, name, pages):
    """pages：每页一组 (x, y, w, h, 填充色, 字)。"""
    prs = Presentation()
    prs.slide_width, prs.slide_height = Cm(33.87), Cm(19.05)
    for boxes in pages:
        sl = prs.slides.add_slide(prs.slide_layouts[6])
        for x, y, w, h, color, text in boxes:
            s = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(x), Cm(y), Cm(w), Cm(h))
            s.fill.solid()
            s.fill.fore_color.rgb = RGBColor.from_string(color)
            s.line.fill.background()
            s.text_frame.text = text
    path = os.path.join(root, name)
    prs.save(path)
    return path


def metric(res, key):
    return next(m for m in res['metrics'] if m['metric'] == key)


ROWS = [(2, y, 4, 2, '1F3864', '栏头文字') for y in (5, 8, 11)] + [(6, y, 20, 2, 'E7EEF4', '内容文字') for y in (5, 8, 11)]


class AuditTest(unittest.TestCase):
    def test_tidy_pages_are_flagged(self):
        with TempDir() as root:
            res = audit(deck_of(root, 'd.pptx', [ROWS, ROWS]))
            self.assertEqual(res['visiblePages'], 2)
            self.assertEqual(metric(res, 'equalSizePages')['value'], 1)
            self.assertEqual(metric(res, 'labelRowPages')['value'], 1)
            self.assertEqual(metric(res, 'similarAdjacentPages')['value'], 1)
            self.assertEqual(metric(res, 'noImagePages')['value'], 1)
            flags = res['pages'][1]['flags']
            self.assertTrue(any('等尺寸' in f for f in flags))
            self.assertTrue(any('左色块' in f for f in flags))
            self.assertTrue(any('上一页' in f for f in flags))
            self.assertTrue(metric(res, 'labelRowPages')['outside'])

    def test_varied_pages_are_not_flagged(self):
        with TempDir() as root:
            a = [(2, 5, 12, 3, 'FFC000', '黄色块'), (15, 5, 16, 9, 'ED7D31', '橙色块'), (2, 9, 8, 6, '00B050', '绿色块')]
            b = [(2, 4, 8, 1, 'C00000', '短条'), (22, 12, 9, 5, '808080', '灰色块'), (12, 15, 6, 2, '2F5597', '蓝色块')]
            res = audit(deck_of(root, 'd.pptx', [a, b]))
            self.assertEqual(metric(res, 'equalSizePages')['value'], 0)
            self.assertEqual(metric(res, 'labelRowPages')['value'], 0)
            self.assertEqual(metric(res, 'similarAdjacentPages')['value'], 0)
            self.assertEqual(metric(res, 'officeColorShare')['value'], 1)

    def test_language_clues(self):
        with TempDir() as root:
            res = audit(deck_of(root, 'd.pptx', [[(2, 5, 10, 2, 'FFFFFF', '动了法定内容的走规划修改'),
                                                  (2, 8, 10, 2, 'FFFFFF', '闲置办公转商业，比例待定'),
                                                  (2, 11, 10, 2, 'FFFFFF', '调整为商业用地，按法定程序修改控规；不动产登记')]]))
            kinds = {(x['kind'], x['word']) for x in res['language']}
            self.assertIn(('讨论口语', '动了'), kinds)
            self.assertIn(('单字动词接术语', '转商业'), kinds)
            self.assertFalse([x for x in res['language'] if '调整为商业' in x['text']])
            self.assertEqual([(x['kind'], x['word']) for x in res['problems']], [('待定和核查用语', '待定')])

    def test_compare_text_ignores_layout(self):
        with TempDir() as root:
            a = deck_of(root, 'a.pptx', [[(2, 5, 10, 2, 'FFFFFF', '甲'), (2, 8, 10, 2, 'FFFFFF', '乙')], [(2, 5, 10, 2, 'FFFFFF', '丙')]])
            b = deck_of(root, 'b.pptx', [[(5, 9, 20, 3, 'FFC000', '乙'), (2, 2, 4, 1, 'ED7D31', '甲')], [(2, 5, 10, 2, 'FFFFFF', '丁')]])
            res = compare_text(a, b)
            self.assertEqual(res['changedPages'], [2])
            self.assertEqual(res['changes'][0]['onlyHere'], ['丙'])
            self.assertEqual(res['changes'][0]['onlyThere'], ['丁'])

    def test_audit_does_not_change_the_file(self):
        with TempDir() as root:
            path = deck_of(root, 'd.pptx', [ROWS])
            with open(path, 'rb') as f:
                before = f.read()
            audit(path)
            with open(path, 'rb') as f:
                self.assertEqual(before, f.read())

    def test_cli(self):
        with TempDir() as root:
            path = deck_of(root, 'd.pptx', [ROWS])
            out = io.StringIO()
            with redirect_stdout(out):
                code = main(['audit', path, '--against', path, '--brief'])
            res = json.loads(out.getvalue())
            self.assertEqual(code, 0)
            self.assertNotIn('pages', res)
            self.assertEqual(res['textCompare']['changedPages'], [])

    def test_inspect_warns_check_words(self):
        with TempDir() as root:
            deck = Deck(EXAMPLE_STYLE)
            p = deck.page('evidence', '2.1 现状用地', parent='二、现状评估')
            p.text(*deck.style.col(0, 6)[:1], deck.style.y0, deck.style.col(0, 6)[1], ['这组数据待核，涉及调整的走备案。'])
            res = inspect(deck.save(os.path.join(root, 'd.pptx')))
            self.assertIn('待定和核查用语“待核”', ' '.join(q['problem'] for q in res['problems']))
            self.assertIn('讨论时的口语', ' '.join(w['warning'] for w in res['warnings']))

    def test_inspect_and_audit_share_word_lists(self):
        """同一批词两边都查：inspect 和 audit 不再各维护一份。"""
        with TempDir() as root:
            deck = Deck(EXAMPLE_STYLE)
            p = deck.page('evidence', '2.1 现状用地', parent='二、现状评估')
            x, w = deck.style.col(0, 6)
            for k, text in enumerate(('状态列已删。', '这段碰红线。', '靠控规修改。')):      # inspect 每个文本框每类只报第一处
                p.text(x, deck.style.y0 + 2 * k, w, [text])
            path = deck.save(os.path.join(root, 'd.pptx'))
            ins, aud = inspect(path), audit(path)
            warned = ' '.join(v['warning'] for v in ins['warnings'])
            found = {x['word'] for x in aud['language']}
            for word in ('碰红线', '靠控规'):
                self.assertIn(word, warned)
                self.assertIn(word, found)
            self.assertIn('状态列', ' '.join(q['problem'] for q in ins['problems']))
            self.assertEqual([x['word'] for x in aud['problems']], ['状态列'])

    def test_banned_words(self):
        """系列统一的清单：每类都报；数据许可署名行放行解说类，跟了免责照样报；平常的规划用语不报。"""
        hits = {
            '示意图，边界以正式图层为准': ['示意图，', '边界以正式图层为准'],
            '范围以批复为准': ['范围以批复为准'],
            '本图仅供参考': ['仅供参考', '本图'],
            '图上界线不作为划界依据': ['不作为划界依据'],
            '不作审批依据': ['不作审批依据'],
            '口径说明：按 2023 年变更调查': ['口径说明'],
            '用地范围经概括处理': ['经概括'],
            '未经核实的数据': ['未经核实'],
            '底图精度有限': ['精度有限'],
            '这一组暂不讨论': ['暂不讨论'],
            '官网标有效': ['官网标'],
            '这条和上一条不冲突': ['不冲突'],
            '这不是分区统计的数字': ['不是分区统计的数字'],
            '注：红线为规划范围': ['注：'],
            '  说明:按现状统计': ['说明:'],
            '此图表达规划结构': ['此图'],
            '图中红色表示规划范围': ['图中红色表示'],
            '数据来源：本图据自然资源局资料编制': [],
            '© OpenStreetMap contributors': [],
            '底图来源：天地图，仅供参考': ['仅供参考'],
        }
        for text, words in hits.items():
            self.assertEqual([w for _, w, _ in banned(text)], words, text)
        for text in ('旧产业区更新潜力空间为 12.6km²', '该图斑现状为工业用地', '用地布局规划图', '符合清单的项目视为符合详细规划',
                     '各地政策在适用情形、核算标准和审批程序等方面规定较细', '注册资本不低于 500 万元',
                     '暂不推荐', '该图斑面积', '不以申请人自报为准', '以人为本'):
            self.assertEqual(banned(text), [], text)
        self.assertEqual([w for _, w, _ in banned('第一行正文\n注：第二行')], ['注：'])

    def test_banned_words_fail_inspect_and_audit(self):
        """页面上的报为问题、退出码 1；备注、隐藏页、版面外不查。"""
        with TempDir() as root:
            deck = Deck(EXAMPLE_STYLE)
            x, w = deck.style.col(0, 6)
            p = deck.page('evidence', '2.1 现状用地', parent='二、现状评估')
            p.text(x, deck.style.y0, w, ['用地范围示意，以正式图层为准。'])
            p.notes('待确认：口径说明放在备注里')
            q = deck.page('evidence', '2.2 用地结构', parent='二、现状评估')
            q.text(x, deck.style.y0, w, ['现状建设用地 86.5 公顷。'])
            q.notes('仅供参考')
            path = deck.save(os.path.join(root, 'd.pptx'))
            for cmd in ('inspect', 'audit'):
                out = io.StringIO()
                with redirect_stdout(out):
                    code = main([cmd, path, '--brief'])
                res = json.loads(out.getvalue())
                self.assertEqual(code, 1, cmd)
                self.assertFalse(res['ok'])
                pages = {q['page'] for q in res['problems']}
                self.assertEqual(pages, {1}, res['problems'])
            prs = Presentation(path)
            prs.slides[0]._element.set('show', '0')
            prs.save(path)
            self.assertEqual(audit(path)['problems'], [])
            self.assertFalse([q for q in inspect(path)['problems'] if '页面绝不出现' in q['problem']])


if __name__ == '__main__':
    unittest.main()
