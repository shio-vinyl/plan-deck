"""compare：两版配对、分类、差异框和审核 PDF。渲染图用假图，不调 PowerPoint。"""
import io
import json
import os
import unittest
from contextlib import redirect_stdout

from helpers import TempDir
from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.util import Cm

from plan_deck import CompareError, compare
from plan_deck.cli import main
from plan_deck.compare import diff_boxes, kind, pair, slides


def deck(root, name, pages, hidden=()):
    """pages：每页一组段落文字；第一段放在标题位置。"""
    prs = Presentation()
    prs.slide_width, prs.slide_height = Cm(33.87), Cm(19.05)
    for n, lines in enumerate(pages, 1):
        sl = prs.slides.add_slide(prs.slide_layouts[6])
        box = sl.shapes.add_textbox(Cm(2), Cm(2), Cm(28), Cm(14))
        box.name = 'pd:header:title'
        box.text_frame.text = lines[0]
        for t in lines[1:]:
            box.text_frame.add_paragraph().text = t
        if n in hidden:
            sl._element.set('show', '0')
    path = os.path.join(root, name)
    prs.save(path)
    return path


def fake_renders(root, name, texts):
    """每张可见页一张图，图上画一块随文字变化的色块，模拟 render 的输出名。"""
    d = os.path.join(root, name)
    os.makedirs(d, exist_ok=True)
    for k, t in enumerate(texts, 1):
        im = Image.new('RGB', (400, 225), 'white')
        h = sum(map(ord, t)) % 150
        ImageDraw.Draw(im).rectangle([40, 40 + h % 60, 200, 120 + h % 60], fill=(200, 30, 30))
        im.save(os.path.join(d, f'deck.powerpoint-{k:02d}.png'))
    return d


OLD = [['1.1 背景', '现状用地以工业为主'],
       ['2.1 定义', '兼容不超过 30%', '备案制办理'],
       ['2.2 边界', '30% 以内兼容'],
       ['2.3 清单', '产业升级清单'],
       ['备用 广州做法', '四张正面清单', '临时转换联审']]
NEW = [['1.1 背景', '现状用地以工业为主'],
       ['2.1 定义', '兼容 30% 以内可永久转换', '备案制办理'],     # 改过
       ['2.2 边界', '30% 以内兼容'],
       ['2.3 永久与临时', '超过 30% 只能临时转换'],              # 新增
       ['2.4 清单', '产业升级清单'],                            # 只改编号
       ['备用 广州做法', '四张正面清单', '临时转换由区政府联审']]  # 改过（同位置）


class CompareTest(unittest.TestCase):
    def test_pair_and_kind(self):
        with TempDir() as root:
            so, sn = slides(deck(root, 'a.pptx', OLD)), slides(deck(root, 'b.pptx', NEW))
            pairs, added, deleted = pair(so, sn)
            got = {(o['no'], n['no']): kind(o, n) for o, n in pairs}
            self.assertEqual(got[(1, 1)], 'same')
            self.assertEqual(got[(2, 2)], 'changed')
            self.assertEqual(got[(4, 5)], 'renumber')
            self.assertEqual(got[(5, 6)], 'changed')
            self.assertEqual([n['no'] for n in added], [4])
            self.assertEqual(deleted, [])

    def test_moved_page(self):
        with TempDir() as root:
            old = OLD + [['备用 深圳做法', '适建用途表', '功能转换正面清单']]
            new = [OLD[0], ['备用 深圳做法', '适建用途表', '功能转换正面清单（按是否补缴分档）']] + OLD[1:]
            pairs, added, deleted = pair(slides(deck(root, 'a.pptx', old)), slides(deck(root, 'b.pptx', new)))
            self.assertIn((6, 2), {(o['no'], n['no']) for o, n in pairs})
            self.assertEqual((added, deleted), ([], []))

    def test_deleted_page(self):
        with TempDir() as root:
            pairs, added, deleted = pair(slides(deck(root, 'a.pptx', OLD)),
                                         slides(deck(root, 'b.pptx', OLD[:2] + OLD[3:])))
            self.assertEqual([o['no'] for o in deleted], [3])

    def test_diff_boxes(self):
        a = Image.new('RGB', (300, 200), 'white')
        b = a.copy()
        ImageDraw.Draw(b).rectangle([100, 50, 160, 90], fill='black')
        boxes = diff_boxes(a, b)
        self.assertEqual(len(boxes), 1)
        x0, y0, x1, y1 = boxes[0]
        self.assertTrue(x0 <= 100 and y0 <= 50 and x1 >= 160 and y1 >= 90)
        self.assertEqual(diff_boxes(a, a.copy()), [])

    def test_pdf_with_hidden_page(self):
        with TempDir() as root:
            old = deck(root, 'v4.pptx', OLD, hidden={5})
            new = deck(root, 'v5.pptx', NEW, hidden={6})
            ro = fake_renders(root, 'r4', ['\n'.join(p) for p in OLD[:4]])
            rn = fake_renders(root, 'r5', ['\n'.join(p) for p in NEW[:5]])
            out = os.path.join(root, '审核', '对比.pdf')
            res = compare(old, new, out, ro, rn, labels=('v4', 'v5'))
            self.assertTrue(os.path.isfile(out))
            self.assertEqual([(c['old'], c['new']) for c in res['changed']], [(2, 2), (5, 6)])
            self.assertEqual([a['new'] for a in res['added']], [4])
            self.assertEqual(res['renumbered'], [[4, 5]])
            self.assertEqual(res['sheets'], 1 + 2 + 1)                   # 清单、两页改动、一页新增

    def test_render_count_mismatch(self):
        with TempDir() as root:
            old, new = deck(root, 'a.pptx', OLD), deck(root, 'b.pptx', NEW)
            ro = fake_renders(root, 'ra', ['x'] * 4)                     # 少一张
            rn = fake_renders(root, 'rb', ['x'] * 6)
            with self.assertRaises(CompareError):
                compare(old, new, os.path.join(root, 'o.pdf'), ro, rn)

    def test_cli(self):
        with TempDir() as root:
            old, new = deck(root, 'a.pptx', OLD), deck(root, 'b.pptx', NEW)
            ro = fake_renders(root, 'ra', ['\n'.join(p) for p in OLD])
            rn = fake_renders(root, 'rb', ['\n'.join(p) for p in NEW])
            out = os.path.join(root, 'o.pdf')
            buf = io.StringIO()
            with redirect_stdout(buf):
                code = main(['compare', old, new, '--out', out, '--old-renders', ro, '--new-renders', rn, '--labels', 'v4,v5'])
            res = json.loads(buf.getvalue())
            self.assertEqual(code, 0)
            self.assertTrue(res['ok'])
            self.assertEqual(len(res['changed']), 2)
            with redirect_stdout(io.StringIO()):
                self.assertEqual(main(['compare', old, new, '--out', out, '--old-renders', ro, '--new-renders', rn,
                                       '--labels', 'v4']), 1)


if __name__ == '__main__':
    unittest.main()
