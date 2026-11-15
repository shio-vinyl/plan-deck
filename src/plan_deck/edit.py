"""改已有 pptx 用的小工具：保留格式替换文字、取备注文本框。"""
import copy

A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'


def _replace_in_par(p, old, new):
    """在一个段落里替换。文字跨 run 时，新文字放进第一个涉及的 run，其余 run 删去相应部分。"""
    runs = p.findall(A + 'r')
    count = start = 0
    while True:
        ts = [r.find(A + 't') for r in runs]
        full = ''.join(t.text or '' for t in ts)
        at = full.find(old, start)
        if at < 0:
            break
        end, pos, placed = at + len(old), 0, False
        for r, t in zip(runs, ts):
            x = t.text or ''
            lo, hi = max(at - pos, 0), min(end - pos, len(x))
            pos += len(x)
            if lo >= hi:
                continue
            t.text = x[:lo] + ('' if placed else new) + x[hi:]
            placed = True
            rpr = r.find(A + 'rPr')
            if rpr is not None:
                rpr.set('lang', 'zh-CN')
                rpr.attrib.pop('err', None)
        count += 1
        start = at + len(new)
    if count:
        for r in runs:
            if not r.find(A + 't').text and len(p.findall(A + 'r')) > 1:
                p.remove(r)
    return count


def replace_text(target, old, new):
    """替换一页（或一个形状）里的文字，保留格式，返回替换次数。

    target 是 slide 或 shape，按段落查找，所以文字跨 run 也能替换；替换过的 run 标 lang="zh-CN"，
    否则 PowerPoint 不按中文规则断行，标点会落到行首。没找到时返回 0，由调用方决定是否报错。
    """
    root = target._element
    return sum(_replace_in_par(p, old, new) for p in root.iter(A + 'p'))


def notes_frame(prs, slide):
    """取一页的备注文本框。备注页没有正文占位符时（notes_text_frame 是 None），从别页复制一个空的。"""
    ns = slide.notes_slide
    if ns.notes_text_frame is not None:
        return ns.notes_text_frame
    for other in prs.slides:
        if other.has_notes_slide and other.notes_slide.notes_text_frame is not None:
            el = copy.deepcopy(other.notes_slide.notes_placeholder._element)
            break
    else:
        raise ValueError('全册没有带正文占位符的备注页，无处复制')
    tree = ns.shapes._spTree
    ids = [int(x) for x in tree.xpath('.//p:cNvPr/@id')]
    el.xpath('.//p:cNvPr')[0].set('id', str(max(ids + [1]) + 1))
    tree.append(el)
    ns.notes_text_frame.text = ''
    return ns.notes_text_frame
