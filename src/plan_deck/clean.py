"""另存交付副本：清掉备注里的过程信息，页面不动。讲稿保留，隐藏页默认保留。

备注按行处理：带标签的行（[改稿说明] 等）和“待确认：”开头的行整行删；其余行按句号分句，删掉含内部说法的句子。
多页相同的讲稿要对照页面判断属于哪页，这里只列出来；页面文字里的过程信息也只列出来，不改页面。
"""
import os
import re

from pptx import Presentation

from .offcanvas import placements

TAG = re.compile(r'^\s*\[[^\]\n]{1,16}\]')
WHOLE = re.compile(r'^\s*待确认[:：]')
INTERNAL = '负责人|甲方|待核|待确认|需核对|草拟'
ON_PAGE = re.compile(r'待确认|待核|负责人|甲方|改稿|讨论稿|旧页|同事意见')
DUP_MIN = 20        # 这么多字以上的备注行才算重复，短句相同可能只是巧合
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'


def clean_line(line, internal):
    if TAG.match(line) or WHOLE.match(line):
        return ''
    return ''.join(x for x in re.findall(r'[^。]+。?', line) if not internal.search(x))


def _clean_frame(tf, internal):
    changed = 0
    for par in list(tf.paragraphs):
        old = par.text
        new = clean_line(old, internal)
        if new == old:
            continue
        changed += 1
        if new:
            runs = par.runs
            runs[0].text = new
            for r in runs[1:]:
                par._p.remove(r._r)
        elif len(tf.paragraphs) > 1:
            par._p.getparent().remove(par._p)
        else:
            for r in par.runs:
                par._p.remove(r._r)
    return changed


def duplicate_notes(prs):
    """可见页里相同的备注行：[{'text', 'pages'}]。"""
    seen = {}
    for n, sl in enumerate(prs.slides, 1):
        if sl._element.get('show') == '0' or not sl.has_notes_slide or sl.notes_slide.notes_text_frame is None:
            continue
        for line in {x.strip() for x in sl.notes_slide.notes_text_frame.text.split('\n')}:
            if len(line) >= DUP_MIN:
                seen.setdefault(line, []).append(n)
    return [{'text': k[:60], 'pages': v} for k, v in seen.items() if len(v) > 1]


def _page_hits(prs):
    hits = []
    for n, sl in enumerate(prs.slides, 1):
        outside = {el for el, place in placements(sl.shapes, prs.slide_width, prs.slide_height).items() if place == 'outside'}
        for p in sl._element.iter(A + 'p'):
            text = ''.join(t.text or '' for t in p.iter(A + 't'))
            m = ON_PAGE.search(text)
            if m:
                hit = {'page': n, 'hidden': sl._element.get('show') == '0', 'word': m.group(), 'text': text[:60]}
                if any(a in outside for a in p.iterancestors()):
                    hit['offCanvas'] = True     # 不在渲染图上，但会跟着文件发出去
                hits.append(hit)
    return hits


def clean(src, out, pattern=None, drop_hidden=False, force=False):
    if os.path.abspath(src) == os.path.abspath(out):
        raise ValueError('副本不能和工作稿同名，工作稿要留着继续改')
    if os.path.exists(out) and not force:
        raise ValueError(f'{out} 已存在，对方可能开着；换个名字，或确认后加 --force')
    internal = re.compile(INTERNAL + (f'|{pattern}' if pattern else ''))
    prs = Presentation(src)
    changed = 0
    for sl in prs.slides:
        if sl.has_notes_slide and sl.notes_slide.notes_text_frame is not None:
            changed += _clean_frame(sl.notes_slide.notes_text_frame, internal)
    dropped = 0
    if drop_hidden:
        ids = prs.slides._sldIdLst
        for sld_id, sl in list(zip(ids, prs.slides)):
            if sl._element.get('show') == '0':
                prs.part.drop_rel(sld_id.rId)
                ids.remove(sld_id)
                dropped += 1
    if 'python-pptx' in (prs.core_properties.comments or ''):
        prs.core_properties.comments = ''
    report = {'file': str(out), 'slides': len(ids) if drop_hidden else len(prs.slides._sldIdLst),
              'hidden': sum(1 for s in prs.slides if s._element.get('show') == '0'), 'hiddenDropped': dropped,
              'notesChanged': changed, 'duplicateNotes': duplicate_notes(prs), 'pageHits': _page_hits(prs)}
    prs.save(out)
    return report
