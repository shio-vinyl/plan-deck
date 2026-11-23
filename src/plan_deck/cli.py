"""plan-deck 命令行。输出 JSON（--text 时输出给人看的摘要）；有错时退出码 1。

  plan-deck style check <style.json>
  plan-deck style init <style.json> --preset banner [--page 52.8x29.7] [--name 项目名] [--force]
  plan-deck maps check <composer 项目> --release <版本> [--strict]
  plan-deck inspect <deck.pptx> [--style <style.json>] [--brief]
  plan-deck audit <deck.pptx> [--against <另一版.pptx>] [--brief]
  plan-deck render <deck.pptx> --out <目录> [--engine powerpoint|libreoffice] [--style <style.json>]
  plan-deck clean <deck.pptx> <副本.pptx> [--pattern <正则>] [--drop-hidden] [--force]
  plan-deck compare <旧.pptx> <新.pptx> --out <审核对比.pdf> [--old-renders <目录>] [--new-renders <目录>] [--labels v4,v5] [--engine powerpoint|libreoffice]
  plan-deck offcanvas <deck.pptx>… --out <目录> [--render 32,58|all] [--engine libreoffice|powerpoint] [--crop 0.2]
"""
import argparse
import json
import sys

from . import __version__
from .audit import audit
from .clean import clean
from .compare import CompareError, compare
from .inspect import inspect
from .maps import MapError, MapRelease
from .offcanvas import CROP_MIN, run as run_offcanvas
from .render import RenderError, render
from . import preset
from .style import Style, StyleError


def _utf8_output():
    """The installed `plan-deck` command also writes UTF-8, not the Windows code page; a test's StringIO has no reconfigure."""
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, 'reconfigure'):
            stream.reconfigure(encoding='utf-8', errors='backslashreplace')


def _out(obj):
    print(json.dumps(obj, ensure_ascii=False, indent=2))


def main(argv=None):
    _utf8_output()
    ap = argparse.ArgumentParser(prog='plan-deck', description='规划汇报 PPT 运行时')
    ap.add_argument('--version', action='version', version=__version__)
    sub = ap.add_subparsers(dest='cmd', required=True)
    st = sub.add_parser('style', help='项目视觉系统').add_subparsers(dest='sub', required=True)
    sc = st.add_parser('check', help='校验 style.json，列出解析后的颜色、类别配色、栏位和警告')
    sc.add_argument('path')
    si = st.add_parser('init', help='从预设生成项目的 style.json：页眉、字号阶梯、按含义命名的配色都写好，换页面尺寸时等比换算')
    si.add_argument('path', help='要写出的 style.json；已存在时不覆盖')
    si.add_argument('--preset', required=True, choices=preset.names())
    si.add_argument('--page', help='页面尺寸 宽x高（cm），如 52.8x29.7；不给就用预设的 33.867x19.05')
    si.add_argument('--name', help='写进 name 的项目名（只作记录）')
    si.add_argument('--force', action='store_true', help='已存在时覆盖')
    mp = sub.add_parser('maps', help='地图发布版').add_subparsers(dest='sub', required=True)
    mc = mp.add_parser('check', help='逐张核对发布版的图：哈希、legend.json 版本、PNG 像素与 print')
    mc.add_argument('project')
    mc.add_argument('--release', required=True)
    mc.add_argument('--strict', action='store_true', help='发布版里有未过 QA（或没有 QA 结论）的图时报错，默认只警告')
    ins = sub.add_parser('inspect', help='列出 pptx 的页面与形状，并检查页面绝不出现的用语、小字、重名、画布外对象、悬空的动画引用；警告放不下和被盖住的内容')
    ins.add_argument('path')
    ins.add_argument('--style')
    ins.add_argument('--brief', action='store_true', help='只列每页摘要和问题，不列形状')
    au = sub.add_parser('audit', help='只读体检：页面绝不出现的用语报为问题；全册版面和文字的“过于整齐”线索，逐页列出；给 --against 时比对两版文字')
    au.add_argument('path')
    au.add_argument('--against', help='另一版 pptx：逐页比文字，确认去 AI 味前后文字没变，或找出工作稿改过哪些页')
    au.add_argument('--brief', action='store_true', help='只列全册指标、用语问题、文字线索和比对结果，不列每页')
    cl = sub.add_parser('clean', help='另存交付副本：清掉备注里的过程信息，列出多页相同的讲稿和页面上的过程信息')
    cl.add_argument('path')
    cl.add_argument('out')
    cl.add_argument('--pattern', help='另外要清的内部说法（正则），按句删除')
    cl.add_argument('--drop-hidden', action='store_true', help='删除隐藏页；隐藏页的去留要先问用户')
    cl.add_argument('--force', action='store_true', help='副本已存在时覆盖')
    rd = sub.add_parser('render', help='转成 pdf 和逐页 png，列出实际用到的字体（交付前用 PowerPoint 引擎核对）')
    rd.add_argument('path')
    rd.add_argument('--out', required=True, help='输出目录（渲染图可随时删）')
    rd.add_argument('--engine', choices=('powerpoint', 'libreoffice'), default='powerpoint')
    rd.add_argument('--dpi', type=int, default=80)
    rd.add_argument('--style', help='给了就核对 style.json 的页面字体是否被替换')
    cp = sub.add_parser('compare', help='改稿审核：两版逐页配对，改过的页出“旧 | 新”并排图（红框标出变化），新增页、删除页单列，合成一份 PDF 给人审')
    cp.add_argument('old', help='上一版 pptx')
    cp.add_argument('new', help='这一版 pptx')
    cp.add_argument('--out', required=True, help='审核 PDF；不给渲染目录时，渲染图放在它旁边的 <同名>.renders/')
    cp.add_argument('--old-renders', help='上一版已有的渲染目录（render 的输出，只放这一版的逐页 png）')
    cp.add_argument('--new-renders', help='这一版已有的渲染目录')
    cp.add_argument('--labels', default='旧版,新版', help='两版的叫法，逗号分开，如 v4,v5')
    cp.add_argument('--engine', choices=('powerpoint', 'libreoffice'), default='powerpoint')
    cp.add_argument('--dpi', type=int, default=80)
    oc = sub.add_parser('offcanvas', help='版面外审计：接手已有 pptx 时先跑。逐页列版面外和跨边对象的全部内容，导出图片，原 pptx 不改')
    oc.add_argument('path', nargs='+', help='一份或多份 pptx；读不了的文件报告没审计，退出码 1')
    oc.add_argument('--out', required=True, help='输出目录：offcanvas.json、版面外清单.md、全文提取.md、assets/；多份时各一个子目录')
    oc.add_argument('--render', metavar='页', help='连版面外一起渲染这些页，红框是版面边界：32,58 或 24-31；all 是有版面外或跨边内容的页')
    oc.add_argument('--engine', choices=('libreoffice', 'powerpoint'), default='libreoffice', help='--render 用的引擎')
    oc.add_argument('--dpi', type=int, default=100)
    oc.add_argument('--crop', type=float, default=CROP_MIN, help='裁掉的面积占原图这个比例以上的图片列为线索，默认 0.2')
    a = ap.parse_args(argv)

    try:
        if a.cmd == 'style' and a.sub == 'init':
            page = None
            if a.page:
                try:
                    page = tuple(float(v) for v in a.page.lower().replace('×', 'x').split('x'))
                    assert len(page) == 2 and min(page) > 0
                except (ValueError, AssertionError):
                    _out({'ok': False, 'error': 'BAD_PAGE', 'message': f'--page 写成 宽x高（cm），如 52.8x29.7，实际是 {a.page!r}'})
                    return 1
            try:
                s = preset.init(a.preset, a.path, page, a.name, a.force)
            except FileExistsError as e:
                _out({'ok': False, 'error': 'EXISTS', 'message': str(e)})
                return 1
            _out({'ok': True, 'written': a.path, 'preset': a.preset, **s.summary()})
            return 0
        if a.cmd == 'style':
            s = Style.load(a.path)
            _out({'ok': True, **s.summary()})
            return 0
        if a.cmd == 'maps':
            rel = MapRelease(a.project, a.release, strict=a.strict)
            rows = rel.check()
            ok = bool(rows) and all(r['status'] == 'ok' for r in rows)
            _out({'ok': ok, 'release': a.release, 'status': rel.status, 'sourceRun': rel.source_run, 'strict': a.strict,
                  'warnings': rel.warnings, 'maps': rows,
                  **({} if rows else {'error': 'NO_MAPS', 'message': f'发布版 {a.release} 的 output/ 里没有可读的图，什么都没核对'})})
            return 0 if ok else 1
        if a.cmd == 'inspect':
            res = inspect(a.path, Style.load(a.style) if a.style else None)
            if a.brief:
                for p in res['pages']:
                    p.pop('items')
            _out({'ok': not res['problems'], **res})
            return 0 if not res['problems'] else 1
        if a.cmd == 'audit':
            res = audit(a.path, a.against)
            if a.brief:
                res.pop('pages')
            _out({'ok': not res['problems'], **res})
            return 0 if not res['problems'] else 1
        if a.cmd == 'clean':
            try:
                res = clean(a.path, a.out, a.pattern, a.drop_hidden, a.force)
            except ValueError as e:
                _out({'ok': False, 'error': 'CLEAN_REFUSED', 'message': str(e)})
                return 1
            _out({'ok': True, **res})
            return 0
        if a.cmd == 'compare':
            labels = [s.strip() for s in a.labels.split(',')]
            if len(labels) != 2 or not all(labels):
                _out({'ok': False, 'error': 'BAD_LABELS', 'message': f'--labels 写成 旧,新，如 v4,v5，实际是 {a.labels!r}'})
                return 1
            try:
                res = compare(a.old, a.new, a.out, a.old_renders, a.new_renders, a.engine, a.dpi, tuple(labels))
            except (CompareError, RenderError) as e:
                _out({'ok': False, 'error': 'COMPARE_FAILED', 'message': str(e)})
                return 1
            _out({'ok': True, **res})
            return 0
        if a.cmd == 'offcanvas':
            res = run_offcanvas(a.path, a.out, a.crop, a.render, a.engine, a.dpi)
            _out(res)
            return 0 if res['ok'] else 1
        if a.cmd == 'render':
            fonts = None
            if a.style:
                with open(a.style, encoding='utf-8') as f:
                    fonts = json.load(f).get('fonts')         # 只取字体，不解析主题
            res = render(a.path, a.out, a.engine, a.dpi, fonts=fonts)
            _out({'ok': not res['warnings'], **res})
            return 0 if not res['warnings'] else 1
    except RenderError as e:
        _out({'ok': False, 'error': 'RENDER_FAILED', 'message': str(e)})
        return 1
    except StyleError as e:
        _out({'ok': False, 'error': 'STYLE_INVALID', 'problems': e.problems})
        return 1
    except MapError as e:
        _out({'ok': False, 'error': 'MAP_INVALID', 'message': str(e)})
        return 1
    return 2


if __name__ == '__main__':
    sys.exit(main())
