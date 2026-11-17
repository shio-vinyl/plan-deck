"""渲染检查：把 pptx 转成 pdf 和逐页 png，并列出 pdf 里实际用到的字体。

两种引擎：
- powerpoint：macOS 上用 Microsoft PowerPoint（AppleScript）另存为 PDF，就是汇报现场的样子。
  PowerPoint 是沙盒应用，读写沙盒外的文件会弹授权框，所以先把 pptx 复制进它的容器目录，导出后再取回。
- libreoffice：soffice --headless 转 PDF。缺字体时会换字或漏字，看图前先确认字体可用。
逐页 png 用 pdftoppm（poppler），字体清单用 pdffonts；没装时只给 pdf，并说明。
"""
import os
import re
import shutil
import subprocess
import sys
import tempfile
import uuid

from .fonts import checkable, pdf_font_is

POWERPOINT_APP = '/Applications/Microsoft PowerPoint.app'
POWERPOINT_DATA = os.path.join('~', 'Library', 'Containers', 'com.microsoft.Powerpoint', 'Data')
SCRIPT = '''on run argv
  with timeout of {timeout} seconds
    tell application "Microsoft PowerPoint"
      open POSIX file (item 1 of argv)
      set p to presentation (item 3 of argv)
      save p in POSIX file (item 2 of argv) as save as PDF
      close p saving no
    end tell
  end timeout
end run'''


class RenderError(RuntimeError):
    pass


def _powerpoint_pdf(pptx, pdf, timeout):
    if sys.platform != 'darwin' or not os.path.isdir(POWERPOINT_APP):
        raise RenderError('powerpoint 引擎要在装了 Microsoft PowerPoint 的 macOS 上用；没有时用 --engine libreoffice，'
                          '交付前再请用户在 PowerPoint 里打开核对')
    key = uuid.uuid4().hex
    work = os.path.join(os.path.expanduser(POWERPOINT_DATA), 'plan-deck-render', key)
    os.makedirs(work)
    try:
        # 文件名唯一，按名字找回这份稿子：别的会话同时用 PowerPoint 时，“当前演示文稿”可能是别人的
        src, out = os.path.join(work, f'{key}.pptx'), os.path.join(work, f'{key}.pdf')
        shutil.copyfile(pptx, src)
        res = subprocess.run(['osascript', '-e', SCRIPT.format(timeout=timeout), src, out, f'{key}.pptx'],
                             capture_output=True, text=True, timeout=timeout + 30)
        if res.returncode != 0 or not os.path.isfile(out):
            raise RenderError(f'PowerPoint 导出 PDF 失败：{(res.stderr or res.stdout).strip()[-600:]}'
                              '（第一次运行时 macOS 可能要求允许终端控制 PowerPoint）')
        shutil.move(out, pdf)
    finally:
        shutil.rmtree(work, ignore_errors=True)


def _libreoffice_pdf(pptx, pdf, timeout):
    soffice = shutil.which('soffice') or shutil.which('libreoffice')
    if not soffice:
        raise RenderError('找不到 soffice（LibreOffice）')
    with tempfile.TemporaryDirectory(prefix='plan-deck-render-') as tmp:
        src = os.path.join(tmp, 'deck.pptx')
        shutil.copyfile(pptx, src)
        res = subprocess.run([soffice, '--headless', '--convert-to', 'pdf', '--outdir', tmp, src],
                             capture_output=True, text=True, timeout=timeout)
        out = os.path.join(tmp, 'deck.pdf')
        if res.returncode != 0 or not os.path.isfile(out):
            raise RenderError(f'LibreOffice 转 PDF 失败：{(res.stderr or res.stdout).strip()[-600:]}')
        shutil.move(out, pdf)


def parse_pdffonts(text):
    """pdffonts 的输出 → 去掉子集前缀（ABCDEF+）的字体名，按出现顺序去重。"""
    names = []
    for line in text.splitlines()[2:]:
        if not line.strip():
            continue
        name = re.sub(r'^[A-Z]{6}\+', '', line.split()[0])
        if name not in names:
            names.append(name)
    return names


def missing_fonts(fonts, used):
    """核对 style.json 的页面字体（fonts.cjk、fonts.latin），按 fonts.py 的字体名表比：
    - 在 PDF 里一个字形都没用上的，多半被换成了别的字体（思源黑体换成思源宋体也算）；
    - PDF 里有 style.json 没定的字体，多半是哪里没设字体、落回了默认字体（如图表中文落回宋体）。
    比不了的字体名（见 fonts.checkable）不核对。"""
    out = []
    names = {key: fonts[key] for key in ('cjk', 'latin') if fonts.get(key)}
    for key, name in names.items():
        if checkable(name) and not any(pdf_font_is(name, u) for u in used):
            out.append(f'fonts.{key} {name} 没出现在 PDF 里（PDF 用的是 {"、".join(used) or "无"}），可能被替换')
    if names and all(checkable(n) for n in names.values()):
        extra = [u for u in used if not any(pdf_font_is(n, u) for n in names.values())]
        if extra:
            out.append(f'PDF 里有 style.json 没定的字体：{"、".join(extra)}；多半是哪里没设字体、落回了默认字体，逐页找出来')
    return out


def render(pptx, out_dir, engine='powerpoint', dpi=80, timeout=300, fonts=None):
    """返回 {engine, pdf, pages: [png…], fonts: [字体名…], warnings: [...], notes: [说明…]}。
    fonts 给 style.json 的 fonts 时，核对页面字体是否真的用上了。"""
    if not os.path.isfile(pptx):
        raise RenderError(f'找不到 {pptx}')
    os.makedirs(out_dir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(pptx))[0]
    pdf = os.path.join(out_dir, f'{stem}.{engine}.pdf')
    {'powerpoint': _powerpoint_pdf, 'libreoffice': _libreoffice_pdf}[engine](os.path.abspath(pptx), pdf, timeout)
    result = {'engine': engine, 'pdf': pdf, 'pages': [], 'fonts': None, 'warnings': [], 'notes': []}
    if shutil.which('pdftoppm'):
        prefix = os.path.join(out_dir, f'{stem}.{engine}')
        subprocess.run(['pdftoppm', '-r', str(dpi), '-png', pdf, prefix], check=True, capture_output=True)
        result['pages'] = sorted((os.path.join(out_dir, n) for n in os.listdir(out_dir)
                                  if n.startswith(f'{stem}.{engine}-') and n.endswith('.png')),
                                 key=lambda p: int(re.search(r'-(\d+)\.png$', p).group(1)))
    else:
        result['notes'].append('没装 pdftoppm（poppler），只出了 pdf')
    if shutil.which('pdffonts'):
        res = subprocess.run(['pdffonts', pdf], capture_output=True, text=True, check=True)
        result['fonts'] = parse_pdffonts(res.stdout)
        if fonts:
            result['warnings'] = missing_fonts(fonts, result['fonts'])
            result['notes'] += [f'fonts.{k} {fonts[k]} 认不出 PDF 里的名字，没核对' for k in ('cjk', 'latin')
                                if fonts.get(k) and not checkable(fonts[k])]
    else:
        result['notes'].append('没装 pdffonts（poppler），列不出实际用到的字体')
    return result
