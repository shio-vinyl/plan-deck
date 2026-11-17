"""按字宽估文字排出来的高度，后面的元素据此往下排。

给了字体名且本机找得到字体文件时，按真实字宽折行（含中文避头：句末标点不落行首，带上一个字下行）；
找不到时按估算：汉字和全角符号 1 个字宽，ASCII 0.6 个字宽。估低了 PowerPoint / LibreOffice 会把文本框撑高、
压到下面的元素，所以宁可略高；结果只用来定位，不改字号。
环境变量 PLAN_DECK_MEASURE=estimate 时一律按估算（要在不同机器上结果一致时用）。
"""
import os
import re
import sys

PT_CM = 2.54 / 72          # 1 pt 合多少 cm
LINE = 1.2                 # 行高 / 字号
MARK = re.compile(r'(\*\*|__|\^\^)')    # 加粗记号，见 Page.runs
NO_LINE_START = set('，。、；：！？）》”’〕】」』％,.;:!?)]}%')
# 常见页面字体的文件名（常规, 粗体）；PLAN_DECK_FONT_DIR 指定的目录先找
FONT_FILES = {
    '微软雅黑': (('msyh.ttc', 'msyh.ttf'), ('msyhbd.ttc', 'msyhbd.ttf')),
    'microsoft yahei': (('msyh.ttc', 'msyh.ttf'), ('msyhbd.ttc', 'msyhbd.ttf')),
    '思源黑体': (('SourceHanSansSC-Regular.otf', 'SourceHanSansCN-Regular.otf', 'NotoSansCJKsc-Regular.otf', 'NotoSansSC-Regular.otf'),
             ('SourceHanSansSC-Bold.otf', 'SourceHanSansCN-Bold.otf', 'NotoSansCJKsc-Bold.otf', 'NotoSansSC-Bold.otf')),
    '等线': (('Deng.ttf',), ('Dengb.ttf',)),
    '黑体': (('simhei.ttf',), ('simhei.ttf',)),
}
_FONTS = {}
PX = 400                   # 量宽时取的字号（px），量出来的宽度按它换算


def text_units(text):
    """一段文字合多少个汉字宽（估算）。加粗记号不占宽。"""
    return sum(0.6 if ord(ch) < 0x2000 else 1.0 for ch in MARK.sub('', text))


def _font_dirs():
    home = os.path.expanduser('~')
    dirs = [os.environ.get('PLAN_DECK_FONT_DIR', '')]
    if sys.platform == 'darwin':
        dirs += [os.path.join(home, 'Library', 'Fonts'), '/Library/Fonts', '/System/Library/Fonts',
                 '/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts']
    elif os.name == 'nt':
        windir = os.environ.get('WINDIR', r'C:\Windows')
        dirs += [os.path.join(windir, 'Fonts'), os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Microsoft', 'Windows', 'Fonts')]
    else:
        dirs += [os.path.join(home, '.fonts'), os.path.join(home, '.local', 'share', 'fonts'), '/usr/share/fonts',
                 '/usr/local/share/fonts']
    return [d for d in dirs if d and os.path.isdir(d)]


def font(family, bold=False):
    """字体名 → PIL 字体（取 PX 号字）；不认识的字体、本机找不到或要求估算时为 None。"""
    if not family or os.environ.get('PLAN_DECK_MEASURE') == 'estimate':
        return None
    key = (family, bold)
    if key not in _FONTS:
        _FONTS[key] = None
        names = FONT_FILES.get(family) or FONT_FILES.get(str(family).casefold())
        if names:
            from PIL import ImageFont
            wanted = names[1 if bold else 0]
            for d in _font_dirs():
                for root, _, files in os.walk(d):
                    hit = next((f for f in wanted if f in files), None)
                    if not hit:
                        continue
                    try:
                        _FONTS[key] = ImageFont.truetype(os.path.join(root, hit), PX)
                        break
                    except OSError:
                        continue
                if _FONTS[key] is not None:
                    break
    return _FONTS[key]


def _segments(para, bold):
    """[(字, 是否加粗)]：** __ ^^ 之间的字加粗，记号本身不占宽。"""
    out, on = [], False
    for k, tok in enumerate(MARK.split(para)):
        if k % 2:
            on = not on
            continue
        out += [(ch, bold or on) for ch in tok]
    return out


def text_lines(text, width_cm, size_pt, family=None, bold=False):
    """排在宽 width_cm 里要几行。给了 family 且找得到字体时按真实字宽，否则估算。"""
    usable = max(width_cm - 0.15, 0.5)
    regular = font(family, False)
    if regular is None:
        em = size_pt * PT_CM
        return int(sum(max(1, -(-text_units(para) * em // usable)) for para in text.split('\n')))
    heavy = font(family, True) or regular
    limit = usable / PT_CM * PX / size_pt * 0.98        # 留 2% 余量：PowerPoint 和 LibreOffice 折行略有出入
    total = 0
    for para in text.split('\n'):
        chars = [(ch, heavy if b else regular) for ch, b in _segments(para, bold)]
        lines, cur = 1, 0.0
        for i, (ch, f) in enumerate(chars):
            w = f.getlength(ch)
            if cur + w > limit and cur > 0:
                lines += 1
                cur = 0.0
                if ch in NO_LINE_START and i:      # 标点不落行首，带上一个字下行
                    cur = chars[i - 1][1].getlength(chars[i - 1][0])
            cur += w
        total += lines
    return total


def text_height(text, width_cm, size_pt, spacing=1.15, family=None, bold=False):
    """一段文字排在宽 width_cm 里的高度（cm），含上下各约 0.05 cm 的边距。"""
    return text_lines(text, width_cm, size_pt, family, bold) * size_pt * PT_CM * LINE * spacing + 0.1


def text_width(text, size_pt, family=None, bold=False):
    """一行文字的宽度（cm），不折行；找不到字体时估算。"""
    f = font(family, bold)
    if f is None:
        return text_units(text) * size_pt * PT_CM
    heavy = font(family, True) or f
    return sum((heavy if b else f).getlength(ch) for ch, b in _segments(text, bold)) / PX * size_pt * PT_CM
