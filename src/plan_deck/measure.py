"""按字宽估文字排出来的高度，后面的元素据此往下排。

汉字和全角符号按 1 个字宽，ASCII 按 0.6 个字宽。估低了 PowerPoint / LibreOffice 会把文本框撑高、压到下面的元素，
所以宁可略高；结果只用来定位，不改字号。
"""

import re

PT_CM = 2.54 / 72          # 1 pt 合多少 cm
LINE = 1.2                 # 行高 / 字号
MARK = re.compile(r'(\*\*|__|\^\^)')    # 加粗记号，见 Page.runs


def text_units(text):
    """一段文字合多少个汉字宽。加粗记号不占宽。"""
    return sum(0.6 if ord(ch) < 0x2000 else 1.0 for ch in MARK.sub('', text))


def text_lines(text, width_cm, size_pt):
    em = size_pt * PT_CM
    usable = max(width_cm - 0.15, 0.5)
    total = 0
    for para in text.split('\n'):
        total += max(1, -(-text_units(para) * em // usable))
    return int(total)


def text_height(text, width_cm, size_pt, spacing=1.15):
    """一段文字排在宽 width_cm 里的高度（cm），含上下各约 0.05 cm 的边距。"""
    return text_lines(text, width_cm, size_pt) * size_pt * PT_CM * LINE * spacing + 0.1
