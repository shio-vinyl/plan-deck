"""字体名比对：style.json 的 fonts、地图主题的字体名、发布版记录的字体文件，是不是同一种字体。

同一种字体有中英文名、地区后缀和字重写法（思源黑体 / Source Han Sans CN / SourceHanSansCN-Medium），
按名字归一后比；常见字体的不同名字列在 SAME_FONT 里。gis-composer 给出字体文件 name 表里的名字时
（theme resolve 的 typography.fonts[].names、metrics.json 的 render.fonts[].names），主题里的别名（如内置主题的
“GC Sans”）按这些名字认；没有 name 表的旧记录才认不出，只能说“无法确认”，不能说“不一致”。
"""
import os
import re

from PIL import ImageFont

# 同一字体的不同名字，写成归一后的前缀（小写、去空格和连字符）。思源黑体与 Noto Sans CJK 是同一套字形。
SAME_FONT = {
    '思源黑体': ('思源黑体', 'sourcehansans', 'notosanscjk', 'notosanssc', 'notosanstc'),
    '思源宋体': ('思源宋体', 'sourcehanserif', 'notoserifcjk', 'notoserifsc', 'notoseriftc'),
    '微软雅黑': ('微软雅黑', 'microsoftyahei', 'msyh'),
    '苹方': ('苹方', 'pingfang'),
    '黑体': ('黑体', 'simhei'),
    '宋体': ('宋体', 'simsun', 'nsimsun'),
    '仿宋': ('仿宋', 'fangsong', 'simfang'),
    '楷体': ('楷体', 'kaiti', 'simkai'),
    '等线': ('等线', 'dengxian'),
    '阿里巴巴普惠体': ('阿里巴巴普惠体', 'alibabapuhuiti'),
    'HarmonyOS Sans': ('harmonyossans',),
}
STYLE_WORDS = re.compile(r'(regular|normal|medium|semibold|demibold|extrabold|ultrabold|bold|extralight|ultralight|light|'
                         r'thin|heavy|black|italic|oblique|it|ui|常规|粗体|细体|中等)$')
GENERIC = {'serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui', 'ui-sans-serif', 'ui-serif',
           'ui-monospace', '-apple-system', 'blinkmacsystemfont'}


def _norm(name):
    key = re.sub(r'[\s"\'_\-]+', '', str(name)).casefold()
    while True:
        stripped = STYLE_WORDS.sub('', key)
        if stripped == key or not stripped:
            return key
        key = stripped


def font_key(name):
    """字体名 → 比对用的键：SAME_FONT 里认得的给组名，其余给归一后的名字（去掉字重、斜体写法）。"""
    key = _norm(name)
    for group, prefixes in SAME_FONT.items():
        if any(key.startswith(p) for p in prefixes):
            return group
    return key


def known_font(name):
    return font_key(name) in SAME_FONT


def same_font(a, b):
    return bool(a) and bool(b) and font_key(a) == font_key(b)


def pdf_font_is(name, pdf_font):
    """PDF 里嵌入的字体（Arial-BoldMT、SourceHanSerifCN-Medium）是不是 style.json 写的这种。PDF 字体名后面常带厂商后缀，再按前缀认。"""
    return same_font(name, pdf_font) or _norm(pdf_font).startswith(_norm(name))


def checkable(name):
    """能不能和 PDF 字体名比：PDF 里只有英文名，不在 SAME_FONT 里的中文名对不上。"""
    return known_font(name) or str(name).isascii()


def families(value):
    """CSS 字体栈（'"A", B, sans-serif'）或 MapLibre text-font 列表 → 字体名列表，去掉通用族名。"""
    names = value if isinstance(value, (list, tuple)) else str(value or '').split(',')
    out = [str(n).strip().strip('"\'').strip() for n in names]
    return [n for n in out if n and n.casefold() not in GENERIC]


def theme_font_uses(typography):
    """地图主题（theme resolve 的 typography）里每处用字：[(哪里, [字体名, …])]。
    mapFontFamily 是注记默认字体，也是 MapLibre 画汉字用的字体（localIdeographFontFamily）。"""
    typo = typography or {}
    uses = []
    default = typo.get('mapFontFamily') or ((typo.get('fonts') or [{}])[0].get('family'))
    if default:
        uses.append(('mapFontFamily', families(default)))
    for role, stack in sorted((typo.get('map') or {}).items()):
        uses.append((f'map.{role}', families(stack)))
    if typo.get('fontFamily'):
        uses.append(('fontFamily', families(typo['fontFamily'])))
    for role in ('title', 'subtitle', 'legend', 'annotation'):
        spec = typo.get(role)
        if isinstance(spec, dict) and spec.get('family'):
            uses.append((f'{role}.family', families(spec['family'])))
    return [(where, stack) for where, stack in uses if stack]


def name_aliases(names):
    """gis-composer 读出的字体 name 表 → 这个字体的全部名字：英文族名、全名、PostScript 名和各语言的本地化名。"""
    if not isinstance(names, dict):
        return []
    out = [names.get(k) for k in ('family', 'fullName', 'postScriptName')]
    for loc in names.get('localized') or []:
        if isinstance(loc, dict):
            out += [loc.get('family'), loc.get('fullName')]
    return [n for n in dict.fromkeys(out) if n]


def face_name(path):
    """字体文件里的字体名（读得到文件时），否则用文件名去掉字重部分（SourceHanSansCN-Medium.otf → SourceHanSansCN）。"""
    if path and os.path.isfile(path):
        try:
            return ImageFont.truetype(path, 12).getname()[0]
        except OSError:
            pass
    stem = os.path.splitext(os.path.basename(path or ''))[0]
    return stem.split('-')[0] or stem
