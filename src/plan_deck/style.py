"""项目视觉系统：读 style.json，校验，按名字取色值、字号、栏位。

出页代码只用名字（颜色记号、类别配色、字号级别），不写色值和 pt；
写了色值或不存在的名字就报 StyleError，不回退。

颜色只有两个来源：style.json 本身，和 gis-composer `theme resolve` 解析出的地图主题（写了 mapTheme 时）。
地图相关的颜色（强调色、类别配色）引用主题，plan-deck 不自己解析主题的 extends。
"""
import copy
import difflib
import json
import os
import re

from .fonts import known_font, name_aliases, same_font, theme_font_uses
from .theme import ThemeError, check_resolved, project_theme_file, resolve_theme

CONTRACT = 'plan-deck-style'
SCHEMA_VERSION = 1

# 组件的度量。style.json 可以覆盖其中任何一项，不写就用这里的值。
DEFAULTS = {
    'text': {'lineSpacing': 1.15, 'headingMark': '■ ', 'headingSpaceBeforePt': 6, 'bodySpaceBeforePt': 2,
             'marginCm': 0.05},
    'strokes': {'arrowPt': 3, 'arrowColor': 'muted', 'framePt': 1.0, 'axisColor': 'line', 'gridColor': 'gridline'},
    'table': {'headerFill': 'title', 'headerText': 'white', 'text': 'text', 'fill': 'white', 'totalFill': 'fill',
              'highlightFill': 'highlightFill', 'rowHeightCm': 0.55, 'paddingCm': 0.12, 'lineColor': 'line', 'linePt': 0.75},
    'chart': {'text': 'text', 'linePt': 2.25},
    'legend': {'swatchCm': 0.42, 'rowMinCm': 0.55, 'labelOffsetCm': 0.6},
    'node': {'paddingCm': 0.15, 'text': 'white'},
    'ledger': {'opWidthCm': 0.9, 'opColor': 'muted'},
}

# 地图图面内注记的默认下限（pt，按放进页面后的实际尺寸计）：19.05 cm 高的 16:9 页面对象标注 10、次要注记 9，
# 其他页面按页高等比（取到 0.5 pt）。项目在 style.json 的 mapText 里另定。见 docs/visual-language.md“字号层级”
MAP_TEXT_DEFAULT = {'labelPt': 10, 'minorPt': 9}
PX_TO_PT = 0.75          # gis-composer 主题字号是 CSS px；按厘米出的图 1 px = 0.75 pt
REQUIRED_COLORS = ('title', 'emphasis', 'text', 'muted', 'line', 'white')
REQUIRED_TYPE = ('title', 'heading', 'body')
HEX = re.compile(r'^#[0-9A-Fa-f]{6}$')
REF = re.compile(r'^\{(colors|palette|theme)\.([^{}]+)\}$')
COLOR_FORM = '"#RRGGBB"、"{colors.<名字>}"、"{palette.<名字>}" 或 "{theme.<路径>}"'
# style.json 里运行时认识的键。type、palette、colors、classes 下是项目自己起的名字，不查；name、purpose 只作记录，运行时不读
_BOX = ('xCm', 'yCm', 'wCm', 'hCm')
KNOWN_KEYS = {
    '': ('contract', 'schema_version', 'name', 'purpose', 'page', 'body', 'grid', 'fonts', 'type', 'mapText', 'mapTheme',
         'palette', 'colors', 'classes', 'header', 'conclusion', *DEFAULTS),
    'page': ('widthCm', 'heightCm'), 'body': ('leftCm', 'rightCm', 'topCm', 'bottomCm'), 'grid': ('columns', 'gutterCm'),
    'fonts': ('cjk', 'latin', 'map'), 'mapText': tuple(MAP_TEXT_DEFAULT), 'mapTheme': ('project', 'theme'),
    'header': ('lines', 'rule', 'bars'), 'header.lines': ('role', *_BOX, 'size', 'color', 'bold'),
    'header.rule': ('xCm', 'yCm', 'wCm', 'color', 'widthPt'), 'header.bars': (*_BOX, 'color', 'slot'),
    'conclusion': _BOX,
    **{k: tuple(v) for k, v in DEFAULTS.items()},
}
REQUIRED_KEYS = {'page': KNOWN_KEYS['page'], 'body': KNOWN_KEYS['body'], 'grid': KNOWN_KEYS['grid'], 'fonts': ('cjk',),
                 'header.lines': _BOX, 'header.bars': _BOX, 'header.rule': ('xCm', 'yCm', 'wCm', 'color'), 'conclusion': _BOX}


class StyleError(ValueError):
    def __init__(self, problems):
        self.problems = [problems] if isinstance(problems, str) else list(problems)
        super().__init__('；'.join(self.problems))


def _sections(doc, path):
    """style.json 里 path 这一段的每个对象：[(显示的位置, 对象)]；数组逐项给。"""
    node = doc
    for part in path.split('.') if path else ():
        node = node.get(part) if isinstance(node, dict) else None
    items = [(f'{path}[{i}]', v) for i, v in enumerate(node)] if isinstance(node, list) else [(path, node)]
    return [(where, v) for where, v in items if isinstance(v, dict)]


def _unknown_keys(doc):
    """拼错或多写的键写了也不生效，记警告：[(位置, 键, 提示)]。"""
    out = []
    for path, known in KNOWN_KEYS.items():
        for where, node in _sections(doc, path):
            for key in node:
                if key not in known:
                    near = difflib.get_close_matches(key, known, 1)
                    hint = f'是不是 {near[0]}' if near else f'这一段认识：{"、".join(known)}'
                    out.append(f'{where + "." if where else ""}{key} 运行时不读，写了也不生效；{hint}')
    return out


def _merge(base, over):
    out = copy.deepcopy(base)
    for k, v in (over or {}).items():
        out[k] = _merge(out[k], v) if isinstance(v, dict) and isinstance(out.get(k), dict) else v
    return out


class Style:
    """一个项目的视觉系统。用 Style.load(path) 或 Style(dict) 构造；构造时校验，有错就抛 StyleError。
    style.json 写了 mapTheme 时，构造时调 gis-composer `theme resolve` 取地图主题；
    theme 参数可以直接给已解析的主题（resolved-theme 文档），不再调命令。
    warnings 是要处理的问题，info 是已知取舍或无法确认的事，交付前看一遍。"""

    def __init__(self, doc, source=None, theme=None):
        self.source = source
        self.doc = doc
        self.warnings = []
        self.info = []
        problems = []
        if doc.get('contract') != CONTRACT:
            problems.append(f'contract 应为 "{CONTRACT}"，实际是 {doc.get("contract")!r}')
        if doc.get('schema_version') != SCHEMA_VERSION:
            problems.append(f'不认识 schema_version {doc.get("schema_version")!r}，本运行时只读 {SCHEMA_VERSION}')
        if problems:
            raise StyleError(problems)
        for key in ('page', 'body', 'grid', 'fonts', 'type', 'colors', 'header'):
            if key not in doc:
                problems.append(f'缺少 {key}')
        for path, keys in REQUIRED_KEYS.items():
            for where, node in _sections(doc, path):
                problems += [f'{where} 缺少 {key}' for key in keys if key not in node]
        if problems:
            raise StyleError(problems)
        self.unknown_keys = _unknown_keys(doc)
        self.warnings += self.unknown_keys
        self.theme = self._load_theme(doc.get('mapTheme'), theme)
        self.metrics = _merge(DEFAULTS, {k: doc[k] for k in DEFAULTS if k in doc})
        self.page_w, self.page_h = float(doc['page']['widthCm']), float(doc['page']['heightCm'])
        b = doc['body']
        self.x0, self.y0 = float(b['leftCm']), float(b['topCm'])
        self.x1, self.y1 = self.page_w - float(b['rightCm']), self.page_h - float(b['bottomCm'])
        self.columns, self.gutter = int(doc['grid']['columns']), float(doc['grid']['gutterCm'])
        self.col_w = (self.x1 - self.x0 - self.gutter * (self.columns - 1)) / self.columns
        self.fonts = doc['fonts']
        self.type = {k: float(v) for k, v in doc['type'].items()}
        k = self.page_h / 19.05
        self.map_text = {key: round(v * k * 2) / 2 for key, v in MAP_TEXT_DEFAULT.items()}
        self.map_text.update({key: float(v) for key, v in doc.get('mapText', {}).items()})
        self._palette = {}
        for k, v in doc.get('palette', {}).items():
            hexv = self._resolve_value(v, f'palette.{k}', problems, (f'palette.{k}',))
            if hexv:
                self._palette[k] = hexv
        self._colors = self._resolve_colors(doc['colors'], problems)
        self._classes, self.class_sources = {}, {}
        for scheme, codes in doc.get('classes', {}).items():
            self._classes[scheme] = {}
            self.class_sources[scheme] = 'style.json'
            for code, value in codes.items():
                hexv = self._resolve_value(value, f'classes.{scheme}.{code}', problems)
                if hexv:
                    self._classes[scheme][code] = hexv
        # 地图主题里的类别配色全部可用，名字照旧；同名方案在 style.json 里再写一份就是两处定色，报错
        for scheme, codes in ((self.theme or {}).get('classes') or {}).items():
            if scheme in self._classes:
                problems.append(f'classes.{scheme} 在地图主题里已有：类别配色只在一处定，删掉 style.json 里的这一份'
                                f'（页面直接用主题的；要改颜色改地图主题）')
                continue
            self._classes[scheme] = {}
            self.class_sources[scheme] = 'theme'
            for code, value in codes.items():
                if isinstance(value, str) and HEX.match(value):
                    self._classes[scheme][code] = value.upper()
                else:
                    problems.append(f'地图主题 classes.{scheme}.{code} 不是色值：{value!r}')
        problems += self._check(doc)
        if problems:
            raise StyleError(problems)

    # ---------- 加载与校验 ----------
    @classmethod
    def load(cls, path, theme=None):
        with open(path, encoding='utf-8') as f:
            return cls(json.load(f), source=str(path), theme=theme)

    def _load_theme(self, ref, theme):
        self.theme_project = None
        if ref is None:
            return None
        if not isinstance(ref, dict) or not ref.get('project'):
            raise StyleError('mapTheme 要写成 {"project": "<gis-composer 项目目录，相对 style.json>", "theme": "<可选>"}')
        base = os.path.dirname(os.path.abspath(self.source)) if self.source else os.getcwd()
        self.theme_project = os.path.normpath(os.path.join(base, ref['project']))
        try:
            return check_resolved(theme if theme is not None else resolve_theme(self.theme_project, ref.get('theme')))
        except ThemeError as e:
            raise StyleError(f'mapTheme：{e}') from e

    def _theme_color(self, path, where, problems):
        if self.theme is None:
            problems.append(f'{where}：引用了 {{theme.{path}}}，但 style.json 没写 mapTheme（地图主题从哪个 composer 项目取）')
            return None
        # 类别配色的方案名和 code 里可能有点号，classes 只按前两个点切
        parts = path.split('.', 2) if path.startswith('classes.') else path.split('.')
        node = self.theme
        for i, part in enumerate(parts):
            if not isinstance(node, dict) or part not in node:
                have = ', '.join(sorted(node)) if isinstance(node, dict) else ''
                problems.append(f'{where}：地图主题里没有 {".".join(parts[:i + 1])}' + (f'；有：{have}' if have else ''))
                return None
            node = node[part]
        if not (isinstance(node, str) and HEX.match(node)):
            problems.append(f'{where}：地图主题的 {path} 不是 #RRGGBB 色值：{node!r}')
            return None
        return node.upper()

    def _resolve_value(self, value, where, problems, seen=()):
        if not isinstance(value, str):
            problems.append(f'{where}：颜色要写成 {COLOR_FORM}，实际是 {value!r}')
            return None
        if HEX.match(value):
            return value.upper()
        m = REF.match(value)
        if not m:
            problems.append(f'{where}：颜色要写成 {COLOR_FORM}，实际是 {value!r}')
            return None
        kind, name = m.groups()
        if kind == 'theme':
            return self._theme_color(name, where, problems)
        raw = (self.doc.get(kind) or {}).get(name)
        if raw is None:
            problems.append(f'{where}：引用了不存在的颜色 {name}（{kind}）')
            return None
        key = f'{kind}.{name}'
        if key in seen:
            problems.append(f'{where}：颜色引用成环（{" → ".join(seen + (key,))}）')
            return None
        return self._resolve_value(raw, key, problems, seen + (key,))

    def _resolve_colors(self, colors, problems):
        out = {}
        for k, v in colors.items():
            hexv = self._resolve_value(v, f'colors.{k}', problems, (f'colors.{k}',))
            if hexv:
                out[k] = hexv
        for k in REQUIRED_COLORS:
            if k not in colors:
                problems.append(f'colors 缺少 {k}')
        return out

    def _check(self, doc):
        problems = []
        for k in REQUIRED_TYPE:
            if k not in self.type:
                problems.append(f'type 缺少 {k}')
        if problems:
            return problems
        body = self.type['body']
        # 页面上没有比正文小的字（地图图面内的注记另由 mapText 定，见 docs/visual-language.md）
        for k, v in self.type.items():
            if v < body:
                problems.append(f'type.{k} = {v:g} pt 小于正文 {body:g} pt：页面上没有比正文小的字')
        # 正文以投影可读为准：19.05 cm 高的 16:9 页面不小于 12 pt，其他尺寸按页高等比（留 5% 余量，29.7 cm 高约 18 pt）
        floor = 12 * self.page_h / 19.05
        if body < floor * 0.95:
            problems.append(f'type.body = {body:g} pt，页高 {self.page_h:g} cm 时正文至少约 {floor:.0f} pt')
        # 红字结论：19.05 cm 高的 16:9 页面 20—24 pt，其他尺寸按页高等比（留 5% 余量）
        if 'conclusion' in self.type:
            k = self.page_h / 19.05
            if not 20 * k * 0.95 <= self.type['conclusion'] <= 24 * k * 1.05:
                self.warnings.append(f'type.conclusion = {self.type["conclusion"]:g} pt，页高 {self.page_h:g} cm 时红字结论宜为 '
                                     f'{20 * k:.0f}—{24 * k:.0f} pt')
        levels = sorted(set(self.type.values()))
        for a, b in zip(levels, levels[1:]):
            if b / a < 1.2:
                self.warnings.append(f'字号 {a:g} 与 {b:g} pt 只差 {b / a:.2f} 倍，相邻两级宜差约 1.25 倍')
        if self.theme is not None and 'map' not in self.fonts:
            self.warnings.append('写了 mapTheme 却没写 fonts.map：整册字体在 style.json 的 fonts 一处定，'
                                 '地图图面用什么字体也写在这里（与页面同就写同一字体，不同就在 style.md 写理由）')
        self._check_fonts()
        for k, v in self.map_text.items():
            if v > body:
                self.warnings.append(f'mapText.{k} = {v:g} pt 比正文还大，地图注记下限一般小于正文')
        # 地图主题的注记字号（按厘米出图、原尺寸放进页面时就是页面上的 pt）不低于项目定的下限
        for name, px in sorted(((self.theme or {}).get('constants') or {}).items()):
            if not (name.startswith('label') and name.endswith('Size') and isinstance(px, (int, float))):
                continue
            key = 'labelPt' if name == 'labelPrimarySize' else 'minorPt'
            floor = self.map_text.get(key)
            if floor and px * PX_TO_PT < floor - 0.05:
                self.warnings.append(f'地图主题 {name} = {px:g} px（{px * PX_TO_PT:.1f} pt）低于 mapText.{key} {floor:g} pt：'
                                     f'图上用到这类注记时会太小，改地图主题')
        if not (0 <= self.x0 < self.x1 <= self.page_w and 0 <= self.y0 < self.y1 <= self.page_h):
            problems.append('body 的边距超出页面')
        header = doc['header']
        for i, bar in enumerate(header.get('bars', [])):
            if bar.get('color') not in self._colors:
                problems.append(f'header.bars[{i}].color 不是 colors 里的名字：{bar.get("color")!r}')
        for i, line in enumerate(header.get('lines', [])):
            if line.get('role') not in ('title', 'parent', 'lead'):
                problems.append(f'header.lines[{i}].role 应为 title / parent / lead')
            if line.get('size') not in self.type:
                problems.append(f'header.lines[{i}].size 不是 type 里的级别：{line.get("size")!r}')
            if line.get('color') not in self._colors:
                problems.append(f'header.lines[{i}].color 不是 colors 里的名字：{line.get("color")!r}')
        if not any(line.get('role') == 'title' for line in header.get('lines', [])):
            problems.append('header.lines 里要有 role 为 title 的一行（本页节点名）')
        for section, keys in (('table', ('headerFill', 'headerText', 'text', 'fill', 'totalFill', 'highlightFill', 'lineColor')),
                              ('chart', ('text',)),
                              ('strokes', ('arrowColor', 'axisColor', 'gridColor')), ('node', ('text',)), ('ledger', ('opColor',))):
            for key in keys:
                if self.metrics[section][key] not in self._colors:
                    problems.append(f'{section}.{key} 要用 colors 里的名字，colors 里没有 {self.metrics[section][key]!r}')
        return problems

    def _check_fonts(self):
        """地图用字：fonts.map 与页面字体不同给 info；地图主题实际用的字体与 fonts.map 不一致记警告。"""
        page, mapf = self.fonts.get('cjk'), self.fonts.get('map')
        if page and mapf and not same_font(page, mapf):
            self.info.append(f'fonts.map “{mapf}”与页面 fonts.cjk “{page}”不同：放地图的页上会有两种黑体。'
                             f'只有字体许可或字形覆盖不允许地图用页面字体时才这样定，理由写进 style.md')
        if self.theme is None or not mapf:
            return
        typo = self.theme.get('typography') or {}
        # 主题自带字体（typography.fonts）的名字可以是别名；composer 给了 name 表（names）就按真名认
        declared = {f.get('family'): name_aliases(f.get('names')) for f in typo.get('fonts') or [] if isinstance(f, dict)}
        differ, unsure = [], {}
        for where, stack in theme_font_uses(typo):
            if any(same_font(n, mapf) for name in stack for n in (name, *declared.get(name, ()))):
                continue
            if all(name in declared and not declared[name] and not known_font(name) for name in stack):
                unsure.update(dict.fromkeys(stack))
            else:
                shown = [f'{name}（即 {declared[name][0]}）' if declared.get(name) and declared[name][0] != name else name
                         for name in stack]
                differ.append(f'{where} 用 {"、".join(shown)}')
        if differ:
            self.warnings.append(f'地图主题的字体与 fonts.map “{mapf}”不一致：{"；".join(differ)}。'
                                 f'地图主题的字体照 fonts.map 设（改地图主题的 typography），'
                                 f'或改 fonts.map 并在 style.md 写理由')
        if unsure:
            self.info.append(f'地图主题用的字体 {"、".join(unsure)} 是主题自带字体（typography.fonts）的名字，'
                             f'theme resolve 没给它们的 name 表，认不出是不是 fonts.map “{mapf}”：放图时按发布版记录的字体文件核对')

    # ---------- 取值 ----------
    def color(self, ref):
        """颜色记号（"emphasis"）或类别配色 (方案名, code) → "#RRGGBB"。写色值会报错。"""
        if isinstance(ref, tuple):
            return self.cls(*ref)
        if isinstance(ref, DataColor):
            return str(ref)
        if isinstance(ref, str) and ref.startswith('#'):
            raise StyleError(f'出页代码里不写色值 {ref}：写进 style.json 的 colors 或 classes，再按名字引用')
        if ref not in self._colors:
            if ref in self._palette:
                raise StyleError(f'{ref!r} 是 palette 里的基础色，出页代码不直接用：在 colors 里按含义起名'
                                 f'（如 "ledgerGross": "{{palette.{ref}}}"），一组同类的放进 classes')
            raise StyleError(f'colors 里没有 {ref!r}；可用：{", ".join(sorted(self._colors))}')
        return self._colors[ref]

    def cls(self, scheme, code):
        if scheme not in self._classes:
            raise StyleError(f'classes 里没有配色方案 {scheme!r}；可用：{", ".join(sorted(self._classes)) or "（无）"}')
        if code not in self._classes[scheme]:
            raise StyleError(f'配色方案 {scheme} 里没有 {code!r}；有：{", ".join(self._classes[scheme])}')
        return self._classes[scheme][code]

    def classes(self, scheme):
        if scheme not in self._classes:
            raise StyleError(f'classes 里没有配色方案 {scheme!r}；可用：{", ".join(sorted(self._classes)) or "（无）"}')
        return dict(self._classes[scheme])

    @property
    def palettes(self):
        """全部类别配色：{方案名: {code: 色值}}（引用已解析）。"""
        return {k: dict(v) for k, v in self._classes.items()}

    def size(self, role):
        if role not in self.type:
            raise StyleError(f'type 里没有字号级别 {role!r}；可用：{", ".join(self.type)}')
        return self.type[role]

    def col(self, i, n=1):
        """第 i 栏起（0 起算）、跨 n 栏：返回 (左边 cm, 宽 cm)。"""
        if i < 0 or n < 1 or i + n > self.columns:
            raise StyleError(f'栏位越界：第 {i} 栏起跨 {n} 栏，网格只有 {self.columns} 栏')
        return self.x0 + i * (self.col_w + self.gutter), n * self.col_w + (n - 1) * self.gutter

    @property
    def body(self):
        """版心 (x0, y0, x1, y1)，cm。"""
        return self.x0, self.y0, self.x1, self.y1

    @property
    def theme_file(self):
        """地图主题链里项目自己的主题文件 (项目内路径, sha256)；没写 mapTheme 或用内置主题时为 None。"""
        return project_theme_file(self.theme) if self.theme else None

    def summary(self):
        out = {'source': self.source, 'page': [self.page_w, self.page_h], 'body': [round(v, 3) for v in self.body],
               'grid': {'columns': self.columns, 'gutterCm': self.gutter, 'columnCm': round(self.col_w, 3)},
               'type': self.type, 'mapText': self.map_text, 'fonts': self.fonts, 'palette': self._palette,
               'colors': self._colors, 'classes': self._classes, 'classSources': self.class_sources}
        if self.theme is not None:
            typo = self.theme.get('typography') or {}
            out['mapTheme'] = {'project': self.theme_project, 'theme': self.theme.get('theme'), 'id': self.theme.get('id'),
                               'chain': [c.get('id') for c in self.theme.get('chain') or []],
                               'file': dict(zip(('path', 'sha256'), self.theme_file)) if self.theme_file else None,
                               'mapFontFamily': typo.get('mapFontFamily')}
        out['warnings'] = self.warnings
        out['info'] = self.info
        return out


class DataColor(str):
    """来自数据约定（如 legend.json）的色值：地图上是什么色，页面就画什么色，不经过 style。"""
