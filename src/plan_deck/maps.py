"""地图衔接：从 gis-composer 项目的 releases/<版本>/ 取图和 legend.json（约定 map-legend v1）。

- 只读发布版（`project release <run> --version vN` 产出的目录），不读某次 run 的 output/：
  run 目录会被 gc、重跑覆盖，出页脚本引用它就会悄悄换图。
- 读文件前按 release.json 记的 sha256 核对，发布后被改过的图拒绝使用。
- 放图尺寸以 legend.json 的 `print` 为准（gis-composer 按厘米出的图，图上字号即页面 pt），
  PNG 像素与 print 对不上、槽位放不下时报错。
- 发布版可以带着未过的 QA 发布（gis-composer 的已知问题发布：交付物齐全、只是检查没过，发布时写明原因）。
  读到这样的图默认记警告、列出未过的检查，不阻塞；strict=True 时报错。
  QA 没检查到的不算通过：一张图在发布版里找不到 QA 结论，也按未过处理。
- 图面实际用的字体读发布版 qa/ 下的 <图>.metrics.json（render.fonts：加载的字体文件），放图时与 style.json 的
  fonts.map 核对；没有记录的旧发布版不核对。
"""
import hashlib
import json
import os
from dataclasses import dataclass, field

from PIL import Image

from .fonts import face_name, name_aliases, same_font
from .svgshape import LibrarySymbol

LEGEND_CONTRACT = 'map-legend'
LEGEND_VERSION = 1


class MapError(ValueError):
    pass


def _sha256(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(1 << 20), b''):
            h.update(chunk)
    return h.hexdigest()


def load_legend(path):
    with open(path, encoding='utf-8') as f:
        doc = json.load(f)
    if doc.get('contract') != LEGEND_CONTRACT:
        raise MapError(f'{path}：contract 应为 "{LEGEND_CONTRACT}"，实际是 {doc.get("contract")!r}')
    if doc.get('schema_version') != LEGEND_VERSION:
        raise MapError(f'{path}：不认识 map-legend schema_version {doc.get("schema_version")!r}，本运行时只读 {LEGEND_VERSION}')
    return doc


def _ids(issue):
    """检查点到的要素、标注、注记 id（composer 的 issue 里是 featureId / labelId / annotationId）。"""
    return [str(issue[k]) for k in ('featureId', 'labelId', 'annotationId') if issue.get(k) is not None]


def describe_issues(issues):
    """[{code, panel?, ids}] → "IMPORTANT_LABEL_MISSING（zoom 面板：CB-021）、…"。"""
    parts = []
    for it in issues:
        where = '、'.join(([f'{it["panel"]} 面板'] if it.get('panel') else []) + it.get('ids', []))
        parts.append(it['code'] + (f'（{where}）' if where else ''))
    return '、'.join(parts) or '（未列出具体检查）'


@dataclass
class MapQa:
    """一张图在发布版里的 QA 结论。status：ok / failed / unknown（发布版里找不到结论）。"""
    status: str
    issues: list = field(default_factory=list)      # [{code, panel, ids}]
    reason: str = None                              # 带着问题发布时写的原因（release.json accepted）
    source: str = None                              # release.json / qa/ 报告

    def for_panel(self, panel):
        """多面板图的一个面板：只看这个面板的问题和不分面板的问题；都没有时这个面板算通过。"""
        if panel is None or self.status != 'failed' or not any(i.get('panel') for i in self.issues):
            return self
        mine = [i for i in self.issues if i.get('panel') in (None, panel)]
        return MapQa('failed' if mine else 'ok', mine, self.reason if mine else None, self.source)

    def message(self, map_id, version):
        if self.status == 'unknown':
            return (f'{map_id} 在发布版 {version} 里没有 QA 结论（release.json 没记，qa/ 下也没有报告）：'
                    f'QA 没检查到的不算通过')
        return (f'{map_id} 带着未过的 QA 发布（发布版 {version}）：{describe_issues(self.issues)}'
                + (f'；发布原因：{self.reason}' if self.reason else '')
                + '。交付前核对图面，或修好后重新出图发布')


@dataclass
class MapAsset:
    """发布版里的一张图（或多面板图的一个面板）及其图例文档。"""
    id: str
    image: str
    legend_path: str = None
    legend: dict = None
    release: str = None
    pixels: tuple = field(default=None)
    qa: MapQa = None
    theme_sha256: dict = None      # release.json inputs 里各主题文件的 sha256：{项目内路径: sha256}
    fonts: list = None             # 图面实际加载的字体 [{family, file, name, aliases, missing}]；发布版没记录时为 None

    @property
    def print(self):
        return (self.legend or {}).get('print')

    def check_print(self, tolerance_px=2):
        """核对 PNG 像素与 legend.json 的 print（宽高 cm × dpi）一致，返回 (宽 cm, 高 cm, dpi) 或 None。"""
        p = self.print
        if not p:
            return None
        w, h, dpi = float(p['widthCm']), float(p['heightCm']), float(p['dpi'])
        ew, eh = w / 2.54 * dpi, h / 2.54 * dpi
        iw, ih = self.pixels
        if abs(iw - ew) > tolerance_px or abs(ih - eh) > tolerance_px:
            raise MapError(f'{self.id}：PNG {iw}×{ih} px 与 legend.json 的 print {w:g}×{h:g} cm @ {dpi:g} dpi'
                           f'（应为 {ew:.0f}×{eh:.0f} px）不符，图和图例文档不是同一次出的')
        return w, h, dpi

    def main_fonts(self):
        """画图面大部分文字的字体：缺字最少的那些（中文图上是中文字体，只有西文字形的不算）。"""
        if not self.fonts:
            return []
        least = min(f['missing'] for f in self.fonts)
        return [f for f in self.fonts if f['missing'] == least]

    def font_mismatch(self, expected, key='fonts.map'):
        """图面主要字体与 style.json 定的地图字体（expected）不是同一种时返回说明，否则 None。没有字体记录时不核对。"""
        main = self.main_fonts()
        if not expected or not main or any(same_font(n, expected) for f in main for n in (f['name'], f['family'], *f['aliases'])):
            return None
        by_name = {}
        for f in main:
            fam, files = by_name.setdefault(f['name'], ({}, {}))
            fam[f['family']] = files[f['file']] = True
        used = '、'.join(f'{name}（主题里叫 {"、".join(fam)}；{"、".join(files)}）' for name, (fam, files) in by_name.items())
        return (f'{self.id} 图面实际用的字体是 {used}，style.json 的 {key} 是“{expected}”：'
                f'地图主题的字体照 {key} 设后重新出图发布，或改 {key} 并在 style.md 写理由')


class MapRelease:
    """gis-composer 项目的一个发布版：<项目>/releases/<版本>/{release.json, output/, qa/}。
    strict=True 时，未过 QA（或没有 QA 结论）的图取用即报错；默认只记警告（warnings，放图时也会记进 deck.warnings）。"""

    def __init__(self, project_root, version, strict=False):
        self.project_root = os.path.abspath(project_root)
        releases = os.path.join(self.project_root, 'releases')
        have = sorted(d for d in os.listdir(releases) if os.path.isfile(os.path.join(releases, d, 'release.json'))) \
            if os.path.isdir(releases) else []
        if not version:
            raise MapError(f'要指定发布版本；{releases} 里有：{", ".join(have) or "（无）"}')
        self.version = version
        self.root = os.path.join(releases, version)
        record = os.path.join(self.root, 'release.json')
        if not os.path.isfile(record):
            hint = '先用 gis-composer `project release <run-id> --version <版本>` 发布确认过的那次出图' if not have else f'现有：{", ".join(have)}'
            raise MapError(f'没有发布版 {version}（{record} 不存在）；{hint}')
        with open(record, encoding='utf-8') as f:
            self.record = json.load(f)
        self.files = {e['path']: e['sha256'] for e in self.record.get('files', [])}
        self._verified = {}
        self._fonts = {}
        self.strict = strict
        self.warnings = []
        self._qa = self._read_qa()

    # ---------- QA 结论 ----------
    def _read_qa(self):
        """release.json 的逐图结论：maps（每张图 ok / failed 和未过的检查）、accepted（带着问题发布的图和原因），
        见 gis-composer 的已知问题发布规则。"""
        out = {}
        entries = self.record.get('maps') or []
        if isinstance(entries, dict):
            entries = [dict(v, id=k) for k, v in entries.items()]
        for e in entries:
            status = e.get('qa', e.get('status'))
            status = status.get('status') if isinstance(status, dict) else status
            issues = [{'code': i.get('code'), 'panel': i.get('panel') or i.get('panelId'), 'ids': _ids(i)}
                      for i in e.get('issues') or []]
            out[e.get('id') or e.get('map')] = MapQa(status if status in ('ok', 'failed') else 'unknown', issues,
                                                     source='release.json')
        for a in self.record.get('accepted') or []:
            q = out.setdefault(a.get('map'), MapQa('failed', source='release.json'))
            q.status, q.reason = 'failed', a.get('reason')
            known = {i['code'] for i in q.issues}
            q.issues += [{'code': c, 'panel': None, 'ids': []} for c in a.get('codes') or [] if c not in known]
        return out

    def _qa_from_reports(self, map_id):
        where = self._dir('qa', map_id)
        rel = f'{where}{map_id}.qa.json'
        if rel not in self.files:
            return MapQa('unknown')
        with open(self._file(rel), encoding='utf-8') as f:
            doc = json.load(f)
        if (doc.get('summary') or {}).get('ok', False):
            return MapQa('ok', source='qa/')
        issues = []
        for i in doc.get('issues') or []:
            if i.get('severity') != 'error':
                continue
            panel = i.get('panelId')
            if i.get('code') == 'SHEET_PANEL_QA_FAILED' and panel:
                sub = f'{where}panels/{panel}.qa.json'
                if sub in self.files:
                    with open(self._file(sub), encoding='utf-8') as f:
                        issues += [{'code': j.get('code'), 'panel': panel, 'ids': _ids(j)}
                                   for j in json.load(f).get('issues') or [] if j.get('severity') == 'error']
                    continue
                issues += [{'code': c, 'panel': panel, 'ids': []} for c in (i.get('evidence') or {}).get('codes') or []]
                continue
            issues.append({'code': i.get('code'), 'panel': panel, 'ids': _ids(i)})
        return MapQa('failed', issues, source='qa/')

    def qa(self, map_id, panel=None):
        """这张图（或面板）的 QA 结论 MapQa。release.json 有逐图结论（maps）时没列的图算没有结论；
        没有 maps 时看 qa/ 下的报告。"""
        if map_id not in self._qa:
            self._qa[map_id] = MapQa('unknown') if 'maps' in self.record else self._qa_from_reports(map_id)
        return self._qa[map_id].for_panel(panel)

    # ---------- 图面字体 ----------
    def fonts(self, map_id, panel=None):
        """图面实际加载的字体（gis-composer 写在 qa/<图>/<图>.metrics.json 的 render.fonts）：
        [{family: 主题里的字体名, file: 字体文件名, name: 文件里的字体名, aliases: name 表里的全部名字, missing: 图上文字里这个字体缺的字数}]。
        字体名优先取 gis-composer 记下的 name 表（names）；旧发布版没有时读字体文件，读不到再按文件名猜。
        多面板图的整页读各面板的记录；发布版里没有记录时返回 None。"""
        where = self._dir('qa', map_id)
        if panel:
            rels = [f'{where}panels/{panel}.metrics.json']
        else:
            rels = [f'{where}{map_id}.metrics.json'] + sorted(
                p for p in self.files if p.startswith(f'{where}panels/') and p.endswith('.metrics.json'))
        found, faces = False, {}
        for rel in rels:
            if rel not in self.files:
                continue
            if rel not in self._fonts:
                with open(self._file(rel), encoding='utf-8') as f:
                    self._fonts[rel] = (json.load(f).get('render') or {}).get('fonts')
            for face in self._fonts[rel] or []:
                found = True
                sub = face.get('substitution')
                path = sub or face.get('path')     # 换用了备用字体文件时，画图的是备用的（names 记的是原字体）
                file = os.path.basename(path or '')
                aliases = [] if sub else name_aliases(face.get('names'))
                cov = face.get('coverage') or {}
                missing = cov.get('faceMissingCount', len(cov.get('missing') or []))
                key = (face.get('family'), file)
                if key not in faces or missing < faces[key]['missing']:
                    faces[key] = {'family': face.get('family'), 'file': file, 'name': aliases[0] if aliases else face_name(path),
                                  'aliases': aliases, 'missing': missing}
        return list(faces.values()) if found else None

    @property
    def theme_inputs(self):
        """release.json inputs 里的主题文件：{项目内路径: sha256}（与 theme resolve 的 chain 对照）。"""
        return {i['ref']: i.get('sha256') for i in self.record.get('inputs') or []
                if str(i.get('ref', '')).endswith('.theme.json')}

    @property
    def source_run(self):
        return self.record.get('sourceRun')

    @property
    def status(self):
        """发布版状态：ok，或 known-issues（带着未过的 QA 发布，见 accepted）。旧发布版没有这一项，按 ok。"""
        return self.record.get('status', 'ok')

    def _file(self, rel):
        path = os.path.join(self.root, rel)
        if rel not in self.files:
            raise MapError(f'发布版 {self.version} 的 release.json 没有登记 {rel}')
        if not os.path.isfile(path):
            raise MapError(f'发布版 {self.version} 缺文件 {rel}')
        if rel not in self._verified:
            if _sha256(path) != self.files[rel]:
                raise MapError(f'发布版 {self.version} 的 {rel} 与 release.json 记的 sha256 不符：发布后被改过。'
                               f'重新出图请发布成新版本，不要改旧版本')
            self._verified[rel] = True
        return path

    def _dir(self, kind, map_id):
        """一张图的文件在 output/、qa/ 下的位置：整目录运行是 <kind>/<图>/，单图运行直接放在 <kind>/。"""
        return f'{kind}/' if f'output/{map_id}.png' in self.files else f'{kind}/{map_id}/'

    def map_ids(self):
        flat = [p[len('output/'):-len('.png')] for p in self.files
                if p.startswith('output/') and p.endswith('.png') and p.count('/') == 1]
        if flat:     # 单图运行：output/ 下的子目录只有 panels/，不是图
            return sorted(flat)
        out = os.path.join(self.root, 'output')
        return sorted(d for d in os.listdir(out) if os.path.isdir(os.path.join(out, d))) if os.path.isdir(out) else []

    def map(self, map_id, panel=None):
        """取一张图；panel 为多面板图的面板名（读 <图的 output 目录>/panels/<面板>.png）。"""
        if map_id not in self.map_ids():
            raise MapError(f'发布版 {self.version} 里没有图 {map_id}；有：{"、".join(self.map_ids()) or "（无）"}')
        where = self._dir('output', map_id)
        base = f'{where}panels/{panel}' if panel else f'{where}{map_id}'
        image = self._file(base + '.png')
        legend_path = legend = None
        if base + '.legend.json' in self.files:
            legend_path = self._file(base + '.legend.json')
            legend = load_legend(legend_path)
        with Image.open(image) as im:
            pixels = im.size
        full_id = map_id + (f'/{panel}' if panel else '')
        qa = self.qa(map_id, panel)
        if qa.status != 'ok':
            msg = qa.message(full_id, self.version)
            if self.strict:
                raise MapError(msg + '（严格模式）')
            if msg not in self.warnings:
                self.warnings.append(msg)
        return MapAsset(id=full_id, image=image, legend_path=legend_path, legend=legend, release=self.version,
                        pixels=pixels, qa=qa, theme_sha256=self.theme_inputs, fonts=self.fonts(map_id, panel))

    def check(self):
        """逐张核对：文件哈希、legend.json 约定版本、PNG 像素与 print。返回每张图的摘要。"""
        rows = []
        for map_id in self.map_ids():
            prefix = f'{self._dir("output", map_id)}panels/'
            panels = sorted({p[len(prefix):-len('.png')] for p in self.files if p.startswith(prefix) and p.endswith('.png')})
            for panel in [None] + panels:
                row = {'map': map_id + (f'/{panel}' if panel else '')}
                try:
                    a = self.map(map_id, panel)
                    pr = a.check_print()
                    row.update(pixels=list(a.pixels), legend=bool(a.legend),
                               print=None if pr is None else {'widthCm': pr[0], 'heightCm': pr[1], 'dpi': pr[2]},
                               drawn=(a.legend or {}).get('drawn'), status='ok', qa=a.qa.status,
                               fonts=None if a.fonts is None else sorted({f['name'] for f in a.main_fonts()}))
                    if a.qa.status != 'ok':
                        row.update(qaIssues=a.qa.issues, qaReason=a.qa.reason, warning=a.qa.message(row['map'], self.version))
                    if a.legend is None:
                        row['note'] = '没有 legend.json：页面画不了可编辑图例'
                    elif pr is None:
                        row['note'] = '没按厘米出图（legend.json 无 print）：放进页面后图上字号不可控'
                except MapError as e:
                    row.update(status='error', error=str(e))
                rows.append(row)
        return rows


# ---------------- legend.json → 页面图例行 ----------------

def legend_rows(doc, short=True, context=True):
    """把 legend.json 的 items 展开成页面图例行：
    ('title', 文字) 或 ('row', 色值, 文字, 符号, 不透明度)；符号为 fill / circle / frame，
    或 LibrarySymbol（gis-composer 符号库的行，带 librarySymbol 和 svg：页面照 svg 画同一个符号）。
    - 分类行：先出标题（行的 title，没有时用文档 title），再逐类一行；
    - 尺寸行：一行"圆的大小：<标签>（小–大）"；
    - 普通行（无 model，如范围线、示意面、放大位置）：排在分类行和尺寸行之后，线和圈画成框；
    - role 为 context 的行，context=False 时不画；
    - short=True 时有 shortLabel 就用 shortLabel（地图上的长名放到页面上偏长）。
    doc 也可以是几份 legend.json 的列表（如放大图面板 + 整页图例），按顺序合并。"""
    docs = doc if isinstance(doc, (list, tuple)) else [doc]
    rows, plain, skipped = [], [], []
    pick = (lambda d: d.get('shortLabel') or d.get('label')) if short else (lambda d: d.get('label'))
    for d in docs:
        for item in d.get('items', []):
            if item.get('role') == 'context' and not context:
                continue
            model = item.get('model')
            kind = 'circle' if item.get('type') == 'circle' else 'fill'
            if model == 'classified':
                title = (item.get('shortTitle') if short else None) or item.get('title') or d.get('title')
                rows.append(('title', title or ''))
                for c in item.get('classes', []):
                    rows.append(('row', c['color'], pick(c), kind, float(item.get('opacity', 1))))
            elif model == 'size':
                sym = item.get('symbols', [])
                rng = f"（{sym[-1]['label']}–{sym[0]['label']}）" if sym else ''
                rows.append(('row', item.get('color', '#808080'), f'圆的大小：{pick(item)}{rng}', 'circle', 1.0))
            elif model is None and item.get('svg') and isinstance(item.get('librarySymbol'), dict) and item.get('label'):
                plain.append(('row', item.get('color') or '#000000', pick(item),
                              LibrarySymbol(item['librarySymbol'].get('name', ''), item['svg']), 1.0))
            elif model is None and item.get('color') and item.get('label'):
                frame = item.get('type') == 'line' or item.get('symbol') == 'ring'
                plain.append(('row', item['color'], pick(item), 'frame' if frame else kind,
                              1.0 if frame else float(item.get('opacity', 1))))
            else:
                skipped.append(f'{item.get("label")}（model={model}）')
    return rows + plain, skipped


def class_mismatches(doc, classes):
    """页面类别配色（style.json 的 classes）与地图图例里同一 code 的颜色不一致的地方。
    legend.json 不写方案名，按 code 认：一个分类行的全部 code 都在某个方案里，就拿这个方案比。"""
    out = []
    for d in (doc if isinstance(doc, (list, tuple)) else [doc]):
        for item in d.get('items', []):
            codes = {c.get('code'): c['color'].upper() for c in item.get('classes', []) if c.get('code') is not None}
            if not codes:
                continue
            for scheme, palette in classes.items():
                if set(codes) <= set(palette):
                    out += [(scheme, code, palette[code], color) for code, color in codes.items() if palette[code] != color]
    return out
