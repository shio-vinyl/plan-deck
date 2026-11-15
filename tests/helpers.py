"""测试用的虚构数据：示例视觉系统、按 gis-composer 发布版结构搭的假地图项目。产物都写进仓库外的临时目录。"""
import copy
import hashlib
import json
import os
import shutil
import sys
import tempfile

from PIL import Image

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, 'src'))
EXAMPLE_STYLE = os.path.join(ROOT, 'examples', 'style.json')


def style_doc(**over):
    with open(EXAMPLE_STYLE, encoding='utf-8') as f:
        doc = json.load(f)
    for k, v in over.items():
        doc[k] = v
    return copy.deepcopy(doc)


class TempDir:
    def __enter__(self):
        self.path = tempfile.mkdtemp(prefix='plan-deck-test-')
        return self.path

    def __exit__(self, *exc):
        shutil.rmtree(self.path, ignore_errors=True)


def legend_doc(image, print_cm=None, items=None, drawn=False):
    doc = {'contract': 'map-legend', 'schema_version': 1, 'image': image, 'drawn': drawn, 'title': None,
           'items': items if items is not None else [
               {'label': '更新单元范围', 'type': 'line', 'color': '#194461', 'width': 1, 'dasharray': [4, 3]},
               {'label': 'phases', 'type': 'fill', 'opacity': 1, 'model': 'classified', 'title': '实施时序',
                'classes': [{'code': '近期', 'label': '近期（2026–2027 年，3 处）', 'shortLabel': '近期', 'color': '#C00000'},
                            {'code': '远期', 'label': '远期', 'color': '#B4C7E7'}]},
               {'label': '更新规模 / 万㎡', 'type': 'circle', 'color': '#6a7378', 'model': 'size',
                'symbols': [{'label': '12', 'radius': 10}, {'label': '1', 'radius': 3}]},
               {'label': '江屿市域', 'type': 'fill', 'color': '#dcdcdc', 'opacity': 0.32, 'role': 'context'},
           ]}
    if print_cm:
        doc['print'] = {'widthCm': print_cm[0], 'heightCm': print_cm[1], 'dpi': print_cm[2]}
    return doc


def make_release(root, version='v1', maps=None, panels=None, failed=None, record=None, fonts=None, flat=False):
    """在 root 下搭一个 gis-composer 项目的发布版：releases/<版本>/{release.json, output/<图>/..., qa/<图>/...}。
    maps: {图 id: (宽 cm, 高 cm, dpi) 或 None（像素画幅、无 print）}
    failed: {图 id: [(检查 code, 要素 id, 面板或 None)]}，这些图的 QA 报告写成未过（照 composer 的 qa.json 结构）
    record: 合并进 release.json 的字段（如 gis-composer 已知问题发布的 maps、accepted）
    fonts: {图 id 或 "图 id/面板": [(主题里的字体名, 字体文件, 图上文字里缺的字数[, name 表 names])]}，写进 qa/ 下的 metrics.json 的 render.fonts
    flat: 照单图运行的布局，文件直接放在 output/、qa/ 下，不分图的子目录"""
    maps = maps if maps is not None else {'m1-site': (10.0, 8.0, 300), 'm2-pixel': None}
    rel = os.path.join(root, 'releases', version)
    os.makedirs(rel, exist_ok=True)
    sub =(lambda map_id: ()) if flat else (lambda map_id: (map_id,))
    for map_id, pr in maps.items():
        d = os.path.join(rel, 'output', *sub(map_id))
        os.makedirs(d, exist_ok=True)
        size = (round(pr[0] / 2.54 * pr[2]), round(pr[1] / 2.54 * pr[2])) if pr else (800, 600)
        Image.new('RGB', size, (241, 241, 239)).save(os.path.join(d, f'{map_id}.png'))
        with open(os.path.join(d, f'{map_id}.legend.json'), 'w', encoding='utf-8') as f:
            json.dump(legend_doc(f'{map_id}.png', pr), f, ensure_ascii=False)
    for map_id, (panel, pr) in (panels or {}).items():
        d = os.path.join(rel, 'output', *sub(map_id), 'panels')
        os.makedirs(d, exist_ok=True)
        Image.new('RGB', (round(pr[0] / 2.54 * pr[2]), round(pr[1] / 2.54 * pr[2]))).save(os.path.join(d, f'{panel}.png'))
        with open(os.path.join(d, f'{panel}.legend.json'), 'w', encoding='utf-8') as f:
            json.dump(legend_doc(f'{panel}.png', pr), f, ensure_ascii=False)
    for map_id in maps:
        issues = []
        for code, fid, panel in (failed or {}).get(map_id, []):
            if panel:
                issues.append({'code': 'SHEET_PANEL_QA_FAILED', 'severity': 'error', 'panelId': panel, 'evidence': {'codes': [code]}})
                write_qa(os.path.join(rel, 'qa', *sub(map_id), 'panels', f'{panel}.qa.json'),
                         [{'code': code, 'severity': 'error', 'featureId': fid}])
            else:
                issues.append({'code': code, 'severity': 'error', 'featureId': fid})
        write_qa(os.path.join(rel, 'qa', *sub(map_id), f'{map_id}.qa.json'), issues)
    for key, faces in (fonts or {}).items():
        map_id, _, panel = key.partition('/')
        path = os.path.join(rel, 'qa', *sub(map_id), *(['panels', f'{panel}.metrics.json'] if panel else [f'{map_id}.metrics.json']))
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, 'w', encoding='utf-8') as f:
            json.dump({'render': {'fonts': [{'family': fam, 'path': f'assets/fonts/{file}', 'sha256': 'd' * 64, 'status': 'loaded',
                                             'substitution': None, 'coverage': {'missing': [], 'faceMissingCount': n},
                                             **({'names': rest[0]} if rest else {})}
                                            for fam, file, n, *rest in faces]}}, f, ensure_ascii=False)
    files = []
    for dirpath, _, names in os.walk(rel):
        for n in sorted(names):
            full = os.path.join(dirpath, n)
            relp = os.path.relpath(full, rel).replace(os.sep, '/')
            with open(full, 'rb') as f:
                files.append({'path': relp, 'sha256': hashlib.sha256(f.read()).hexdigest(), 'bytes': os.path.getsize(full)})
    with open(os.path.join(rel, 'release.json'), 'w', encoding='utf-8') as f:
        json.dump({'schemaVersion': 1, 'version': version, 'sourceRun': '20260101-000000-render-demo', 'files': files,
                   **(record or {})}, f, ensure_ascii=False)
    return rel


def write_qa(path, issues):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump({'schemaVersion': 1, 'kind': 'qa', 'summary': {'ok': not issues, 'errors': len(issues)}, 'issues': issues},
                  f, ensure_ascii=False)


def resolved_theme(highlight='#C00000', classes=None, theme_sha='a' * 64):
    """虚构的 gis-composer `theme resolve` 输出（kind resolved-theme v1）。"""
    return {'schemaVersion': 1, 'kind': 'resolved-theme', 'theme': 'style/map.theme.json', 'id': 'demo-map', 'name': '示例地图主题',
            'mode': {'mode': 'light', 'from': 'research'},
            'chain': [{'id': 'base', 'sha256': 'b' * 64, 'builtIn': True, 'file': 'themes/base.theme.json'},
                      {'id': 'research', 'sha256': 'c' * 64, 'builtIn': True, 'file': 'themes/research.theme.json'},
                      {'id': 'demo-map', 'sha256': theme_sha, 'builtIn': False, 'file': 'style/map.theme.json'}],
            'palette': {'highlight': highlight, 'text': '#194461', 'land': '#f1f1ef'},
            'classes': classes if classes is not None else {'更新方式': {'拆除重建': '#E46C3C', '综合整治': '#33A3DA', '有机更新': '#0d7578'}},
            'constants': {'labelPrimarySize': 13.4}, 'typography': {'mapFontFamily': 'GC Sans', 'fonts': [{'family': 'GC Sans'}]},
            'sheet': {'accent': highlight}}
