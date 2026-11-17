"""视觉系统预设：起稿时从预设生成项目的 style.json，页眉几何和字号阶梯不用手写。

预设按 33.867×19.05 cm（PowerPoint 默认 16:9）定。换页面尺寸时，横向的位置和宽度按页宽等比、纵向按页高等比，
字号和线宽按页高等比（取到 0.5 pt），见 docs/architecture.md“预设”。
"""
import copy
import json
import os

from .style import Style

HERE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'presets')
BASE_W, BASE_H = 33.867, 19.05
X_KEYS, Y_KEYS = ('xCm', 'wCm', 'leftCm', 'rightCm', 'gutterCm'), ('yCm', 'hCm', 'topCm', 'bottomCm')    # 其他 …Cm 按页高


def names():
    return sorted(f[:-5] for f in os.listdir(HERE) if f.endswith('.json'))


def load(name):
    path = os.path.join(HERE, f'{name}.json')
    if not os.path.isfile(path):
        raise ValueError(f'没有预设 {name!r}；可用：{"、".join(names())}')
    with open(path, encoding='utf-8') as f:
        return json.load(f)


def _scale(node, kx, ky, kt):
    if isinstance(node, list):
        return [_scale(v, kx, ky, kt) for v in node]
    if not isinstance(node, dict):
        return node
    out = {}
    for k, v in node.items():
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            if k in X_KEYS:
                v = round(v * kx, 2)
            elif k in Y_KEYS or k.endswith('Cm'):
                v = round(v * ky, 2)
            elif k.endswith('Pt'):
                v = round(v * kt * 2) / 2
        out[k] = _scale(v, kx, ky, kt)
    return out


def build(name, page=None, title=None):
    """预设 → style.json 的内容（dict）。page=(宽 cm, 高 cm) 时按页面尺寸换算。"""
    doc = copy.deepcopy(load(name))
    w, h = page or (BASE_W, BASE_H)
    kx, ky = w / BASE_W, h / BASE_H
    if page:
        doc = _scale(doc, kx, ky, ky)
        doc['page'] = {'widthCm': w, 'heightCm': h}
        doc['type'] = {k: round(v * ky * 2) / 2 for k, v in doc['type'].items()}
    if title:
        doc['name'] = title
    Style(doc)      # 生成的要能通过校验
    return doc


def init(name, out, page=None, title=None, force=False):
    """写出项目的 style.json；已存在时不覆盖（force=True 才覆盖）。返回校验后的 Style。"""
    if os.path.exists(out) and not force:
        raise FileExistsError(f'{out} 已存在，不覆盖；确要重写加 --force')
    doc = build(name, page, title)
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(doc, f, ensure_ascii=False, indent=2)
        f.write('\n')
    return Style.load(out)
