"""把 gis-composer 符号库图例行的 SVG（legend.json 的 svg）拆成可画成 PowerPoint 自由曲线的几何。

只认 gis-composer 写出的那一小套 SVG：viewBox，<path d>（绝对坐标 M L C Z），<rect>，
fill / stroke / stroke-width / fill-opacity / stroke-opacity / stroke-dasharray，
fill="url(#…)" 指向的斜线 <pattern>。<clipPath> 不画（符号的裁切框只防溢出，图例尺寸下看不出）。
曲线按段长拆成折线，画出来的自由曲线可以在 PowerPoint 里改色、改线。"""
import math
import re
import xml.etree.ElementTree as ET
from typing import NamedTuple

SVG = '{http://www.w3.org/2000/svg}'


class LibrarySymbol(NamedTuple):
    """图例行里的符号库符号：名称和 legend.json 给的 SVG。"""
    name: str
    svg: str


class Part(NamedTuple):
    contours: list          # [(points, closed)]，points 是 viewBox 坐标
    fill: str | None        # '#RRGGBB'
    fill_opacity: float
    stroke: str | None
    width: float            # viewBox 单位
    stroke_opacity: float
    dash: list | None       # viewBox 单位
    hatch: dict | None      # {'angle', 'color'}


def _num(text, default=0.0):
    try:
        return float(text)
    except (TypeError, ValueError):
        return default


def _path(d, step=0.6):
    tokens = re.findall(r'[MLCZ]|-?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?', d or '', re.I)
    out, cur, i, x, y = [], None, 0, 0.0, 0.0

    def take():
        nonlocal i
        i += 2
        return float(tokens[i - 2]), float(tokens[i - 1])
    while i < len(tokens):
        c = tokens[i].upper()
        i += 1
        if c == 'M':
            x, y = take()
            cur = [[(x, y)], False]
            out.append(cur)
        elif c == 'L':
            x, y = take()
            cur[0].append((x, y))
        elif c == 'C':
            (ax, ay), (bx, by), (ex, ey) = take(), take(), take()
            reach = math.hypot(ax - x, ay - y) + math.hypot(bx - ax, by - ay) + math.hypot(ex - bx, ey - by)
            n = max(2, min(24, math.ceil(reach / step)))
            for s in range(1, n + 1):
                t = s / n
                u = 1 - t
                cur[0].append((u ** 3 * x + 3 * u * u * t * ax + 3 * u * t * t * bx + t ** 3 * ex,
                               u ** 3 * y + 3 * u * u * t * ay + 3 * u * t * t * by + t ** 3 * ey))
            x, y = ex, ey
        elif c == 'Z':
            cur[1] = True
        else:
            raise ValueError(f'SVG 路径里有不支持的命令 {c}')
    return [(pts, closed) for pts, closed in out if len(pts) > 1]


def _colour(value):
    return value if value and re.fullmatch(r'#[0-9A-Fa-f]{6}', value) else None


def parse(svg):
    """SVG 文本 → (视框 (x, y, w, h), [Part])，按 SVG 的先后（先画的在下）。"""
    root = ET.fromstring(svg)
    box = [float(v) for v in root.get('viewBox').replace(',', ' ').split()]
    patterns = {}
    for p in root.iter(f'{SVG}pattern'):
        m = re.search(r'rotate\(([-\d.]+)\)', p.get('patternTransform', ''))
        line = p.find(f'{SVG}path')
        patterns[p.get('id')] = {'angle': -float(m.group(1)) if m else 0.0, 'color': _colour(line.get('stroke') if line is not None else None) or '#000000'}
    parts = []
    skip = {id(e) for d in root.iter(f'{SVG}defs') for e in d.iter()}
    for e in root.iter():
        if id(e) in skip or e.tag not in (f'{SVG}path', f'{SVG}rect'):
            continue
        if e.tag == f'{SVG}rect':
            x, y, w, h = (_num(e.get(k)) for k in ('x', 'y', 'width', 'height'))
            contours = [([(x, y), (x + w, y), (x + w, y + h), (x, y + h)], True)]
        else:
            contours = _path(e.get('d'))
        fill_attr = e.get('fill')
        hatch = None
        url = re.fullmatch(r'url\(#(.+)\)', fill_attr or '')
        if url:
            hatch, fill_attr = patterns.get(url.group(1)), None
        stroke = _colour(e.get('stroke'))
        dash = [float(v) for v in e.get('stroke-dasharray', '').replace(',', ' ').split()] or None
        parts.append(Part(contours, _colour(fill_attr), _num(e.get('fill-opacity'), 1.0), stroke, _num(e.get('stroke-width'), 1.0),
                          _num(e.get('stroke-opacity'), 1.0), dash, hatch))
    return (box[0], box[1], box[2], box[3]), parts
