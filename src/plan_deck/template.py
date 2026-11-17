"""默认模板随项目（决策 4、特点 9）：同事在 PowerPoint 里新建文本框、新建页时用的是主题和版式。

python-pptx 的默认模板是 4:3、Calibri、黑字。主题字体和主题色换成 style.json 的，版式占位符按页面尺寸缩放。
组件的字体和颜色都显式写，所以这里不影响运行时画出来的页面。
"""
from lxml import etree
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Cm

# 主题色槽 → style.json 的颜色名（都是必填的颜色）。accent5、6 不留 Office 的橙绿，新建图表的第 5、6 个系列也在视觉系统里
THEME_COLORS = {'dk1': 'text', 'lt1': 'white', 'dk2': 'title', 'lt2': 'line', 'accent1': 'title', 'accent2': 'emphasis',
                'accent3': 'muted', 'accent4': 'line', 'accent5': 'title', 'accent6': 'emphasis', 'hlink': 'title',
                'folHlink': 'muted'}


def apply_template(prs, style):
    sx, sy = Cm(style.page_w) / prs.slide_width, Cm(style.page_h) / prs.slide_height
    prs.slide_width, prs.slide_height = Cm(style.page_w), Cm(style.page_h)
    master = prs.slide_master
    for part in (master, *master.slide_layouts):
        for xfrm in part._element.iter(qn('a:xfrm')):
            off, ext = xfrm.find(qn('a:off')), xfrm.find(qn('a:ext'))
            off.set('x', str(round(int(off.get('x')) * sx)))
            off.set('y', str(round(int(off.get('y')) * sy)))
            ext.set('cx', str(round(int(ext.get('cx')) * sx)))
            ext.set('cy', str(round(int(ext.get('cy')) * sy)))
    theme = master.part.part_related_by(RT.THEME)      # python-pptx 不解析主题，只能改它的字节（没有公开接口）
    root = etree.fromstring(theme.blob)
    for kind in ('a:majorFont', 'a:minorFont'):
        font = root.find(f'.//{qn(kind)}')
        font.find(qn('a:latin')).set('typeface', style.fonts.get('latin', style.fonts['cjk']))
        font.find(qn('a:ea')).set('typeface', style.fonts['cjk'])
        for script in font.findall(qn('a:font')):      # PowerPoint 的中文优先取这里（默认宋体），只改 a:ea 不生效 [实测]
            if script.get('script') in ('Hans', 'Hant'):
                script.set('typeface', style.fonts['cjk'])
    scheme = root.find(f'.//{qn("a:clrScheme")}')
    for slot, name in THEME_COLORS.items():
        el = scheme.find(qn(f'a:{slot}'))
        el.clear()
        etree.SubElement(el, qn('a:srgbClr'), val=style.color(name).lstrip('#'))
    # 主题的效果样式（effectStyleLst）默认带外阴影：python-pptx 直接 add_shape、同事在 PowerPoint 里新建的形状都引用它，
    # 无填充的框会把阴影画在字上。形状不加阴影（visual-language.md“框和卡片”），三档都清成空效果
    effects = root.find(f'.//{qn("a:effectStyleLst")}')
    for es in effects if effects is not None else ():
        for child in list(es):
            es.remove(child)
        etree.SubElement(es, qn('a:effectLst'))
    theme._blob = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
