"""plan-deck 最小运行时：项目视觉系统、页面组件、地图衔接。

    from plan_deck import Deck, Style, MapRelease
    deck = Deck('style.json')
    maps = MapRelease('<composer 项目>', 'v1')
    p = deck.page('evidence', '2.1 现状用地', parent='二、现状评估')
    x, w = deck.style.col(0, 5)
    p.text(x, deck.style.y0, w, [('h', '小标题'), '正文，**强调色加粗**、__正文色加粗__、^^标题色加粗^^'])
    deck.save('deck.pptx')
"""
__version__ = '0.1.0'

from .audit import audit, compare_text  # noqa: E402
from .clean import clean  # noqa: E402
from .deck import Box, Deck, Page  # noqa: E402
from .edit import notes_frame, replace_text  # noqa: E402
from .maps import MapAsset, MapError, MapQa, MapRelease, load_legend  # noqa: E402
from .measure import text_height  # noqa: E402
from .offcanvas import NotAudited, offcanvas, placements  # noqa: E402
from .style import Style, StyleError  # noqa: E402
from .theme import ThemeError, resolve_theme  # noqa: E402

__all__ = ['Box', 'Deck', 'Page', 'audit', 'clean', 'compare_text', 'notes_frame', 'replace_text', 'MapAsset', 'MapError', 'MapQa', 'MapRelease', 'load_legend', 'NotAudited', 'offcanvas', 'placements', 'text_height', 'Style',
           'StyleError', 'ThemeError', 'resolve_theme', '__version__']
