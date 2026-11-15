# 架构

本文只写现状。最小运行时（`src/plan_deck/`，v0.1）有项目视觉系统（含从 gis-composer 地图主题取色）、页面组件、地图衔接，加上 `inspect`、`audit`、`offcanvas`、`render`、`clean`。还没做的能力记在 GitHub Issue，“运行时”一节末尾列了名字。

## 分层

```
SKILL.md 和 docs/       工作流和判断规则
src/plan_deck/
  style、theme、fonts   视觉系统：读 style.json，从 gis-composer 取地图主题，核对字体
  deck、template        页眉和组件；新建稿子的模板
  maps                  读地图发布版、放图、图例
  inspect、words        每次出页跑的检查；页面文字的线索词表
  audit                 定稿前的只读体检和两版文字比对
  offcanvas             版面外审计；对象在页面上的位置，inspect、clean 也用它判断版面内外
  edit、clean           改已有 pptx；出交付副本
  render                渲染和字体核对
```

## 项目视觉系统

每个项目一份 `style.json`，起稿时定，全册共用；数值只在 `style.json` 里，人看的说明写在旁边的 `style.md`。出页代码只按名字取颜色、字号和栏位，写色值或不存在的名字都报错（`StyleError`），不回退。

**理由写在哪里。** `style.md` 写汇报对象、用途，以及版式、栏位、字号层级这些页面独有的取舍和理由。颜色和字体的取向与理由写在项目的 `style/STYLE.md`：它是 house-style skill 定义的项目样式记录，地图、图表和汇报共用一份（有地图时在 gis-composer 项目里，与地图主题同目录）；`style.md` 不再写一遍，链接到它。项目没有 `STYLE.md` 时（只做汇报、不涉及地图），颜色和字体的理由照旧写在 `style.md`。plan-deck 不要求项目有 `STYLE.md`，运行时也不读 `style.md` 和 `STYLE.md`。

页面上的颜色只有两个来源：`style.json`，和 gis-composer `theme resolve` 解析出的地图主题（写了 `mapTheme` 时）。地图也用的颜色（强调色、类别配色）定在地图主题里，页面从主题读；只在页面上用的颜色（图表系列、账算式、漏斗、浅色底、表格高亮）按含义在 `style.json` 里起名。出页脚本里没有色值。

```json
{
  "contract": "plan-deck-style", "schema_version": 1,
  "purpose": "practical",
  "page": {"widthCm": 33.87, "heightCm": 19.05},
  "body": {"leftCm": 1.4, "rightCm": 1.4, "topCm": 4.0, "bottomCm": 0.7},
  "grid": {"columns": 12, "gutterCm": 0.4},
  "fonts": {"cjk": "微软雅黑", "latin": "Arial", "map": "思源黑体"},
  "type": {"parent": 24, "title": 20, "conclusion": 24, "heading": 15, "body": 12, "figure": 30},
  "mapText": {"labelPt": 10, "minorPt": 9},
  "mapTheme": {"project": "地图"},
  "palette": {"navy": "#194461", "blue": "#4F81BD", "grey": "#A6A6A6"},
  "colors": {"title": "{palette.navy}", "emphasis": "{theme.palette.highlight}", "text": "#262626", "muted": "#646464",
             "line": "#BFBFBF", "gridline": "#E5E5E5", "fill": "#F2F2F2", "highlightFill": "#FFF2CC", "white": "#FFFFFF", "tint": "#DEEBF7", "ledgerGross": "{palette.blue}", "ledgerDeduct": "{palette.grey}"},
  "classes": {"实施时序": {"近期": "{colors.emphasis}", "中期": "#FFC000", "远期": "#B4C7E7"}},
  "header": {
    "bars": [{"slot": "header-bar", "xCm": 0, "yCm": 0, "wCm": "page", "hCm": 0.36, "color": "title"}],
    "lines": [{"role": "parent", "xCm": 1.4, "yCm": 0.55, "wCm": 31.07, "hCm": 1.0, "size": "parent", "color": "title", "bold": true},
              {"role": "title", "xCm": 1.4, "yCm": 1.55, "wCm": 31.07, "hCm": 0.9, "size": "title", "color": "title", "bold": true}],
    "rule": {"xCm": 1.4, "yCm": 2.55, "wCm": 31.07, "color": "line", "widthPt": 0.75}
  },
  "conclusion": {"xCm": 1.4, "yCm": 2.65, "wCm": 31.07, "hCm": 1.1}
}
```

完整的虚构示例见 [examples/style.json](../examples/style.json)（52.8×29.7 cm，页眉按 `visual-language.md` 的默认样式：上级标题在上、本页标题在下、细线）。

- **长度一律厘米，字号一律 pt。** `body` 是版心的四边边距，`grid` 把版心分成等宽栏，`style.col(i, n)` 取第 i 栏起跨 n 栏的左边和宽度。
- **`type`**：页面上的字号阶梯，键名就是组件里用的级别（正文 `body`、小标题 `heading`、上级标题 `parent`、本页标题 `title`、红字结论 `conclusion`、副标题 `lead`、大号数字 `figure`，可以加别的级别）。校验：任何一级小于 `body` 报错（页面上没有比正文小的字）；`body` 低于按页高换算的投影下限（19.05 cm 高约 12 pt，29.7 cm 高约 18 pt）报错；相邻两级差不到 1.2 倍、红字结论不在 20—24 pt（按页高等比）只警告，由项目决定。
- **`mapText`**：项目定的地图图面内注记下限（`labelPt` 对象标注、`minorPt` 次要注记），按放进页面后的实际尺寸计，只用于核对地图，不是页面字号。不写时按页高换算默认值（19.05 cm 高 10 / 9 pt，见 `visual-language.md`“字号层级”）。写了 `mapTheme` 时，`style check` 把主题的 `label*Size`（px × 0.75 = pt）和它比，低了就警告。
- **`fonts`**：整册字体一处定，`cjk`、`latin` 是页面用字，`map` 是地图图面用字（理由见 `visual-language.md`“字体”）。`map` 与 `cjk` 不同时 `style check` 给 info（同页两种黑体）；写了 `mapTheme` 时把主题 `typography` 里各处用字（`mapFontFamily`、`map.*`、`fontFamily`、`title` / `subtitle` / `legend` / `annotation` 的 `family`）和 `fonts.map` 比，没有的警告。比的是字体名：同一字体的中英文名、地区后缀、字重写法归成一种（`src/plan_deck/fonts.py`）；主题自带字体的别名（内置主题的 “GC Sans”）按 `theme resolve` 给的 name 表（`typography.fonts[].names`：英文族名、全名、本地化名）认；旧版 composer 没给 name 表时认不出，只给 info，放图时按发布版记录的字体再核对。
- **`mapTheme`**：地图主题从哪里取，`{"project": "<gis-composer 项目，相对 style.json>", "theme": "<可选，默认项目主题>"}`。加载 `style.json` 时运行时调 composer 的命令行 `theme resolve`（命令取自环境变量 `GIS_COMPOSER_CMD`，没有时读 gis-plugin 的 `~/.config/gis-plugin/engines.env`），读解析后的主题；plan-deck 不自己解析 `extends`。命令不可用或主题出错就报错，不回退到旧值。
- **`palette`**（可选）：基础色板，只供 `colors`、`classes` 引用（`{palette.<名字>}`），出页代码不能直接用。色相名（“蓝”“灰”）放这里，页面上按含义起名。
- **`colors`**：页面的颜色记号，按含义起名。`title`、`emphasis`、`text`、`muted`、`line`、`white` 必须有；值是 `#RRGGBB`，或引用 `{colors.<名字>}`、`{palette.<名字>}`、`{theme.<路径>}`。`{theme.…}` 的路径照 `theme resolve` 输出的字段走：`{theme.palette.highlight}`（地图强调色）、`{theme.sheet.accent}`、`{theme.classes.<方案名>.<code>}`。
- **`classes`**：命名的类别配色，`{方案名: {code: 颜色}}`，同一含义全册同色（分期、更新方式、权属……）。组件里写 `(方案名, code)` 取色。地图主题里的 `classes` 全部可以直接用，名字不变；`style.json` 只写页面独有的方案（比如图表系列），和主题同名的方案再写一份就报错（两处定色）。
- **`header`**：页眉几何。`bars` 是色条，`lines` 按顺序画上级标题（`parent`）、本页标题（`title`）和可选的副标题（`lead`），可选 `rule` 细线。各项目的页眉不同（研究稿照原稿、新起稿按默认样式），所以写成数据，组件只照着画。
- **`conclusion`**（要用红字结论时写）：`p.conclusion()` 的默认位置，`xCm`、`yCm`、`wCm`、`hCm`，放在标题线下。不写时每次调用都要给坐标。它不是页眉的一部分，去 AI 味版可以逐页换位置。
- **组件度量**（`text`、`strokes`、`table`、`chart`、`legend`、`node`、`ledger` 段）有默认值，`style.json` 写了就覆盖，比如表头色、表格行线、图表文字色（`chart.text`，深色底上写浅色）、箭头线宽、图例色块尺寸。
- **其他键**：`name`、`purpose` 只作记录，运行时不读。运行时不认识的键（拼错的、以为有的能力）写了也不生效，`style check` 和 `deck.warnings` 都给警告；`page`、`body`、`grid`、`fonts.cjk` 和页眉的几何缺项时报错。

**为什么用 JSON**：读 YAML 要多装 PyYAML，Python 3.10 也没有标准库的 TOML；gis-composer 的项目主题、map-legend 等约定都是 JSON，形状一致（`widthCm`、`palette.classes`）便于对照；agent 读写 JSON 不易出错。JSON 不能写注释，取舍和理由写在 `style.md` 或 `STYLE.md`（见本节开头）。

### 从地图主题取色

**地图也用的颜色以地图主题为准，页面从主题读**：`mapTheme` 指向 composer 项目，强调色写 `{theme.palette.highlight}`，类别配色直接用主题的方案。改主题的强调色或一个类别色，重出页面就跟着变；地图要重新出图并发布。页面放的发布版若是用旧主题出的图（`release.json` 的 `inputs` 里主题文件的 sha256 与 `theme resolve` 的 `chain` 不同），放图时警告。

`legend()` 画图例时，把 legend.json 里每个分类行的 code 和 `classes` 里含有这些 code 的方案比一遍，同一类不同色就记警告（改了主题、地图还没重出时就会看到）。

## 结构大纲

每份 pptx 旁边放一份同名的 `.outline.md`（设计决策 11），由 agent 维护，供人浏览结构、指挥修改，人不直接改它：

```markdown
# 江屿市北岸老城片区城市更新统筹规划｜第二次阶段汇报

## 二、现状评估与问题识别 · p7–14
### 2.1 现状用地 · p8
### 2.2 主要问题 · p9–12 · 总览 p9
#### 问题一：老旧小区量大面广 · p10
#### 问题二：城中村消防隐患突出 · p11
#### 问题三：纺织厂停产闲置 · p12 · 待补

## 备用
- 2.1 现状用地（旧版） · p31 · 隐藏
```

- 标题层级就是结构树层级；叶节点后写这一页的页码，中间节点写页码范围，带总览页的再注明“总览 pN”。
- 只记节点名、页码和状态（待补、隐藏），不记内容；术语表例外，放在开头（`## 术语`，一行一个词和一句定义）。
- 封面、目录、章节页、致谢页由树生成，不列。
- pptx 每次修改后同步大纲（现在由 agent 手工维护）：节点信息按标题、面包屑、目录反推。

## 对象命名约定

工具生成的形状统一用 `pd:` 前缀，便于定位，也和人工添加的元素区分：

```
pd:<page-type>:<slot>[:<index>]
例：pd:evidence:title   pd:evidence:map   pd:evidence:stat:2   pd:plan-map:legend
```

运行时自动命名：同一页同一槽位第二次出现起自动编号（第一个改成 `:0`）；显式给的序号与本页已有的名字重复时报错。`inspect` 也会报同页重名。

## 地图衔接

现在走的是 gis-composer 已有的两件东西，不另立格式：

- **发布版**：确认过的一次出图用 gis-composer `project release <run-id> --version vN` 发布成 `<composer 项目>/releases/vN/`（`release.json` 记来源运行和每个文件的 sha256）。出页只读发布版（`MapRelease(项目, 'vN').map('<图 id>'[, panel=...])`），不读 `runs/<id>/output/`：运行目录会被 gc、重跑，引用它就会悄悄换图。读文件前按 `release.json` 核对哈希，发布后被改过就拒绝。整目录运行（`render maps/`）的文件在 `output/<图>/`、`qa/<图>/`；单图运行的文件直接在 `output/`、`qa/`，多面板图的面板在 `output/panels/`、`qa/panels/`。下文的路径都按整目录运行写，运行时两种都读。发布版里一张可读的图都没有时，`maps check` 报 `NO_MAPS` 并以退出码 1 结束，不算通过。
- **带着已知问题的发布版**（gis-composer 的已知问题发布）：默认只发布 QA 全过的运行；交付物齐全、只是检查没过的运行，可以写明原因后发布（`--accept-failed-checks --reason`），`release.json` 的 `status` 为 `known-issues`，`maps` 逐图记 `qa`（ok / failed）和未过的检查（`issues[].code`、`panelId`、`featureId`…），`accepted` 记 `{map, codes, reason}`。运行时读到未过的图：`maps check` 和放图都记警告，列出未过的检查和发布原因，不阻塞；`MapRelease(..., strict=True)` 或 `maps check --strict` 时报错。没有 `maps` 的旧发布版看发布版里的 `qa/<图>/<图>.qa.json`。找不到 QA 结论的图也按未过处理（QA 没检查到的不算通过）。多面板图的问题只在某个面板时，只有放这个面板才警告。
- **图面字体**：发布版带着运行的 `qa/`，其中 `<图>.metrics.json`（多面板图在 `qa/<图>/panels/<面板>.metrics.json`）的 `render.fonts` 记了出图时加载的字体：主题里的字体名 `family`、字体文件 `file` / `path`（可移植的相对路径；换用备用文件时为 `substitution`）、name 表 `names`、`coverage.faceMissingCount`（图上文字里这个字体缺的字数）。放图时取缺字最少的那些字体（中文图上是中文字体，只有西文字形的不算），字体名取 `names`（换用了备用文件时不用，`names` 记的是原字体）；旧发布版没有 `names` 时读字体文件里的名字，文件不在本机时用文件名，与 `fonts.map`（没写时 `fonts.cjk`）不是同一种就警告；`maps check` 每张图列出 `fonts`。没有这份记录的发布版不核对。多面板图读整页（Sheet）自己的记录和各面板的记录。
- **map-legend v1**：每张图旁的 `<图>.legend.json`（约定由 gis-composer 维护，gis-plugin 登记）。运行时用它做两件事：
  - **放图**：有 `print`（`widthCm`、`heightCm`、`dpi`）时按原尺寸放，图上的字就是主题定的 pt。先核对 PNG 像素等于 `print` 换算的像素；槽位比 `print` 小时报错（改地图清单的 `layout.widthCm/heightCm` 重新出图，或换大槽位），明确要缩时写 `shrink=True`，并警告缩小后图上注记约多少 pt、是否低于 `mapText`。没有 `print` 的图按比例撑满槽位并警告。每次放图记一条放置记录（槽位、实际尺寸、实际 dpi）。
  - **图例**：画成可编辑形状。分类行出标题加各类；尺寸行一行；普通行（范围线、示意面、放大位置）排在后面，线和圈画成虚线框，半透明面按不透明度画；`role: context` 的行可以不画；有 `shortLabel` 时默认用短名。几份 legend.json 可以合并成一个图例（放大图面板 + 整页的放大位置）。图例已画进 PNG（`drawn: true`）时警告重复。

## 运行时

`src/plan_deck/`，依赖 python-pptx、pillow。项目脚本用库入口，检查用命令行。

```python
from plan_deck import Deck, MapRelease
deck = Deck('style.json')
maps = MapRelease('<composer 项目>', 'v1')
p = deck.page('evidence', '2.1 现状用地', parent='二、现状评估与问题识别')      # 副标题 lead= 默认不写
x0, y0, x1, y1 = deck.style.body
x, w = deck.style.col(0, 4)
box = p.text(x, y0, w, [('h', '小标题'), '正文，**强调**'])      # 返回 Box，box.bottom 接着往下排
mx, mw = deck.style.col(4, 5)
p.map(maps.map('m1-landuse'), mx, y0, mw, y1 - y0)
lx, lw = deck.style.col(9, 3)
lg = p.legend(maps.map('m1-landuse'), lx, y0, lw)
p.table(lx, lg.bottom + 0.4, lw, rows, widths=[2, 1, 1], total=True)
deck.save('deck.pptx'); print(deck.warnings, deck.placements)
```

| 组件 | 说明 |
|---|---|
| `page(页型, title, parent, lead, header=True)` | 新页并按 `style.header` 画页眉；`header=False` 不画（满版图、章小结页），放图后调 `p.header()` 可以把页眉画在图上；什么字都没有时不画页眉线 |
| `conclusion` | 红字结论：一行带事实的判断，字号 `conclusion`、强调色加粗，默认位置取 `style.json` 的 `conclusion`，可以给坐标；放不下时记警告 |
| `text` / `label` | 正文块（条目：正文、`('h', 小标题)`、`('m', 次要)`、`('big', 大号数字)`；`**…**` 强调色加粗，`__…__` 所在段落的颜色加粗，`^^…^^` 标题色加粗，不嵌套；`__`、`^^` 没配对或中间为空时报错并指出那段文字，`**` 没配对不报错，后面的字都算强调）/ 单段标签；按字宽估高度（汉字 1、ASCII 0.6），返回 Box；放不下或排出页面时记进 `deck.warnings`（按字宽估算，窄栏混排可能差一两行，只警告，不改版面）。`label(shadow=True)` 加柔和投影，只用于照片上的字 |
| `table` | 表头深色、行间细线（`table.lineColor`、`linePt`）、数字右对齐、合计行、提示行；按单元格换行估行高 |
| `chart` | 原生图表：`bar`、`column`（簇状、堆积）、`line`、`area`、`doughnut`、`pie`，或同类的 `XL_CHART_TYPE` 成员；系列色按名字取，`doughnut`、`pie` 按类别给色；文字色取 `chart.text`，折线线宽取 `chart.linePt`；`color_point`、`label_series` 微调。单系列的柱形、条形图 PowerPoint 会拿系列名自动出标题，运行时没有处理 |
| `ledger` | 账算式：同尺寸色块装数字，中间放运算符 |
| `node` / `arrow` / `shape` / `line` / `band` | 图示节点、全册一个样式的箭头、标记形状（`rect`、`oval`、`diamond`，其他直接传 `MSO_SHAPE` 成员）、连线、照片上垫字的渐隐暗带 |
| `map` / `legend` / `picture` | 发布版里的地图（按 `print` 放、核对，未过 QA、主题已变或图面字体不是 `fonts.map` 时警告）/ 可编辑图例 / 照片截图（默认撑满槽位可能留白；`fit='cover'` 铺满并裁掉多出的部分，用于满版图） |
| `notes` | 讲稿备注 |
| `Deck.open(路径, style)` / `deck.at(n)` | 打开已有 pptx 接着改：模板和同事的形状不动，`at(n)` 取已有的页，页上已有的名字都算占用，新增形状不重名；`page()` 加新页 |

新建的稿子，模板也随项目：主题字体、主题色取 `style.json`，版式占位符按页面尺寸缩放，同事在 PowerPoint 里新建文本框、新建页时不是 Calibri 黑字。主题里 `script="Hans"` 的字体要一起改，否则 PowerPoint 的中文仍是宋体。组件的字体和颜色都显式写，不受主题影响。

**什么进运行时：** 否则会静默出错、在真实 PowerPoint 上踩过且每个脚本都会忘、或需要跨页跨文件的状态，这三种才进；其余写文档或配方，用 `p.slide` 和 python-pptx 自己写。

命令行（`python3 <仓库>/bin/plan-deck …`，或 `pip install -e <仓库>` 后用 `plan-deck`），输出 JSON，有问题时退出码 1：

```
plan-deck style check style.json              校验视觉系统，列出解析后的颜色（含来源）、类别配色、栏位、地图主题、警告和 info
plan-deck maps check <composer 项目> --release v1 [--strict]   逐张核对发布版：哈希、legend.json 版本、PNG 像素与 print、QA 结论
plan-deck offcanvas deck.pptx… --out 目录 [--render 32,58|24-31|all] [--engine libreoffice|powerpoint] [--dpi 100] [--crop 0.2]   版面外审计，见“画布外对象”；有文件没审计、对象没读全或渲染失败时退出码 1
plan-deck inspect deck.pptx [--style style.json] [--brief]   页面与形状清单（含组内形状，每个形状标 placement）；problems 影响退出码（页面绝不出现的用语、小字、重名、画布外对象、悬空动画），warnings 不影响（版面和文字线索，不查隐藏页），每条说明了要看什么
plan-deck audit deck.pptx [--against 另一版.pptx] [--brief]   只读体检（docs/humanize.md）：全册指标和参考范围、逐页标记（框等尺寸、左色块右底框、和上一页版面相似、没有图）、页面绝不出现的用语（problems，影响退出码）、文字线索（讨论口语、单字动词接术语）；--against 逐页比两版文字
plan-deck clean deck.pptx 副本.pptx [--pattern 正则] [--drop-hidden] [--force]   另存副本，清掉备注里的标签行、待确认和内部说法；列出多页相同的讲稿和页面文字里的过程信息
plan-deck render deck.pptx --out renders/ [--engine powerpoint|libreoffice] [--style style.json]   转 pdf 和逐页 png，列出实际嵌入的字体
```

**还没有的**：起稿 spec、结构大纲生成、页面池操作（取页、重排、隐藏）、按名字 patch（文字替换有 `replace_text`）、页面元数据、版式清单的槽位模板、结构页（封面、目录、章节页）、导出地图主题、地图槽位请求。这些在出页脚本里手写或用常规库做，规则照样遵守；设想记在 Issue #8。

## 改已有 pptx

接手旧稿、局部改字时容易出的问题，运行时能处理的用 `plan_deck.replace_text`、`plan_deck.notes_frame`，其余自己留意：

- **文字跨 run**：同一句话常被拆成几个 run，按 run 查找会漏。`replace_text` 按段落查找，新文字放进第一个涉及的 run，其余 run 删去相应部分，所以要保留强调时只替换同一 run 里的短语。
- **强调色和模板**：整段重写时别拿加粗标红的 run 当模板，否则整段变红。
- **中文标点落行首**：python-pptx 写入的 run 没标语言时，PowerPoint 不按中文规则断行。`replace_text` 会标 `lang="zh-CN"`，自己写 run 时也要标。
- **备注页没有正文占位符**：`notes_slide.notes_text_frame` 会是 `None`，用 `notes_frame` 补；从别页复制占位符时要清掉里面的文字，否则带过来别页的备注。
- **合并单元格**：合并后下边线可能丢失，要给下一行的上边线补线。
- **缩放形状**：线宽和字号不随形状缩放，圆环一类靠线宽成形的图形要同时改线宽。
- **批量替换形状位置**：按 `pd:` 名字前缀筛，不按坐标筛，否则会误移页码。

## 画布外对象

团队习惯把页面任务说明、批注要用的表和暂存素材放在幻灯片可见范围之外（见 `design-principles.md` 特点 5）。渲染图看不到它们，不分位置的文字提取又把它们和正文混在一起。脚本不得删除或移动这些对象；工具不往画布外放东西。

`plan-deck offcanvas` 读 pptx，只写输出目录，原文件不改：

| 输出 | 内容 |
|---|---|
| `版面外清单.md` | 人看的清单：索引表，逐页列渲染图上看不到的对象和它们的全部内容 |
| `全文提取.md` | 每页全部文字和备注、批注，每个对象标 `[版面内]`、`[版面外]`、`[跨边]`、`[窗格隐藏]` |
| `offcanvas.json` | 每页 `tableVersions`（见“内外两版”）；全部对象：`placement`、`cm`（页面上的外接矩形）、`text`、`rows`（表格）、`chart`（系列、类别、数值，外部工作簿路径 `workbookLink`）、`pictures`（`crop`、`cropClue`、导出的 `file`）、`hiddenInPane`、`group`、`hasContent`；每页有 `hidden`、`notes`、`comments`、`counts`；`layouts` 是母版和版式上看不到的有内容对象 |
| `assets/` | 版面外、跨边、窗格隐藏的图片和裁剪线索的原图（同一内容只存一份）；这些图表的内嵌工作簿、嵌入对象 |
| `renders/` | `--render` 的结果：`第NNN页_连版面外.png` 和 pdf |
| `未审计.md` | 有文件没审计时才有，写在 `--out` 根目录 |

多个文件时各写一个子目录，放在 `--out` 下。

- **位置**（`placement`）：按页面上的绝对坐标算外接矩形，含组的缩放、平移、翻转和旋转，也含形状自身的旋转。`outside` 完全在版面外，`crossing` 出界超过 0.25 cm（满版图略出一点不算），`inside` 版面内，`unknown` 读不到位置（列出来，不跳过）。`mc:AlternateContent` 里的形状 python-pptx 不进，运行时自己取，优先 Fallback。
- **线索**：图片裁掉的面积占原图 `--crop`（默认 20%）以上时标 `cropClue`，导出原图，裁掉的部分可能有内容；版面内但选择窗格里隐藏的对象也列出来。隐藏页的版面内内容在 `全文提取.md`，清单的索引只列版面外的线索。
- **连版面外渲染**：在临时副本上放大画布（超过 PowerPoint 的 56 英寸上限时收窄两侧并警告）、平移页和母版版式上的形状、取消隐藏、画红框，渲染后每页裁成自己的范围。`--engine` 默认 libreoffice，图只用来看位置关系，文字以 `offcanvas.json` 为准。
- **内外两版**：同一内容版面内外各有一版、对不上时，默认以版面内为准，版面外视为旧版或暂存（决策 19）。运行时比同一页的表格：版面外的表和版面内最像的那张，列数相同、行标题八成以上相同、首行至少一半相同，数据格一半以上相同（或数据格不少于 6 个）才算同结构，小表数字全不同更可能是另一组数据；千分位逗号和空白不算差异。对不上的写进 `pages[].tableVersions`（`inside`、`outside`、`sameShare`、`cells`、`rowsOnlyInside`、`rowsOnlyOutside`），版面外的表标 `olderVersionOf`，清单和全文提取里写“疑似旧版”。只是线索：内容照样列出，什么都不改。文字和图表不比，自己对照。
- **没审计的文件**：打不开的文件列进 `notAudited`，写明格式和怎么办：`flat-xml`（Flat OPC，整份是一个 XML）、`ole-compound`（旧版 .ppt 或带密码）、`encrypted-wrapper`（文件头异或 0x5C 后是 `DG_FILE` 的加密外壳）、`broken-package`、`unknown`、`directory`、`missing`。单个对象读不了时记进 `errors`，也让退出码为 1。

其他命令对版面外的处理：`inspect` 给每个形状标 `placement`，版面外的对象进 `problems`，跨边的图片给警告；`audit` 的用语问题和文字线索不查版面外，`--against` 比文字时给版面外的文字加 `[版面外]` 前缀；`clean` 的 `pageHits` 给版面外的命中标 `offCanvas`。

## 渲染检查

渲染图只有版面内：版面外、跨边的部分、裁掉的部分不在图上，隐藏页 LibreOffice 也不出（见“画布外对象”）。`plan-deck render` 把 pptx 转成 pdf，再用 pdftoppm 出逐页 png 交给 agent 看，并用 pdffonts 列出 pdf 里实际嵌入的字体；给了 `--style` 时核对页面字体有没有被换掉。两种引擎：

- **libreoffice**：随时可用，出页过程中反复看。缺字体会换字或漏字，看图前先确认字体可用或已做映射（见 `AGENTS.md`）。表格行高、图表标签位置、字体替换和 PowerPoint 不完全一样。
- **powerpoint**：macOS 上装了 Microsoft PowerPoint 时，用 AppleScript 让 PowerPoint 另存为 PDF，就是汇报现场的样子，**交付前必须过一遍**。PowerPoint 是沙盒应用，读写沙盒外的文件会弹授权框，所以运行时先把 pptx 复制进它的容器目录（`~/Library/Containers/com.microsoft.Powerpoint/Data/plan-deck-render/`），导出后取回并删除临时文件；第一次运行时 macOS 可能要求允许终端控制 PowerPoint。Windows 上可以用同样思路走 PowerPoint 的 COM 接口（`Presentation.SaveAs(..., 32)` 存 PDF），还没做；没有 PowerPoint 的机器上自动化不了，交付时请用户在 PowerPoint 里打开逐页看，并在交付说明里写明没有核对过。

PowerPoint 和 LibreOffice 渲染不一样、已在运行时处理的两处（在真实项目里用 powerpoint 引擎查出来的）：

- 图表文字只写西文字体时，PowerPoint 的中文落回默认字体（宋体）；组件给图表的标题、图例、刻度、数据标签都同时写西文和东亚字体。`render --style` 发现 PDF 里有 `style.json` 没定的字体会警告。
- python-pptx 的自动形状带主题样式引用（`effectRef`），PowerPoint 会画出阴影；组件去掉了样式引用，线和填充都显式写。
