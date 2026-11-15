"""地图主题：调 gis-composer `theme resolve`，读解析后的主题（extends 和记号都已展开）。

plan-deck 不自己解析主题链，只走 composer 的命令行：
  <GIS_COMPOSER_CMD> --project <composer 项目> theme resolve [<主题>]
GIS_COMPOSER_CMD 是命令行字符串（例如 "node <composer 运行时>/src/agent-entry.js"），先看同名环境变量，
再看 gis-plugin 的共用文件 ~/.config/gis-plugin/engines.env。
"""
import json
import os
import shlex
import subprocess

ENGINES_ENV = os.path.join('~', '.config', 'gis-plugin', 'engines.env')
KIND = 'resolved-theme'
SCHEMA_VERSION = 1


class ThemeError(ValueError):
    pass


def split_command(command, windows=os.name == 'nt'):
    """按 shell 规则拆分；Windows 上反斜杠是路径分隔符，不是转义符。引号都会去掉。"""
    lexer = shlex.shlex(command, posix=True)
    lexer.whitespace_split = True
    lexer.commenters = ''
    if windows:
        lexer.escape = ''
    return list(lexer)


def composer_cmd():
    """GIS_COMPOSER_CMD 拆成参数列表；没配置时返回 None。"""
    value = os.environ.get('GIS_COMPOSER_CMD')
    if not value:
        path = os.path.expanduser(ENGINES_ENV)
        if os.path.isfile(path):
            with open(path, encoding='utf-8') as f:
                for line in f:
                    key, _, raw = line.strip().partition('=')
                    if key == 'GIS_COMPOSER_CMD':
                        tokens = split_command(raw)
                        value = tokens[0] if len(tokens) == 1 else raw     # 整条加了引号时去掉外层引号
    return split_command(value) if value else None


def resolve_theme(project, theme=None, cmd=None):
    """返回 composer `theme resolve` 的 JSON（kind resolved-theme）。命令不可用、退出码非 0、输出不认识时抛 ThemeError。"""
    cmd = cmd or composer_cmd()
    if not cmd:
        raise ThemeError('style.json 写了 mapTheme，但找不到 gis-composer：设置环境变量 GIS_COMPOSER_CMD'
                         '（如 "node <composer 运行时>/src/agent-entry.js"），或写进 ~/.config/gis-plugin/engines.env')
    if not os.path.isdir(project):
        raise ThemeError(f'mapTheme.project 不是目录：{project}')
    argv = [*cmd, '--project', project, 'theme', 'resolve', *([theme] if theme else [])]
    try:
        res = subprocess.run(argv, capture_output=True, text=True, encoding='utf-8', timeout=120)
    except (OSError, subprocess.TimeoutExpired) as e:
        raise ThemeError(f'调用 gis-composer theme resolve 失败：{e}') from e
    if res.returncode != 0:
        detail = (res.stdout or '').strip() or (res.stderr or '').strip()
        raise ThemeError(f'gis-composer theme resolve 退出码 {res.returncode}：{detail[-800:]}')
    try:
        doc = json.loads(res.stdout)
    except json.JSONDecodeError as e:
        raise ThemeError(f'gis-composer theme resolve 的输出不是 JSON：{e}') from e
    return check_resolved(doc)


def check_resolved(doc):
    if doc.get('kind') != KIND or doc.get('schemaVersion') != SCHEMA_VERSION:
        raise ThemeError(f'不认识的主题解析结果：kind={doc.get("kind")!r}、schemaVersion={doc.get("schemaVersion")!r}，'
                         f'本运行时只读 {KIND} v{SCHEMA_VERSION}')
    return doc


def project_theme_file(doc):
    """主题链里项目自己的主题文件（chain 最后一项，不是内置主题时）：(项目内路径, sha256) 或 None。"""
    chain = doc.get('chain') or []
    leaf = chain[-1] if chain else None
    if not leaf or leaf.get('builtIn'):
        return None
    return leaf.get('file'), leaf.get('sha256')
