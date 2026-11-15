# AGENTS.md

## 语言

所有回复和文档使用简体中文。

## 项目定位

规划汇报 PPT 工具，见 [README.md](README.md)。动手前先读 `docs/design-principles.md`，它记录了已定的决策；改动若与其冲突，先和用户确认。

## 不绑定 harness

- 机制（pptx 读写、版式生成、渲染检查）放进本仓库的 CLI；skill 文档（`SKILL.md`）只写工作流和判断规则，规则全文留在 `docs/`，两边改动要保持一致。
- 不依赖某个 harness 独有的工具或 API；Claude Code 与 Codex 都应能按同一份说明使用。

## 渲染检查

LibreOffice 渲染时缺“微软雅黑”等字体会漏字，看缩略图判断版式前先确认字体可用或已做映射。

## 协作约定

- 任务、计划、进度、待决问题记在 GitHub Issue/PR，不写进仓库 md。
- `docs/` 只放做完后仍然成立的结论：架构、决策及理由、调研结论。
- 试做产物和临时分析放在 `.work/`（已忽略），不进仓库。
- 推送、合并、开 PR、关 Issue 需要用户授权。
