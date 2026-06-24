# Changelog

## 2026-06-24

### 合同审查

- 新增 Word 红线稿执行链路：合同审查在用户要求“修订模式 / 红线稿 / Word 修订格式”时，先生成 `redline-plan.json`，再运行 `scripts/redline/apply_redline_plan.py`，以原合同为只读来源生成新的审核修订稿。
- 新增真实 DOCX 修订结构：红线执行器会写入 `w:trackRevisions`，并通过 `w:ins` / `w:del` 生成真实 Word 修订痕迹；批注写入 `word/comments.xml`，同时补齐关系文件和内容类型。
- 新增红线结构 QA：`scripts/redline/qa_redline.py` 检查修订开关、插入、删除、批注数量和 comments relationship。
- 新增 `references/修订策略.md`：将“必须修改 / 建议修改 / 需客户确认 / 可优化”映射到 `replace / insert / delete / comment / report-only`，强调轻微确定性修订通常只留修订痕迹，重大修订才保留解释性批注。
- 新增 `references/redline-plan-protocol.md`：固化 `redline-plan.json` 的字段、路径、动作分流、审查人信息、审查时限 / 上线期、正式审查意见路径和解析方式。
- 更新 `references/完整流程.md` 和 `references/审查交付规范.md`：将红线计划、执行器、QA、执行日志、版本记录和正式意见书路径纳入合同审查交付流程。
- 新增单元测试：覆盖修订与批注生成、重复文本未指定 `occurrence` 时失败、指定 `occurrence` 后精准命中、`selector.contains` 定位和 `comment_required` 批注控制。

### 致谢

本轮合同审查红线执行链路参考了 [cat-xierluo/contract-copilot.skill](https://github.com/cat-xierluo/contract-copilot.skill) 的开源项目设计，尤其是“审查计划 -> 动作执行 -> DOCX 修订/批注 -> 报告/日志/归档/校验”的稳定技术路径。感谢杨卫薪律师及该项目的开源启发。
