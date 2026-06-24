# Changelog

## 2026-06-24

### 合同审查

- 新增 Word 红线稿执行链路：合同审查在用户要求“修订模式 / 红线稿 / Word 修订格式”时，先生成 `redline-plan.json`，再运行 `scripts/redline/apply_redline_plan.py`，以原合同为只读来源生成新的审核修订稿。
- 新增真实 DOCX 修订结构：红线执行器会写入 `w:trackRevisions`，并通过 `w:ins` / `w:del` 生成真实 Word 修订痕迹；批注写入 `word/comments.xml`，同时补齐关系文件和内容类型。
- 新增红线结构 QA：`scripts/redline/qa_redline.py` 检查修订开关、插入、删除、批注数量和 comments relationship。
- 新增审查立场硬闸门：用户未明确确认甲方、乙方或中立审查前，不得开始实质审查、生成问题清单、飞书报告、流程图、Word 红线稿或 `redline-plan.json`；红线执行器同步校验 `meta.party_role_confirmed` 和确认来源。
- 增强红线执行保护：禁止默认整段替换/删除，要求 `target_text` 使用最小修改片段；新增 `allow_full_paragraph_replace`、`allow_unconfirmed_revision`、`allow_direct_revision` 显式授权字段。
- 增强动作降级规则：待确认、占位、待补事实默认转为批注；`handling_advice=可优化` 默认转为 `report-only`，不直接写入合同正文。
- 增强 QA：检测单处插入/删除文本长度，通过 `--max-del-chars` 和 `--max-ins-chars` 识别整段替换噪音。
- 新增 `references/修订策略.md`：将“必须修改 / 建议修改 / 需客户确认 / 可优化”映射到 `replace / insert / delete / comment / report-only`，强调轻微确定性修订通常只留修订痕迹，重大修订才保留解释性批注。
- 新增 `references/redline-plan-protocol.md`：固化 `redline-plan.json` 的字段、路径、动作分流、审查人信息、审查时限 / 上线期、正式审查意见路径和解析方式。
- 更新 `references/完整流程.md` 和 `references/审查交付规范.md`：将红线计划、执行器、QA、执行日志、版本记录和正式意见书路径纳入合同审查交付流程。
- 新增单元测试：覆盖修订与批注生成、重复文本未指定 `occurrence` 时失败、指定 `occurrence` 后精准命中、`selector.contains` 定位、`comment_required` 批注控制、整段替换阻断、待确认项降级和附件引用确定性修订。
- 修复 DOCX XML 写出命名空间：关系文件和 Content Types 不再出现 `ns0:` 前缀；新增未确认审查立场时禁止生成红线稿的测试。

### Legal Skills 总控与子 Skill

- 法律工作总控新增标准响应骨架，要求稳定披露 Skill 路径、完整读取、关键数据提取与校验、读取复查摘要、法规校验摘要、来源边界和用户确认事项。
- 正式交付物门禁进一步收紧：拒绝跳过材料读取、OCR 兜底、法规名称编号内容核验和现行有效核验；拦截说明必须写出对应门禁名称。
- 法律咨询助手新增 quick reply 硬门禁：明显微信客户咨询只输出一条短微信式回复，不展开正式法律意见或交付流程。
- 劳动关系认定与经济补偿计算新增计算门禁：N、N+1、2N、月工资基数、加班费、双倍工资等计算前必须完整读取劳动合同、工资流水、社保记录、解除通知、考勤或沟通记录。
- 诉讼文书起草新增证据目录格式门禁：默认使用含 `第一组证据` 和 `证明目的` 的分组文本段落结构，未经用户明确覆盖不得改成表格版。
- 产品法务、监管合规监测、案件材料生成专业文档同步强化法规核验底线：不得用模型记忆替代法律法规核验。

### 致谢

本轮合同审查红线执行链路参考了 [cat-xierluo/contract-copilot.skill](https://github.com/cat-xierluo/contract-copilot.skill) 的开源项目设计，尤其是“审查计划 -> 动作执行 -> DOCX 修订/批注 -> 报告/日志/归档/校验”的稳定技术路径。感谢杨卫薪律师及该项目的开源启发。
