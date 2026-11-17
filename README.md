# report-iterate

一套用于迭代单篇论文和多篇关联报告的 Codex 技能。它把评审意见和证据核对转成可审阅、可续跑的修改工单，并在修改 Word 时核对题注、书签和交叉引用。

- [技能入口](SKILL.md)：说明审读、修改和交付流程。普通问题直接处理；主旨、框架或核心结论的变化先给原文位置与备选方案。
- [调研与路线探索](references/research-and-ideas.md)：区分启发性例子和明确要求，比较不同机制及其证据、代价和验证条件。
- [Word 执行说明](references/runtime.md)：使用 Python 3.10+ 和 `lxml` 检查 DOCX、执行工单和恢复中断。实际更新目录和域需要 Windows 上的 Word 和 PowerShell 7。
- [项目配置](references/project-profile.md)：可按研究项目维护偏好，另附可选的航天调研配置。

当前版本：v2.1。仓库包含技能源码和使用指引。
