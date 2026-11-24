# Word 工单与验证

需要 Python 3.10+ 和 `lxml`。不需要 Python 版 Office API 或绑定用户目录。脚本从技能自身目录运行；以下命令中的 `scripts/` 相对于技能根目录。首次使用在隔离副本验证，不向原稿或已存在的输出目录写入。

```text
python scripts/iterate.py inspect INPUT.docx --out inspect.json
python scripts/iterate.py run plan.json run-storage --stop-after 2
python scripts/iterate.py run plan.json run-storage
python scripts/iterate.py export run-storage trial-output
python scripts/iterate.py compare INPUT.docx trial-output/INPUT.docx
```

`inspect` 返回以原有 Word 样式、自动编号、域、书签为依据的记录，包括跨段 TOC、分段的域代码、正文、表格的全部段落和可用的定位锚点。正文开头提到“图3-1”不是判断题注的依据。使用 `--summary` 可省略全文。Word 显示的域结果、实际域指令和引用目标分开看，未用编号域的可疑题注要人工核实。`compare` 只比较保护结构及报告现存的检查结果；不宣称语义一致、图片可编辑或原生 Word 自动编号已经更新。

## 工单计划（JSON，format 1）

- `documents`：`{"id":"r1","path":"input/报告.docx","sha256":"..."}`。`path` 可为绝对路径或相对于计划，源文件名互异。用 `inspect` 输出的真实 `file_sha256`。
- `items`：每项填 `id`（英文数字和 `_.-`，唯一）、`title`、`rationale`、`risk`（`ordinary`/`major`）、`evidence`（`not_required`/`verified`/`pending`）。`verified` 附 `sources`（来源及原页、节、图表或数据定位）；脚本不验证来源里的语义。重大项填写 `impact`、`options`（至少两种，给可理解的具体影响）。
- 每项 `changes` 用数组放同一问题的全部联动。每个元素形如 `{"document":"r1","anchor":{...},"old":"待替换的唯一原片段","new":"经审读的改法"}`。`anchor` 从对应 `inspect.json` 的 `anchors` 原样复制，包含 `location`、`section`、`paragraph`、`before`、`after`、`para_id`。请用紧邻原段上下文区分相似句。一个工单内同一段合成一个替换；不同工单若竞争同一段，先合并或另建已核验的新基线。`new` 允许空字符串（只删选定普通文字）。跨段、删整段、表格结构、图像、标题、实际题注、超链接、域结果或书签变更改走对应编辑方式。

例如同一术语在两份文档改名时，用一个 `ordinary` 工单列两处。两处都通过定位、写入和结构复核后才提交，避免半改。近似命中、真实结构变化及无法验证的项不硬改。无需结构变更的重大选择一旦由用户确认，可以在新计划中用已确认范围处理；实际结构改动仍由专用编辑方式完成并复核，不把“已批准”解释成可以绕过文本脚本保护。

## 执行状态

多轮自动审查按 [autonomous-review.md](autonomous-review.md) 管理审读、核验及复审。这里的 `run` 仅执行已形成的文字工单，不自行调用模型、检索来源或生成下一轮意见。每轮从已核验导出建新计划和新运行目录；同一轮中断后继续使用原目录。循环索引记录各轮的计划、运行目录和已核验结果，不能代替批次的原子提交状态。新一轮没有有效工单时直接记录复审结果，不创建虚假的替换。

`run` 的退出码：`0` 全部普通项处理完成且无挂起；`3` 有 `review`、`evidence_pending` 或 `blocked`；`75` 按 `--stop-after` 有意暂停；`2` 计划不合法、卷损坏或工具失败。`applied` 是写入和静态结构已验，原生文档另验。报告 `new_edits`、`new_records`、`reused_records`；第二次执行同一批应当 `new_edits=0` 且状态文件与成稿 SHA-256 不变。

各文件先写成经验证的不可变快照，`state.json` 的原子更新是唯一提交点。多文档工单在同一状态更新中提交。中途关闭程序后使用原命令续跑；未提交的快照即使留在存储也不会被导出。保留它们到交付完成，勿逐个手删。脚本锁阻止两个进程同时修改同一批。文件系统需支持同卷原子替换和锁；跨设备同步或突然断电的持久性不在本技能承诺内。更改计划内容会改变哈希；从已核验的导出开新批，不编辑状态文件。`export` 可供审读待定项的中间结果，`receipts.json` 会如实保留待定状态，不将其当成全部通过。

## 用 Word 检查实际更新结果

批次工具的 `export` 生成可读的工作副本。小修可将它放在工作区，完成内容及适用的原生检查后，把已核验文件的路径和哈希写进累计记录。下一项修改在该文件上重新取锚点、开新批次；同一项断点续跑仍用原计划和存储目录。正式发布从累计记录选取已核验的最新稿，逐项确认已纳入发布范围。

在装有 Microsoft Word 的 Windows 中用 PowerShell 7。`-Inputs` 仅指本次已导出的工作副本；使用与输入父目录不重叠的新输出目录。Word 只打开核验副本。遇到需要外部取数的域或外链资源先报告，走隔离、已授权的专用流程，避免自动更新外部内容。

```powershell
& 'scripts/word-check.ps1' -Inputs @('trial-output/报告.docx') -OutputDirectory 'word-verification'
```

检查 `word-validation.json` 的版本、域和目录更新、重开、各类 REF 目标及题注变化。阅读重点页的实际显示（标题层级、分页、表头、题注相邻、图件可读性等）；不能由脚本或页数代替这一环节。若 Word 不可用，只记录 OOXML 验收，并列出原生待测事项。若工具遇到不明 `Bad Request` 或已完成部分操作，先检查实际保存路径、状态和失败步骤，再从未提交项继续，不自动反复重跑。不根据新版 Word 或文件扩展名推断旧版 Office 通过。
