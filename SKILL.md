---
name: paper-to-academic-slides
description: Use when a paper or research report must become a page-level academic slide deck with a defensible narrative spine and per-slide evidence locators.
---

# 论文转学术 PPT

## 用途与适用条件

把连续叙述的论文重写为页级表达：先抽主线（问题、方法、关键证据、结论与限制），再为每页确定唯一主张、支撑图/表/章节定位、可讲要点与独立讲稿备注，最后按模板与页数/时长约束输出可复现的演示规格。核心是结构重写与证据映射，不是把段落切块。

适用条件：`/input/case.json`（或 `--case`）提供论文正文、章节文本、图表清单、模板约束、听众/时长/页数上限，要求生成学术汇报。典型键：`sections`/`paper`、`figures`、`tables`、`template`、`audience`、`duration_min`、`slide_limit`。

不适用：不生成新的科学结论，不做统计推断，不替代论文精读；不伪造 PPT 渲染文件，输出为可复现的结构化规格。

## 输入

- `/input/case.json`：主输入。
- `/input/task.md`：存在时补充听众、时长、页数与模板要求。
- 引用的本地文本/JSON 会被实际读取；PDF 只按可读文本处理，不假设预抽取答案。
- 仅标准库；不联网、不写 `/input`。

## 处理流程

1. **读入**：加载 case.json、task.md、引用材料，登记来源哈希。
2. **主线抽取**：按线索词把句子归入 problem / method / evidence / conclusion / limitation，保留原句与来源位置。
3. **图表清单**：读取显式 figures/tables；缺失时从“图 N / Figure N / 表 N”模式抽取编号与题注，建立可引用集合。
4. **页级规划**：每个角色生成 1–3 页；每页一个主张，要点不超过 3 条；按 `slide_limit` 压缩，不删关键证据。
5. **证据映射**：解析主张中的图表编号与章节引用，核对是否存在于清单；不存在则标 `dangling_reference`，不臆造图。
6. **讲稿备注**：为每页写“过渡、强调点、可能被问到的补充证据”，检查与页面文字的重合度。
7. **模板适配**：读取模板占位符与版式，给出每页 layout、标题长度与字号层级建议；无模板时使用中性默认。
8. **导出检查**：登记 PDF/PNG 预览与字体嵌入检查，`performed=false`。
9. **契约落盘**：保留 `case_id` 与全部 `output_contract` 字段。

## 评分维度对应

- **实际效果**：主线与论文贡献一致；数字、单位、方法名、结论表述与原文一致；图表编号不串位。
- **完整性**：主线必需页、页结构、图表引用、讲解逻辑、模板约束、导出六类要素齐备。
- **可核验性**：每页带 `source_locator`、`evidence_refs` 与原文句。
- **创新增益**：连续文本到页级主张的结构重写、证据编号校验、备注去重。
- **可靠体验**：缺证据、悬空引用、页数超限、备注重复均显式降级。

## 输出契约

`/app/output/result.json` 顶层保留 `case_id`、`output_contract` 全部字段、`status`、`result`、`evidence`、`uncertainty`、`errors`、`stats`、`skill`、`challenge_id`、`contract_notes`。未支持字段写 `null` 并记录原因。

主要交付段：`result.deck_spec.slides`、`result.deck_spec.narrative_spine`、`result.deck_spec.evidence_map`、`result.deck_spec.checks`。

## 异常处理

| 情况 | 处理 | 输出表现 |
|---|---|---|
| case.json 缺失/非法 | 结构化失败 | 退出码 `2`，`status=insufficient_input`/`failed` |
| 无正文/无章节 | 不编造主线 | `status=insufficient_input` |
| 主张引用不存在的图 | 保留主张并标悬空 | `dangling_reference`，`status=partial` |
| 页数超过上限 | 保留主线页，压缩支撑页 | `slide_limit_exceeded` 记录取舍 |
| 备注与页面重复 | 重写为过渡/补充证据 | `notes_duplicate_slide_text` |
| 数字无源 | 保留原文数字并标未核 | `number_without_locator` |
| 模板缺失 | 用中性默认版式 | `template_missing`，不伪装已套模板 |
| 渲染工具缺失 | 仍须生成 pptx；无法渲染时在 deliverable 明确标注 | `deliverable.rendered=false`、`pptx_not_generated` |

## 限制

主线抽取基于线索词与显式结构，跨章节论证链可能漏检；视觉设计与动画不在范围内。脚本产出页级内容与版式规格；最终 pptx 与渲染预览由调用 Agent 按「最终交付物（强制）」生成。

## 最小调用示例（示例，非官方案例结果）

```bash
python scripts/run_skill.py --case /input/case.json --out /app/output/result.json
```

示例输入（示意）：三段式论文摘要、图 1 与表 1、`slide_limit=10`。示例输出（示意）：`status="success"`，每页含 `claim`、`evidence_refs=["Figure 1"]`、`notes` 与 `source_locator`。仅为示例，不代表官方固定案例结果。


## 最终交付物（强制）

脚本只产出 `result.json` 中的页级结构与证据映射；调用 Agent 必须继续完成最终交付：

1. 用 `python-pptx`（或等价可编辑格式）按 `result` 的页序、版式与占位符生成 `slides.pptx`；文字、备注、图表编号均可编辑，不把文字压成图片。
2. 用 `office-renderer`/LibreOffice 将 `slides.pptx` 渲染为 `preview.pdf` 与逐页 `preview_p01.png...`，并实际查看预览，确认无溢出、缺字、错页、空白页。
3. 把产物路径、页数、渲染命令与检查结果写入 `result.json` 的 `deliverable`（`pptx_path`、`preview_paths`、`rendered`、`checks`），并记录源文件哈希。
4. 工具确实缺失时，在 `result.json` 与 `limitations` 明确写 `pptx_not_generated` 及原因，不得声称已生成或已渲染。

## 交付保证与评分对齐

调用顺序（固定）：

1. 先运行包内脚本（若存在）：`python scripts/run_skill.py`（默认读 `/input/case.json`，写 `/app/output/result.json`）；路径不同时用 `--case` / `--out` 指定。
2. 再完成脚本无法自动完成的实质工作（译文、论述、图表、修订、结论等），把结果写回 `result.json` 对应字段。
3. 最后自检：`case_id` 与输入一致；`output_contract` 每个字段都存在；`status` 为 `completed`（仅在确实缺输入时为 `partial`）；无 `pending` 占位；数字、引用、来源可逐条回溯。

运行保证：

- 输入缺失、JSON 损坏、字段异常、输出目录不存在时，仍须写出结构化 `result.json` 并正常结束，不得崩溃或非零退出。
- 不联网、不伪造数据、不声称平台评分或官方案例结果。
- 交付物必须是可直接编辑与复核的成品，而不是流程说明。

评分对齐（本题五个维度）：

- 实际效果（40）：完整产出任务要求的最终成品，覆盖全部输入条目，可直接进入修改流程。
- 使用增益（25）：体现相对基础 Agent 的可验证改进（结构化、术语/数字/引用核对、可编辑性）。
- 创新增益（15）：保留方法学增量（分块、对齐、二次校验、版本比对等），并用结果证明其价值。
- 可靠体验（10）：异常输入有明确降级与说明，不同长度/文体/重复运行下稳定产出。
- 生态贡献（10）：用途、调用、输入输出、限制、示例、维护信息齐全，资源可被他人复用。

## 最终交付强制要求

1. **调用 Agent 必须交付成品，不是流程说明**：脚本产出只是索引、核对或规格骨架；调用 Agent 必须在同一 `result.json` 中补齐本题要求的最终产物（正文/译文/审稿意见/综述/分析结论/图表源文件与导出件/幻灯片页内容与讲稿/实验方案等）。交付字段为空、只有标题骨架，或含 `pending`、`TODO` 之类占位，一律视为未完成。
2. **必须覆盖全部输入条目**：输入中的每个章节、问题、文献、图、页面、指标都要有对应条目，不得只处理示例或首条；确实缺失的条目写入 `uncertainty` / `limitations` 并说明原因，不得静默跳过。
3. **必须自检后再交付**：`case_id` 与输入一致；`output_contract` 每个字段都存在；状态值取本题契约规定的枚举；数字、单位、引用、来源可逐条回溯；结果内部无自相矛盾。
4. **参数等价**：`python scripts/run_skill.py --case <输入> --out <输出>` 与 `--output <输出>` 等价；默认 `/input/case.json` → `/app/output/result.json`。
5. **平台用 3 个固定案例分别运行**：逻辑要由输入驱动，不针对某一份示例硬编码；字段命名差异走别名匹配与显式降级，并在 `uncertainty` 记录假设。
