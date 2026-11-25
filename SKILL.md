---
name: bohr-slides
description: 当需要汇报科研材料并产出可复核的结构化结果时使用。 本题范围：论文转学术 PPT。 题目家族：论文转学术 PPT。 产出 result.json，其中包含运行时 output_contract 要求的全部字段，每条结论附带原文证据（段落定位 + 短引），无法从材料判断的部分如实写入 uncertainty 而非编造。
---

# 论文转学术 PPT

## 用途

本题属于**slides** 赛道。目标能力：把科研内容组织成逐页汇报大纲，每页给出标题、要点与其来源段落，并产出可编辑的演示文稿与真实渲染的预览图。

本 Skill 在 DeepSeek Harness 中被加载到 `/skills/weekly-skill`。
Harness 会在 3 个固定科研案例上分别调用它；每个案例的输入是
`/input/case.json`（含 `task_instruction` 与 `output_contract`）与同目录下的
科研材料文件。

## 任务要求（来自题面）

(见 /input/case.json 的 task_instruction)

本题可预知的契约键（来自题面「Skill 运行后的输出要求」表格，另有部分键按**同赛道兄弟题的公开接口推断**；**实际以运行时 `/input/case.json` 为准**，下表仅供预热）：

- `case_id`
- `compliance_checks`
- `content_map`
- `editable_source`
- `exceptions`
- `export_artifact`
- `page_evidence_map`
- `render_and_version_log`
- `slide_outline`
- `status`

## 交付物

result.json（契约齐全）、逐页大纲、可编辑 PPT、渲染预览 PNG

交付物的**内容契约**由运行时 `output_contract` 决定（见下）；下面「交付物落盘
要求」给出的是必须真实存在的**文件**。

## 执行方式

入口脚本：`scripts/sk_entry.py`。它会读取 `/input`，产出 `/app/output/result.json`。

```bash
# 在案例环境中执行（Harbor 会自动调用；手动复现同一条命令）
python /skills/weekly-skill/scripts/sk_entry.py \
  --input /input --work /app/work --out /app/output
```

执行结束后必须存在且可解析：

- `/app/output/result.json` —— 契约字段齐全的最终结果
- `/app/output/run-review.md` —— 人类可读的证据索引与不确定性清单（复核用）

### 单次运行即产出完整交付（判断内置，可选覆盖）

脚本是**确定性**的，但**不是只做取证的半成品**：一次运行就完成
「读材料 → 判断每句的叙事角色 → 按角色组织页面 → 生成可编辑 PPTX 与逐页
矢量/位图 → 写页级证据映射与分离的讲稿 → 给出实质判断（claims / answers）」。
正常用法只有一条命令：

```bash
python /skills/weekly-skill/scripts/sk_entry.py \
  --input /input --work /app/work --out /app/output
```

- **幻灯片家族**：页面按叙事角色成页（研究问题 / 方法路线 / 关键结果 /
  讨论与含义 / 局限与结论），**页标题 = 该页主张**（角色前缀 + 材料原句片段），
  同角色多页标题不重复；判不出角色的句子按段落顺序兜底成「研究内容」页。
- **图位**：材料点名的图/表登记为结构化图位（图号 + 图注/引用句 + 段落号），
  源图未随材料提供时留图位并标注不重绘，不编造图片。
- **判断内置**：每条 claim 就是该页主张本身（页级主张 + 段落 id + 逐字短引），
  交付物成形即 `status=answered`；真正的边界（运行时无 PDF 转换器、材料不足）
  逐条写入 `uncertainty`。
- 材料不足或渲染失败时 `status` 才停在 `partial`，并在 `status_reason`
  写明**具体**缺什么（而不是"等 Agent 判断"）。

**单次运行**：判断在本次运行内完成，不存在「先跑一遍、再跑一遍」的两段式。
执行 Agent 若要给出更强判断，可把 `agent_judgement.json`（结构见
`references/judgement-override.md`）放进输入目录或 work 目录；脚本在**同一次运行**里
逐条校验引文后合并，未通过校验的条目退回并记录原因。不提供该文件也不影响交付完整性。

> 判断可以自由，**证据不可以自由**：每条 quote 必须能在材料原文里逐字找回。

### 脚本职责

| 脚本 | 职责 |
|---|---|
| `scripts/sk_entry.py` | 入口：定位输入、发现材料、执行 `topic.run()`、装配 `result.json`、自检并写复核产物 |
| `scripts/sk_base.py` | 公共库：契约解析、材料发现、证据构造、结果装配、自检工具 |
| `scripts/topic.py` | 题目逻辑：读材料 → 建证据索引 → 按契约键名语义产出字段 → 校验并合并你的判断 |
| `scripts/specialize.py` | reading / writing 引擎：逐 `field_id` 抽取（四态状态）；可编辑稿件 + 改动记录 + 来源映射 |
| `scripts/artifacts.py` | graphics 引擎：实体—连接—步骤 IR → 分层 SVG → 渲染 PNG + 来源记录 |
| `scripts/pptxgen.py` | slides 引擎：stdlib-only 真实 PPTX + 逐页 SVG/PNG + 独立讲稿备注 |

**这些引擎会真的写文件。** 每条赛道具体落盘哪些文件、每个文件里有什么、
如何回挂到契约键，见 `references/artifact-pipeline.md`。

### 运行时契约是唯一权威

`output_contract` **只在运行时可见**（`/input/case.json`）。
不要硬编码任何字段名，也不要凭题面示例臆造字段：

1. 读取 `/input/case.json` 的 `output_contract`；
2. 对每个契约键，按**键名语义**判断它应该承载什么（证据 / 结论 / 不确定性 /
   交付物 / 覆盖度 / 变更记录 …）；
3. 产出对应内容，**不允许**留下 `null` 或"未产出"占位；
4. 若该键确实无法从材料得出，填入结构化的"无法判断"条目（带 `kind`、`item`、
   `detail`、`resolution`），而不是留空。

## 交付物落盘要求

除 `result.json` 外，**可编辑的交付物源文件与其渲染预览必须真实存在**：

| 赛道 | 必须落到 `/app/output` 的文件 |
|---|---|
| slides | `result.json`（契约齐全的结构化结果 + 逐页大纲） |
| slides | `deck.pptx` —— 真实可编辑演示文件（OOXML） |
| slides | `slides/slide-NN.svg` / `slides/slide-NN.png` —— 每页矢量源与渲染预览 |
| slides | `speaker-notes.md` —— 与页面文字分离的讲解备注 |
| slides | `slide-evidence-map.json` —— 页级证据映射 |
| slides | `slide-export-checks.json` —— 导出前合规检查：页数与结构、要点覆盖、讲稿是否混入页面文字、字体嵌入与中文回退、预览分辨率、PDF 导出状态、人工确认清单 |

平台提供的 `office-renderer` / `pdf-render-tool` / `latex-compiler` 可用来渲染预览。
**只有规格描述、没有真实文件，等于没交付** —— 评审要打开文件看。

### 导出前合规检查（pre-export compliance）

导出件落盘后，本 Skill 会**自动核对**并在 `*-export-checks.json` 里留档，
让评审不必逐个打开文件猜：

| 检查维度 | 记什么 |
|---|---|
| 目标格式 | 期望的 SVG / PNG / PDF 是否齐全、各自字节数、是否是**真矢量** |
| 画布与分辨率 | 像素尺寸 → pt → mm 换算，实测 dpi，目标 dpi |
| 字体 | 字体清单、嵌入情况、中文回退策略、是否存在方框/缺字风险 |
| 色彩空间 | 色彩空间标注（sRGB 等）与是否有嵌入配置 |
| 线宽 / 分层 | 线宽是否过细、图层是否可拆、图层清单 |
| 数据一致性 | 图上出现的数值与源数据表的比对结果 |
| 讲稿隔离 | 备注文字是否混进页面文字（逐页最大重叠片段） |
| 人工确认 | 需人眼确认的项一律标 `pending_human_review`，**不代签** |

失败的检查项会**原样保留为 fail**，不会被改写成通过；不可用的工具（如运行时
没有矢量转换器或 LibreOffice）也如实记为 `unavailable`，不做假成功。

## 质量红线

- 每页要点都能回指原文段落
- 页数与内容密度符合任务卡要求，不灌水也不跳页
- 预览图与可编辑文件内容一致
- 讲稿只进备注区，不混进页面文字（导出前逐页比对）
- 字体方案写全 latin/ea/cs 与中文回退，换机打开不被替换
- 导出前把页数、要点覆盖、字体嵌入、预览分辨率、PDF 导出状态逐项核对并留档

通用红线（违反即失分）：

- **不编造**：材料没有的数字、引用、图表、方法、结论，一律不写。写不出来就进
  `uncertainty`，并在 `status_reason` 里说明。
- **可回溯**：每条结论必须能指回原文的段落 id 与短引；引文必须是原文片段，
  不做改写、不做拼接。
- **不提前交卷**：材料可能读不全。若只读到部分文件，`status` 应为 `partial`，
  并在 `coverage.not_read` 写明未读的部分。
- **不自评**：不要在结果里写"质量良好""得分很高"这类自评语句。

## 失败与边界

| 情况 | 正确处理 |
|---|---|
| 材料是扫描件、无文本层 | 用 `scripts/sk_base.py` 的 PDF 解析；仍无文本则 `status=unsupported` 并在 `uncertainty` 写明"无文本层" |
| 材料是图片/图表 | 用平台提供的 `pdf-render-tool` / `ocr-tool`；OCR 结果标注为推断 |
| 材料格式未知 | 记录文件路径与可读性，不猜测内容 |
| 契约键无法解读 | 产出该键的结构化说明而非空值，并记录在 `uncertainty` |

详见 `references/failure-modes.md` 与 `references/evidence-discipline.md`。

## 自检

`result.json` 写完后由 `scripts/sk_entry.py` 自动自检（结构、契约覆盖、
证据完整性）。自检失败时以 `partial` 提交并在 `status_reason` 写明原因，
不要伪造 `answered`。

```bash
python /skills/weekly-skill/scripts/sk_selftest.py --input /input
```

## 参考

- `references/evidence-discipline.md` —— 证据与不确定性纪律
- `references/contract-handling.md` —— 运行时契约的读法与字段语义表
- `references/judgement-override.md` —— 可选判断覆盖文件的结构（同一次运行内读取与合并）
- `references/failure-modes.md` —— 常见失败模式与正确退路
- `references/artifact-pipeline.md` —— 每条赛道真实落盘哪些文件、怎么回挂契约键
- `assets/result-schema.md` —— `result.json` 的字段说明
- `assets/policy.json` —— 不编造承诺与来源声明
- `examples/demo-case/` —— 一个可离线复现的最小案例
