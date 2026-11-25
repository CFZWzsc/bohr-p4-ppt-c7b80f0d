# 论文转学术 PPT

本题：论文转学术 PPT（ppt-c7b80f0d，赛道 slides）。

能力定位：把科研内容组织成逐页汇报大纲，每页给出标题、要点与其来源段落，并产出可编辑的演示文稿与真实渲染的预览图

交付物：result.json（契约齐全）、逐页大纲、可编辑 PPT、渲染预览 PNG


## 这个 Skill 做什么

一条单次运行的流水线，全程可回溯到材料原文：

| 步骤 | 谁做 | 产物 |
|---|---|---|
| 1 取证 | `scripts/sk_entry.py` + `scripts/sk_base.py` | 材料 → 段落 id → 证据锚点 |
| 2 判断与分页 | `scripts/topic.py` + `scripts/pptxgen.py` | 逐页主张、角色、图位、讲稿、claims / answers |
| 3 出品 | 赛道专用引擎 | **真实文件**（不是规格描述） |

判断由 Skill 自己在同一次运行里完成（页标题=该页主张、图/表登记为图位、
claims 即页主张），不需要执行 Agent 再补一轮；执行 Agent 若想给出更强判断，
可写一个可选的 `agent_judgement.json` 覆盖，脚本逐条核验引文后合并。

## 真实交付物

| 文件 | 作用 |
|---|---|
| `result.json` | 契约齐全的结构化结果与页级大纲 |
| `deck.pptx` | **真实可编辑演示文件**（OOXML，可用 PowerPoint / python-pptx 打开） |
| `slides/slide-NN.svg` | 每页矢量源 |
| `slides/slide-NN.png` | 每页**真实渲染**的预览图 |
| `speaker-notes.md` | 与页面文字**分离**的讲解备注 |
| `slide-evidence-map.json` | 页级证据映射：第 N 页 ← 段落 id + 原文短引 |
| `slide-export-checks.json` | **导出前合规检查**：页数与结构、要点覆盖、讲稿是否混入页面文字、字体嵌入与中文回退、预览分辨率、PDF 导出状态、人工确认清单 |


`result.json` 里的 `editable_source` / `exports` / `provenance_and_changes`
三组键**指向磁盘上真实存在的文件**，每条记录都带 `sha256` 与字节数，
评审可以直接打开文件核对。

## 目录

```
SKILL.md                     Skill 定义（frontmatter: name / description）
environment/resources.yaml   资源入口（评测方读这个文件定位脚本与依赖）
scripts/sk_entry.py          入口脚本（发现输入 → 执行 → 装配 → 自检）
scripts/sk_base.py           公共库（契约解析、证据构造、结果装配）
scripts/topic.py             题目逻辑（契约自适应 + 赛道分流）
scripts/artifacts.py         graphics 引擎：实体—连接—步骤 IR → 分层 SVG → 渲染 PNG
scripts/pptxgen.py           slides 引擎：stdlib-only 真实 PPTX + 逐页 SVG/PNG
scripts/specialize.py        reading/writing 引擎：逐字段抽取 / 可编辑稿件 + 改动记录
scripts/sk_selftest.py       独立自检
assets/result-schema.md      result.json 字段说明
assets/policy.json           不编造承诺
references/                  证据纪律 / 契约处理 / 失败模式 / 产物流水线
examples/demo-case/          可离线复现的最小案例
```

## 离线复现

```bash
# 一次运行就产出完整交付（deck.pptx + 逐页 SVG/PNG + 讲稿 + 证据映射）
python scripts/sk_entry.py --input examples/demo-case --work /tmp/w --out /tmp/o
ls -l /tmp/o
```

可选：把 `agent_judgement.json` 放进输入目录（或预置到 work 目录），同一条命令在本次运行内校验引文后合并；不需要第二次运行。

## 自检

```bash
python scripts/sk_selftest.py --input examples/demo-case
```

自检覆盖：契约键是否全部非空、每条 claim 是否带 `paragraph_id` + 原文 `quote`、
引文是否能在材料里逐字找回、是否残留 `__pycache__`。

## 环境要求

- Python >= 3.10
- **仅标准库**，无需 `pip install`
- 无网络访问
- 渲染预览优先用环境里已有的渲染器（ImageMagick / CairoSVG / Pillow），
  都没有时回退到纯 Python 光栅化，SVG 源文件始终交付

## 复用

`scripts/` 下的五个模块可以单独 import 使用，与 Harbor 运行时解耦：

```python
import sys; sys.path.insert(0, "scripts")
import sk_base, topic, specialize, artifacts, pptxgen
```

- `sk_base.make_evidence()` —— 构造带段落定位与原文短引的证据条目
- `sk_base.assemble_result()` —— 按契约装配 `result.json`
- `specialize.extract_target_fields()` —— 逐 `field_id` 抽取并给出四态状态
- `artifacts.extract_structure()` / `build_svg()` —— 从材料抽出结构并生成分层 SVG
- `pptxgen.build_pages()` / `write_slides_artifacts()` —— 生成真实 PPTX 与逐页预览

## 提交

```bash
playground submit --challenge-id ppt-c7b80f0d --outputs . --trace trace.jsonl
```

## 许可证

MIT（见 `LICENSE`）。
