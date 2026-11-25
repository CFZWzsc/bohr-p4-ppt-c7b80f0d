# 单次运行与可选覆盖协议

## 默认：一次运行就完成判断

脚本是**确定性**的，但判断不再外包给第二个 Agent 回合：读材料 → 建证据索引
→ 按叙事角色分页（页标题 = 该页主张）→ 产出可编辑交付物 → 把交付物逐项回溯
到原文 → 生成证据锚定的 `claims` / `answers`。默认只需一条命令；交付物成形即
`status=answered`，剩余边界（读不出的文件、缺渲染器、PDF 是否矢量）逐条写进
`uncertainty`。材料不足或渲染失败时才停在 `partial`，并写明具体缺什么。

## 可选覆盖（在同一份工作目录上多跑一次）

如果执行 Agent 想用自己的领域判断覆盖脚本判断，只需在 `/app/work` 下写一个
`agent_judgement.json`，再执行**同一条命令**：

```
第一次运行（脚本）  材料 → 证据索引 → result.json + agent_tasks.json
      ↓（可选）
写覆盖文件（你）    参考 agent_tasks.json 的证据锚点 → agent_judgement.json
      ↓
再运行一次（脚本）  逐条校验引文 → 合并进 result.json，覆盖条目优先
```

覆盖条目要过引文校验闸门；未通过的条目被退回并记录原因，脚本保留自己的
确定性判断作为兜底。**跳过覆盖不会降级** —— 脚本自己的判断就是最终判断。

## 第一次运行的产物：agent_tasks.json

| 字段 | 用途 |
|---|---|
| `contract_keys` | 本案例的真实契约键（运行时读 `/input/case.json` 得到） |
| `contract_semantics` | 每个键被判为什么语义（claims / deliverable / uncertainty …） |
| `evidence_anchors` | 全部段落 id + 原文短引 —— **你的证据只能从这里取** |
| `draft_claims` | 脚本机械抽到的候选论断（不保证是结论） |
| `open_questions` | 任务卡里解析出的问句 |
| `what_to_judge` | 需要你判断的四件事 |
| `how_to_write_the_judgement` | 判断文件的结构与规则 |

## 覆盖文件：agent_judgement.json（或直接写 result.json）

两种写法任选其一：

- **(a)** 写 `/app/work/agent_judgement.json`（推荐）；
- **(b)** 直接把判断写进 `/app/output/result.json` —— 脚本会先把已有
  的 `result.json` 留底，合并时把你写的契约键取出来，**不会覆盖你的内容**。

判断文件的结构：

```json
{
  "case_id": "必须与 case.json 一致",
  "claims": [
    {"text": "你判断为结论的陈述",
     "kind": "reported 或 inferred",
     "evidence": [{"paragraph_id": "P002", "quote": "原文逐字片段"}]}
  ],
  "answers": [
    {"question": "被问的问题", "answer": "你的回答",
     "evidence": [{"paragraph_id": "P002", "quote": "原文逐字片段"}]}
  ],
  "deliverable": "最终交付物正文",
  "contract_fields": {"契约键名": "该键内容"},
  "uncertainties": [
    {"kind": "category", "item": "对象", "detail": "客观描述",
     "resolution": "你做了什么 / 为何无法继续"}
  ]
}
```

## 合并前的校验闸门

脚本在合并前逐条检查，**不通过就退回**：

| 检查 | 不通过时的动作 | 记录到 |
|---|---|---|
| `quote` 能在材料原文里逐字找回 | 退回该条目 | `uncertainty`（`judgement_quote_not_found`） |
| claim / answer 至少有一个 evidence | 退回该条目 | `uncertainty`（`judgement_evidence_missing`） |
| `contract_fields` 的值不是 `null` / 空容器 | 退回该字段 | `uncertainty`（`judgement_field_empty`） |
| 判断文件本身可解析且是对象 | 保留纯确定性结果 | `uncertainty`（`judgement_unreadable`） |

退回**不是静默丢弃**：每条退回都带原因写进 `uncertainty`，
`change_log` 里也有 `reject_unverifiable_judgement` 记录，
`run-review.md` 里能看到"接受了几条、退回了几条"。

> 判断可以自由，**证据不可以自由**。

## 合并后的状态

- 有合规判断被合并 → 覆盖条目优先，`status_reason` 说明"判断已合并且证据已逐字校验"；
- 有条目被退回 → `status` 保持 `answered`（脚本自己的判断仍在），
  但退回条目在 `uncertainty` 里可见；
- 没有判断文件 → 脚本自己的确定性判断就是最终判断：交付物成形即
  `status=answered`，剩余边界写进 `uncertainty`；只有交付物未成形
  （材料不足/渲染失败）才停在 `partial` 并写明具体原因。

## 自查

```bash
python - <<'PY'
import json
r = json.load(open("/app/output/result.json", encoding="utf-8"))
q = r.get("quality", {})
print("status:", r["status"])
print("agent claims :", q.get("agent_judged_claims", 0))
print("agent answers:", q.get("agent_judged_answers", 0))
print("rejected     :", q.get("judgement_entries_rejected", 0))
print("unmapped     :", r["contract"]["unmapped_fields"])
PY
```
