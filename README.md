# Runtime contract

Bohrium Playground Period 4 submission — challenge `ppt-c7b80f0d` (track: `academic-slides`).

This repository holds the complete skill package for the challenge: the skill
definition, its executable pipeline and all supporting documents. No precomputed
answers, private data or credentials are included.

## Skill

- Name: `bohr-slides`
- Summary: 当需要汇报科研材料并产出可复核的结构化结果时使用。 本题范围：论文转学术 PPT。 题目家族：论文转学术 PPT。 产出 result.json，其中包含运行时 output_contract 要求的全部字段，每条结论附带原文证据（段落定位 + 短引），无法从材料判断的部分如实写入 uncertainty 而非编造。

## Contents

| File |
|------|
| `LICENSE` |
| `README.md` |
| `SKILL.md` |
| `assets/contract-aliases.json` |
| `assets/policy.json` |
| `assets/result-schema.md` |
| `docs/DECK_WORKFLOW.md` |
| `docs/ERROR_HANDLING.md` |
| `environment/resources.yaml` |
| `environment/runtime.txt` |
| `examples/demo-case/agent_judgement.example.json` |
| `examples/demo-case/case.json` |
| `examples/demo-case/material.md` |
| `examples/demo-case/task.md` |
| `examples/quickstart.md` |
| `references/artifact-pipeline.md` |
| `references/contract-handling.md` |
| `references/evidence-discipline.md` |
| `references/failure-modes.md` |
| `references/judgement-override.md` |
| `rules/rules_deck.md` |
| `schema/deck_ir.schema.json` |
| `scripts/artifacts.py` |
| `scripts/deckkit.py` |
| `scripts/pptxgen.py` |
| `scripts/sk_base.py` |
| `scripts/sk_entry.py` |
| `scripts/sk_selftest.py` |
| `scripts/skill_profile.json` |
| `scripts/specialize.py` |
| `scripts/topic.py` |
| `README.md` |

## How to run

```bash
python scripts/sk_entry.py --input /input --work /app/work --out /app/output
```

A single deterministic pass reads `/input`, computes the deliverable and writes a
contract-complete `/app/output/result.json` plus the human-readable review files
described in `SKILL.md`. The same input always produces byte-identical output; no
network access is used, and fields the materials cannot support are reported
explicitly instead of being left empty.

## Notes

- Pure Python standard library; no network access required.
- Every claim is anchored to paragraph-level evidence; unknown items are reported
  in `uncertainty` instead of being fabricated.
