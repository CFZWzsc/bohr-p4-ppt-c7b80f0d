# 快速上手

本 Skill 在 DeepSeek Harness 中被加载为 `/skills/weekly-skill`。
Harness 在每个案例上执行**一条**命令：

```bash
python /skills/weekly-skill/scripts/sk_entry.py \
  --input /input --work /app/work --out /app/output
```

## 一次运行就完成判断

脚本不是只做取证的半成品：同一条命令完成「读材料 → 建证据索引 → 按叙事角色
分页（页标题 = 该页主张）→ 生成可编辑 PPTX 与逐页 SVG/PNG → 写页级证据映射
与分离的讲稿 → 给出实质判断 claims / answers」。交付物成形即 `status=answered`；
只有材料不足或渲染失败才停在 `partial`，并在 `status_reason` 写明具体原因。

## 可选覆盖（跳过也不会降级）

想用自己的领域判断覆盖脚本判断：把 `/app/work/agent_judgement.json` 写好后再跑
**同一条命令**，脚本逐条校验引文后合并（未通过校验的条目退回并记原因）。
不做这一步，脚本自己的判断就是最终判断。

## 手动复现

```bash
cd /skills/weekly-skill
python scripts/sk_entry.py --input examples/demo-case --work /tmp/work --out /tmp/out
cat /tmp/out/result.json
```

## 产出

- `/app/output/result.json` —— 契约字段齐全的结果
- `/app/output/run-review.md` —— 证据索引 + 不确定性清单
- `/app/output/deck.pptx`、`/app/output/slides/slide-NN.svg|.png`、
  `speaker-notes.md`、`slide-evidence-map.json`、`slide-export-checks.json`
- `/app/work/agent_tasks.json` —— 证据锚点索引（写可选覆盖时参考）

## 自检

```bash
python scripts/sk_selftest.py --input examples/demo-case
```
