# Phase 5 · 400 条关系人工审签工作量预判表

> 生成脚本：`research/analysis/ai_assisted_relation_review.py` 加本表后处理

> 配套排序 CSV：`research/drafts/reports/phase5_human_review_queue.csv`（按 batch + 核心人物排序）

> 截至：2026-07-30


## 1. 总览

| 维度 | 数量 |
| --- | ---: |
| 总关系 | 400 |
| AI 预审 plausible | 148 |
| AI 预审 needs_human | 252 |
| 涉及 Top20 核心人物 | 185 条（46%）|
| 已在 critical 降级清单 | 17 条 |

## 2. 148 条 plausible —— 10% 抽查

AI 预审判定为可信的关系。不需主动审签，仅抽查。


| ai_method | 数量 | 预估准确率 |
| --- | ---: | --- |
| `org_corroborated` | 67 | 0.90-0.97 高 |
| `context_match` | 81 | 0.70-0.85 中 |
| 合计 | **148** | 加权 ≈ 0.83-0.90 |

**抽查方案**：从中随机抽 15 条（10%）逐条看，记录实测准确率。

- `org_corroborated` 抽 6-7 条（含 1 条 critical 风险）
- `context_match` 抽 8-9 条（覆盖 4 个风险等级）

## 3. 252 条 needs_human —— 按体量从上到下安排

| Batch | 数量 | 核心人物 | 处置策略 | 预计工时 |
| --- | ---: | ---: | --- | ---: |
| **B1**  |  88 |  36 | 批量：保留边，标 needs_manual_review | 30 分钟 |
| **B2**  | 103 |  38 | 按关系类型分桶决策 | 50 分钟 |
| **B3**  |  22 |  10 | 逐条看：上下文 5 秒判定 | 30 分钟 |
| **B4**  |  39 |  13 | 逐条 + 17 条可套用降级建议 | 150 分钟 |
| **总计** | **252** | | | **≈4.5 小时** |

> **上表中“核心人物”** 是不能批量的那部分。例如：B1 里有 36 条涵盖鲁迅、茅盾等，这些需逐条看，不能随 B1 一起批量。

## 4. B1 处置规则（low 风险 + 无佐证）


**默认批量规则**：
1. `human_verdict = "保留(evidence thin)"`
2. `relation_risk_level` 保持 low
3. `confidence = "low"`
4. `needs_manual_review = "yes"`
5. `correction_reason` 填"低风险无佐证，保留待核验"

**例外**：涵盖 Top20 核心人物的 36 条 → 不走批量，跳到 B1c 逐条看。

## 5. B2 处置规则（type_flag 关系类型）


**按关系类型分桶决策**：

| 关系类型 | 建议决策 |
| --- | --- |
| 交往 / 交游 / 同人 / 同事 / 同属组织 | `human_verdict = "保留(待核验)"` |
| 文学论战 / 笔战 | 逐条看（涉及立场）|
| 合作 / 翻译 / 编辑 / 签名联署 | 逐条看（事实可查）|
| 时空共现 / 空间共现 / 其他 | 逐条看 |

**例外**：涵盖 Top20 核心人物的 38 条 → 不走分桶，跳到 B2c 逐条看。

## 6. B3 逐条看（context_match_partial）


**22 条**：上下文同时出现双方姓名，但被 AI 标为 critical 风险。


**判定要点**：
1. 看 `context` 字段前 30 字，姓名是否真的属于同一段叙述
2. 是 = `human_verdict = "成立"`
3. 否 = `human_verdict = "误判"`, `correction_reason` 填"上下文误读"

**速度目标**：5 秒/条 → 22 条 ≈ 2 分钟

## 7. B4 重点：high/medium/critical 风险关系


**39 条**，内部还分几个优先级：


| 子批 | 数量 | 处置 |
| --- | ---: | --- |
| **B4c** 核心人物 critical（非降级目标） | 5 | 逐条，优先级最高 |
| **B4d** 已有降级建议（17 条原列表）| 17 | 直接套用降级：critical → medium |
| **B4** 其他 critical | 0 | 逐条 |
| **B4** 高/中风险 | 17 | 逐条 |

**处置脚本（仅限 B4d 17 条）**：
```python
# 可以跑这个脚本套用降级建议，但要先审阅 17 条原始 context
# 脚本路径：research/analysis/apply_phase5_downgrades.py（未存在，需新写）
import pandas as pd
dg = pd.read_csv("research/drafts/reports/phase5_critical_downgrade_recs.csv")
rel = pd.read_csv("data/processed/person_relations.csv")
# 仅处理 dg 里的 17 条
for _, row in dg.iterrows():
    rel.loc[rel["relation_id"] == row["relation_id"], "relation_risk_level"] = "medium"
    rel.loc[rel["relation_id"] == row["relation_id"], "confidence"] = "low"
    rel.loc[rel["relation_id"] == row["relation_id"], "needs_manual_review"] = "yes"
    rel.loc[rel["relation_id"] == row["relation_id"], "correction_reason"] = row["correction_reason"]
rel.to_csv("data/processed/person_relations.csv", index=False)
```

## 8. 工时排程建议


| 段 | 任务 | 时长 | 建议时段 |
| --- | --- | ---: | --- |
| 1 | B1 批量（低风险无佐证）| 30 分钟 | 工作日上午 |
| 2 | B1c 核心人物逐条 | 30 分钟 | 与段 1 同时 |
| 3 | B2 按关系类型分桶 | 50 分钟 | 工作日上午 |
| 4 | B2c 核心人物逐条 | 30 分钟 | 与段 3 同时 |
| 5 | 148 plausible 10% 抽查 | 30 分钟 | 段 3-4 同时 |
| 6 | B3 22 条逐条 | 30 分钟 | 段 5 同时 |
| 7 | B4d 17 条套用降级 | 30 分钟 | 单独作业 |
| 8 | B4c 核心人物 critical 逐条 | 60 分钟 | 不要一次做完 |
| 9 | B4 其他 critical/high/medium | 60 分钟 | 段 8 后 |
| **总计** | | **≈5 小时** | 分 3-4 次，最好朊晚拆分 |

## 9. 完成后回填 `phase5_accuracy_report.md`


审签完成后，把以下数字更新到 `phase5_accuracy_report.md`：


| 指标 | 当前预估 | 填入实测 |
| --- | --- | --- |
| `org_corroborated` 准确率 | 0.90-0.97 | ___ |
| `context_match` 准确率 | 0.70-0.85 | ___ |
| plausible 总体准确率 | 0.83-0.90 | ___ |
| needs_human 拒判正确率 | — | ___ |
| critical 风险最终保留率 | — | ___ |

## 10. 本次生成的 4 件数据


1. 本 Markdown 报告（`phase5_human_review_workload_plan.md`）
2. 排序后的人工审签 CSV（`phase5_human_review_queue.csv`）— 252 条按 batch + 核心人物优先级排序
3. 148 条 plausible 子集（CSV 中 `batch=S` 的行）— 抽查不主动审
4. 17 条 critical 降级建议（`phase5_critical_downgrade_recs.csv` 已存在，本次仅引用）
