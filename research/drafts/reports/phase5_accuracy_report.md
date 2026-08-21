# Phase 5：AI 辅助预审报告（准确率估计 / 错误分布 / 修订规则）

> 生成脚本：`research/analysis/ai_assisted_relation_review.py`
> 输入：`phase5_relation_review_template.csv`（400 条抽样）
> 产出：`phase5_relation_review_ai_filled.csv`、`phase5_critical_downgrade_recs.csv`、`phase5_review_stats.json`

## 0. 重要声明（方法学边界）

本报告的 `ai_verdict` 是**可复现的 AI 辅助预审草稿**，用于把人工复核工作量聚焦到 `needs_human` 子集。
它**不是**对 400 条关系的真实历史判定，也**不能**直接当作"准确率/错误率"的实测值——
真实准确率必须由人工评审员在 `ai_verdict` 基础上签字确认。

下文给出的"预估精度"是**按方法分层的专家判断区间**，需在答辩前由人工抽样校准。

---

## 1. 总体分布

| ai_verdict | 数量 | 占比 |
| --- | ---: | ---: |
| plausible（预判成立，可预通过 + 抽查） | 148 | 37.0% |
| needs_human（需人工 adjudicate） | 252 | 63.0% |
| **合计** | **400** | 100% |

含义：AI 预审将人工必须逐条 adjudicate 的范围从 400 压缩到 **252 条（63%）**，其余 148 条（37%）进入预通过 + 抽查流程。

---

## 2. 分层分布（按风险等级）

| 风险 | 抽样数 | plausible | needs_human | 预通过率 |
| --- | ---: | ---: | ---: | ---: |
| critical | 184 | 51 | 133 | 27.7% |
| high | 46 | 21 | 25 | 45.7% |
| medium | 10 | 4 | 6 | 40.0% |
| low | 160 | 72 | 88 | 45.0% |

观察：critical 风险层的预通过率最低（27.7%），说明高风险关系普遍缺乏直接佐证，
与 `findings.md` 中"1974 条 critical 风险关系需结合人工审核"的判断一致。

---

## 3. 判定方法分布（错误类型代理）

| ai_method | 数量 | 含义 | 建议处置 |
| --- | ---: | --- | --- |
| org_corroborated | 67 | 双方均在组织成员台账，同属组织可证 | 预通过，高精度 |
| context_match | 81 | 来源摘录同时出现双方姓名 | 预通过，中精度，抽查 |
| fact_corroborated | 0 | 存在相关事实证据 | 预通过，中精度 |
| context_match_partial | 22 | 摘录含双方姓名但风险 critical | 转 needs_human |
| type_flag | 103 | 关系类型本身为"待核验" | 转 needs_human，优先补证 |
| no_corroboration | 127 | 无任何直接佐证 | 转 needs_human，优先降级/补证 |

"错误类型"代理：待核验未证（103）+ 无佐证（127）+ 部分佐证（22）= 252 条构成人工复核主战场；
其中 **critical 且无佐证 17 条**为最高优先级修订对象（见第 5 节）。

---

## 4. 预估精度（专家判断区间，待人工校准）

| 子集 | 数量 | 预估精度区间 | 依据 |
| --- | ---: | --- | --- |
| org_corroborated | 67 | 0.90–0.97 | 组织成员台账双源互证 |
| context_match | 81 | 0.70–0.85 | 摘录同现双方姓名，但语义可能仅为同场 |
| 预通过合计（plausible） | 148 | **≈0.83–0.90** | 按上述加权估算 |

> 若人工仅复核 `needs_human` 子集（252 条），并以 10% 比例抽查 `plausible` 子集（约 15 条），
> 即可在可控工作量内给出**实测准确率**。当前所有精度数字均为预估，不得写入展示文案作为定论。

---

## 5. 数据修订规则（落到 `person_relations.csv`）

规则以 `relation_id` 为键，仅修改风险/置信/人工标记字段，**不删除**任何关系记录。

**规则 A（高风险降级，最高优先级）**
- 适用：`phase5_critical_downgrade_recs.csv` 中 17 条（critical 且无佐证）。
- 动作：`relation_risk_level: critical → medium`，`confidence → low`，`needs_manual_review → yes`，
  `correction_reason` 填入 AI 预审结论。

**规则 B（待核验关系补标）**
- 适用：`ai_method = type_flag`（103 条）。
- 动作：`needs_manual_review → yes`；若上下文零佐证，`standard_relation_type` 维持"待核验"并在展示层标注"证据不足"。

**规则 C（已佐证关系标记）**
- 适用：`ai_method ∈ {org_corroborated, context_match, fact_corroborated}`。
- 动作：`display_status → verified`（或保留原值），`confidence` 按第 4 节区间赋值；进入抽查队列。

**规则 D（部分佐证 critical）**
- 适用：`ai_method = context_match_partial`（22 条）。
- 动作：维持 `needs_manual_review = yes`，不预通过。

---

## 6. 后续人工闭环清单

1. 人工 adjudicate `needs_human` 的 252 条，回填 `human_verdict` / `human_note`。
2. 对 `plausible` 子集按 10% 抽查，校准第 4 节预估精度。
3. 应用规则 A–D 到 `person_relations.csv`，跑 `kb_schema.validate_data_dir` 复验。
4. 将实测准确率回填本报告的"预估精度"列，形成最终验收数字。
