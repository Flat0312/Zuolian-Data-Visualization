# P1.2 / P1.3 收尾队列处理总结

> 生成脚本：`research/analysis/build_p1_queue_priorities.py`、`research/analysis/ai_assisted_relation_review.py`
> 关联输入：`phase4_review_queue.csv`、`phase5_relation_review_ai_filled.csv`

---

## P1.2 事件与地点审核队列（phase4_review_queue.csv，165 条）

按 `confidence` + `date_precision` + 复核原因打分，优先级分布：

| 优先级 | 数量 | 处置策略 |
| --- | ---: | --- |
| 高 | 14 | 低置信 + 日期仅到日级 → 优先补权威来源并核实具体日期 |
| 中 | 56 | 单一风险信号 → 本轮内复核 |
| 低 | 95 | 维持待核，定期滚动复核 |

产出：`phase4_priority_recs.csv`（高优先级在前，含 recommended_action）。
高优先级样本：鲁迅与内山书店/内山完造相关事件（EVT-00016/00017/00019/00020）、鲁迅与柔石会面（EVT-00006）等——均为 low 置信且日期仅到日级，需补权威来源（如《鲁迅日记》具体卷页）并提升日期精度。

---

## P1.3 critical / high 风险关系修订与降级

来自 P0.1 的 AI 辅助审查（400 条抽样）：

| 风险 | 抽样数 | plausible（预通过） | needs_human（需人工签名） |
| --- | ---: | ---: | ---: |
| critical | 184 | 51 | 133 |
| high | 46 | 21 | 25 |

### 已生成的修订建议

**A. critical 且无佐证 → 降级（最高优先级）**
- 文件：`phase5_critical_downgrade_recs.csv`（17 条）
- 动作：`relation_risk_level: critical → medium`，`confidence → low`，`needs_manual_review → yes`。
- 这些关系在现有数据中既无组织台账佐证、也无来源摘录双名、也无事实证据，置为 critical 会高估可信度，故降级。

**B. high 风险 needs_human（25 条）**
- 策略：不自动降级，但**必须在展示前由人工签名**；其中 `ai_method = type_flag`（关系类型本身为"待核验"）者，展示层须标注"证据不足"。
- 完整清单见 `phase5_relation_review_ai_filled.csv`（筛选 `relation_risk_level=high` 且 `ai_verdict=needs_human`）。

### 落地方式（建议，需人工确认后执行）

修订以 `relation_id` 为键，仅改 `relation_risk_level` / `confidence` / `needs_manual_review` / `correction_reason` **四字段，不删除任何关系**。提供两种执行路径：

1. **人工路径**：审阅 `phase5_critical_downgrade_recs.csv` → 在 `person_relations.csv` 逐条修改。
2. **脚本路径**（待编写）：读取降级建议 CSV，批量 patch `person_relations.csv`，随后运行
   `python -c "from kb_schema import validate_data_dir; validate_data_dir('data/processed')"` 复验。

> 依据 AGENTS.md「修改保持最小范围」，本阶段**仅产出建议清单，未自动改写生产数据**；落地需人工复核签字。

---

## 与 P0 的衔接

- P0.1 审查产出的 `needs_human` 子集（252 条）即 P1.3 的人工复核对象。
- P1.1 补证（生卒年/角色/标志事件）可反向提升 P0.1 的 `fact_corroborated` 命中率，未来重跑审查时 `plausible` 预通过率应上升。
- 三条收尾线（P0 结论闭环 / P1 证据补强 / P1 风险降级）共用 `relation_id` 与 `evidence_id` 两套主键，可端到端追溯。
