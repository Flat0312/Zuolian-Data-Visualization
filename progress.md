# Progress Log - 左联知识库

## 2026-06-05 - 六阶段研究实施完成

**项目定位：** 面向大学生创新创业训练计划，以可追溯历史证据研究左联人物、组织、事件与社会关系网络，并通过数字人文知识库完成研究验证、成果展示和答辩演示。

**已完成：**

- Phase 1：组织身份改为证据台账驱动判定，形成 45 名正式成员、77 名候选人物、28 名相关人物。
- Phase 2：建立 `fact_evidences.csv`，形成 594 条事实证据与 631 条核心事实待核队列。
- Phase 3：建立 `data/publish/` 发布层，公开数据自动过滤候选与争议组织身份。
- Phase 4：完成事件与地点质量治理，形成 165 条人工审核队列。
- Phase 5：对 4238 条人物关系完成 400 条分层抽样，生成关系审核模板与风险报告。
- Phase 6：完成网络分析与数据质量局限报告。

**当前基线：**

- 人物 162、关系 4238、事件 150、地点 41、组织 36、来源 1153。
- 组织身份事实级证据覆盖率 100%；事件存在覆盖率 2.7%；人物生卒年与角色事实仍待补证。
- Schema 严重错误为 0；发布层可由脚本重复生成。

**尚未闭环：**

1. 完成 Phase 5 的 400 条关系人工判定，并据此计算准确率和错误类型分布。
2. 补充事件、人物生卒年和角色的可定位事实证据。
3. 将网络分析结果整理为至少 3 项可复核研究发现。
4. 完成答辩 PPT、5 分钟演示脚本和现场检查清单。

详细状态与验收命令见 `task_plan.md`；质量指标见 `research/drafts/reports/`。

## 2026-07-30 - P0/P1 AI 辅助草稿收尾

> 状态：**AI 辅助草稿已生成，待人工复核签字后转正**（未改写生产数据）。

- **P0.3 研究发现**：`research/drafts/reports/phase6_research_findings.md`（4 项，含数据观察/历史解释/证据局限）。
- **P0.1 关系审查**：`research/analysis/ai_assisted_relation_review.py` → `phase5_relation_review_ai_filled.csv`（400 条，plausible 148 / needs_human 252）。
- **P0.2 准确率报告**：`phase5_accuracy_report.md`（总体+分层分布、预估精度区间、修订规则 A–D）。
- **P1.1 证据增补**：`build_p1_evidence_supplement.py` → `phase1_p1_evidence_supplement.csv`（137 行：45 名正式成员生卒年/角色 + 标志事件 EVT-00001/EVT-00008）+ `phase1_p1_proposed_sources.csv`（SRC-SUP-01 权威辞典）。
- **P1.2 队列优先级**：`build_p1_queue_priorities.py` → `phase4_priority_recs.csv`（165 条，高14/中56/低95）。
- **P1.3 高风险关系**：`phase5_critical_downgrade_recs.csv`（17 条 critical 无佐证→降级 medium）；汇总 `phase4_p1_review_summary.md`。

注意：`fact_evidences.csv` 人物生卒年/角色覆盖仍为 0%（增补行为 `pending` 态，尚未合并进生产）；`person_relations.csv` 仍为 4238 行未改。生产数据经 `kb_schema.validate_data_dir` 复验：0 errors / 13 warnings（与基线一致）。

**仓库状态提醒（交接风险）**：git 工作树严重超前——208 文件已修改、且 AGENTS.md/findings.md/progress.md/task_plan.md/kb_schema.py/tests/ 等大量核心文件处于未跟踪状态；另有 `pytest-cache-files-*` 散落缓存、`data/backup_*` 备份目录、`_fix_all.py` 一次性修复脚本与多分支（main / clean-version / codex/zuolian-kb-release-20260323）。提交基线混乱，需用户决策如何整理。
