# 左联知识库 - 收尾计划

> 最后更新：2026-06-07  
> 项目根目录：`D:\1大创\左联知识库项目`  
> 主分支：`main`

## 总体目标

把已完成的六阶段工程与研究基础，收敛为可复核、可演示、可答辩的大创成果。

## 当前状态

| 阶段 | 状态 | 核心产物 |
| --- | --- | --- |
| Phase 1 组织身份重建 | 已完成 | 证据台账、分层组织身份、回归测试 |
| Phase 2 事实级证据层 | 已完成基础设施，持续补证 | 事实证据表、覆盖率报告、待核队列 |
| Phase 3 研究层与发布层分离 | 已完成 | 发布生成器、发布清单、门禁测试 |
| Phase 4 事件与地点质量治理 | 已完成基础治理，待人工复核 | 质量报告、165 条审核队列 |
| Phase 5 人物关系抽样审计 | 已完成抽样，待人工判定 | 400 条审核模板、风险报告 |
| Phase 6 研究分析与答辩交付 | 已完成网络分析，待形成答辩成果 | 网络分析报告、质量与局限报告 |

## 收尾任务

### P0 - 研究结论闭环

- [ ] 人工审核 `research/drafts/reports/phase5_relation_review_template.csv` 的 400 条关系。
- [ ] 根据人工审核结果生成总体与分层准确率、错误率和修订规则。
- [ ] 从网络分析中形成至少 3 项研究发现，每项区分数据观察、历史解释和证据局限。

> 2026-07-30 状态：以上三项均已生成 **AI 辅助草稿**（见 `research/drafts/reports/phase6_research_findings.md`、`phase5_relation_review_ai_filled.csv`、`phase5_accuracy_report.md`），尚待人工复核签字后转正，故未勾选。

### P1 - 证据与质量补强

- [ ] 按 `phase2_core_fact_review_queue.csv` 补充核心事件、人物生卒年和角色证据。
- [ ] 处理 `phase4_review_queue.csv` 中优先级最高的事件与地点记录。
- [ ] 对 critical/high 风险关系优先执行修订或降级。

> 2026-07-30 状态：P1.1 证据增补与 P1.2 队列优先级已生成草稿（`phase1_p1_evidence_supplement.csv`、`phase4_priority_recs.csv`、`phase1_p1_proposed_sources.csv`），均为 `pending` 待合并；P1.3 的 17 条 critical 无佐证降级建议见 `phase5_critical_downgrade_recs.csv`，待人工确认后落地。均未勾选。

### P2 - 答辩交付

- [ ] 制作答辩 PPT，覆盖问题、方法、发现、价值与局限。
- [ ] 编写并演练 5 分钟演示脚本。
- [ ] 准备离线可运行版本与现场检查清单。

## 验收命令

```powershell
python -m pytest -v
python -c "from kb_schema import validate_data_dir; r=validate_data_dir('data/processed'); print(len(r.errors), len(r.warnings))"
python research/analysis/build_publish_data.py
python build_static_site.py
```

新增或修改的 Python 文件还应通过定向 `ruff` 检查。全仓历史 lint 债务不在收尾任务中无边界清理。

## 关联文档

- 项目说明：`README.md`
- 当前进度：`progress.md`
- 数据观察：`findings.md`
- 完整实施方案：`docs/superpowers/plans/2026-06-05-full-project-implementation-roadmap.md`
- 双层架构设计：`docs/superpowers/specs/2026-06-05-research-and-presentation-dual-layer-design.md`
- 阶段报告：`research/drafts/reports/`
