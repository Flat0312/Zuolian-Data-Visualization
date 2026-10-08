# 左联知识库 - 收尾计划

> 最后更新：2026-10-08
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
| Phase 4 事件与地点质量治理 | 已完成基础治理，待人工复核 | 当前时空审计及待核队列；83个事件泛化挂接上海 |
| Phase 5 人物关系抽样审计 | 已完成抽样、回源核查、400 条裁决与生产层落地；重捕候选 28 条中 4 条已于 2026-10-08 逐条独立裁决落地（公开关系 0→5→7），剩余 24 条待处置 | 裁决表与实测准确率报告、两批落地台账与报告、`phase5_quote_recapture_queue.csv`、`phase5_quote_recapture_adjudicated.csv`、`quote_attestation.py` 双方佐证门 |
| Phase 6 研究分析与答辩交付 | 已校正数据观察与解释边界，已更新答辩文字稿 | 首篇专题初稿；旧PPTX待重制，未演练 |
| Phase 7 关系证据候选 | 12条冻结样本，8条支持性候选、4条不足，全部待人工复核；主Agent独立审计完成（8赞成/1维持/2赞成升级/1引文问题） | 候选表、原文定位与引文校正、`phase7_candidate_independent_audit.csv` |

## 收尾任务

### P0 - 研究结论闭环

- [x] 400 条关系裁决完成（2026-09-20 授权按建议执行，路径 C；成立率 40.2%／类型准确率 33.8%）。规范审核包 `phase5_relation_review_package.csv` 保持空裁决，如需逐条独立复核可另起并覆盖。
- [x] 总体与分层准确率、错误率与 19 条修订规则已生成（`phase5_review_accuracy_report.md`）。
- [x] 裁决落地生产层（2026-09-28）：公开关系 0→5，发布层/静态站/四口径全部重建。
- [x] 裁决 `phase5_quote_recapture_candidates.csv` 的 4 条重捕候选（2026-10-08 用户逐条独立裁决：
      REL-00622 周扬—邵荃麟直接过、REL-01368 郁达夫—陈望道直接过并改类型 签名联署、
      REL-01161 阳翰笙—林淡秋成立但不进公开层、REL-01891 叶紫—萧军判「证据不足」零改动；
      已由幂等脚本 `apply_phase5_recapture_landing.py` 落地，公开关系 5→7，
      裁决另出 `phase5_quote_recapture_adjudicated.csv`，候选包与队列文件原样未动）。
- [ ] 其余 24 条重捕候选按 `recapture_status` 分类处置（16 条罗列级、5 条重捕被拒、3 条无同窗共现）。
- [x] 界定引文缺陷范围：`research/drafts/reports/evidence_verbatim_audit_2026-09-28.md`——夜间轮 182 条引文逐字全真、选段错误约 56%；事件/人物事实证据层 442 条可核 0 条未命中，覆盖率口径不受影响。
- [x] 校正研究观察与历史解释的边界，撤下证据不足的强结论。
- [x] 完成首篇专题初稿、10个人物片段与5个地点或机构内容卡。
- [ ] 逐项审核专题及关系候选，补充可支持历史解释的史料（AI 侧独立审计已完成，人工裁决待做）。

> 2026-09-20 状态：P0 三项的 AI 侧全部完成——①审核包 `phase5_relation_review_package.csv`（夜间轮 sample400 400/400 并入，建议 correct 135 / wrong_type 26 / not_supported 239，`human_verdict` 留空待签）＋签核单（三条签核路径）；②`analyze_relation_review.py` 就绪，当前 0 裁决如实输出「尚无人工裁决」，无预估数字；③Phase7 12 候选独立审计（REL-00523 二轮 support 升级因引文摘错段落判 issue_found，换引文前不得转正；REL-01219/01743 升级证据核实赞成）＋T2–T5 专题独立审核（T5 过，T2/T3/T4 各有引文静默校改待整改，整改前维持 limited_report）。待人工事项登记 BLOCKED.md 首部。

> 2026-09-05 状态：网络分析已独立复算，可信关系为0，不具备稳定排名样本。当前研究观察见 `research/drafts/reports/phase6_research_findings.md`；专题见 `research/drafts/topics/1928-correspondence/专题初稿.md`。400条人工审计仍未完成；旧AI预审和预估精度不作为实测准确率。

### P1 - 证据与质量补强

- [ ] 按 `phase2_core_fact_review_queue.csv` 补充核心事件、人物生卒年和角色证据。
- [ ] 处理 `phase4_review_queue.csv` 中优先级最高的事件与地点记录。
- [ ] 对 critical/high 风险关系优先执行修订或降级。

> 2026-07-30 状态：P1.1 证据增补与 P1.2 队列优先级已生成草稿（`phase1_p1_evidence_supplement.csv`、`phase4_priority_recs.csv`、`phase1_p1_proposed_sources.csv`），均为 `pending` 待合并；P1.3 的 17 条 critical 无佐证降级建议见 `phase5_critical_downgrade_recs.csv`，待人工确认后落地。均未勾选。

### P2 - 答辩交付

- [ ] 制作答辩 PPT，覆盖问题、方法、发现、价值与局限。
- [x] 更新九页答辩文字大纲、五分钟口述稿与内容检查清单。
- [ ] 按新版文字重制PPT并实际演练五分钟脚本。
- [ ] 准备离线可运行版本，并完成现场检查清单各项检查。

新版答辩文字保存在 `research/drafts/defense/`；`docs/superpowers/specs/` 中留下的旧入口存根已指向新版。`docs/` 是被Git忽略的站点输出目录，不作为任何文字或设计文档的唯一存放处——设计与方案文档已于 2026-09-28 迁至受版本管理的 `research/design/`。

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
- 完整实施方案：`research/design/plans/2026-06-05-full-project-implementation-roadmap.md`
- 双层架构设计：`research/design/specs/2026-06-05-research-and-presentation-dual-layer-design.md`
- 关系发布状态五态与门禁规则：`research/analysis/relation_publish_status.py`（模块 docstring 即规范）
- 引文双方佐证门：`research/analysis/quote_attestation.py`
- Phase 5 落地口径与未做事项：`research/drafts/reports/phase5_relation_landing_report.md`
- 阶段报告：`research/drafts/reports/`
