# Progress Log - 左联知识库

## 2026-08-21 - 合并管线幂等化与口径统一（任务书执行）

**开工回执**：目标=两个合并脚本幂等化+白名单文档口径统一（608/1165/14/150/9.3%）；顺序=任务0基线核验→任务1幂等与隔离测试（含红绿反向验证）→任务2文档同步→验收提交；最大风险=修改合并逻辑时误伤已转正生产数据，以"早退不写盘"为第一保障。基线实测与任务书一致：264db29、43 passed、0 err/13 warn、2个UP009、608/1165/150、origin 12+2。

**任务1完成**：两脚本以 origin_evidence_id 识别已转正证据，全部转正时早退不写盘；来源复用映射=已转正证据回推+生产表URL查重兜底；注记追加均带标记检查；移除2处UP009编码头。新增 `tests/test_merge_scripts_idempotency.py` 共5项测试（已合并副本二跑零变化×2、剥离痕迹后首跑增量+二跑零变化×2、生产origin编号唯一性），只读复制生产CSV、全程隔离临时目录。反向验证：临时破坏 batch2 的 pending 过滤→测试红（二跑重复追加610→612被拦截）→还原→48 passed 全绿。过程中修复两处自引入缺陷（证据行误取source_url列、复用映射误传sources行）。

**任务2完成**：README/findings 现行口径统一为 608条事实证据/1165来源/事件覆盖14/150=9.3%；`phase2_evidence_coverage_report.md` 顶部加注"历史快照"及新旧口径对照。progress.md 中 2026-06-05 条目内的 594/1153/2.7% 属带日期阶段语境的历史快照，按任务书保留。验收命令中的 `rg` 本机不存在且任务书禁止安装依赖，改用 Select-String 等价执行并逐条核对剩余旧数字归属。

**最终验收**：pytest 48 passed（>43，0 skipped）；Ruff 三文件 All checks passed；schema 0 errors / 13 warnings（未超基线）；生产数据保持 608/1165/150，origin 编号 12+2 各自唯一；工作树仅含白名单文件改动。

**补充修复：来源挂接**：修复 merge_longhua_roster.py 从旧基线合并时交叉核对来源不挂接事件的缺陷（第二批合并时曾以手工补挂掩盖）。根因有二：事件挂接只收集证据引用的来源，遗漏无证据直接引用的 SRC-1165；且注册循环原地改写 row["source_id"] 致批次映射查空。改为留存试点ID清单、按批次全量挂接。新增第6项测试"旧基线合并后 schema 0错误、≤13警告、SRC-1165已引用"，TDD 先红后绿验证。最终 pytest 49 passed，Ruff 全绿，schema 0 err/13 warn，生产 EVT-00148 挂接含 SRC-1164;SRC-1165。

## 2026-08-21 - 版权合规、仓库基线重建与事件证据合并转正

**版权合规：**

- `research/raw_texts/` 三部受版权保护史料全文（含 Z-Library 来源的《鲁迅日记》整理版）停止 git 跟踪并从全部历史中清除（git filter-repo 重写 35 个提交，强推 main）。
- 基线 tag：`copyright-clean-baseline`；本地文件保留供研究管线使用。
- 待办：联系 GitHub Support 清除旧提交缓存视图。

**仓库基线重建：**

- 208 个修改文件与全部未跟踪核心资产分 11 批提交入库（工程设施、研究脚本、生产数据、发布层、报告、文档）。
- 解除跟踪 CI 自建的 `docs/` 生成产物；删除已合并分支 clean-version、codex/zuolian-kb-release-20260323。
- 修复本地环境缺依赖导致的 3 个测试失败；ruff 自动修复 434 处并对历史研究脚本定向豁免。

**事件证据补证试点合并转正（Phase 2）：**

- 10 个候选来源网页逐条核验通过后合并为 SRC-1154..SRC-1163（失效的中国军网来源由纪录小康工程·广东数据库替代）。
- 12 条候选证据以 FE-EVI-* 正式 ID 合并进 `fact_evidences.csv`（review_status=reviewed，origin_evidence_id 保留原编号）。
- 决策落地：EVT-00029 日期由 1922-02 降为 1922（精度年）；EVT-00143 保留"被捕"并注记来源原词"秘密绑架"；EVT-00260 置信度 low→medium。
- P1 草稿 FE-SUP-0137（EVT-00008 无定位候选）标记废弃，避免双计数。
- 事件事实级证据覆盖率：4/150（2.7%）→ **14/150（9.3%）**。
- 验收：schema 0 errors / 13 warnings；pytest 43 passed；publish 与静态站重新生成。

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
