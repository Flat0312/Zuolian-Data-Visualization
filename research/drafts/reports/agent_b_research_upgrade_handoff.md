# Agent B 研究升级交接报告（codex/research-upgrade）

> 任务：核心内容补证候选、时空治理审计、可信网络分析。
> 边界：只新增 `research/analysis/build_core_upgrade_candidates.py`、`research/analysis/audit_event_place_normalization.py`、`research/analysis/build_trustworthy_network_analysis.py`、`research/drafts/reports/` 本任务新增文件与 `tests/test_research_upgrade_outputs.py`；不修改生产数据、发布层、Schema、前台与既有测试。

## 一、基线核验（2026-09-04 实测）

| 项目 | 任务书预期 | 实测 | 判定 |
| --- | --- | --- | --- |
| HEAD | 53ccf9c | 53ccf9c（开工时共享 HEAD 位于 Agent A 的 `codex/data-governance` 分支） | 相符 |
| 工作树 | 干净（隐含） | 不干净：Agent A 并行施工中，`data/processed/*.csv`、`kb_schema.py` 等持续变化 | **差异** |
| pytest | 74 passed | 开工时 **67 passed / 7 failed**（Agent A 进行中改动的中间态） | **差异** |
| 人物/关系/事件 | 162 / 4238 / 147 | 162 / 4238 / 147 | 相符 |
| 事件三口径 | 26/147、21/147、26/147 | 26/147、21/147、26/147 | 相符 |

### 差异说明（已按任务书记录，未改生产数据）

1. **共享工作树与 Agent A 并行**：本仓库为共享工作树。施工期间 Agent A 的未提交状态多次变化（`events.csv` 时间列与 `event_participants` 角色归一一度出现又回退、`person_relations.publish_status` 与 `relation_evidences.csv` 保持在场）。本 Agent 全部脚本按"列/表存在与否"双态兼容设计，对 Agent A 改动前后两种数据状态均可运行；测试对 `data/processed` 做只读快照拷贝后再断言，隔离并行改动。
2. **开工时 7 个既有测试失败**：全部源于 Agent A 当时的中间态（新 Schema 校验撞旧快照重放数据），属 Agent A 范围，本 Agent 未触碰。**收尾时全量复跑为 94 passed / 0 failed**（Agent A 已修复其失败并新增自有测试；本任务新增 11 项）。
3. **分支布局**：按任务书建议创建 `codex/research-upgrade`（与 53ccf9c 同源），提交仅含本任务白名单文件。

## 二、交付物清单（全部完成）

- [x] `research/analysis/build_core_upgrade_candidates.py`（只读、幂等、双态兼容）
- [x] `research/drafts/reports/core_person_evidence_candidates.csv`（30 人）
- [x] `research/drafts/reports/core_event_evidence_candidates.csv`（20 事件）
- [x] `research/drafts/reports/core_place_review_candidates.csv`（10 地点）
- [x] `research/drafts/reports/core_upgrade_selection_report.md`
- [x] `research/analysis/audit_event_place_normalization.py`
- [x] `research/drafts/reports/event_normalization_candidates.csv`（104 条）
- [x] `research/drafts/reports/place_normalization_candidates.csv`（39 条）
- [x] `research/drafts/reports/participant_role_mapping_candidates.csv`
- [x] `research/drafts/reports/temporal_spatial_audit_report.md`
- [x] `research/analysis/build_trustworthy_network_analysis.py`
- [x] `research/drafts/reports/trustworthy_network_analysis.json`
- [x] `research/drafts/reports/trustworthy_network_analysis_report.md`
- [x] `research/drafts/reports/updated_research_findings_candidates.md`
- [x] `tests/test_research_upgrade_outputs.py`（11 项守门测试）

## 三、结果摘要

### 1. 核心补证候选包

- 选择规则全部由数据 + 显式种子常量决定，可复现（公式见选择报告）；同分按 ID 升序决断；重跑字节一致。
- Top30 人物含鲁迅、丁玲、五烈士全部 5 人、冯雪峰、楼适夷、周扬等；泛化主题优先（命中任一优先主题者整体优先）。
- **90 个人物事实字段（30 人 × 生/卒/角色）全部为 0 条事实级证据**（fact_evidences 无 person birth/death/role 谓词）；本轮只读检索到 **20 个候选证据**（柔石/胡也频/冯铿/殷夫/李伟森/丁玲/楼适夷/冯雪峰/周扬，全部 `pending_human_review`，含 URL/访问日期/定位/逐字引文），其余 **70 个如实记 missing**。
- 20 个关键事件全部缺直接支持证据（补证队列按缺口优先选取）；检索到 3 个候选证据（萌芽月刊创刊 1930-01-01、楼适夷 1933 被捕、鲁迅 1930 避居内山书店）；洛阳书店大会等未获第二独立来源，如实 missing。
- 10 个地点候选：内山书店（含旧址）、龙华刑场、东方旅社 31 号房间等；3 个地址候选证据（内山书店旧址四川北路2050号、龙华烈士陵园龙华西路180号）。

### 2. 时空与枚举审计（只出建议，零改动）

- 事件候选 104 条：**泛化挂接"上海"83 个**、同名不同日期 8 条、重复 canonical_event_key 7 条、无 place_id 5 条、时间列缺失（表级）1 条。
- 地点候选 39 条：合并候选 8 组（内山书店/旧址、东方旅社系列、光华书局系列等，全部 pending）、同坐标异名 8 组、坐标质量 23 条。
- participant_role 现值：直接参与者 111 / **unclear 90** / 待核 17 / 关联人物 4——unclear 给出映射建议（→不明），非预设值人工归类。
- 注：以上为 2026-09-04 快照实测；Agent A 若落地其归一迁移，重复条数会相应减少，审计脚本按现状如实输出。

### 3. 可信网络分析

- 数据状态：关系 4238 → **可信 1760**（41.5%）；三套网络边数：全量 4227 / 可信 1758 / 加权 1758（同拓扑、异权重）。
- **核心发现（数据观察）**：可信口径下度中心性 Top10 与全量口径重合仅 **2/10（Jaccard 0.11）**；**鲁迅由全量第 1 降至可信第 5**，可信榜首为高尔基。现行 `phase6_research_findings.md` 发现一（鲁迅结构性中心）**对低质量关系高度敏感，须降级表述**。
- 《鲁迅日记》单一来源边 14 条，降权后 Top10 加权度重合 0.82；C/D 级最强证据边 0 条（现有可信关系至少有 B 级支撑），降权无影响（重合 1.0）。
- 时间切片（事件共参与口径）：1928–1930 边 41 / 1931–1933 边 80 / 1934–1936 边 27；切片互不重复（测试锁定）。关系表无逐条日期，切片不代表该期完整关系网（报告已声明局限）。
- 社区数：全量 44 → 可信 53；最大分量 122 → 114 人。
- `updated_research_findings_candidates.md` 给出 4 条候选发现（候选稿，未经人工复核不得替代现行 phase6 版本），逐条区分数据观察/历史解释/证据局限。

### 4. 验收与反向验证

- `python -m pytest -q`：**94 passed / 0 failed / 0 skipped**（含本任务新增 11 项；开工时 Agent A 中间态造成的 7 个既有失败已由 Agent A 修复）。
- `python -m ruff check`（3 个新脚本 + 新测试文件）：**All checks passed**。
- 三个脚本默认输出重跑：结果与快照一致；双跑字节级一致性由测试 `test_rerun_outputs_identical` 锁定（11 个输出文件逐一比对）。
- **反向验证**：临时禁用可信关系过滤四条排除规则 → `test_trusted_network_excludes_pending_relations` 失败（合成注入的待审核关系混入可信网络被拦截），红灯日志 `%TEMP%\agentb_red_validate.log`（1 failed）；恢复后 11 项全绿。
- `git diff --check`：通过。

## 四、需要人工审核的项目

1. **殷夫生年两说**：维基百科 1909-06-11 vs 生产值 1910（同济档案馆等同记 1910）——候选证据与生产值冲突，需裁决。
2. **周扬生年两说**：维基百科 1907-11-07 vs 生产值 1908（《周扬同志年谱》封面作 1908—1989）——需裁决。
3. **EVT-00183『1928年《萌芽月刊》文学活动记录』**：维基记载创刊 1930-01-01，与生产日期 1928 冲突，建议改期并具体化为创刊事件。
4. **EVT-00138『楼适夷因支持《文学》月刊被捕』**：维基记载 1933 年被捕（1937 出狱），与生产年份 1934 冲突。
5. **EVT-00078/00079『周扬 1932/1933 遭拘押』**：检索未见权威佐证（其被捕记录为留日时期），已按 missing 记录，需人工复查后再决定保留/改写/删除。
6. **内山书店沿革**：现址四川北路2050号为 1929 年迁入；1929 年及以前"内山书店"事件的地点归属需人工核对（涉及 EVT-00019 等）。
7. **地点合并 8 组候选**：全部 pending_human_review，需确认沿革后另行执行；"上海东方旅社、中山旅社秘密会议"是事件名还是地点名需先厘清语义。
8. **83 个泛化挂接"上海"的事件**：建议按证据细化到街区/建筑级（清单见 event_normalization_candidates.csv）。
9. **phase6 发现一表述**：建议按可信网络结果降级为"全量数据集内的中心地位"，替换与否由人工决定（候选稿见 updated_research_findings_candidates.md）。
10. **20 个候选证据转正**：全部 pending_human_review，URL/访问日期/定位/引文已登记，人工复核后方可立证。

## 五、给下一个执行 Agent 的接力说明

本任务白名单内工作已全部完成并本地提交（不推送）。若需继续，剩余工作均属**人工裁决落地**或 **Agent A 范围**：

- 上表第四节 10 项人工审核裁决后，由获授权的执行者把批准项落入生产层（走既有批次审计-裁决-落地流程，勿直接改 CSV）。
- Agent A 的生产数据/Schema/发布层工作以其自身任务书为准；本 Agent 的脚本对 Agent A 改动前后状态均兼容，无需返工。
- 复验命令：`python -m pytest -q tests/test_research_upgrade_outputs.py`（11 项）、三个脚本各自 `--data-dir/--out-dir` 可重跑。
