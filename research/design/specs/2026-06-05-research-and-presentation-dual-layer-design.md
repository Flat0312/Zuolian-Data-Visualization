# 左联知识库研究与展示双层架构设计

## 1. 目标与定位

本项目面向大学生创新创业训练计划，定位为：

> 以可追溯历史证据为基础，研究左联人物、组织、事件与社会关系网络，并通过数字人文知识库完成研究验证、成果展示和答辩演示。

项目同时服务两类成果：

1. **研究成果**：可复核的数据、证据链、质量评估和网络分析结论。
2. **答辩成果**：可操作的知识库、可解释的图表、清晰的方法与局限说明。

项目应满足：

- 重要结论可追溯到具体来源、定位和原文摘录。
- 研究层保留探索性数据，不因展示过滤而丢失。
- 展示层明确区分已确认、相关、候选和争议状态。
- 数据构建、判定和发布流程可重复执行。
- 数据质量指标可被自动验证，并能写入研究报告。

## 1.1 总体完成标准

结项时必须满足：

- 每个核心研究结论都能回溯到具体事实和来源。
- 事实、线索、推断和争议不会在展示层混为一谈。
- 数据生成与发布流程能够从脚本重复执行。
- 数据质量有量化报告，不用"数据很多"代替质量证明。
- 至少形成 3 个有证据支撑的研究发现。
- 答辩演示能够在 5 分钟内展示"问题、方法、发现、价值、局限"。

## 2. 总体架构

数据分为四层运行态：

1. **原始资料层**：原始表格、文本、OCR 和网页来源。
2. **研究证据层**：事实证据、审核状态、推断与争议结论。主真值。
3. **展示发布层**：由研究层脚本生成，公开模式使用此层。
4. **前端运行层**：公开模式默认读取 `data/publish/`，研究模式读取 `data/processed/`。

所有层共享现有实体 ID。研究层是结论和证据的主真值；展示层不得单独维护另一套历史结论。

## 3. 证据模型

### 3.1 事实证据总表

`data/processed/fact_evidences.csv`，用于记录字段级、关系级和事件级事实证据。核心字段为：

- `evidence_id`、`subject_type`、`subject_id`、`predicate`、`object_value`
- `source_id`、`locator`、`quote`
- `evidence_support`、`source_level`、`review_status`、`reviewer_note`

### 3.2 组织成员证据台账

`data/processed/org_membership_evidences.csv`，作为 `fact_evidences.csv` 的领域先行版本。它记录人物与组织关系的具体证据、来源等级、支持类型和审核状态。

`org_memberships.csv` 是由证据台账和判定规则生成的结论表，不再依据 `persons.role` 自动确认正式成员。

## 4. 来源等级和组织身份判定

### 4.1 来源等级

- **A**：档案、一手史料、日记、书信、正式组织名单。
- **B**：权威研究论著、政府、纪念馆、党史机构资料。
- **C**：普通研究文章和媒体专题，仅作辅助佐证。
- **D**：百科、无定位表格和来源不明记录，仅作线索。

### 4.2 正式成员判定

满足以下任一条件可判定为 `confirmed_member`：

1. 一条 A 级来源明确支持正式成员身份。
2. 两条相互独立的 B 级来源明确支持正式成员身份。

其他状态：

- `related_person`：有明确组织关联，但不足以证明正式成员身份。
- `candidate`：仅有待核线索或证据不足。
- `disputed`：支持与反对证据冲突。

百科不得单独确认正式成员身份。

## 5. 研究层与展示层

研究层保留全部状态及其证据。展示层分层展示全部组织关系：

- 正式成员：已确认标签。
- 相关人士：关联标签。
- 候选关系：待核标签。
- 争议关系：争议标签。

任何候选或争议关系都不得在页面文案中被陈述为确定史实。

## 6. 六阶段实施路线与执行状态

### Phase 1：重建 ORG-001 组织关系

**目标**：解决"所有组织身份从 `persons.role` 自动推断"的根本缺陷。

**交付物**：

- `data/processed/org_membership_evidences.csv`（581 条证据记录）
- `data/processed/org_memberships.csv`（150 条结论，含 45/77/28 分层）
- `research/analysis/rebuild_org_memberships.py`
- `tests/test_org_membership_rebuild.py`（13 项测试）
- 组织身份 Schema 校验和回归测试

**当前指标**：

| 身份状态 | 数量 |
| --- | ---: |
| `confirmed_member` | 45 |
| `candidate` | 77 |
| `related_person` | 28 |

**状态：已完成**

**阶段检查点**：

```
python -m pytest tests/test_org_membership_rebuild.py -v
python -c "from kb_schema import validate_data_dir; r=validate_data_dir('data/processed'); print(len(r.errors), len(r.warnings))"
python build_static_site.py --data-dir data/processed --output-dir docs
```

---

### Phase 2：建立事实级证据链

**目标**：把"实体引用过哪些来源"升级为"某条具体事实由什么来源支持"。

**交付物**：

- `data/processed/fact_evidences.csv`（594 条事实证据）
- `research/analysis/build_fact_evidences.py`
- `research/analysis/report_evidence_coverage.py`
- `research/drafts/reports/phase2_evidence_coverage_report.md`
- `research/drafts/reports/phase2_core_fact_review_queue.csv`（631 条待核事实）

**当前指标**：

| 事实类型 | 已覆盖 | 总数 | 覆盖率 |
| --- | ---: | ---: | ---: |
| 组织身份 | 150 | 150 | 100% |
| 事件存在 | 4 | 150 | 2.7% |
| 人物出生年 | 0 | 162 | 0% |
| 人物逝世年 | 0 | 162 | 0% |
| 人物角色 | 0 | 162 | 0% |

**状态：已完成**

**阶段检查点**：

```
python research/analysis/build_fact_evidences.py
python research/analysis/report_evidence_coverage.py
python -m pytest tests/test_fact_evidences.py tests/test_kb_schema.py -v
```

---

### Phase 3：建立研究层与展示发布层

**目标**：展示页面不直接读取研究主表，而由可审计规则生成发布数据。

**发布规则**：

- `confirmed_member`、`related_person`：保留到发布层。
- `candidate`、`disputed`：仅保留在研究层。
- 非 `formal` 关系状态：不发布。
- 发布层引用闭合，Schema 0 严重错误。

**交付物**：

- `research/analysis/build_publish_data.py`
- `data/publish/`（10 张表，由脚本生成）
- `data/publish/publish_manifest.json`
- `research/drafts/reports/phase3_publish_gate_report.md`
- `tests/test_publish_data.py`（2 项测试）
- 前端"公开模式 / 研究模式"切换

**当前指标**：

| 数据表 | 研究层 | 发布层 | 过滤 |
| --- | ---: | ---: | ---: |
| `org_memberships.csv` | 150 | 73 | 77 |
| `org_membership_evidences.csv` | 581 | 438 | 143 |
| `fact_evidences.csv` | 594 | 451 | 143 |

发布层 Schema 严重错误：0

**状态：已完成**

**阶段检查点**：

```
python research/analysis/build_publish_data.py --processed-dir data/processed --publish-dir data/publish --report research/drafts/reports/phase3_publish_gate_report.md
python -m pytest tests/test_publish_data.py tests/test_data_loader.py -v
python build_static_site.py --data-dir data/processed --output-dir docs
```

---

### Phase 4：事件与地点质量治理

**目标**：解决事件定义模糊、日期精度混乱和地点坐标看似精确但实际不精确的问题。

**计划交付**：

- `research/analysis/audit_events.py`
- `research/analysis/audit_places.py`
- `research/drafts/reports/phase4_event_place_quality_report.md`
- `research/drafts/reports/phase4_review_queue.csv`

**当前指标**：

| 指标 | 数值 |
| --- | ---: |
| 审核队列总量 | 165 条（142 events + 23 places） |
| places 精度字段覆盖 | 100% |
| Schema errors | 0 |

**状态：已完成**

**产出文件**：
- `research/analysis/audit_events_and_places.py`
- `research/drafts/reports/phase4_event_place_quality_report.md`
- `research/drafts/reports/phase4_review_queue.csv`
- `tests/test_audit_events_places.py`

---

### Phase 5：人物关系抽样审计

**目标**：用统计抽样回答"人物关系数据有多可信"，而不是依靠个别示例。

**计划交付**：

- `research/analysis/sample_relations.py`
- `research/drafts/reports/phase5_relation_review_template.csv`
- `research/analysis/report_relation_audit.py`
- `research/drafts/reports/phase5_relation_quality_report.md`

**当前指标**：

| 指标 | 数值 |
| --- | ---: |
| 总关系数 | 4238 |
| 抽样数 | 400（9.4%） |
| critical 风险 | 1974（46.6%） |
| low 风险 | 1719（40.6%） |
| high 风险 | 477（11.3%） |
| medium 风险 | 68（1.6%） |
| Top 关系类型 | 交游 1207, 待核验 1117, 同属组织 729 |

**状态：已完成**

**产出文件**：
- `research/analysis/sample_relations.py`
- `research/drafts/reports/phase5_relation_audit_report.md`
- `research/drafts/reports/phase5_relation_review_template.csv`

---

### Phase 6：研究成果与答辩交付

**目标**：从知识库中提取可答辩的研究发现和质量说明。

**研究问题**：

1. 左联人物网络中哪些人物承担跨群体连接作用？
2. 不同时间阶段的人物关系结构如何变化？
3. 正式成员、相关人物和候选人物在网络位置上有何差异？

**计划交付**：

- `research/analysis/build_network_analysis.py`
- `research/drafts/reports/phase6_network_findings.md`
- `research/drafts/reports/phase6_quality_limitations.md`
- `docs/superpowers/specs/2026-06-05-defense-slide-outline.md`
- `docs/superpowers/specs/2026-06-05-defense-script.md`

**当前指标**：

| 指标 | 数值 |
| --- | ---: |
| 节点数 | 162 |
| 边数 | 4227 |
| 连通分量 | 41 |
| 社区数 | 43 |
| 最大连通分量 | 57 人（75.3%） |
| 度中心性 Top3 | 鲁迅 0.72, 郑伯奇 0.68, 夏衍 0.66 |
| 中介中心性 Top3 | 鲁迅 0.032, 郑伯奇 0.011, 茅盾 0.008 |

**状态：已完成**

**产出文件**：
- `research/analysis/build_network_analysis.py`
- `research/drafts/reports/phase6_network_analysis_report.md`
- `research/drafts/reports/quality_and_limitations_report.md`

## 7. 自动化质量门禁

每阶段结束必须运行：

```powershell
python -m pytest -v
python -c "from kb_schema import validate_data_dir; r=validate_data_dir('data/processed'); print(len(r.errors), len(r.warnings))"
python build_static_site.py --data-dir data/processed --output-dir docs
```

新增或修改的脚本还必须通过定向 `ruff` 检查。全仓历史 lint 债务单独记录，不在阶段任务中无边界清理。

## 8. 阶段验收原则

每个阶段都必须同时交付：

1. 可运行的数据或功能。
2. 明确的生成或处理脚本。
3. 自动测试。
4. 验收报告或可量化指标。
5. 对研究报告和答辩展示有直接价值的产物。

## 9. 项目风险与控制

| 风险 | 控制措施 |
| --- | --- |
| OCR 文字误识别 | 保存原文定位，自动提取结果标注审核状态 |
| 百科被当成权威证据 | 百科仅作为 D 级线索 |
| 网络图把数量误当价值 | 研究发现必须经过证据等级敏感性分析 |
| 候选结论被误展示 | 研究层与发布层分离 |
| 人工审核不可重复 | 固定抽样规则、记录审核决定和说明 |
| 答辩夸大结论 | 单独提供局限说明与口径清单 |

## 10. 非目标

- 不在当前阶段一次性人工确认全部历史事实。
- 不删除研究层中的低置信度或争议线索。
- 不以百科数量或来源数量替代事实级证据质量。
- 不为未来可能需求提前设计复杂配置系统。
