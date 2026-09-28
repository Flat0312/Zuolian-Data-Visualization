# Phase 2 通用事实级证据层实施计划

**目标：** 建立统一事实证据表，把组织身份和事件证据迁移到同一数据契约，并输出证据覆盖率报告。

**实施原则：** 先复用现有可定位证据，不用实体级 `source_ids` 伪装成事实级证据；缺少原文摘录的事实进入待核队列。

## 文件范围

- 新增 `research/analysis/build_fact_evidences.py`
- 新增 `research/analysis/report_evidence_coverage.py`
- 新增 `data/processed/fact_evidences.csv`
- 新增 `research/drafts/reports/phase2_evidence_coverage_report.md`
- 修改 `kb_schema.py`
- 修改 `app/frontend/data_paths.py`
- 修改 `app/frontend/data_loader.py`
- 修改测试与 README

## Task 1：定义事实证据数据契约

- [x] 为必需字段、允许值和悬挂引用编写失败测试。
- [x] 在 Schema 中注册 `fact_evidences.csv`。
- [x] 验证非法主体类型、空来源、悬挂主体和悬挂来源会失败。

## Task 2：迁移组织身份事实证据

- [x] 编写迁移测试。
- [x] 将 `org_membership_evidences.csv` 转换为通用事实证据。
- [x] 保留原证据 ID 的可追踪映射。
- [x] 验证全部 150 条组织身份结论至少存在一条通用事实证据。

## Task 3：迁移事件事实证据

- [x] 编写事件证据 JSON 转换测试。
- [x] 将 `event_evidences.json` 中的可定位证据写入通用表。
- [x] 跳过指向已删除事件的证据并记录数量。
- [x] 验证事件、来源和定位引用有效。

## Task 4：生成覆盖率报告

- [x] 统计组织身份、事件、人物生卒年和人物角色的证据覆盖率。
- [x] 区分“有来源 ID”和“有事实级证据”。
- [x] 输出待核核心事实清单和 Markdown 报告。

## Task 5：加载、文档与验收

- [x] 前端加载器读取事实证据表。
- [x] README 说明事实级证据与实体级来源的区别。
- [x] 运行 Schema、全量测试、静态构建和定向 lint。

## 验收标准

- `fact_evidences.csv` 每行均包含唯一 ID、主体、谓词、来源、支持类型、来源等级和审核状态。
- 组织身份与事件证据可由同一接口查询。
- 所有引用均无悬挂。
- 覆盖率报告明确显示待核范围。
- 全量测试、Schema 与静态构建通过。
