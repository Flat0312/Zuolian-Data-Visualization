# 返修红→绿验证记录（任务九）

生成时间：2026-09-04（本地执行，未推送）
基线 HEAD：012f10c
验证方式：先构造故障态（红），再恢复/修复（绿），逐项保存命令与结果。

## 1. supported 改 critical 后发布门禁（任务九-1 / 二-6）

- 红：旧逻辑透传已有 `publish_status=supported`，即使 `risk=critical` 仍返回 supported 并保留在发布层。
  复现（旧 `derive_relation_publish_status` 首行 `if existing in statuses: return existing`）：
  `publish_status=supported, final_relation_type=待核验/needs_manual_review=yes/risk=critical/confidence=low` 仍返回 supported（任务书实测）。
- 绿：新逻辑每次按行 + 证据重算，不透传；仅 human_adjudication 的 verified/rejected 保留。
  ```text
  tests/test_remediation_governance.py::test_old_supported_not_passthrough_recompute_on_risk_change PASSED
  tests/test_relation_publication_governance.py::test_pending_high_risk_unverified_excluded_from_default_publish PASSED
  ```
  端到端：R1 配合格 support 先公开，改 `relation_risk_level=critical` 后重建发布，R1 从发布层消失（测试内断言）。

## 2. associated 伪装成 support（任务九-2 / 三）

- 红：若把 `evidence_support=associated` 误判为支持，`associated+pending+locator+quote` 将得到 supported（旧启发式不看证据即为 supported）。
- 绿：新 `has_qualifying_support_evidence` 要求 `support + 未 rejected + locator + (quote|context)`；associated 永不产生 supported。
  ```text
  tests/test_remediation_governance.py::test_associated_pending_cannot_be_supported PASSED
  tests/test_remediation_governance.py::test_fake_support_with_associated_fails_evidence_gate PASSED
  tests/test_remediation_governance.py::test_missing_quote_and_locator_cannot_be_supported PASSED
  ```
  生产现状：10249 条证据全部 `associated+pending+quote为空`，故 supported=0（见下 §5），不再冒充“材料直接支持”。

## 3. 新增 source 不加映射（任务九-3 / 四-8）

- 红：只新增 `sources.csv` 一行（SRC-9999）不同步 `source_works/source_passages` 时：
  ```text
  missing_source_passage: sources.SRC-9999 缺少 passage 映射
  ```
  `validate_data_dir` 报 error（测试 `test_source_incremental_requires_sync` 第一段断言）。
- 绿：调用统一入口 `sync_source_layer(work)` 后：
  ```text
  passages == len(sources)，validate 0 errors，既有 work_id/passage_id 前缀不变，二跑字节一致，citation_count 按实际 passage 数重算
  ```
  ```text
  tests/test_remediation_governance.py::test_source_incremental_requires_sync PASSED
  ```
  生产脚本 `merge_longhua_roster / merge_batch3_event_review / merge_event_evidence_pilot` 已统一调用 `sync_source_layer`；Schema 对单表存在、缺映射、悬空引用、citation_count 错误一律报 error（不再静默漂移）。

## 4. 连续重建 manifest 变化（任务九-4 / 七）

- 红：旧 `build_publish_data` 每次写入 `generated_at=datetime.now(UTC)`，双跑 `publish_manifest.json` 必然不同，`git status` 变脏。
- 绿：默认 `generated_at=unstamped`（仅 `--stamp` 显式写入时间），表按主键排序、`lineterminator="\n"`、`sort_keys=True`，CSV/manifest/report 双跑字节一致。
  ```text
  tests/test_remediation_governance.py::test_publish_double_run_byte_identical PASSED
  python research/analysis/build_publish_data.py (x2) → tables 相同
  python build_static_site.py (x2) → Static site generated: 162 people, 0 relation cards, 147 events (x2 一致)
  git status 第二次无新增变脏（仅本次返修预期变更）
  ```

## 全量回归（任务十）

```text
python -m pytest -q → 103 passed (94 基线 + 9 新增返修)，0 failed，0 skipped
python -m ruff check app.py build_static_site.py kb_schema.py app research/analysis → All checks passed!
validate_data_dir('data/processed') → 0 errors, 13 warnings
validate_data_dir('data/publish') → 0 errors, 1080 warnings (见发布报告分类：过滤后预期 54 孤立人物 + 1013 孤儿来源，研究层基线 12+1 为真正孤立)
git diff --check → 通过（仅 CRLF 警告，无空白错误）
```
