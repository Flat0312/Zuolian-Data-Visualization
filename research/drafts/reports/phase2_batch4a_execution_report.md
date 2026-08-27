# Phase 2 第四批A：五项人工裁决生产执行报告

执行日期：2026-08-27  
执行脚本：`research/analysis/apply_batch4a_decisions.py`  
固定执行前基线：提交 `30d5f77`，四表 `1177/628/148/224`  
授权原话：**「5项全部按推荐方案执行」**（2026-08-27）

本报告记录已获授权的生产落地，不开展新史料搜索；本批只访问和使用既有审计包及生产表中的来源、引文和事件。原始文本、来源、人物、人物关系和组织数据未改动。

## 1. 逐项执行结果

### 1.1 EVT-00001 中国左翼作家联盟成立大会

- 保留并设为 `reviewed`：`FE-EVI-3B84F7AC63`、`FE-EVI-489692805D`、`FE-EVI-58338B3D55`、`FE-EVI-8DB5BD3637`、`FE-EVI-FD8EBDD93E`（对应 AUD-B4A-003/004/005/007/008）。
- 设为 `rejected` 并留在研究层：`FE-EVI-21007B9057`、`FE-EVI-6883382B0E`（对应 AUD-B4A-002/006）；不新建北方分盟筹组事件。
- `FE-EVI-0528D7CD44` 从 `support/machine_extracted` 降为 `lead/pending`（AUD-B4A-001）。
- 事件字段、参与者和来源挂接不变。

### 1.2 EVT-00004

- 事件名：`鲁迅与柔石会面` → `鲁迅与柔石赴北四川路看屋未成`。
- canonical key：`鲁迅与柔石会面|ZLH-001` → `EVT-00004|鲁迅与柔石赴北四川路看屋未成|1930-03-28`。
- `needs_manual_review`：`yes` → `no`；事件来源由 `SRC-1122;SRC-1143;SRC-1144` 补入 `SRC-1167`。
- `FE-EVI-0357C10A69`（AUD-B4A-009）保留原主体但设为 `rejected`。
- `FE-EVI-7792AD0E80`（原主体 `EVT-00006`、原状态 `conflict/reviewed`）改指 `EVT-00004`，变为 `support/reviewed`，并清空 `adjudication_status`；其 `quote`、`locator`、`source_id` 未改。
- 两条既有参与者 `EVP-00006`、`EVP-00007` 仅补入 `SRC-1167`，未新增参与者。

### 1.3 EVT-00005

- 物理删除事件 `EVT-00005`。
- 删除两条参与者：`EVP-00009`、`EVP-00008`。
- 删除事实行 `FE-EVI-04DD852F1C`（AUD-B4A-010）和 `FE-EVI-43ECB964FE`（AUD-B4A-011）；两行的审计结论永久保留在第四批A审计表。
- `FE-EVI-D016A6A994`（AUD-B4A-012）改指 `EVT-00008`，`object_value` 改为“柔石等左联五烈士于1931年2月7日夜或2月8日凌晨在上海龙华警备司令部遇害，鲁迅约于2月10日获悉”，状态为 `support/reviewed`。
- 删除后事实主体和参与者引用无悬空记录。

### 1.4 EVT-00017

- 保留事件，`confidence=low`、`needs_manual_review=yes` 不变。
- `FE-EVI-2EC5596E2B`（AUD-B4A-013）设为 `rejected`，不重复改挂 `EVT-00019`。

### 1.5 EVT-00029

- 事件名：`丁玲抵沪就读平民女校` → `丁玲就读平民女校`。
- canonical key：`ZLH-021|丁玲抵沪就读平民女校|1922` → `EVT-00029|丁玲就读平民女校|1922`。
- 保留 `event_date=1922`、年份精度和现有来源；说明收窄为“1922年丁玲在平民女校就读”，不再把未经本条证据支持的迁入动作作为当前主张。
- `FE-EVI-CDAFBAFA48` 从 `lead/reviewed` 改为 `support/reviewed`（AUD-B4A-014）。
- 既有参与者 `EVP-00043` 仅由 `SRC-1126` 同步为 `SRC-1126;SRC-1154`，未新增参与者。

所有未删除证据的 `reviewer_note` 只追加本批裁决出处；`quote`、`locator`、`source_id` 未改。

## 2. 规模与覆盖率

| 指标 | 执行前 | 执行后 | 变化 |
| --- | ---: | ---: | ---: |
| sources | 1177 | 1177 | 0 |
| fact_evidences | 628 | 626 | -2 |
| events | 148 | 147 | -1 |
| event_participants | 224 | 222 | -2 |
| 已挂接事件 | 28/148 | 26/147 | -2 个事件 |
| 直接支持事件 | 22/148 | 21/147 | -1 个事件 |
| 已确认事件 | 23/148 | 26/147 | +3 个事件 |

三种事件口径均排除 `review_status=rejected`。`EVT-00017` 仅剩拒绝证据，因此重新进入核心待核队列；`EVT-00004`、`EVT-00005`、`EVT-00029` 已从两个 Phase 4 队列移除。

## 3. 发布层与静态站

- 发布层重建结果：`events=147`、`event_participants=222`、`fact_evidences=479`、`sources=1177`；发布层不含任何 `rejected` 事实证据。
- `phase2_evidence_coverage_report.md` 和 `phase2_core_fact_review_queue.csv` 已按新口径生成。
- `phase3_publish_gate_report.md`、`publish_manifest.json` 已按新生产数据生成。
- 静态站重建结果：`162 people, 4227 relation cards, 147 events`。

## 4. 幂等与门禁证据

- 首轮脚本输出：`已执行：sources +0、fact_evidences -2、events -1、event_participants -2；终值 1177/626/147/222。`
- 二轮脚本输出：`无新增/已完成：五项第四批A裁决均已落地，跳过写入。`
- 二轮前后三张研究表 SHA256 一致。
- 隔离重放测试使用 `git archive` 提取 `30d5f77`，首轮执行后 `6 passed`；异常基线测试确认前置校验失败时不写盘。
- Schema：`0 errors / 13 warnings`。

本报告与候选审计报告的追加记录只登记授权和执行路径，不改写第四批A原审计结论。


## 附：验收轮测试修复记录（2026-08-27）

经授权将两个既有测试纳入本次修复范围：

1. `tests/test_batch4a_candidate_audit.py`：「审计表覆盖14条证据」的期望集合改为从审计基线提交 `2a87781` 推导（git show），不再与执行后的实时生产表比较；仍严格断言历史快照恰好14条、无重复、无遗漏，其余全部审计断言原样保留。
2. `tests/test_merge_batch3_review.py::test_batch3_fresh_add_counts_merge_and_remap`：移除硬编码的"148事件"，改为记录运行前生产副本基数并断言合并后等于基数减去运行前实际存在的待删重复事件数（当前生产已无 EVT-00007/EVT-00119，期望增量为0）；EVT-00007/EVT-00119 不存在、参与者与证据无悬空、二次运行字节级幂等等断言全部保留。第三批真实基线重放（150→148）由 test_merge_batch3_real_baseline.py 单独锚定，未做任何削弱。

修复后全量 **74 passed / 0 failed / 0 skipped**（未新增 skip/xfail，未恢复 EVT-00005，未制造占位事件）；Schema 0 错误 / ≤13 警告；第四批A目标计数 1177/626/147/222、覆盖率 26/147、21/147、26/147 复核一致。
