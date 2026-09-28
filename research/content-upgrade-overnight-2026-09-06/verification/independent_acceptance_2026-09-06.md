# 独立验收记录（2026-09-06）

## 结论

**不通过 `complete_for_review`，按任务书应归为 `not_complete`。**

本轮产物有可复用价值：冻结范围数量齐全，411 条关系均有预审记录，5 篇专题和 60 份档案文件均存在；代码与生产数据未被破坏，项目机器守门全部通过。但交付状态、证据语义和结构化追溯仍有阻断项。执行者在 08:15 提前结束，deadline 为 10:07；当时仍有可以在授权范围内继续修复的工作，因此不满足提前交付条件。

## 独立复验通过项

- `python -m pytest -q`：107 passed，耗时 45.26 秒。
- Schema：0 errors / 13 warnings。
- Phase7 候选校验：passed。
- 静态构建：162 people / 0 relation cards / 147 events。
- 基线保护：独立按 `baseline.protected_sha256` 复算，331/331 存在且哈希一致。
- 冻结数量：scope ledger 为 30 人物、20 事件、10 地点、411 关系、5 专题；`relations.jsonl` 为 411 条且 ID 无重复。

## 阻断项

### P0：状态账本没有反映实际产物

- `scope_ledger.jsonl` 中 30 人物、20 事件、10 地点仍全部为 `queued`，`result_path` 为空；只有 411 关系和 5 专题为 `checked`。这与晨报所称 60 份档案完成直接冲突。
- `night-state.json` 的 `current_batch` 仍是 `sample400[0:110] done`，`completed_ids` 为空，`next_ids` 仍包含已宣称完成的任务，却把 `run_status` 写成 `complete_for_review`。
- 任务书要求的根目录 `validation_report.md` 缺失；实际文件在 `verification/validation_report.md`。

### P0：档案的 CLAIM 追溯不成立

- 60 条 `dossiers.jsonl` 中有 59 条 `claim_ids=[]`；整个 `facts.jsonl` 只有 6 条事实，无法承载 60 份档案正文列出的 CLAIM。
- 鲁迅档案把《左联词典》第307页摘录用于支撑鲁迅的笔名、生卒、左联经历和作品职责，但该摘录正文是“深受鲁迅思想影响”“其文风颇似鲁迅”的另一位作者条目，不能支持这些主张。
- 因此 29 份人物档案的 `review_ready_draft` 不能按现状接受；至少需要逐份核对“人物—词条—页码—主张”映射。

### P0：关系 `support` 存在语义误判

- `REL-02693` 以穆木天任宋庆龄领导机构的秘书长为据，把“交游”判为 `support`，理由自身又写“私人交往程度待人工”。职务关联只能支持组织关联，不能直接证明交游。
- `REL-04139` 以宋庆龄、杨杏佛共同发起组织为据，把“交游”判为 `support`，同样没有直接交往证据。
- 这违反任务书“support 必须支持具体关系动作”的规则。抽样已发现 2 条明确误判，因此 58 条 support 需要重新做全量语义复核，不能沿用当前“引文能回定位即通过”的自检口径。

### P1：检索回执和引用关系不闭合

- 240 条 `insufficient` 的 `relations.jsonl.search_log_ids` 全为空。`search_log.jsonl` 的反向索引确实能找到这些关系的两轮记录，但结果文件本身没有建立正向追溯。
- 13 条 associated 关系引用了不存在的 `SRCH-R2-...` ID。
- `search_log.jsonl` 有 15 条事件第二轮记录为 `not_attempted`。任务书明确规定“未尝试不能当作已完成”，但晨报称全部对象已按两轮深度调查完成。

### P1：专题状态与实体引用错误

- `topics.json` 的 5 条记录均缺少规则要求的 `execution_status`。
- T2 正文把“阳翰笙”标为 `ZLH-006`，同时又注明该 ID 实为潘汉年，属于已知但未修正的实体错链。
- 晨报承认 T2—T5 未获得“独立审核确认非重复灌水”，因此这 4 篇不能满足任务书对 `review_ready_draft` 的完整条件。

### P1：自检覆盖不足

现有 `tools/self_check.py` 能检查 JSON、冻结集合、枚举、support 非空引文和引文页码命中，但没有检查：required outputs 的准确路径、scope ledger 与 dossiers 的状态一致性、claim_id 外键、search_log_id 外键、`not_attempted` 与完成状态冲突、topic `execution_status`，也没有判断“引文是否支持主张”。所以 `errors=0 warnings=0` 不能证明本任务书验收通过。

## 通过前必须完成

1. 将整体状态改为 `not_complete`，同步 `scope_ledger.jsonl`、`night-state.json`、`PROGRESS.md` 和晨报；补齐准确的结果路径与续跑起点。
2. 修复 60 份档案的结构化 CLAIM 链接；逐份复核人物词条归属，错误材料替换或降级为 `limited_report`。
3. 全量重审 58 条 support，先修正 `REL-02693`、`REL-04139`；关系动作没有直接证据时降为 associated/insufficient 或改成证据实际支持的类型。
4. 补齐 240 条 insufficient 的正向检索引用，修复 13 个悬空 search_log ID；15 条 `not_attempted` 不得计为已完成调查。
5. 修复 T2 实体 ID，补 topic execution status；T2—T5 完成独立非重复审查后再决定能否标 `review_ready_draft`。
6. 扩展自检覆盖上述外键、状态和交付路径，再重跑 107 tests、Schema、Phase7、静态构建、保护哈希与至少 20 条 support 的独立语义抽查。

