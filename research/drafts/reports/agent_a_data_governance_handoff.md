# Agent A 数据治理交接报告（关系发布状态 + 关系证据 + 来源层级）

分支：`codex/data-governance`｜基线 HEAD：`53ccf9c`

## 一、基线核验

- 接手时 HEAD=`53ccf9c`；工作树初始干净，`pytest 74 passed`，Schema `0 errors/13 warnings`（与任务书一致）。
- 本分支工作完成后：`pytest 83 passed（74 既有 + 9 新增），0 failed，0 skipped`；Schema `0 errors/13 warnings`（无新增 warning）。

## 二、状态规则（互斥五态，语义独立）

`research/analysis/relation_publish_status.py:derive_relation_publish_status`

- `verified`：人工确认且有可定位证据（当前生产 0，须人工裁决后方可标记，本分支不自动生成）。
- `supported`：材料直接支持、尚未人工确认（唯一可进入默认公开层的两态之一）。
- `inferred`：同属组织 / 空间共现 / 时空共现或 `confidence=low` 的规则推断（不公开）。
- `pending_review`：`final_relation_type=待核验` 或 `needs_manual_review=yes` 或 `risk in (critical,high)`（不公开）。
- `rejected`：预留（`review_status=rejected` 透传；当前生产 0）。

判定顺序：待核验 → 需审核 → 高风险 → 低置信 → 推断类型 → supported。
已人工标记 `verified/rejected` 的予以保留，其余重算；**风险列原样保留，无反向降险**
（`critical` 仍为 1974，未改成 medium；风险/置信/审核/展示四者语义独立）。

生产分布（4238）：`pending_review 2451 / supported 1760 / inferred 27 / verified 0 / rejected 0`。
公开层（`data/publish`）：`person_relations 4238→1760（过滤 2478）`，不再把全部 4238 当正式已证实展示。

## 三、数据前后变化

| 项目 | 改前 | 改后 |
| --- | --- | --- |
| `person_relations` 列 | 无 `publish_status`，`display_status` 全 `formal` | 新增 `publish_status` 五态；`display_status` 保留兼容（存量仍全 formal，新管线非公开记 `review`） |
| `relation_evidences.csv` | 不存在 | 新增 10249 行（`REL×source` 逐条拆分，`evidence_support=associated`，`review_status` 全 `pending`，`quote_or_context` 由 `context` 迁移，不把共现/同组织标成直接支持） |
| `sources.source_family` | 无 | 新增 30 个族；`luxun_diary` 431 条（含 428 本地 + 3 维基文库转录）同族 |
| `sources.source_path` | 4 种仓库内绝对路径（`D:\1大创\...`） | 全部改为 `research/...` 相对路径；空值保持空；仓库外 URL 原样保留 |
| `source_works.csv` | 不存在 | 新增 65 行（按 title/path/url 去重；`author` 仅鲁迅日记填“鲁迅”，其余待核，不编造） |
| `source_passages.csv` | 不存在 | 新增 1177 行（每 `source_id` 一条，`file_hash` 仅本地存在文件才有，否则留空） |
| 发布层关系证据 | 不存在 | `relation_evidences 10249→5224`（仅公开关系子集 + 排除 rejected） |
| 发布清单 | 仅表行数 | 新增 `source_summary：引文 1177 / 作品 65 / 家族 30`，报告同步输出 |
| 前台/静态站 | 全量展示、无可信标注 | 默认仅 `verified/supported`；待审核/推断须主动开启并明确标注；静态站关系索引加公开子集说明 |

`source_id` 未删除、未重编号；旧 ID（`REL-00001`/`SRC-0001` 等）与外键经 `validate_data_dir` 确认有效。

## 四、反向验证

- 备份 `person_relations.csv` 后将一条 `pending_review（REL-00001，high）` 强行改为 `supported`。
- `test_production_supported_subset_never_contains_risky_records` 以 `1761 == 1760` 失败（红灯符合预期）。
- 恢复备份后 9 项新增测试全绿，全量 `83 passed`。

## 五、验收输出

- `python -m pytest -q`：`83 passed`（74 既有 + 9 新增）。
- `python -m ruff check app.py build_static_site.py kb_schema.py app research/analysis`：`All checks passed!`
- `validate_data_dir('data/processed')`：`0 errors / 13 warnings`（与基线一致，无新增 warning）。
- `python research/analysis/build_publish_data.py`：关系 `4238→1760`，关系证据 `10249→5224`，`source_summary 1177/65/30`。
- `python build_static_site.py`：`162 people, 1758 relation cards, 147 events`（关系卡片为可信子集聚合）。
- `git diff --check`：通过。

## 六、仍需人工决定的问题

1. `verified` 当前为 0：`supported 1760` 只是“材料直接支持”，转 `verified` 须人工逐条确认可定位证据后裁决，本分支未代签。
2. `inferred 27`（同属组织/共现）与 `pending_review 2451` 的升级路径：需人工补证或明确降级展示，不自动转正。
3. `source_works` 的 `author/version/publication_info` 多为“待人工核录”：需人工补权威版本信息，不编造。
4. 增量合并脚本（`merge_longhua_roster` 等）暂不同步 `source_passages/works`：schema 对此不告警（保既有测试），新增测试覆盖生产映射完整；后续合并需补同步逻辑。
5. 时空统一（`canonical_event_key` 去重、`participant_role` 枚举、`date_certainty`）归 Agent B，本分支 schema 对此时空项保持静默，未改动 `events/places/participants` 数据。
