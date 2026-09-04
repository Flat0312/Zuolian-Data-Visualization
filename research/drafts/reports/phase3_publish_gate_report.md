# Phase 3 发布门禁报告（返修版）

发布层由研究层自动生成，研究层原始结论未被删除或覆盖。
关系门禁每次按当前证据全量重算 derived 状态，不透传旧 supported；
仅 human_adjudication 的 verified/rejected 保留（须带 reviewer/reviewed_at/review_note）。
supported 要求 support + 未 rejected + locator + (quote|context)；
associated/pending、无 quote/locator、critical/high、待核验、low、needs_manual_review=yes 均不公开。

| 数据表 | 输入 | 发布 | 过滤 |
| --- | ---: | ---: | ---: |
| `event_participants.csv` | 222 | 222 | 0 |
| `events.csv` | 147 | 147 | 0 |
| `fact_evidences.csv` | 626 | 479 | 147 |
| `org_membership_evidences.csv` | 581 | 438 | 143 |
| `org_memberships.csv` | 150 | 73 | 77 |
| `organizations.csv` | 36 | 36 | 0 |
| `person_relations.csv` | 4238 | 0 | 4238 |
| `persons.csv` | 162 | 162 | 0 |
| `places.csv` | 41 | 41 | 0 |
| `relation_evidences.csv` | 10249 | 0 | 10249 |
| `source_passages.csv` | 1177 | 1177 | 0 |
| `source_works.csv` | 65 | 65 | 0 |
| `sources.csv` | 1177 | 1177 | 0 |

- 研究层 Schema 严重错误：0；警告：13（{'isolated_person': 12, 'orphan_source': 1}）。
- 发布层 Schema 严重错误：0；警告：1080（{'isolated_person': 66, 'orphan_source': 1014}）。
- 发布层警告分类：`isolated_person` 增加主要为过滤非公开关系后预期产生（人物失去公开边）；
`orphan_source` 增加主要为非公开关系证据被过滤后、其来源在发布层暂无公开引用（研究层仍保留）。
- 真正孤立数据（研究层即孤立/孤儿）见研究层警告明细，不得只隐藏 warning。
- 公开组织身份仅保留 `confirmed_member` 与 `related_person`。
- `candidate` 与 `disputed` 仅保留在研究层。
- `fact_evidences.csv` 中 `review_status=rejected` 的事实证据不进入发布层。
- 人物关系仅保留 `publish_status` 为 `verified/supported` 的记录；`pending_review/inferred/rejected`（含 critical/high、待核验、low、needs_manual_review=yes、associated/pending 证据、无 quote/locator、证据冲突）仅保留在研究层。
- `relation_evidences.csv` 中 `review_status=rejected` 的关系证据不进入发布层；且仅保留发布层关系的外键闭合子集（公开关系为 0 时为空表头）。
- 来源层级：引文 1177 条 / 作品 65 种 / 独立来源族 30 个（引用条数≠独立来源作品数，同一来源族不重复计数）。
