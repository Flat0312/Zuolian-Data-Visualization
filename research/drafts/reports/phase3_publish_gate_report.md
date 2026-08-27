# Phase 3 发布门禁报告

发布层由研究层自动生成，研究层原始结论未被删除或覆盖。

| 数据表 | 输入 | 发布 | 过滤 |
| --- | ---: | ---: | ---: |
| `persons.csv` | 162 | 162 | 0 |
| `organizations.csv` | 36 | 36 | 0 |
| `places.csv` | 41 | 41 | 0 |
| `events.csv` | 147 | 147 | 0 |
| `person_relations.csv` | 4238 | 4238 | 0 |
| `org_memberships.csv` | 150 | 73 | 77 |
| `org_membership_evidences.csv` | 581 | 438 | 143 |
| `fact_evidences.csv` | 626 | 479 | 147 |
| `event_participants.csv` | 222 | 222 | 0 |
| `sources.csv` | 1177 | 1177 | 0 |

- Schema 严重错误：0
- Schema 警告：37
- 公开组织身份仅保留 `confirmed_member` 与 `related_person`。
- `candidate` 与 `disputed` 仅保留在研究层。
- `fact_evidences.csv` 中 `review_status=rejected` 的事实证据不进入发布层。
