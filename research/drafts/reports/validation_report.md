# validation_report

## 标准表产出
- persons.csv: 150
- organizations.csv: 3
- places.csv: 77
- events.csv: 262
- person_relations.csv: 4238
- org_memberships.csv: 150
- event_participants.csv: 361
- sources.csv: 1142
- review_queue.csv: 2084
- correction_log.csv: 7218

## 校验结果
- duplicate_person_ids: 0
- duplicate_place_ids: 0
- duplicate_event_ids: 0
- duplicate_relation_ids: 0
- duplicate_membership_ids: 0
- duplicate_event_participant_ids: 0
- missing_relation_people: 0
- missing_membership_people: 0
- missing_event_places: 0
- missing_event_participants: 0
- missing_relation_sources: 0
- missing_event_sources: 0
- missing_review_sources: 0

## 高风险联网核验
- 中国左翼作家联盟成立大会：1930-03-02，中华艺术大学教室（今址上海市虹口区多伦路201弄2号）。
- 左联五烈士遇难：1931-02-07，上海龙华淞沪警备司令部刑场。
- 内山书店秘密会议：保留为 1931 年级别，地点统一为内山书店旧址 / 四川北路2050号，具体日期继续待核。

## app.py 验证
- headless_streamlit_run: yes
- validation_log:
```text

  You can now view your Streamlit app in your browser.

  URL: http://127.0.0.1:8521

```

## 兼容层归档
- 本次运行前未检测到旧版兼容 CSV。

## 产物文件清单
- data\processed\event_evidences.json
- data\processed\event_participants.csv
- data\processed\events.csv
- data\processed\org_memberships.csv
- data\processed\organizations.csv
- data\processed\person_relations.csv
- data\processed\persons.csv
- data\processed\places.csv
- data\processed\runtime_sources\左联史.txt
- data\processed\runtime_sources\左联回忆录_ocr_text.json
- data\processed\runtime_sources\左联词典.txt
- data\processed\sources.csv
- research\logs\correction_log.csv
- research\logs\review_queue.csv
- research\logs\standard_pipeline.log
- research\drafts\reports\reorganization_report.md
- research\drafts\reports\validation_report.md
- research\archive\legacy\old_outputs_旧输出\edges_old.csv
- research\archive\legacy\old_outputs_旧输出\events_old.csv
- research\archive\legacy\old_outputs_旧输出\legacy_kb_compat_20260320_182013\edges.csv
- research\archive\legacy\old_outputs_旧输出\legacy_kb_compat_20260320_182013\edges_audited.csv
- research\archive\legacy\old_outputs_旧输出\legacy_kb_compat_20260320_182013\events.csv
- research\archive\legacy\old_outputs_旧输出\legacy_kb_compat_20260320_182013\merged_events.csv
- research\archive\legacy\old_outputs_旧输出\legacy_kb_compat_20260320_182013\nodes.csv
- research\archive\legacy\old_outputs_旧输出\nodes_old.csv
- research\archive\legacy\old_outputs_旧输出\review_needed_old.csv