# validation_report

## 标准表产出
- persons.csv: 150
- organizations.csv: 3
- places.csv: 86
- events.csv: 279
- person_relations.csv: 4238
- org_memberships.csv: 150
- event_participants.csv: 371
- sources.csv: 1130
- review_queue.csv: 2261
- correction_log.csv: 7051

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

## output 文件清单
- output_输出结果\cleaned_data_清洗数据\isolated_members_found.xlsx
- output_输出结果\cleaned_data_清洗数据\isolated_members_ocr_result.xlsx
- output_输出结果\cleaned_data_清洗数据\review_needed.csv
- output_输出结果\cleaned_data_清洗数据\workbook_corrected.xlsx
- output_输出结果\cleaned_data_清洗数据\《左联相关档案资源目录》.xlsx
- output_输出结果\cleaned_data_清洗数据\《左联相关档案资源目录》_修改日志.xlsx
- output_输出结果\cleaned_data_清洗数据\《左联相关档案资源目录》_修正版.xlsx
- output_输出结果\cleaned_data_清洗数据\最终验证报告.xlsx
- output_输出结果\kb_data_知识库数据\event_evidences.json
- output_输出结果\kb_data_知识库数据\event_participants.csv
- output_输出结果\kb_data_知识库数据\events.csv
- output_输出结果\kb_data_知识库数据\org_memberships.csv
- output_输出结果\kb_data_知识库数据\organizations.csv
- output_输出结果\kb_data_知识库数据\person_relations.csv
- output_输出结果\kb_data_知识库数据\persons.csv
- output_输出结果\kb_data_知识库数据\places.csv
- output_输出结果\kb_data_知识库数据\sources.csv
- output_输出结果\logs_日志\correction_log.csv
- output_输出结果\logs_日志\review_queue.csv
- output_输出结果\logs_日志\standard_pipeline.log
- output_输出结果\reports_报告\llm_relation_rerank_report.md
- output_输出结果\reports_报告\reorganization_report.md
- output_输出结果\reports_报告\source_data_audit.md
- output_输出结果\reports_报告\validation_report.md
- output_输出结果\reports_报告\verification_report_2026-03-20.md