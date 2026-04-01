# reorganization_report

## 当前工程结构分析
- 输入仍保留在 input_输入/；清洗脚本位于 data_cleaning_数据清洗/scripts/；中间结果保留在 work_处理中间数据/；唯一生产数据目录锁定为 output_输出结果/kb_data_知识库数据/。
- 现有最可复用的数据资产是 output_输出结果/cleaned_data_清洗数据/《左联相关档案资源目录》_修正版.xlsx，其中 Sheet2_corrected / Sheet3_corrected 已包含关系纠偏、事件校核与待复核标记。
- 旧版 knowledge base 兼容 CSV（nodes/edges/events）仅作为本次标准化输入参考，不再作为最终生产数据源。

## 识别出的主数据流程
1. input_输入/raw_excel_原始表格/《左联相关档案资源目录》.xlsx -> clean_zolian_excel.py
2. clean_zolian_excel.py -> output_输出结果/cleaned_data_清洗数据/《左联相关档案资源目录》_修正版.xlsx
3. build_standard_kb_pipeline.py -> output_输出结果/kb_data_知识库数据/*.csv
4. knowledge_base_知识库构建/app/app.py 从标准表优先加载，并在内存中构造 UI 兼容视图。

## 被复用的脚本列表
- data_cleaning_数据清洗/scripts/clean_zolian_excel.py
- data_cleaning_数据清洗/scripts/build_standard_kb_pipeline.py
- knowledge_base_知识库构建/app/data_paths.py
- knowledge_base_知识库构建/app/app.py

## 被废弃/归档的脚本
- 已在 archive_归档/scripts_废弃脚本/ 中归档的旧脚本继续保持归档状态。
- 不再纳入主 pipeline 的旧链路脚本：process_sheet2.py、expand_sheet3.py、verify_evidence.py、verify_with_llm.py。

## 数据去重说明
- 原有 nodes.csv / edges.csv / edges_audited.csv / merged_events.csv 以及旧 schema 的 events.csv 已从 kb_data_知识库数据/ 移出，归档到 archive_归档/old_outputs_旧输出/。
- 最终知识库数据只保留标准表：persons / organizations / places / events / person_relations / org_memberships / event_participants / sources。

## 路径修复说明
- knowledge_base_知识库构建/app/data_paths.py 只解析 output_输出结果/kb_data_知识库数据/ 一个目录。
- knowledge_base_知识库构建/app/app.py 改为优先读取标准表，不再要求旧版 nodes/edges/events 作为入口。

## app.py 运行结果
- headless_streamlit_run: yes
- 详情见 validation_report.md。