# reorganization_report

## 当前工程结构分析
- 原始输入已集中到 research/raw_excel 与 research/raw_texts；清洗与验证脚本集中到 research/analysis；中间结果集中到 research/intermediate；唯一生产数据目录锁定为 data/processed/。
- 现有最可复用的数据资产是 research/intermediate/cleaned_data/《左联相关档案资源目录》_修正版.xlsx，其中 Sheet2_corrected / Sheet3_corrected 已包含关系纠偏、事件校核与待复核标记。
- 旧版 knowledge base 兼容 CSV（nodes/edges/events）仅作为本次标准化输入参考，不再作为最终生产数据源。

## 识别出的主数据流程
1. research/raw_excel/《左联相关档案资源目录》.xlsx -> clean_zolian_excel.py
2. clean_zolian_excel.py -> research/intermediate/cleaned_data/《左联相关档案资源目录》_修正版.xlsx
3. build_standard_kb_pipeline.py -> data/processed/*.csv
4. app/frontend/app.py 从标准表优先加载，并在内存中构造 UI 兼容视图。

## 被复用的脚本列表
- research/analysis/clean_zolian_excel.py
- research/analysis/build_standard_kb_pipeline.py
- app/frontend/data_paths.py
- app/frontend/app.py

## 被废弃/归档的脚本
- 已在 research/archive/legacy/scripts_废弃脚本/ 中归档的旧脚本继续保持归档状态。
- 不再纳入主 pipeline 的旧链路脚本：process_sheet2.py、expand_sheet3.py、verify_evidence.py、verify_with_llm.py。

## 数据去重说明
- 原有 nodes.csv / edges.csv / edges_audited.csv / merged_events.csv 以及旧 schema 的 events.csv 已从生产数据目录移出，归档到 research/archive/legacy/old_outputs_旧输出/。
- 最终知识库数据只保留标准表：persons / organizations / places / events / person_relations / org_memberships / event_participants / sources。

## 路径修复说明
- app/frontend/data_paths.py 只解析 data/processed/ 一个目录。
- app/frontend/app.py 改为优先读取标准表，不再要求旧版 nodes/edges/events 作为入口。

## app.py 运行结果
- headless_streamlit_run: yes
- 详情见 validation_report.md。