# data_cleaning 数据清洗

## 概述

本目录包含所有数据处理脚本，负责将原始输入数据转化为知识库可用数据。

## 数据流

```
../input_输入/raw_excel_原始表格/   →  清洗  →  ../output_输出结果/cleaned_data_清洗数据/
../input_输入/raw_texts_原始文本/   →  提取  →  ../work_处理中间数据/extracted_抽取结果/
../work_处理中间数据/               →  规范化 →  ../output_输出结果/kb_data_知识库数据/
```

## 脚本说明

### 提取类（`scripts/`）

| 脚本 | 功能 |
|------|------|
| `extract_zuolian.py` | 从左联史/词典提取实体和关系 |
| `extract_context.py` | 提取上下文片段 |
| `extract_from_cidian_shi.py` | 从词典和史书提取结构化数据 |
| `extract_relationships.py` | 实体关系提取 |
| `aggregate_luxun_data.py` | 汇总鲁迅日记数据 |
| `filter_luxun_diary.py` | 过滤鲁迅日记条目 |

### 清洗类（`scripts/`）

| 脚本 | 功能 |
|------|------|
| `clean_zolian_excel.py` | 主清洗：对 Excel 数据进行全量清洗 |
| `clean_sheet2.py` | Sheet2 人物-关系清洗 |
| `fix_data.py` | 通用数据修复 |
| `fix_sheet2.py` / `fix_sheet2_llm.py` | Sheet2 LLM 辅助修复 |
| `fix_birth_death.py` | 生卒年修复 |
| `correct_relations.py` | 关系类型修正 |
| `expand_sheet3.py` | Sheet3 事件扩展 |
| `process_sheet2.py` | Sheet2 标准化处理 |
| `llm_weight_decouple.py` | LLM 辅助关系权重解耦 |

### 工具类（`scripts/`）

| 脚本 | 功能 |
|------|------|
| `ocr_search.py` | OCR 文本搜索 |
| `find_isolated_members.py` / `search_isolated_v2.py` | 孤立成员查找 |
| `write_back_to_xlsx.py` | 将处理结果写回 Excel |
| `verify_evidence.py` / `verify_with_llm.py` | 证据验证 |
| `transfer_data.py` | 数据迁移工具 |
| `export_isolated.py` | 导出孤立节点 |
| `convert_to_txt.py` | 格式转换 |

## 运行说明

```bash
# 切换到脚本目录
cd data_cleaning_数据清洗/scripts

# 运行主清洗脚本（需先配置 API key）
python clean_zolian_excel.py
```

输入路径：`../../input_输入/raw_excel_原始表格/`
输出路径：`../../output_输出结果/cleaned_data_清洗数据/`
