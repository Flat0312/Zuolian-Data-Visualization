# 项目重组报告（reorganization_report）

**日期**：2026-03-20  
**对象**：`D:\1大创` → `D:\1大创\左联知识库项目`

---

## 一、原始问题总结

| 问题 | 描述 |
|------|------|
| 目录结构混乱 | 根目录散落脚本、CSV、Excel，多层中文嵌套，可读性差 |
| 多数据源 | `知识库/data/`、`数据/输出结果/`、`数据修正/output/` 均含 CSV，存在数据重复入口 |
| 路径硬编码 | `data_paths.py` 写死了 `data/`、`数据/输出结果/` 两个候选路径 |
| 无清晰数据流 | 输入→清洗→输出→应用的链路不明确 |
| 备份/归档混同 | .locktest 版本、调试日志与生产数据混在一起 |

---

## 二、新目录结构

```
左联知识库项目/
├── input_输入/
│   ├── raw_excel_原始表格/      # 原始 Excel（唯一输入）
│   ├── raw_texts_原始文本/      # 原始文本（左联史、词典、鲁迅日记）
│   └── backups_备份/            # 原始数据备份
├── work_处理中间数据/
│   ├── extracted_抽取结果/      # OCR JSON、提取中间产物
│   ├── normalized_标准化/       # 标准化中间数据
│   ├── review_待复核/           # 待人工复核的 CSV
│   └── temp_临时/               # 临时文件
├── output_输出结果/
│   ├── kb_data_知识库数据/      # ★ 知识库唯一数据源
│   ├── cleaned_data_清洗数据/   # 清洗后的 Excel
│   ├── reports_报告/            # 验证报告、审计报告
│   └── logs_日志/               # （保留空目录）
├── data_cleaning_数据清洗/
│   ├── scripts/                 # 所有清洗/提取/工具脚本（25个）
│   ├── config/
│   └── README.md
├── knowledge_base_知识库构建/
│   ├── app/                     # Streamlit 应用（运行目录）
│   │   ├── app.py
│   │   ├── data_paths.py        # ★ 已修复
│   │   ├── assets/              # 图片资源
│   │   ├── lib/                 # JS 前端库
│   │   └── ...
│   ├── assets/
│   ├── audit/
│   ├── config/
│   └── README.md
├── docs_文档/
│   ├── project_docs_项目文档/   # 申请书、通讯稿
│   └── archive_docs_归档文档/
├── archive_归档/
│   ├── old_outputs_旧输出/      # 旧版 nodes/edges/events CSV
│   ├── debug_logs_调试日志/     # 所有 .log 文件
│   ├── locktest_锁测试/         # .locktest 版本文件
│   └── env_snapshot_环境快照/   # skills/ 等元数据归档
└── README.md
```

---

## 三、文件移动清单

### input_输入/raw_excel_原始表格/
- `《左联相关档案资源目录》.xlsx`（来自 数据修正/input/raw_excel/）
- `大创数据收集(1).xlsx`（来自 数据/输出结果/）

### input_输入/raw_texts_原始文本/
- `左联史.txt`, `左联词典.txt`, `日记全编：全2册 (鲁迅 著).txt`（来自 数据/原始文本/）

### input_输入/backups_备份/
- 5 个 xlsx 备份文件（来自 数据/备份/）

### output_输出结果/kb_data_知识库数据/ ★ 唯一数据源
- `nodes.csv`, `edges.csv`, `edges_audited.csv`, `events.csv`, `merged_events.csv`（来自 知识库/data/）

### output_输出结果/cleaned_data_清洗数据/
- `《左联相关档案资源目录》_修正版.xlsx`（主数据表）
- `《左联相关档案资源目录》_修改日志.xlsx`
- `《左联相关档案资源目录》.xlsx`
- `isolated_members_found.xlsx`, `isolated_members_ocr_result.xlsx`, `最终验证报告.xlsx`

### output_输出结果/reports_报告/
- `verification_report_2026-03-20.md`, `source_data_audit.md`

### work_处理中间数据/review_待复核/
- `review_needed.csv`（来自 根目录）

### work_处理中间数据/extracted_抽取结果/
- `左联史_ocr_text.json`, `左联词典_ocr_text.json`

### data_cleaning_数据清洗/scripts/（25 个脚本）
- 清洗类：`clean_zolian_excel.py`, `clean_sheet2.py`, `fix_*.py`, `correct_relations.py`, `expand_sheet3.py`, `llm_weight_decouple.py`, `process_sheet2.py`
- 提取类：`extract_*.py`, `aggregate_luxun_data.py`, `filter_luxun_diary.py`
- 工具类：`ocr_search.py`, `find_isolated_members.py`, `write_back_to_xlsx.py`, `verify_*.py`, `transfer_data.py`, `export_isolated.py`, `convert_to_txt.py`, `crop_image.py`

### knowledge_base_知识库构建/app/
- `app.py`, `data_paths.py`（已修复）, `historical_map.py`, `research_findings.py`, `relation_evidence.py`, `audit_source_data.py`, `requirements.txt`, 3 个 JSON mock 文件

---

## 四、归档/处理说明

| 类型 | 处理 | 位置 |
|------|------|------|
| `*.locktest.*` 版本文件 | 归档，不删除 | `archive_归档/locktest_锁测试/` |
| 旧版 `nodes/edges/events.csv`（来自 数据/输出结果/） | 归档 | `archive_归档/old_outputs_旧输出/` |
| 所有 `.log` 日志 | 归档 | `archive_归档/debug_logs_调试日志/` |
| `skills/`（元数据MD文件组） | 归档 | `archive_归档/env_snapshot_环境快照/skills/` |
| `test_patch_file.txt` | 归档 | `archive_归档/debug_logs_调试日志/` |

> **未删除任何原始数据**。原始 Excel、txt、CSV 均已完整保留，旧版本移入 archive。

---

## 五、路径修复说明

### data_paths.py（核心修复）

**修改前**：
```python
candidate_data_dirs → [root/"data", root/"数据"/"输出结果", ...]
```

**修改后**：
```python
# app/ 位于 knowledge_base_知识库构建/app/
# project_root = app.parent.parent = 左联知识库项目/
candidate_data_dirs → [project_root/"output_输出结果"/"kb_data_知识库数据"]
```

- 去除了对旧 `data/` 目录的引用
- 使用 `Path(__file__)` 动态解析，无硬编码绝对路径
- 完全相对路径，可在任意机器运行

---

## 六、数据源统一说明

```
旧（多源）：                   新（唯一）：
知识库/data/    ─┐             output_输出结果/kb_data_知识库数据/
数据/输出结果/  ─┤ →  ★         │── nodes.csv
数据修正/work/  ─┘             │── edges.csv
                               │── edges_audited.csv
                               └── events.csv
```

知识库应用（app.py）**只通过 data_paths.resolve_data_dir() 访问一个路径**，不存在多数据源歧义。

---

## 七、app.py 运行验证

**验证命令**：
```bash
cd 左联知识库项目/knowledge_base_知识库构建/app
python -c "from data_paths import resolve_data_dir; print(resolve_data_dir())"
```

**验证结果**：
- ✅ 所有 5 个 Python 文件语法检查通过（AST parse）
- ✅ `data_paths` 模块正常导入
- ✅ `resolve_data_dir()` 返回正确路径：`...output_输出结果/kb_data_知识库数据`
- ✅ `nodes.csv` ：存在
- ✅ `edges.csv` ：存在
- ✅ `edges_audited.csv` ：存在
- ✅ `events.csv` ：存在
- ✅ `assets/` 目录：存在（banner.png 等 5 个资源文件）

**启动命令**：
```bash
cd 左联知识库项目/knowledge_base_知识库构建/app
streamlit run app.py
```

---

## 八、生成文档清单

| 文件 | 说明 |
|------|------|
| `左联知识库项目/README.md` | 项目总说明、数据流、快速启动 |
| `data_cleaning_数据清洗/README.md` | 脚本说明、输入输出约定 |
| `knowledge_base_知识库构建/README.md` | 应用说明、数据源配置 |
| `reorganization_report.md`（本文件） | 完整重组记录 |
