# 脚本清理报告

## 1. 扫描概览

- 扫描范围：项目内全部 Python 脚本，共 33 个
- 分析方法：AST import 解析 + 文本级脚本名引用 + `README.md` / `data_cleaning_数据清洗/README.md` / `knowledge_base_知识库构建/README.md` 交叉核对
- 清理原则：
  - `app.py` 直接 import 的模块全部保留
  - README 中明确属于当前提取/清洗主流程的独立脚本优先保留
  - 未被任何代码引用、且明显是一次性旧工具/调试脚本/失效脚本的文件，移动到 `archive_归档/scripts_废弃脚本/`
  - 对存在功能重叠但仍可能人工使用的脚本，仅标记为“可能冗余”，不移动

## 2. 关键依赖结论

- `knowledge_base_知识库构建/app/app.py` 直接依赖：
  - `data_paths.py`
  - `historical_map.py`
  - `relation_evidence.py`
  - `research_findings.py`
- 数据清洗目录中的脚本基本都是“独立执行脚本”，脚本之间没有 import 级联依赖
- 仅发现少量文本级流程提示关系：
  - `extract_context.py` 提示下一步运行 `write_back_to_xlsx.py`
  - `write_back_to_xlsx.py` 注释说明复用了 `llm_weight_decouple.py` 的阶段逻辑
  - `extract_from_cidian_shi.py` 注释说明复用了 `extract_zuolian.py` 的人物映射思路

## 3. 全部脚本清单

### 3.1 知识库应用脚本

| 脚本 | 用途判断 | 被谁引用/调用 | 它引用/调用谁 | 分类 |
| --- | --- | --- | --- | --- |
| `knowledge_base_知识库构建/app/app.py` | Streamlit 主应用入口 | 无 | `data_paths.py`, `historical_map.py`, `relation_evidence.py`, `research_findings.py` | 核心脚本 |
| `knowledge_base_知识库构建/app/data_paths.py` | 知识库数据目录解析 | `app.py` | 无 | 核心脚本 |
| `knowledge_base_知识库构建/app/historical_map.py` | 历史地图数据建模与转换 | `app.py` | 无 | 核心脚本 |
| `knowledge_base_知识库构建/app/relation_evidence.py` | 关系证据索引构建 | `app.py` | 无 | 核心脚本 |
| `knowledge_base_知识库构建/app/research_findings.py` | 研究发现分析模块 | `app.py` | 无 | 核心脚本 |
| `knowledge_base_知识库构建/app/audit_source_data.py` | 知识库数据审计工具 | 无代码调用，README 列为 app 辅助工具 | 无 | 核心脚本 |

### 3.2 数据清洗与提取脚本

| 脚本 | 用途判断 | 被谁引用/调用 | 它引用/调用谁 | 分类 |
| --- | --- | --- | --- | --- |
| `data_cleaning_数据清洗/scripts/clean_zolian_excel.py` | 当前 Excel 全量主清洗脚本 | 无代码调用，`data_cleaning_数据清洗/README.md` 指定为“主清洗” | 无 | 核心脚本 |
| `data_cleaning_数据清洗/scripts/clean_sheet2.py` | Sheet2 清洗 | 无代码调用，README 列为清洗类脚本 | 无 | 核心脚本 |
| `data_cleaning_数据清洗/scripts/correct_relations.py` | 修正 Sheet2 关系类型 | 无代码调用，README 列为清洗类脚本 | 无 | 核心脚本 |
| `data_cleaning_数据清洗/scripts/extract_context.py` | 从原始文本重提取 Context | 无 import 调用 | 文本提示下一步运行 `write_back_to_xlsx.py` | 核心脚本 |
| `data_cleaning_数据清洗/scripts/fix_sheet2.py` | 对 `context_extracted.csv` 做结构性修复并写回 xlsx | 无代码调用，README 列为清洗类脚本 | 无 | 核心脚本 |
| `data_cleaning_数据清洗/scripts/llm_weight_decouple.py` | 关系权重重评与标签解耦 | 无 import 调用 | 被 `write_back_to_xlsx.py` 注释引用 | 核心脚本 |
| `data_cleaning_数据清洗/scripts/write_back_to_xlsx.py` | 将清洗结果写回 Excel | 无 import 调用，被 `extract_context.py` 作为下一步提示 | 文本引用 `llm_weight_decouple.py` | 核心脚本 |
| `data_cleaning_数据清洗/scripts/extract_zuolian.py` | 从《左联回忆录》提取关系 | 无 import 调用 | 被 `extract_from_cidian_shi.py` 注释引用 | 核心脚本 |
| `data_cleaning_数据清洗/scripts/extract_from_cidian_shi.py` | 从《左联词典》《左联史》提取关系 | 无 import 调用 | 注释引用 `extract_zuolian.py` | 核心脚本 |
| `data_cleaning_数据清洗/scripts/extract_relationships.py` | 从鲁迅日记提取关系 | 无代码调用，README 列为提取类脚本 | 无 | 核心脚本 |
| `data_cleaning_数据清洗/scripts/aggregate_luxun_data.py` | 汇总鲁迅数据 | 无代码调用，README 列为提取类脚本 | 无 | 核心脚本 |
| `data_cleaning_数据清洗/scripts/filter_luxun_diary.py` | 过滤鲁迅日记记录 | 无代码调用，README 列为提取类脚本 | 无 | 核心脚本 |
| `data_cleaning_数据清洗/scripts/convert_to_txt.py` | OCR/PDF 转 TXT 工具 | 无代码引用 | 无 | 可能冗余 |
| `data_cleaning_数据清洗/scripts/ocr_search.py` | OCR 检索孤岛成员 | 无代码引用 | 无 | 可能冗余 |
| `data_cleaning_数据清洗/scripts/export_isolated.py` | 导出孤岛成员名单 | 无代码引用 | 无 | 可能冗余 |
| `data_cleaning_数据清洗/scripts/find_isolated_members.py` | 孤岛成员识别与 PDF 检索 | 无代码引用 | 无 | 可能冗余 |
| `data_cleaning_数据清洗/scripts/search_isolated_v2.py` | 孤岛成员检索优化版 | 无代码引用 | 无 | 可能冗余 |
| `data_cleaning_数据清洗/scripts/process_sheet2.py` | 旧版 Sheet2 处理链脚本，输出 `modified_zuolian.xlsx` | 无代码引用 | 无 | 可能冗余 |
| `data_cleaning_数据清洗/scripts/expand_sheet3.py` | 旧版 Sheet3 扩展脚本，依赖 `final_fixed_zuolian.xlsx` | 无代码引用 | 无 | 可能冗余 |
| `data_cleaning_数据清洗/scripts/verify_evidence.py` | 基于原文片段的人工校验证据工具 | 无代码引用 | 无 | 可能冗余 |
| `data_cleaning_数据清洗/scripts/verify_with_llm.py` | 旧版 LLM 抽样验证脚本，输入为 `modified_zuolian.xlsx` | 无代码引用 | 无 | 可能冗余 |

### 3.3 已归档的明确废弃脚本

| 原脚本 | 用途判断 | 被谁引用/调用 | 判定理由 | 分类 |
| --- | --- | --- | --- | --- |
| `data_cleaning_数据清洗/scripts/_verify_output.py` | 一次性调试输出检查 | 无 | 无 `main`，只对单个本地 CSV 打印统计，不属于流程 | 明确废弃 |
| `data_cleaning_数据清洗/scripts/crop_image.py` | 图片裁剪小工具 | 无 | 与数据清洗/知识库流程无关，硬编码 `../1.png` | 明确废弃 |
| `data_cleaning_数据清洗/scripts/fix_birth_death.py` | 从本地 Word 向桌面 Excel 回填生卒年 | 无 | 写死桌面路径，属于早期一次性迁移脚本 | 明确废弃 |
| `data_cleaning_数据清洗/scripts/fix_data.py` | 早期桌面 Excel 定点修补脚本 | 无 | 写死桌面路径，功能已脱离当前目录结构 | 明确废弃 |
| `data_cleaning_数据清洗/scripts/fix_sheet2_llm.py` | 旧版 LLM 修正脚本 | 无 | 当前文件存在 `IndentationError`，且依赖旧输出链 | 明确废弃 |
| `data_cleaning_数据清洗/scripts/transfer_data.py` | 从 Word 表格向桌面 Excel 转移数据 | 无 | 写死桌面路径，属于一次性数据迁移工具 | 明确废弃 |

## 4. 已移动脚本列表

已创建目录：

- `archive_归档/scripts_废弃脚本/`

已移动文件：

- `data_cleaning_数据清洗/scripts/_verify_output.py` -> `archive_归档/scripts_废弃脚本/_verify_output.py`
- `data_cleaning_数据清洗/scripts/crop_image.py` -> `archive_归档/scripts_废弃脚本/crop_image.py`
- `data_cleaning_数据清洗/scripts/fix_birth_death.py` -> `archive_归档/scripts_废弃脚本/fix_birth_death.py`
- `data_cleaning_数据清洗/scripts/fix_data.py` -> `archive_归档/scripts_废弃脚本/fix_data.py`
- `data_cleaning_数据清洗/scripts/fix_sheet2_llm.py` -> `archive_归档/scripts_废弃脚本/fix_sheet2_llm.py`
- `data_cleaning_数据清洗/scripts/transfer_data.py` -> `archive_归档/scripts_废弃脚本/transfer_data.py`

## 5. 仍保留的核心脚本

- `knowledge_base_知识库构建/app/app.py`
- `knowledge_base_知识库构建/app/data_paths.py`
- `knowledge_base_知识库构建/app/historical_map.py`
- `knowledge_base_知识库构建/app/relation_evidence.py`
- `knowledge_base_知识库构建/app/research_findings.py`
- `knowledge_base_知识库构建/app/audit_source_data.py`
- `data_cleaning_数据清洗/scripts/clean_zolian_excel.py`
- `data_cleaning_数据清洗/scripts/clean_sheet2.py`
- `data_cleaning_数据清洗/scripts/correct_relations.py`
- `data_cleaning_数据清洗/scripts/extract_context.py`
- `data_cleaning_数据清洗/scripts/fix_sheet2.py`
- `data_cleaning_数据清洗/scripts/llm_weight_decouple.py`
- `data_cleaning_数据清洗/scripts/write_back_to_xlsx.py`
- `data_cleaning_数据清洗/scripts/extract_zuolian.py`
- `data_cleaning_数据清洗/scripts/extract_from_cidian_shi.py`
- `data_cleaning_数据清洗/scripts/extract_relationships.py`
- `data_cleaning_数据清洗/scripts/aggregate_luxun_data.py`
- `data_cleaning_数据清洗/scripts/filter_luxun_diary.py`

## 6. 存疑脚本（需要人工判断）

以下脚本没有发现代码级调用，但存在工具价值或功能重叠，因此未移动：

- `data_cleaning_数据清洗/scripts/convert_to_txt.py`
  - 与现有 `input_输入/raw_texts_原始文本/` 中已存在 TXT 数据有重叠，且与 `ocr_search.py` 部分能力重合
- `data_cleaning_数据清洗/scripts/ocr_search.py`
  - 与 `convert_to_txt.py`、孤岛成员检索脚本有交叉
- `data_cleaning_数据清洗/scripts/export_isolated.py`
  - 与 `find_isolated_members.py`、`search_isolated_v2.py` 的前半段能力重合
- `data_cleaning_数据清洗/scripts/find_isolated_members.py`
  - 与 `search_isolated_v2.py` 功能高度相近
- `data_cleaning_数据清洗/scripts/search_isolated_v2.py`
  - 属于 `find_isolated_members.py` 的优化变体
- `data_cleaning_数据清洗/scripts/process_sheet2.py`
  - 使用旧路径 `数据/输出结果/...`，可能属于旧版处理链
- `data_cleaning_数据清洗/scripts/expand_sheet3.py`
  - 依赖旧文件 `final_fixed_zuolian.xlsx`，疑似旧版后处理脚本
- `data_cleaning_数据清洗/scripts/verify_evidence.py`
  - 读旧路径并输出本地文本报告，像人工校核辅助脚本
- `data_cleaning_数据清洗/scripts/verify_with_llm.py`
  - 依赖旧文件 `modified_zuolian.xlsx`，疑似旧版 LLM 抽样验证链

## 7. 最终验证

### 7.1 知识库应用

- `app.py` 可直接导入，未出现 `ImportError`
- `data_paths.py`、`historical_map.py`、`relation_evidence.py`、`research_findings.py` 均可正常导入
- 执行 `python -m streamlit run app.py --server.headless true --server.address 127.0.0.1 --server.port 8512` 时，Streamlit 成功输出：
  - `You can now view your Streamlit app in your browser.`
  - `URL: http://127.0.0.1:8512`

### 7.2 数据路径

- `data_paths.resolve_data_dir()` 解析结果：
  - `D:\1大创\左联知识库项目\output_输出结果\kb_data_知识库数据`
- 目标目录关键文件存在：
  - `nodes.csv`
  - `edges.csv`
  - `edges_audited.csv`
  - `events.csv`

### 7.3 清洗脚本健康度

- `data_cleaning_数据清洗/scripts/` 归档后的剩余活动脚本已重新做语法级编译检查
- 未发现新的语法错误
- 被归档的 6 个脚本均未被活动脚本 import

## 8. 结论

- 已安全归档 6 个“明确废弃”脚本
- 未改动 `app.py` 及其依赖模块
- 未改动知识库数据路径
- 当前知识库应用加载链和数据路径检查通过
- 尚有 9 个“可能冗余”脚本保留待人工二次确认
