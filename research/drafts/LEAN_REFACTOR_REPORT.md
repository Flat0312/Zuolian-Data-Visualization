# 减负式重构追踪报告

## 1. 主运行链路

### Streamlit 主流程

1. 根入口：`app.py`
2. 实际应用入口：`app/frontend/app.py`
3. 页面模块：
   - `app/frontend/relation_view.py`
   - `app/frontend/event_view.py`
   - `app/frontend/analysis_view.py`
4. 数据入口：
   - `app/frontend/data_loader.py`
   - `app/frontend/data_paths.py`
5. 当前实际生产数据：
   - `output_输出结果/kb_data_知识库数据/persons.csv`
   - `output_输出结果/kb_data_知识库数据/person_relations.csv`
   - `output_输出结果/kb_data_知识库数据/events.csv`
   - `output_输出结果/kb_data_知识库数据/places.csv`
   - `output_输出结果/kb_data_知识库数据/organizations.csv`
   - `output_输出结果/kb_data_知识库数据/org_memberships.csv`
   - `output_输出结果/kb_data_知识库数据/event_participants.csv`
   - `output_输出结果/kb_data_知识库数据/sources.csv`
   - `output_输出结果/kb_data_知识库数据/event_evidences.json`
6. 运行期附加资源：
   - `app/frontend/assets/*`
   - `app/frontend/historical_event_overrides.json`
   - `app/frontend/relation_evidence_mock.json`
   - `app/frontend/shanghai_datav_mock.geojson`
   - `input_输入/raw_texts_原始文本/左联词典.txt`
   - `input_输入/raw_texts_原始文本/左联史.txt`
   - `work_处理中间数据/extracted_抽取结果/左联回忆录_ocr_text.json`

### 静态阅读版链路

1. 生成脚本：`build_static_site.py`
2. 模板资源：`static_site/site.css`、`static_site/site.js`
3. 输出目录：`docs/`
4. Pages 工作流：`.github/workflows/static-pages.yml`

### 启动 / 部署 / 配置

- `start.ps1`
- `render.yaml`
- `requirements.txt`
- `requirements.render.txt`
- `app/frontend/.streamlit/config.toml`

## 2. KEEP（保留）

### 当前运行必须保留

- `app/`
- `app.py`
- `build_static_site.py`
- `static_site/`
- `docs/`
- `output_输出结果/kb_data_知识库数据/`
- `input_输入/raw_texts_原始文本/`
- `work_处理中间数据/extracted_抽取结果/左联回忆录_ocr_text.json`
- `requirements.txt`
- `requirements.render.txt`
- `render.yaml`
- `start.ps1`
- `.github/workflows/`

### 当前数据管线仍应保留

- `data_cleaning_数据清洗/scripts/build_standard_kb_pipeline.py`
- `data_cleaning_数据清洗/scripts/clean_zolian_excel.py`
- `output_输出结果/cleaned_data_清洗数据/《左联相关档案资源目录》_修正版.xlsx`
- `output_输出结果/cleaned_data_清洗数据/《左联相关档案资源目录》_修改日志.xlsx`
- `output_输出结果/cleaned_data_清洗数据/review_needed.csv`

## 3. DELETE（已删除）

### 明确无引用或纯缓存

- `knowledge_base_知识库构建/assets/`
  - 理由：与 `app/frontend/assets/` 内容完全重复，哈希一致，仓库内无任何引用。
- `app/frontend/lib/`
  - 理由：仓库内无任何代码或配置引用这些本地前端库；当前网络图使用 `pyvis` 的内联资源，不依赖该目录。
- `app/frontend/__pycache__/`
  - 理由：Python 缓存。
- `__pycache__/`
  - 理由：Python 缓存。
- `output_输出结果/logs_日志/`
  - 理由：纯日志输出，不参与当前运行。
- `output_输出结果/cleaned_data_清洗数据/workbook_corrected.xlsx`
  - 理由：与 `《左联相关档案资源目录》_修正版.xlsx` 完全重复，文件哈希一致。
- `output_输出结果/cleaned_data_清洗数据/《左联相关档案资源目录》.xlsx`
  - 理由：由管线运行时从原始 Excel 复制生成，可再生，不是当前运行依赖。

## 4. ARCHIVE（已归档到 `archive/legacy/`）

### 历史归档统一收口

- 原 `archive_归档/` 全部内容
- 原 `docs_文档/`
- 根目录历史报告：
  - `reorganization_report.md`
  - `script_cleanup_report.md`

### 原知识库目录中的非主程序内容

- `knowledge_base_知识库构建/README.md`
- `knowledge_base_知识库构建/audit/`
- `knowledge_base_知识库构建/config/`
- `app/frontend/audit_source_data.py`
  - 理由：旧审计脚本，无运行引用，且依赖旧 `data/edges.csv` 结构，已不属于当前主链路。

### 旧输出与研究过程材料

- `output_输出结果/reports_报告/`
  - 理由：生成型报告，不是当前展示链路所需，但有历史参考价值。
- `output_输出结果/cleaned_data_清洗数据/isolated_members_found.xlsx`
- `output_输出结果/cleaned_data_清洗数据/isolated_members_ocr_result.xlsx`
- `output_输出结果/cleaned_data_清洗数据/最终验证报告.xlsx`
  - 理由：一次性分析 / 验证产物，不参与当前运行，但保留参考价值。

## 5. 结构调整

- 主应用从 `knowledge_base_知识库构建/app` 提升到 `app/frontend`
- 历史材料统一收口到 `archive/legacy`
- 根入口、Render 部署、静态站生成、补充脚本和管线脚本已全部改为新路径

## 6. 验证结果

- `python -m compileall app.py build_static_site.py refresh_excerpts.py app data_cleaning_数据清洗/scripts`
  - 结果：通过
- `python build_static_site.py`
  - 结果：通过
- `python -m streamlit run app.py --server.headless true --server.address 127.0.0.1 --server.port 8522`
  - 结果：启动成功，输出 `URL: http://127.0.0.1:8522`

## 7. 未动项说明

- `.venv/`
  - 未删除。虽然属于环境目录，但它是当前本机运行环境，直接删除会影响立即运行能力。
- `.codex_tmp/`
  - 未整体处理。它属于本地临时目录，且存在被占用子目录；不纳入本次主仓库结构调整。
- `data_cleaning_数据清洗/scripts/` 中多数脚本
  - 未批量迁移或删除。当前证据足以确认其中存在非主管线脚本，但不足以对整个目录做无风险裁剪，因此只处理了已明确无引用的项。
