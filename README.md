# 左联知识库项目

中国左翼作家联盟（1930-1936）数字人文知识库，包含数据清洗管线与 Streamlit 可视化应用。

## 功能概览

- 结构化人物、关系、事件、地点等知识库数据
- 统一的数据清洗与标准化脚本
- 基于 Streamlit 的人物关系与事件可视化应用

## 目录结构

```text
左联知识库项目/
├─ input_输入/                  # 输入数据（原始表格等）
├─ data_cleaning_数据清洗/       # 清洗/提取/验证脚本
├─ work_处理中间数据/            # 中间文件（默认不入库）
├─ output_输出结果/              # 输出数据（应用仅读取 kb_data_知识库数据）
├─ knowledge_base_知识库构建/     # Streamlit 应用
├─ docs_文档/                   # 项目说明文档
└─ archive_归档/                 # 历史归档（默认不入库）
```

## 快速开始

### 1) 安装依赖

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

### 2) 启动知识库应用

```bash
cd knowledge_base_知识库构建/app
streamlit run app.py
```

应用数据目录由 [`knowledge_base_知识库构建/app/data_paths.py`](knowledge_base_知识库构建/app/data_paths.py) 统一解析，默认读取：

`output_输出结果/kb_data_知识库数据/`

### 3) 可选：运行清洗脚本

```bash
cd data_cleaning_数据清洗/scripts
python build_standard_kb_pipeline.py
```

## 环境变量

涉及 LLM 的脚本不再存储密钥在代码内，请通过环境变量注入：

```bash
OPENAI_API_KEY=your_api_key
OPENAI_BASE_URL=https://api.openai.com/v1
OPENAI_MODEL=gpt-4o
```

可参考 `.env.example`。

## 仓库发布策略

- 默认提交代码、文档和核心知识库数据（`output_输出结果/kb_data_知识库数据/`）
- 默认忽略缓存、中间产物、归档、日志与本地备份
- 默认忽略 `input_输入/raw_texts_原始文本/` 与 `input_输入/backups_备份/`，避免上传版权和冗余文件

## License

[MIT](LICENSE)
