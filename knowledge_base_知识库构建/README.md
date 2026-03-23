# knowledge_base 知识库构建

## 概述

基于 Streamlit 的左联作家知识库交互应用，包含人物关系图、时空地图、事件时间线等功能。

## 快速启动

```bash
cd app/
pip install -r requirements.txt
streamlit run app.py
```

## 目录结构

```
knowledge_base_知识库构建/
├── app/                    # 应用代码（运行目录）
│   ├── app.py              # Streamlit 主应用
│   ├── data_paths.py       # 数据路径配置（已修复，唯一数据源）
│   ├── historical_map.py   # 历史地图模块
│   ├── research_findings.py # 研究发现模块
│   ├── relation_evidence.py # 关系证据模块
│   ├── audit_source_data.py # 数据审计
│   ├── assets/             # 图片/SVG 资源
│   ├── lib/                # 前端 JS 库
│   ├── .streamlit/         # Streamlit 配置
│   └── requirements.txt
├── assets/                 # 资源备份
├── audit/                  # 数据审计报告
│   └── source_data_audit.md
├── config/
│   └── AGENTS.md           # AI Agent 配置
└── README.md
```

## 数据源配置

应用通过 `data_paths.py` 自动解析数据目录：

- **唯一数据源**：`../../output_输出结果/kb_data_知识库数据/`
- 所需文件：`nodes.csv`, `edges.csv`, `edges_audited.csv`, `events.csv`

不要修改数据路径配置，也不要将数据文件复制到 app/ 目录下。

## 依赖

见 `app/requirements.txt`，主要包括：
- streamlit
- pandas
- pyvis
- folium
- altair
- streamlit-folium
