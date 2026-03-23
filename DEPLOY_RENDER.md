# Render 部署说明

本项目已提供 `render.yaml`，可用 Render Blueprint 一键创建服务。

## 一键部署（推荐）

1. 打开 Render 控制台，选择 `New +` -> `Blueprint`。
2. 连接本仓库：`Flat0312/Zuolian-Data-Visualization`。
3. Render 会自动识别仓库根目录的 `render.yaml`。
4. 点击 `Apply` 创建服务并等待构建完成。

## 当前配置

- 构建命令：`pip install --upgrade pip && pip install -r requirements.render.txt`
- 启动命令：`cd "knowledge_base_知识库构建/app" && streamlit run app.py --server.address 0.0.0.0 --server.port $PORT --server.headless true`
- 运行计划：`free`

## 注意事项

- 首次启动会读取仓库内 `output_输出结果/kb_data_知识库数据/` 的 CSV 数据。
- 若后续要接入私有数据或外部 API，请在 Render 的 Environment 里配置变量，不要写死在代码中。
