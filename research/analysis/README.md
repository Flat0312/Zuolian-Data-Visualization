# research/analysis

本目录集中放置研究阶段使用的数据处理、验证、抽取与重建脚本，不属于前台应用运行目录。

## 当前数据流

```text
research/raw_excel/                     -> clean_zolian_excel.py -> research/intermediate/cleaned_data/
research/raw_texts/                     -> extract_*.py          -> research/intermediate/extracted/
research/intermediate/cleaned_data/     -> build_standard_kb_pipeline.py
research/intermediate/extracted/        -> build_event_evidence.py / refresh_excerpts.py
data/processed/                         -> app/frontend/
```

## 关键脚本

- `build_standard_kb_pipeline.py`: 从研究输入重建 `data/processed/` 标准知识库数据。
- `clean_zolian_excel.py`: 清洗原始 Excel，产出修正版与修改日志。
- `build_event_evidence.py`: 为事件页补充证据索引。
- `refresh_excerpts.py`: 从研究文本与 OCR 缓存刷新前台证据摘录。

## 运行说明

```bash
cd research/analysis
python build_standard_kb_pipeline.py
```

主应用不直接读取本目录；应用只读取 `data/processed/` 与 `data/processed/runtime_sources/`。
