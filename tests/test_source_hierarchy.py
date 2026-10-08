"""Agent A：来源层级（作品/引文/来源族）测试（新增，不修改既有测试）。"""
from __future__ import annotations

from pathlib import Path

import pandas as pd
from conftest import PROJECT_ROOT


def test_source_work_passage_mapping_complete() -> None:
    base = PROJECT_ROOT / "data" / "processed"
    srcs = pd.read_csv(base / "sources.csv", encoding="utf-8-sig", dtype=str).fillna("")
    works = pd.read_csv(base / "source_works.csv", encoding="utf-8-sig", dtype=str).fillna("")
    passages = pd.read_csv(base / "source_passages.csv", encoding="utf-8-sig", dtype=str).fillna("")
    # 2026-09-28 P5-LANDING：为「鲁迅日记 1928年7月1日」注册 1 条同族引文（1177 -> 1178）。
    # 2026-10-08 P5-SWEEP-BATCH1：为左联史 4 页 + 左联词典 1 页注册 5 条同族引文（1178 -> 1183）。
    assert len(srcs) == 1183
    assert len(passages) == 1183
    # 每条引文恰好映射一条 source，且每条 work 存在
    assert set(passages["source_id"].tolist()) == set(srcs["source_id"].tolist())
    assert set(passages["work_id"].tolist()) <= set(works["work_id"].tolist())
    assert len(works) == 65
    assert int(srcs["source_family"].nunique()) == 30
    # 引用条数 ≠ 独立作品数 ≠ 独立来源族数
    assert len(srcs) > len(works) > int(srcs["source_family"].nunique())


def test_same_family_not_counted_as_independent_sources() -> None:
    srcs = pd.read_csv(
        PROJECT_ROOT / "data" / "processed" / "sources.csv", encoding="utf-8-sig", dtype=str
    ).fillna("")
    # 反例：鲁迅日记本地 429 条引文 + 3 条维基文库转录同属一个来源族，不得计为 432 个独立来源
    # （2026-09-28 P5-LANDING 新注册「鲁迅日记 1928年7月1日」1 条，428 -> 429）
    luxun = srcs[srcs["source_family"] == "luxun_diary"]
    assert len(luxun) == 432
    assert luxun["source_family"].nunique() == 1
    # 独立来源族计数必须去重，而非引用条数
    citations = len(srcs)
    families = int(srcs["source_family"].nunique())
    assert families == 30
    assert families < citations
    # 同族内不同 source_id 不得被当作独立来源累加
    per_family = srcs.groupby("source_family")["source_id"].nunique()
    assert int(per_family.loc["luxun_diary"]) == 432
    assert int(per_family.loc["luxun_diary"]) > 1
    # 以族为单位的独立计数为 1，而非 432
    assert 1 < families < citations


def test_source_ids_stable_and_portable_paths() -> None:
    srcs = pd.read_csv(PROJECT_ROOT / "data" / "processed" / "sources.csv", encoding="utf-8-sig", dtype=str).fillna("")
    assert "SRC-0001" in set(srcs["source_id"].tolist())
    assert "SRC-1153" in set(srcs["source_id"].tolist())
    assert srcs["source_id"].duplicated().sum() == 0
    # 仓库内路径已相对化：不得残留 D:\1大创\左联知识库项目 前缀
    for value in srcs["source_path"].tolist():
        assert "D:\\1大创\\左联知识库项目" not in str(value)
        assert "D:/1大创/左联知识库项目" not in str(value)
    # 本地引文仍可定位到仓库相对路径或为空；仓库外 URL 不得伪造为相对路径
    local = srcs[srcs["source_path"] != ""]
    assert (local["source_path"].str.startswith("research/")).all()


def test_publish_manifest_reports_works_citations_families(sandbox_tmp_path: Path) -> None:
    from conftest import create_standard_dataset

    processed = create_standard_dataset(sandbox_tmp_path)
    from research.analysis.build_publish_data import build_publish_data

    publish_dir = sandbox_tmp_path / "data" / "publish"
    manifest = build_publish_data(processed, publish_dir, sandbox_tmp_path / "gate.md")
    # 沙箱无来源层级文件时也应输出 source_summary（作品/家族为 0 而非缺失）
    assert "source_summary" in manifest
    # 生产发布层核对
    import json

    prod_manifest = json.loads((PROJECT_ROOT / "data" / "publish" / "publish_manifest.json").read_text(encoding="utf-8"))
    summary = prod_manifest["source_summary"]
    assert summary["citations"] == 1183
    assert summary["works"] == 65
    assert summary["families"] == 30
