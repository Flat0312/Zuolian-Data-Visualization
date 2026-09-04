"""Agent A：关系发布治理门禁测试（新增，不修改既有测试）。"""
from __future__ import annotations

from pathlib import Path

import pandas as pd
from conftest import PROJECT_ROOT, create_standard_dataset


def _add_relation(frame: pd.DataFrame, **overrides: object) -> pd.DataFrame:
    base = {
        "relation_id": "RX",
        "source_person_id": "P1",
        "target_person_id": "P2",
        "original_relation_type": "通信",
        "standard_relation_type": "通信",
        "raw_relation_type": "通信",
        "llm_suggested_relation_type": "",
        "final_relation_type": "通信",
        "llm_reason": "",
        "llm_confidence": "",
        "display_status": "formal",
        "relation_quality_score": "",
        "relation_risk_level": "low",
        "context": "有直接材料支持。",
        "evidence_ref": "鲁迅日记 1930年3月2日",
        "weight": "1",
        "source_ids": "S1",
        "correction_reason": "",
        "confidence": "medium",
        "needs_manual_review": "no",
    }
    base.update(overrides)  # type: ignore
    return pd.concat([frame, pd.DataFrame([base])], ignore_index=True)


def test_pending_high_risk_unverified_excluded_from_default_publish(sandbox_tmp_path: Path) -> None:
    processed = create_standard_dataset(sandbox_tmp_path)
    rel_path = processed / "person_relations.csv"
    rels = pd.read_csv(rel_path, encoding="utf-8-sig")
    # 基线 R1 为 low/通信/medium/no → 应保留为 supported
    rels = _add_relation(rels, relation_id="R-PENDING", relation_risk_level="low", needs_manual_review="yes")
    rels = _add_relation(rels, relation_id="R-CRITICAL", relation_risk_level="critical")
    rels = _add_relation(rels, relation_id="R-HIGH", relation_risk_level="high")
    rels = _add_relation(rels, relation_id="R-UNVERIFIED", final_relation_type="待核验")
    rels = _add_relation(rels, relation_id="R-LOW", confidence="low")
    rels.to_csv(rel_path, index=False, encoding="utf-8-sig")

    from research.analysis.build_publish_data import build_publish_data

    publish_dir = sandbox_tmp_path / "data" / "publish"
    build_publish_data(processed, publish_dir, sandbox_tmp_path / "gate.md")
    published = pd.read_csv(publish_dir / "person_relations.csv", encoding="utf-8-sig", dtype=str).fillna("")
    pub_ids = set(published["relation_id"].tolist())
    assert "R1" in pub_ids  # 对照组：可信支持关系保留
    for excluded in ("R-PENDING", "R-CRITICAL", "R-HIGH", "R-UNVERIFIED", "R-LOW"):
        assert excluded not in pub_ids
    if "publish_status" in published.columns:
        assert set(published["publish_status"].unique().tolist()) <= {"verified", "supported"}


def test_rejected_relation_evidences_excluded_and_fk_valid(sandbox_tmp_path: Path) -> None:
    processed = create_standard_dataset(sandbox_tmp_path)
    evid_path = processed / "relation_evidences.csv"
    pd.DataFrame(
        [
            {
                "relation_evidence_id": "RELE-00001",
                "relation_id": "R1",
                "source_id": "S1",
                "locator": "loc",
                "quote": "",
                "context": "ctx",
                "quote_or_context": "ctx",
                "evidence_support": "associated",
                "source_level": "A",
                "review_status": "pending",
                "reviewer_note": "test",
            },
            {
                "relation_evidence_id": "RELE-00002",
                "relation_id": "R1",
                "source_id": "S2",
                "locator": "loc",
                "quote": "",
                "context": "ctx",
                "quote_or_context": "ctx",
                "evidence_support": "associated",
                "source_level": "B",
                "review_status": "rejected",
                "reviewer_note": "test",
            },
        ]
    ).to_csv(evid_path, index=False, encoding="utf-8-sig")

    from research.analysis.build_publish_data import build_publish_data

    publish_dir = sandbox_tmp_path / "data" / "publish"
    build_publish_data(processed, publish_dir, sandbox_tmp_path / "gate.md")
    published = pd.read_csv(publish_dir / "relation_evidences.csv", encoding="utf-8-sig", dtype=str).fillna("")
    assert set(published["relation_evidence_id"].tolist()) == {"RELE-00001"}


def test_production_supported_subset_never_contains_risky_records() -> None:
    rels = pd.read_csv(PROJECT_ROOT / "data" / "processed" / "person_relations.csv", encoding="utf-8-sig", dtype=str).fillna("")
    assert len(rels) == 4238
    assert "publish_status" in rels.columns
    counts = rels["publish_status"].value_counts().to_dict()
    # 当前生产无人工作业裁决：verified/rejected 均为 0；supported 为可信公开子集
    assert counts.get("verified", 0) == 0
    assert counts.get("rejected", 0) == 0
    assert counts.get("supported", 0) == 1760
    assert counts.get("pending_review", 0) == 2451
    assert counts.get("inferred", 0) == 27
    public = rels[rels["publish_status"].isin({"verified", "supported"})]
    assert len(public) < len(rels)
    assert not (public["relation_risk_level"].isin(["critical", "high"])).any()
    assert not (public["final_relation_type"] == "待核验").any()
    assert not (public["confidence"] == "low").any()
    assert not (public["needs_manual_review"] == "yes").any()
    # 风险列未被反向降险改写：critical 仍为 1974
    assert int((rels["relation_risk_level"] == "critical").sum()) == 1974


def test_production_relation_evidences_reference_valid_ids() -> None:
    rels = pd.read_csv(PROJECT_ROOT / "data" / "processed" / "person_relations.csv", encoding="utf-8-sig", dtype=str).fillna("")
    srcs = pd.read_csv(PROJECT_ROOT / "data" / "processed" / "sources.csv", encoding="utf-8-sig", dtype=str).fillna("")
    evid = pd.read_csv(PROJECT_ROOT / "data" / "processed" / "relation_evidences.csv", encoding="utf-8-sig", dtype=str).fillna("")
    assert len(evid) == 10249
    assert set(evid["review_status"].unique().tolist()) == {"pending"}
    assert set(evid["evidence_support"].unique().tolist()) == {"associated"}
    assert set(evid["relation_id"].tolist()) <= set(rels["relation_id"].tolist())
    assert set(evid["source_id"].tolist()) <= set(srcs["source_id"].tolist())


def test_legacy_ids_and_foreign_keys_stable() -> None:
    from kb_schema import validate_data_dir

    result = validate_data_dir(PROJECT_ROOT / "data" / "processed")
    assert not result.errors
    rels = pd.read_csv(PROJECT_ROOT / "data" / "processed" / "person_relations.csv", encoding="utf-8-sig", dtype=str).fillna("")
    srcs = pd.read_csv(PROJECT_ROOT / "data" / "processed" / "sources.csv", encoding="utf-8-sig", dtype=str).fillna("")
    assert "REL-00001" in set(rels["relation_id"].tolist())
    assert "SRC-0001" in set(srcs["source_id"].tolist())
    assert rels["relation_id"].duplicated().sum() == 0
    assert srcs["source_id"].duplicated().sum() == 0
