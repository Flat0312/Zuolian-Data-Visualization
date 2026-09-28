"""Agent A：关系发布治理门禁测试（返修版：证据语义 + 来源双态，不修改既有测试语义之外的生产事实）。"""
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


def _write_support_evidence(processed: Path, relation_id: str, evidence_id: str = "RELE-R1") -> None:
    """为指定关系写入一条合格 support 证据（support + locator + quote/context，未 rejected）。"""
    evid_path = processed / "relation_evidences.csv"
    pd.DataFrame(
        [
            {
                "relation_evidence_id": evidence_id,
                "relation_id": relation_id,
                "source_id": "S1",
                "locator": "鲁迅日记 1930年3月2日",
                "quote": "与茅盾通信往来。",
                "context": "鲁迅与茅盾通信往来。",
                "quote_or_context": "与茅盾通信往来。",
                "evidence_support": "support",
                "source_level": "A",
                "review_status": "pending",
                "reviewer_note": "test support",
            }
        ]
    ).to_csv(evid_path, index=False, encoding="utf-8-sig")


def test_pending_high_risk_unverified_excluded_from_default_publish(sandbox_tmp_path: Path) -> None:
    processed = create_standard_dataset(sandbox_tmp_path)
    rel_path = processed / "person_relations.csv"
    rels = pd.read_csv(rel_path, encoding="utf-8-sig")
    # 返修语义：仅合格 support 证据 + 全部门禁通过才可公开；R1 配合格 support 应保留。
    rels = _add_relation(rels, relation_id="R-PENDING", relation_risk_level="low", needs_manual_review="yes")
    rels = _add_relation(rels, relation_id="R-CRITICAL", relation_risk_level="critical")
    rels = _add_relation(rels, relation_id="R-HIGH", relation_risk_level="high")
    rels = _add_relation(rels, relation_id="R-UNVERIFIED", final_relation_type="待核验")
    rels = _add_relation(rels, relation_id="R-LOW", confidence="low")
    rels = _add_relation(rels, relation_id="R-INFERRED", final_relation_type="同属组织")
    rels.to_csv(rel_path, index=False, encoding="utf-8-sig")
    # 仅 R1 拥有合格 support 证据；风险/待核验/low/推断类型即使配 support 也不得公开（此处不配 support，直接验证门禁）。
    _write_support_evidence(processed, "R1")

    from research.analysis.build_publish_data import build_publish_data

    publish_dir = sandbox_tmp_path / "data" / "publish"
    build_publish_data(processed, publish_dir, sandbox_tmp_path / "gate.md")
    published = pd.read_csv(publish_dir / "person_relations.csv", encoding="utf-8-sig", dtype=str).fillna("")
    pub_ids = set(published["relation_id"].tolist())
    assert "R1" in pub_ids  # 对照组：合格 support 证据保留
    for excluded in ("R-PENDING", "R-CRITICAL", "R-HIGH", "R-UNVERIFIED", "R-LOW", "R-INFERRED"):
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
                "locator": "鲁迅日记 1930年3月2日",
                "quote": "与茅盾通信往来。",
                "context": "ctx",
                "quote_or_context": "与茅盾通信往来。",
                "evidence_support": "support",
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
    """生产层公开子集守门：非空后仍不得含任何保守拦截标记，且每条都有合格 support 证据。

    2026-09-28 P5-LANDING 前公开层恒为 0，本测试退化为计数冻结；公开层非空后改为
    直接断言其名称承诺的治理规则（不含 critical/high、待核验、推断类型、low 置信、
    needs_manual_review=yes），并逐条校验合格 support 证据存在。
    """
    rels = pd.read_csv(PROJECT_ROOT / "data" / "processed" / "person_relations.csv", encoding="utf-8-sig", dtype=str).fillna("")
    assert len(rels) == 4238
    assert "publish_status" in rels.columns
    assert "publish_status_origin" in rels.columns
    counts = rels["publish_status"].value_counts().to_dict()
    # P5-LANDING：48 条候选经双方佐证门后 20 条落地 support 证据，其中 5 条通过 derived 门禁进入公开层；
    # 其余 15 条被 critical/high 风险或推断类型挡住，仍留在研究层（未使用 human_adjudication 通道）。
    assert counts.get("verified", 0) == 0
    assert counts.get("rejected", 0) == 0
    assert counts.get("supported", 0) == 5
    assert counts.get("pending_review", 0) == 2451
    assert counts.get("inferred", 0) == 1782
    assert rels["publish_status_origin"].eq("derived").all()
    public = rels[rels["publish_status"].isin({"verified", "supported"})]
    assert len(public) == 5
    assert len(public) < len(rels)
    # 公开子集绝不含风险/待核/低置信/推断类型记录
    assert not public["relation_risk_level"].str.lower().isin(["critical", "high"]).any()
    assert not public["needs_manual_review"].str.lower().eq("yes").any()
    assert not public["confidence"].str.lower().eq("low").any()
    assert not public["final_relation_type"].isin(["待核验", "同属组织", "空间共现", "时空共现"]).any()
    # 风险列未被反向降险改写：critical 仍为 1974
    assert int((rels["relation_risk_level"] == "critical").sum()) == 1974
    # 每条公开关系都必须有合格 support 证据（support + 未 rejected + locator + quote/context）
    evid = pd.read_csv(PROJECT_ROOT / "data" / "processed" / "relation_evidences.csv", encoding="utf-8-sig", dtype=str).fillna("")
    for rid in public["relation_id"].tolist():
        rows = evid[evid["relation_id"] == rid]
        assert len(rows) > 0, f"{rid} 公开但无任何关系证据"
        qualified = rows[
            (rows["evidence_support"] == "support")
            & (rows["review_status"] != "rejected")
            & (rows["locator"].str.strip() != "")
            & ((rows["quote"].str.strip() != "") | (rows["context"].str.strip() != ""))
        ]
        assert len(qualified) >= 1, f"{rid} 公开但缺少合格 support 证据"


def test_production_relation_evidences_reference_valid_ids() -> None:
    rels = pd.read_csv(PROJECT_ROOT / "data" / "processed" / "person_relations.csv", encoding="utf-8-sig", dtype=str).fillna("")
    srcs = pd.read_csv(PROJECT_ROOT / "data" / "processed" / "sources.csv", encoding="utf-8-sig", dtype=str).fillna("")
    evid = pd.read_csv(PROJECT_ROOT / "data" / "processed" / "relation_evidences.csv", encoding="utf-8-sig", dtype=str).fillna("")
    # 2026-09-28 P5-LANDING：新增 20 条经逐字复核 + 双方佐证门的 support 证据；
    # 既有 10249 条机器迁移行保持 associated + pending，未被改判（改判等于凭空提升证据等级）。
    assert len(evid) == 10269
    assert set(evid["review_status"].unique().tolist()) == {"pending", "reviewed"}
    assert set(evid["evidence_support"].unique().tolist()) == {"associated", "support"}
    assert int((evid["evidence_support"] == "support").sum()) == 20
    assert int((evid["review_status"] == "reviewed").sum()) == 20
    assert int((evid["evidence_support"] == "associated").sum()) == 10249
    assert int((evid["review_status"] == "pending").sum()) == 10249
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
