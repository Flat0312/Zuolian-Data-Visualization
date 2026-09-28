from __future__ import annotations

import hashlib
import json
from pathlib import Path

import pandas as pd

from research.analysis.build_phase5_review_package import build

REPO_ROOT = Path(__file__).resolve().parents[1]

TEMPLATE_PATH = REPO_ROOT / "research" / "drafts" / "reports" / "phase5_relation_review_template.csv"
OVERNIGHT_PATH = REPO_ROOT / "research" / "content-upgrade-overnight-2026-09-06" / "relations.jsonl"
DATA_DIR = REPO_ROOT / "data" / "processed"

PACKAGE_COLUMNS = [
    "relation_id",
    "source_person_id",
    "target_person_id",
    "standard_relation_type",
    "relation_risk_level",
    "confidence",
    "context",
    "overnight_evidence_support",
    "overnight_proposed_type",
    "overnight_reason",
    "overnight_evidence_ids",
    "overnight_search_log_ids",
    "ai_suggested_verdict",
    "ai_suggested_type",
    "human_verdict",
    "human_note",
]

VERDICTS = {"correct", "wrong_type", "not_supported", "contradicted"}


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def test_package_covers_template_with_overnight_triage_without_production_writes(tmp_path: Path) -> None:
    watched = [DATA_DIR / name for name in ("person_relations.csv", "sources.csv", "fact_evidences.csv")]
    before = {path: _sha256(path) for path in watched}

    summary = build(TEMPLATE_PATH, OVERNIGHT_PATH, tmp_path)

    assert summary == {
        "rows": 400,
        "support": 48,
        "associated": 113,
        "insufficient": 239,
        "suggested_correct": 135,
        "suggested_wrong_type": 26,
        "suggested_not_supported": 239,
    }
    package = pd.read_csv(tmp_path / "phase5_relation_review_package.csv", encoding="utf-8-sig", dtype=str).fillna("")
    template = pd.read_csv(TEMPLATE_PATH, encoding="utf-8-sig", dtype=str).fillna("")

    assert list(package.columns) == PACKAGE_COLUMNS
    assert len(package) == 400
    assert package["relation_id"].is_unique
    assert package["relation_id"].tolist() == template["relation_id"].tolist()
    for column in ("source_person_id", "target_person_id", "standard_relation_type", "relation_risk_level", "confidence", "context"):
        assert package[column].tolist() == template[column].tolist()
    assert set(package["overnight_evidence_support"]) == {"support", "associated", "insufficient"}
    assert set(package["ai_suggested_verdict"]) <= VERDICTS
    assert before == {path: _sha256(path) for path in watched}


def test_package_never_pre_fills_human_verdict(tmp_path: Path) -> None:
    build(TEMPLATE_PATH, OVERNIGHT_PATH, tmp_path)
    package = pd.read_csv(tmp_path / "phase5_relation_review_package.csv", encoding="utf-8-sig", dtype=str).fillna("")

    assert package["human_verdict"].eq("").all()
    assert package["human_note"].eq("").all()


def test_suggested_verdict_mapping_follows_evidence_and_type_rules(tmp_path: Path) -> None:
    build(TEMPLATE_PATH, OVERNIGHT_PATH, tmp_path)
    package = pd.read_csv(tmp_path / "phase5_relation_review_package.csv", encoding="utf-8-sig", dtype=str).fillna("")

    insufficient = package[package["overnight_evidence_support"] == "insufficient"]
    assert insufficient["ai_suggested_verdict"].eq("not_supported").all()
    assert insufficient["ai_suggested_type"].eq("").all()

    typed = package[package["ai_suggested_verdict"] == "wrong_type"]
    assert (typed["overnight_evidence_support"] != "insufficient").all()
    assert (typed["overnight_proposed_type"] != typed["standard_relation_type"]).all()
    assert (typed["ai_suggested_type"] == typed["overnight_proposed_type"]).all()

    plain_correct = package[
        (package["ai_suggested_verdict"] == "correct") & (package["overnight_evidence_support"] != "insufficient")
    ]
    assert (plain_correct["overnight_proposed_type"] == plain_correct["standard_relation_type"]).all()


def test_build_is_deterministic_across_runs(tmp_path: Path) -> None:
    first = tmp_path / "a"
    second = tmp_path / "b"
    build(TEMPLATE_PATH, OVERNIGHT_PATH, first)
    build(TEMPLATE_PATH, OVERNIGHT_PATH, second)

    for name in ("phase5_relation_review_package.csv", "phase5_relation_review_signoff.md"):
        assert _sha256(first / name) == _sha256(second / name)


def test_signoff_sheet_declares_vocabulary_and_counts(tmp_path: Path) -> None:
    build(TEMPLATE_PATH, OVERNIGHT_PATH, tmp_path)
    signoff = (tmp_path / "phase5_relation_review_signoff.md").read_text(encoding="utf-8")

    for verdict in VERDICTS:
        assert verdict in signoff
    assert "400" in signoff
    assert "pending_human_review" in signoff
    for banned in ("人工已审核", "已签核", "实测准确率"):
        assert banned not in signoff


def test_build_rejects_incomplete_overnight_coverage(tmp_path: Path) -> None:
    template = pd.read_csv(TEMPLATE_PATH, encoding="utf-8-sig", dtype=str).fillna("")
    template.head(3).to_csv(tmp_path / "template.csv", index=False, encoding="utf-8-sig")

    records = [json.loads(line) for line in OVERNIGHT_PATH.read_text(encoding="utf-8").splitlines() if line.strip()]
    dropped = [record for record in records if record["relation_id"] != template.iloc[0]["relation_id"]]
    (tmp_path / "relations.jsonl").write_text(
        "\n".join(json.dumps(record, ensure_ascii=False) for record in dropped) + "\n", encoding="utf-8"
    )

    try:
        build(tmp_path / "template.csv", tmp_path / "relations.jsonl", tmp_path / "out")
    except ValueError as error:
        assert "sample400" in str(error) or "未覆盖" in str(error)
    else:
        raise AssertionError("build should reject templates not fully covered by overnight sample400 records")
