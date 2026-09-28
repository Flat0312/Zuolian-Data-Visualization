from __future__ import annotations

import hashlib
from pathlib import Path

import pandas as pd

from research.analysis.audit_phase7_candidates_independent import audit

REPO_ROOT = Path(__file__).resolve().parents[1]
CANDIDATES_PATH = REPO_ROOT / "research" / "drafts" / "reports" / "phase7_relation_evidence_candidates.csv"
DATA_DIR = REPO_ROOT / "data" / "processed"


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def test_audit_covers_all_candidates_with_verified_anchors_and_no_production_writes(tmp_path: Path) -> None:
    watched = [DATA_DIR / name for name in ("person_relations.csv", "sources.csv", "fact_evidences.csv")]
    before = {path: _sha256(path) for path in watched}

    summary = audit(CANDIDATES_PATH, tmp_path)

    assert summary["rows"] == 12
    assert summary["agree_support"] == 8
    assert summary["agree_insufficient"] == 1
    assert summary["agree_upgrade_to_associated"] == 2
    assert summary["issue_found"] == 1
    assert before == {path: _sha256(path) for path in watched}


def test_audit_csv_gates(tmp_path: Path) -> None:
    audit(CANDIDATES_PATH, tmp_path)
    result = pd.read_csv(tmp_path / "phase7_candidate_independent_audit.csv", encoding="utf-8-sig", dtype=str).fillna("")
    candidates = pd.read_csv(CANDIDATES_PATH, encoding="utf-8-sig", dtype=str).fillna("")

    assert result["candidate_id"].tolist() == candidates["candidate_id"].tolist()
    assert result["review_status"].eq("pending_human_review").all()

    support = result[result["frozen_evidence_support"] == "support"]
    assert len(support) == 8
    assert support["verbatim_located"].eq("True").all()
    assert support["anchor_matches_locator"].eq("True").all()
    assert support["quote_person_tokens"].eq("True").all()

    rel_00113 = result[result["relation_id"] == "REL-00113"].iloc[0]
    assert rel_00113["independent_verdict"] == "agree_insufficient"
    assert int(rel_00113["search_receipts"]) >= 3


def test_audit_flags_rel00523_quote_mismatch_and_confirms_upgrades(tmp_path: Path) -> None:
    audit(CANDIDATES_PATH, tmp_path)
    result = pd.read_csv(tmp_path / "phase7_candidate_independent_audit.csv", encoding="utf-8-sig", dtype=str).fillna("")

    rel_00523 = result[result["relation_id"] == "REL-00523"].iloc[0]
    assert rel_00523["frozen_evidence_support"] == "insufficient"
    assert rel_00523["second_round_support"] == "support"
    assert rel_00523["independent_verdict"] == "issue_found"
    assert rel_00523["new_quote_person_tokens"] == "False"
    assert "EVI-N1-E0960180FF3F" in rel_00523["independent_note"]

    for relation_id in ("REL-01219", "REL-01743"):
        row = result[result["relation_id"] == relation_id].iloc[0]
        assert row["second_round_support"] == "associated"
        assert row["independent_verdict"] == "agree_upgrade_to_associated"
        assert row["new_evidence_verbatim"] == "True"
        assert row["new_quote_person_tokens"] == "True"


def test_audit_report_bans_unauthorized_conclusion_words(tmp_path: Path) -> None:
    audit(CANDIDATES_PATH, tmp_path)
    report = (tmp_path / "phase7_candidate_independent_audit_report.md").read_text(encoding="utf-8")
    csv_text = (tmp_path / "phase7_candidate_independent_audit.csv").read_text(encoding="utf-8-sig")

    for banned in ("已转正", "人工审核通过", "人工复核通过", "已执行", "已合并"):
        assert banned not in report
        assert banned not in csv_text
    assert report.count("pending_human_review") >= 1


def test_audit_is_deterministic(tmp_path: Path) -> None:
    first = tmp_path / "a"
    second = tmp_path / "b"
    audit(CANDIDATES_PATH, first)
    audit(CANDIDATES_PATH, second)

    for name in ("phase7_candidate_independent_audit.csv", "phase7_candidate_independent_audit_report.md"):
        assert _sha256(first / name) == _sha256(second / name)
