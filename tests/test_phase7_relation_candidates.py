from __future__ import annotations

import hashlib
from pathlib import Path

import pandas as pd

from research.analysis.build_phase7_relation_candidates import build
from research.analysis.validate_phase7_relation_candidates import validate


REPO_ROOT = Path(__file__).resolve().parents[1]


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def test_phase7_builds_frozen_candidate_package_without_production_writes(tmp_path: Path) -> None:
    data_dir = REPO_ROOT / "data" / "processed"
    watched = [data_dir / name for name in ("person_relations.csv", "sources.csv", "relation_evidences.csv")]
    before = {path: _sha256(path) for path in watched}

    summary = build(data_dir, tmp_path)

    assert summary == {"selected": 12, "candidates": 12, "receipts": 2, "support": 8, "associated": 0, "conflict": 0, "insufficient": 4}
    assert validate(tmp_path) == []
    assert before == {path: _sha256(path) for path in watched}


def test_phase7_validator_rejects_support_without_verbatim_quote(tmp_path: Path) -> None:
    build(REPO_ROOT / "data" / "processed", tmp_path)
    candidates_path = tmp_path / "phase7_relation_evidence_candidates.csv"
    candidates = pd.read_csv(candidates_path, encoding="utf-8-sig", dtype=str).fillna("")
    candidates.loc[candidates["evidence_support"] == "support", "quote"] = ""
    candidates.loc[candidates["evidence_support"] == "support", "quote_sha256"] = ""
    candidates.to_csv(candidates_path, index=False, encoding="utf-8-sig")

    assert any("support 缺少" in error for error in validate(tmp_path))
