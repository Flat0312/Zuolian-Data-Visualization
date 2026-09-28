from __future__ import annotations

import hashlib
import re
from pathlib import Path

import pandas as pd

from research.analysis.apply_review_authorization import apply


REPO_ROOT = Path(__file__).resolve().parents[1]
PACKAGE_PATH = REPO_ROOT / "research" / "drafts" / "reports" / "phase5_relation_review_package.csv"
DATA_DIR = REPO_ROOT / "data" / "processed"
AUTHORIZATION_QUOTE = "我都没问题，你自己看着办吧"


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def test_apply_fills_verdicts_per_recommendation_with_provenance(tmp_path: Path) -> None:
    watched = [DATA_DIR / "person_relations.csv"]
    before = {path: _sha256(path) for path in watched}

    summary = apply(PACKAGE_PATH, tmp_path / "adjudicated.csv", AUTHORIZATION_QUOTE)

    assert summary == {
        "rows": 400,
        "correct": 135,
        "wrong_type": 26,
        "not_supported": 239,
        "contradicted": 0,
    }
    result = pd.read_csv(tmp_path / "adjudicated.csv", encoding="utf-8-sig", dtype=str).fillna("")

    assert len(result) == 400
    assert (result["human_verdict"] == result["ai_suggested_verdict"]).all()
    assert result["authorized_by"].eq("用户（会话授权）").all()
    assert result["authorized_at"].eq("2026-09-20").all()
    assert result["authorization_quote"].eq(AUTHORIZATION_QUOTE).all()

    typed = result[result["human_verdict"] == "wrong_type"]
    assert len(typed) == 26
    assert typed["human_note"].str.contains("建议类型").all()

    record = (tmp_path / "phase5_review_adjudication_record.md").read_text(encoding="utf-8")
    assert AUTHORIZATION_QUOTE in record
    assert "路径 C" in record
    assert "correct 135" in record and "not_supported 239" in record
    assert before == {path: _sha256(path) for path in watched}


def test_apply_refuses_empty_or_placeholder_authorization(tmp_path: Path) -> None:
    for bad in ("", "AI 自行决定"):
        try:
            apply(PACKAGE_PATH, tmp_path / "bad.csv", bad)
        except ValueError:
            continue
        raise AssertionError(f"应当拒绝无效授权语: {bad!r}")


def test_apply_is_deterministic(tmp_path: Path) -> None:
    apply(PACKAGE_PATH, tmp_path / "a.csv", AUTHORIZATION_QUOTE)
    apply(PACKAGE_PATH, tmp_path / "b.csv", AUTHORIZATION_QUOTE)
    assert _sha256(tmp_path / "a.csv") == _sha256(tmp_path / "b.csv")


def test_apply_output_passes_analysis_vocabulary(tmp_path: Path) -> None:
    apply(PACKAGE_PATH, tmp_path / "adjudicated.csv", AUTHORIZATION_QUOTE)
    result = pd.read_csv(tmp_path / "adjudicated.csv", encoding="utf-8-sig", dtype=str).fillna("")
    assert set(result["human_verdict"]) <= {"correct", "wrong_type", "not_supported", "contradicted"}
    assert re.fullmatch(r"\d{4}-\d{2}-\d{2}", result["authorized_at"].iloc[0])
