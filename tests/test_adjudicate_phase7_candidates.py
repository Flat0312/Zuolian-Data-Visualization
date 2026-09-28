from __future__ import annotations

import hashlib
import unicodedata
from pathlib import Path

import pandas as pd

from research.analysis.adjudicate_phase7_candidates import adjudicate


REPO_ROOT = Path(__file__).resolve().parents[1]
CANDIDATES_PATH = REPO_ROOT / "research" / "drafts" / "reports" / "phase7_relation_evidence_candidates.csv"
DATA_DIR = REPO_ROOT / "data" / "processed"
HISTORY = REPO_ROOT / "data" / "processed" / "runtime_sources" / "左联史.txt"
DICTIONARY = REPO_ROOT / "data" / "processed" / "runtime_sources" / "左联词典.txt"
AUTHORIZATION_QUOTE = "我都没问题，你自己看着办吧"


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _normalize(text: str) -> str:
    text = unicodedata.normalize("NFKC", text)
    return text.translate(str.maketrans("", "", " \u3000\t\r\n·"))


def test_adjudicated_package_replaces_quote_and_upgrades(tmp_path: Path) -> None:
    watched = [DATA_DIR / name for name in ("person_relations.csv", "relation_evidences.csv", "sources.csv")]
    before = {path: _sha256(path) for path in watched}

    summary = adjudicate(CANDIDATES_PATH, tmp_path, AUTHORIZATION_QUOTE)

    assert summary == {"rows": 12, "support": 9, "associated": 2, "insufficient": 1, "quote_corrected": 1}
    result = pd.read_csv(tmp_path / "phase7_relation_evidence_candidates_adjudicated.csv", encoding="utf-8-sig", dtype=str).fillna("")

    assert len(result) == 12
    assert result["review_status"].eq("adjudicated_authorized").all()
    assert result["authorization_quote"].eq(AUTHORIZATION_QUOTE).all()
    assert result["authorized_at"].eq("2026-09-20").all()

    history = _normalize(HISTORY.read_text(encoding="utf-8", errors="ignore"))
    dictionary = _normalize(DICTIONARY.read_text(encoding="utf-8", errors="ignore"))

    rel_00523 = result[result["relation_id"] == "REL-00523"].iloc[0]
    assert rel_00523["evidence_support"] == "support"
    assert rel_00523["adjudicated_verdict"] == "quote_corrected_to_support"
    quote = rel_00523["quote"]
    assert "潘汉年就去看望他们" in quote and "介绍他俩一同加人左联" in quote and "丁玲" in quote
    assert hashlib.sha256(quote.encode("utf-8")).hexdigest() == rel_00523["quote_sha256"]
    assert _normalize(quote) in history
    assert "EVI-N1-E0960180FF3F" not in quote

    for relation_id, source_text, names in (
        ("REL-01219", dictionary, ("冯乃超", "柔石")),
        ("REL-01743", history, ("丁玲", "穆木天")),
    ):
        row = result[result["relation_id"] == relation_id].iloc[0]
        assert row["evidence_support"] == "associated"
        assert row["adjudicated_verdict"] == "agree_upgrade_to_associated"
        assert _normalize(row["quote"]) in source_text
        assert all(name in row["quote"] for name in names)
        assert hashlib.sha256(row["quote"].encode("utf-8")).hexdigest() == row["quote_sha256"]

    rel_00113 = result[result["relation_id"] == "REL-00113"].iloc[0]
    assert rel_00113["evidence_support"] == "insufficient"
    assert rel_00113["adjudicated_verdict"] == "agree_insufficient"

    support_rows = result[result["evidence_support"] == "support"]
    assert len(support_rows) == 9
    frozen = pd.read_csv(CANDIDATES_PATH, encoding="utf-8-sig", dtype=str).fillna("")
    diary_support = support_rows[support_rows["relation_id"] != "REL-00523"]
    for _, row in diary_support.iterrows():
        assert row["quote"] == frozen[frozen["relation_id"] == row["relation_id"]].iloc[0]["quote"]
    assert before == {path: _sha256(path) for path in watched}


def test_adjudication_record_documents_authorization(tmp_path: Path) -> None:
    adjudicate(CANDIDATES_PATH, tmp_path, AUTHORIZATION_QUOTE)
    record = (tmp_path / "phase7_adjudication_record.md").read_text(encoding="utf-8")

    assert AUTHORIZATION_QUOTE in record
    assert "REL-00523" in record
    assert "换引文" in record or "引文替换" in record
    for banned in ("已转正", "人工审核通过", "已执行生产", "已合并入生产"):
        assert banned not in record


def test_adjudicate_is_deterministic(tmp_path: Path) -> None:
    first = tmp_path / "a"
    second = tmp_path / "b"
    adjudicate(CANDIDATES_PATH, first, AUTHORIZATION_QUOTE)
    adjudicate(CANDIDATES_PATH, second, AUTHORIZATION_QUOTE)
    for name in ("phase7_relation_evidence_candidates_adjudicated.csv", "phase7_adjudication_record.md"):
        assert _sha256(first / name) == _sha256(second / name)
