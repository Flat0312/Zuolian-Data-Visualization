from __future__ import annotations

import math
from pathlib import Path

import pandas as pd
import pytest

from conftest import PROJECT_ROOT, create_standard_dataset


# ─── utils.py tests ──────────────────────────────────────────────────────────

def test_clean_text_normalizes_whitespace() -> None:
    from utils import clean_text

    assert clean_text("  hello   world  ") == "hello world"
    assert clean_text("line1\r\nline2") == "line1 line2"
    assert clean_text(None) == ""
    assert clean_text(float("nan")) == ""


def test_clean_text_truncates() -> None:
    from utils import clean_text

    result = clean_text("abcdefghij", limit=5)
    assert result == "abcde..."


def test_clean_text_returns_fallback_for_empty() -> None:
    from utils import clean_text

    assert clean_text("", fallback="N/A") == "N/A"
    assert clean_text(None, fallback="N/A") == "N/A"


def test_split_ids_basic() -> None:
    from utils import split_ids

    assert split_ids("P1;P2;P3") == ["P1", "P2", "P3"]
    assert split_ids("P1；P2、P3") == ["P1", "P2", "P3"]
    assert split_ids(None) == []
    assert split_ids(float("nan")) == []


def test_split_ids_strips_whitespace() -> None:
    from utils import split_ids

    assert split_ids(" P1 ; P2 ") == ["P1", "P2"]


# ─── historical_map.py tests ─────────────────────────────────────────────────

def test_parse_year_from_various_formats() -> None:
    from historical_map import _parse_year

    assert _parse_year("1930") == 1930
    assert _parse_year("约1932年") == 1932
    assert _parse_year("no year") is None
    assert _parse_year(None) is None
    assert _parse_year("") is None


def test_to_float_converts_values() -> None:
    from historical_map import _to_float

    assert _to_float(3.14) == pytest.approx(3.14)
    assert _to_float("2.5") == pytest.approx(2.5)
    assert _to_float(None) is None
    assert _to_float("") is None
    assert _to_float("not_a_number") is None


def test_summary_mentions_location_and_category() -> None:
    from historical_map import _summary

    result = _summary("上海", "成立大会", ["鲁迅", "茅盾"])
    assert "上海" in result
    assert "成立大会" in result
    assert "鲁迅" in result


def test_significance_categorises_events() -> None:
    from historical_map import _significance

    assert "组织化" in _significance("成立大会", [], "上海")
    assert "压迫" in _significance("逮捕", [], "上海")
    assert "联络" in _significance("通信", [], "上海")
    assert "时空锚点" in _significance("其他", [], "上海")


# ─── data_paths.py tests ─────────────────────────────────────────────────────

def test_candidate_data_dirs_returns_project_processed(sandbox_tmp_path: Path) -> None:
    from data_paths import candidate_data_dirs

    app_dir = sandbox_tmp_path / "app" / "frontend"
    app_dir.mkdir(parents=True, exist_ok=True)
    result = candidate_data_dirs(app_dir)
    assert len(result) >= 1
    assert result[0].name == "processed"


def test_resolve_data_dir_returns_path_when_files_exist(sandbox_tmp_path: Path) -> None:
    from data_paths import CORE_DATA_FILES, resolve_data_dir

    create_standard_dataset(sandbox_tmp_path)
    app_dir = sandbox_tmp_path / "app" / "frontend"
    app_dir.mkdir(parents=True, exist_ok=True)
    result = resolve_data_dir(app_dir)
    assert result.exists()
    assert all((result / f).exists() for f in CORE_DATA_FILES)


def test_membership_profile_for_person_returns_layered_identity() -> None:
    from relation_view import membership_profile_for_person

    memberships = pd.DataFrame(
        [
            {
                "person_id": "P1",
                "organization_id": "ORG-001",
                "membership_type": "candidate",
                "membership_role": "成员身份待核",
                "confidence": "medium",
                "decision_rule": "member_evidence_below_threshold",
                "needs_manual_review": "yes",
            }
        ]
    )
    evidence = pd.DataFrame(
        [
            {
                "person_id": "P1",
                "organization_id": "ORG-001",
                "evidence_id": "OME1",
                "evidence_support": "membership",
                "source_work": "左联词典",
                "locator": "第10页",
                "quote": "鲁迅是左联常委。",
            }
        ]
    )

    profile = membership_profile_for_person("P1", memberships, evidence)

    assert profile["status_label"] == "成员身份待核"
    assert profile["evidence_count"] == 1
    assert profile["evidence"][0]["source_work"] == "左联词典"


# ─── relation_view pair_summary tests ────────────────────────────────────────

def test_pair_summary_from_profiles_empty() -> None:
    from app import pair_summary_from_profiles

    df = pair_summary_from_profiles({})
    assert df.empty


def test_pair_summary_from_profiles_basic() -> None:
    from app import pair_summary_from_profiles
    from relation_view import PairProfile

    profile = PairProfile(
        pair_key="P1__P2",
        person_a_id="P1",
        person_a_name="鲁迅",
        person_b_id="P2",
        person_b_name="茅盾",
        relation_types=["通信"],
        relation_count=3,
        max_weight=5.0,
        display_status_counts={"formal": 2, "review": 1},
        evidence_samples=["鲁迅日记 1930年"],
        context_samples=["鲁迅与茅盾通信往来"],
    )
    df = pair_summary_from_profiles({"P1__P2": profile})
    assert len(df) == 1
    assert df.iloc[0]["人物甲"] == "鲁迅"
    assert df.iloc[0]["relation_count"] == 3
    assert df.iloc[0]["formal_count"] == 2
