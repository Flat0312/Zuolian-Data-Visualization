from __future__ import annotations

from pathlib import Path

import pandas as pd
from conftest import create_standard_dataset


def test_validate_data_dir_reports_missing_columns_and_dangling_refs(sandbox_tmp_path: Path) -> None:
    data_dir = create_standard_dataset(sandbox_tmp_path)

    sources_path = data_dir / "sources.csv"
    sources = pd.read_csv(sources_path)
    sources = sources.drop(columns=["evidence_strength"])
    sources.to_csv(sources_path, index=False, encoding="utf-8-sig")

    relations_path = data_dir / "person_relations.csv"
    relations = pd.read_csv(relations_path)
    relations.loc[0, "source_ids"] = "S999"
    relations.to_csv(relations_path, index=False, encoding="utf-8-sig")

    from kb_schema import validate_data_dir

    result = validate_data_dir(data_dir)

    assert result.has_errors
    assert any(issue.code == "missing_columns" and issue.table == "sources.csv" for issue in result.errors)
    assert any(issue.code == "dangling_reference" and issue.table == "person_relations.csv" for issue in result.errors)


def test_validate_data_dir_reports_warning_only_issues(sandbox_tmp_path: Path) -> None:
    data_dir = create_standard_dataset(sandbox_tmp_path)

    relations_path = data_dir / "person_relations.csv"
    relations = pd.read_csv(relations_path)
    relations.loc[0, "target_person_id"] = "P1"
    relations.to_csv(relations_path, index=False, encoding="utf-8-sig")

    sources_path = data_dir / "sources.csv"
    sources = pd.read_csv(sources_path)
    sources.loc[len(sources)] = {
        "source_id": "S3",
        "source_kind": "web_url",
        "title": "无人引用资料",
        "citation": "",
        "source_path": "",
        "source_url": "https://example.com",
        "evidence_layer": "web",
        "availability": "remote",
        "evidence_strength": "转引",
        "evidence_type": "研究论著",
        "needs_manual_review": "yes",
        "review_note": "",
        "classification_rule": "fallback:web",
    }
    sources.to_csv(sources_path, index=False, encoding="utf-8-sig")

    from kb_schema import validate_data_dir

    result = validate_data_dir(data_dir)

    assert not result.has_errors
    assert any(issue.code == "self_loop_relation" for issue in result.warnings)
    assert any(issue.code == "orphan_source" for issue in result.warnings)


def test_validate_data_dir_rejects_invalid_membership_type(sandbox_tmp_path: Path) -> None:
    data_dir = create_standard_dataset(sandbox_tmp_path)
    memberships_path = data_dir / "org_memberships.csv"
    memberships = pd.read_csv(memberships_path)
    memberships.loc[0, "membership_type"] = "member"
    memberships.to_csv(memberships_path, index=False, encoding="utf-8-sig")

    from kb_schema import validate_data_dir

    result = validate_data_dir(data_dir)

    assert any(issue.code == "invalid_membership_type" for issue in result.errors)


def test_validate_data_dir_rejects_dangling_membership_evidence(sandbox_tmp_path: Path) -> None:
    data_dir = create_standard_dataset(sandbox_tmp_path)
    memberships_path = data_dir / "org_memberships.csv"
    memberships = pd.read_csv(memberships_path)
    memberships["evidence_ids"] = "OME-404"
    memberships.to_csv(memberships_path, index=False, encoding="utf-8-sig")

    from kb_schema import validate_data_dir

    result = validate_data_dir(data_dir)

    assert any(
        issue.code == "dangling_reference"
        and issue.table == "org_memberships.csv"
        and "evidence_ids" in issue.message
        for issue in result.errors
    )


def test_validate_data_dir_rejects_invalid_fact_evidence_subject(sandbox_tmp_path: Path) -> None:
    data_dir = create_standard_dataset(sandbox_tmp_path)
    evidences_path = data_dir / "fact_evidences.csv"
    evidences = pd.read_csv(evidences_path)
    evidences.loc[0, "subject_id"] = "P404"
    evidences.to_csv(evidences_path, index=False, encoding="utf-8-sig")

    from kb_schema import validate_data_dir

    result = validate_data_dir(data_dir)

    assert any(
        issue.code == "dangling_fact_subject" and issue.table == "fact_evidences.csv"
        for issue in result.errors
    )


def test_validate_data_dir_rejects_invalid_fact_evidence_values(sandbox_tmp_path: Path) -> None:
    data_dir = create_standard_dataset(sandbox_tmp_path)
    evidences_path = data_dir / "fact_evidences.csv"
    evidences = pd.read_csv(evidences_path)
    evidences.loc[0, "source_level"] = "Z"
    evidences.loc[0, "review_status"] = "done"
    evidences.to_csv(evidences_path, index=False, encoding="utf-8-sig")

    from kb_schema import validate_data_dir

    result = validate_data_dir(data_dir)

    assert any(issue.code == "invalid_fact_source_level" for issue in result.errors)
    assert any(issue.code == "invalid_fact_review_status" for issue in result.errors)


def test_filter_event_participants_removes_rows_for_deleted_events() -> None:
    from _fix_all import filter_event_participants

    events = [{"event_id": "E1"}, {"event_id": "E2"}]
    participants = [
        {"event_participant_id": "EP1", "event_id": "E1"},
        {"event_participant_id": "EP2", "event_id": "E3"},
    ]

    kept, removed = filter_event_participants(events, participants)

    assert [row["event_participant_id"] for row in kept] == ["EP1"]
    assert removed == 1
