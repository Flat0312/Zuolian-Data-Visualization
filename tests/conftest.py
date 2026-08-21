from __future__ import annotations

import sys
import uuid
from pathlib import Path

import pandas as pd
import pytest


PROJECT_ROOT = Path(__file__).resolve().parents[1]
FRONTEND_DIR = PROJECT_ROOT / "app" / "frontend"

if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))
if str(FRONTEND_DIR) not in sys.path:
    sys.path.insert(0, str(FRONTEND_DIR))


def write_csv(path: Path, rows: list[dict], columns: list[str]) -> None:
    frame = pd.DataFrame(rows, columns=columns)
    path.parent.mkdir(parents=True, exist_ok=True)
    frame.to_csv(path, index=False, encoding="utf-8-sig")


@pytest.fixture
def sandbox_tmp_path() -> Path:
    root = PROJECT_ROOT / ".test_sandbox"
    root.mkdir(exist_ok=True)
    path = root / uuid.uuid4().hex
    path.mkdir()
    return path


def create_standard_dataset(base_dir: Path) -> Path:
    data_dir = base_dir / "data" / "processed"
    data_dir.mkdir(parents=True, exist_ok=True)

    write_csv(
        data_dir / "persons.csv",
        [
            {
                "person_id": "P1",
                "standard_name": "鲁迅",
                "aliases": "周树人",
                "birth_year": 1881,
                "death_year": 1936,
                "birth_death": "1881-1936",
                "role": "作家",
                "reliability": 5,
                "source_ids": "S1",
            },
            {
                "person_id": "P2",
                "standard_name": "茅盾",
                "aliases": "沈雁冰",
                "birth_year": 1896,
                "death_year": 1981,
                "birth_death": "1896-1981",
                "role": "作家",
                "reliability": 5,
                "source_ids": "S2",
            },
        ],
        [
            "person_id",
            "standard_name",
            "aliases",
            "birth_year",
            "death_year",
            "birth_death",
            "role",
            "reliability",
            "source_ids",
        ],
    )

    write_csv(
        data_dir / "organizations.csv",
        [
            {
                "organization_id": "ORG1",
                "standard_name": "左联",
                "aliases": "",
                "org_type": "literary_group",
                "start_date": "1930-03-02",
                "end_date": "",
                "source_ids": "S1",
            }
        ],
        [
            "organization_id",
            "standard_name",
            "aliases",
            "org_type",
            "start_date",
            "end_date",
            "source_ids",
        ],
    )

    write_csv(
        data_dir / "places.csv",
        [
            {
                "place_id": "PLC1",
                "place_name": "上海",
                "historical_name": "上海",
                "current_name": "上海",
                "place_type": "city",
                "longitude": 121.47,
                "latitude": 31.23,
                "source_ids": "S1",
                "confidence": "high",
            }
        ],
        [
            "place_id",
            "place_name",
            "historical_name",
            "current_name",
            "place_type",
            "longitude",
            "latitude",
            "source_ids",
            "confidence",
        ],
    )

    write_csv(
        data_dir / "events.csv",
        [
            {
                "event_id": "E1",
                "event_name": "左联成立",
                "event_scope": "organization",
                "canonical_event_key": "zuolian-found",
                "original_event_names": "左联成立大会",
                "event_date": "1930-03-02",
                "date_precision": "day",
                "place_id": "PLC1",
                "historical_location": "上海",
                "current_address": "上海",
                "longitude": 121.47,
                "latitude": 31.23,
                "source_ids": "S1",
                "display_note": "",
                "correction_reason": "",
                "confidence": "high",
                "needs_manual_review": "no",
            }
        ],
        [
            "event_id",
            "event_name",
            "event_scope",
            "canonical_event_key",
            "original_event_names",
            "event_date",
            "date_precision",
            "place_id",
            "historical_location",
            "current_address",
            "longitude",
            "latitude",
            "source_ids",
            "display_note",
            "correction_reason",
            "confidence",
            "needs_manual_review",
        ],
    )

    write_csv(
        data_dir / "person_relations.csv",
        [
            {
                "relation_id": "R1",
                "source_person_id": "P1",
                "target_person_id": "P2",
                "original_relation_type": "通信",
                "standard_relation_type": "通信",
                "raw_relation_type": "通信",
                "llm_suggested_relation_type": "通信",
                "final_relation_type": "通信",
                "llm_reason": "",
                "llm_confidence": 0.9,
                "display_status": "formal",
                "relation_quality_score": 0.9,
                "relation_risk_level": "low",
                "context": "鲁迅与茅盾通信往来。",
                "evidence_ref": "鲁迅日记 1930年3月2日",
                "weight": 5,
                "source_ids": "S1;S2",
                "correction_reason": "",
                "confidence": "high",
                "needs_manual_review": "no",
            }
        ],
        [
            "relation_id",
            "source_person_id",
            "target_person_id",
            "original_relation_type",
            "standard_relation_type",
            "raw_relation_type",
            "llm_suggested_relation_type",
            "final_relation_type",
            "llm_reason",
            "llm_confidence",
            "display_status",
            "relation_quality_score",
            "relation_risk_level",
            "context",
            "evidence_ref",
            "weight",
            "source_ids",
            "correction_reason",
            "confidence",
            "needs_manual_review",
        ],
    )

    write_csv(
        data_dir / "org_memberships.csv",
        [
            {
                "membership_id": "M1",
                "organization_id": "ORG1",
                "person_id": "P1",
                "membership_role": "正式成员",
                "membership_type": "confirmed_member",
                "source_ids": "S1",
                "evidence_ids": "OME1",
                "evidence_status": "evidence_confirmed",
                "evidence_count": "1",
                "decision_rule": "one_a_level_source",
                "confidence": "high",
                "needs_manual_review": "no",
            }
        ],
        [
            "membership_id",
            "organization_id",
            "person_id",
            "membership_role",
            "membership_type",
            "source_ids",
            "evidence_ids",
            "evidence_status",
            "evidence_count",
            "decision_rule",
            "confidence",
            "needs_manual_review",
        ],
    )

    write_csv(
        data_dir / "org_membership_evidences.csv",
        [
            {
                "evidence_id": "OME1",
                "organization_id": "ORG1",
                "person_id": "P1",
                "evidence_support": "membership",
                "source_id": "S1",
                "source_work": "测试史料",
                "source_level": "A",
                "locator": "第1页",
                "quote": "鲁迅为成员。",
                "review_status": "reviewed",
                "reviewer_note": "",
                "extraction_method": "test_fixture",
            }
        ],
        [
            "evidence_id",
            "organization_id",
            "person_id",
            "evidence_support",
            "source_id",
            "source_work",
            "source_level",
            "locator",
            "quote",
            "review_status",
            "reviewer_note",
            "extraction_method",
        ],
    )

    write_csv(
        data_dir / "event_participants.csv",
        [
            {
                "event_participant_id": "EP1",
                "event_id": "E1",
                "person_id": "P1",
                "participant_name": "鲁迅",
                "participant_role": "发起人",
                "source_ids": "S1",
                "confidence": "high",
                "needs_manual_review": "no",
            }
        ],
        [
            "event_participant_id",
            "event_id",
            "person_id",
            "participant_name",
            "participant_role",
            "source_ids",
            "confidence",
            "needs_manual_review",
        ],
    )

    write_csv(
        data_dir / "fact_evidences.csv",
        [
            {
                "evidence_id": "FE1",
                "subject_type": "person",
                "subject_id": "P1",
                "predicate": "organization_membership",
                "object_value": "ORG1",
                "source_id": "S1",
                "locator": "第1页",
                "quote": "鲁迅为成员。",
                "evidence_support": "support",
                "source_level": "A",
                "review_status": "reviewed",
                "reviewer_note": "",
                "origin_evidence_id": "OME1",
            }
        ],
        [
            "evidence_id",
            "subject_type",
            "subject_id",
            "predicate",
            "object_value",
            "source_id",
            "locator",
            "quote",
            "evidence_support",
            "source_level",
            "review_status",
            "reviewer_note",
            "origin_evidence_id",
        ],
    )

    write_csv(
        data_dir / "sources.csv",
        [
            {
                "source_id": "S1",
                "source_kind": "local_text_citation",
                "title": "鲁迅日记",
                "citation": "鲁迅日记 1930年3月2日",
                "source_path": "research/raw_texts/luxun.txt",
                "source_url": "",
                "evidence_layer": "txt_local_evidence",
                "availability": "local",
                "evidence_strength": "一手",
                "evidence_type": "日记",
                "needs_manual_review": "no",
                "review_note": "",
                "classification_rule": "title:diary",
            },
            {
                "source_id": "S2",
                "source_kind": "local_text_citation",
                "title": "左联词典",
                "citation": "左联词典 第96页",
                "source_path": "research/raw_texts/zuolian_cidian.txt",
                "source_url": "",
                "evidence_layer": "txt_local_evidence",
                "availability": "local",
                "evidence_strength": "二手",
                "evidence_type": "研究论著",
                "needs_manual_review": "no",
                "review_note": "",
                "classification_rule": "title:research",
            },
        ],
        [
            "source_id",
            "source_kind",
            "title",
            "citation",
            "source_path",
            "source_url",
            "evidence_layer",
            "availability",
            "evidence_strength",
            "evidence_type",
            "needs_manual_review",
            "review_note",
            "classification_rule",
        ],
    )

    return data_dir
