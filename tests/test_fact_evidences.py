from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
from conftest import create_standard_dataset


def test_build_fact_evidences_migrates_membership_and_event_evidence(sandbox_tmp_path: Path) -> None:
    data_dir = create_standard_dataset(sandbox_tmp_path)
    event_evidence_path = data_dir / "event_evidences.json"
    event_evidence_path.write_text(
        json.dumps(
            [
                {
                    "evidence_id": "EVENT-OLD-1",
                    "event_id": "E1",
                    "source_loc": "第2页",
                    "quote": "1930年3月2日左联成立。",
                    "confidence": 0.91,
                    "source_id": "S2",
                },
                {
                    "evidence_id": "EVENT-DELETED",
                    "event_id": "E404",
                    "source_loc": "第3页",
                    "quote": "已删除事件。",
                    "confidence": 0.8,
                    "source_id": "S2",
                },
            ],
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    from research.analysis.build_fact_evidences import build_fact_evidences

    result = build_fact_evidences(data_dir, event_evidence_path)

    facts = pd.read_csv(data_dir / "fact_evidences.csv").fillna("")
    assert result == {"membership_evidences": 1, "event_evidences": 1, "skipped_event_evidences": 1}
    assert set(facts["predicate"]) == {"organization_membership", "event_occurrence"}
    assert set(facts["origin_evidence_id"]) == {"OME1", "EVENT-OLD-1"}
    assert "E404" not in set(facts["subject_id"])


def test_report_evidence_coverage_counts_fact_level_evidence(sandbox_tmp_path: Path) -> None:
    data_dir = create_standard_dataset(sandbox_tmp_path)
    report_path = sandbox_tmp_path / "coverage.md"
    queue_path = sandbox_tmp_path / "queue.csv"

    from research.analysis.report_evidence_coverage import report_evidence_coverage

    summary = report_evidence_coverage(data_dir, report_path, queue_path)

    assert summary["memberships"]["covered"] == 1
    assert summary["memberships"]["total"] == 1
    assert summary["events"]["covered"] == 0
    assert summary["events"]["total"] == 1
    assert summary["person_birth_year"]["covered"] == 0
    assert report_path.exists()
    assert queue_path.exists()
