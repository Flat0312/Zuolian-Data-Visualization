from __future__ import annotations

import pandas as pd
from pathlib import Path
from research.analysis.audit_events_and_places import audit_events_and_places


def test_audit_events_and_places_adds_coordinate_precision(sandbox_tmp_path: Path) -> None:
    data_dir = sandbox_tmp_path / "data" / "processed"
    data_dir.mkdir(parents=True)

    places = pd.DataFrame([{
        "place_id": "PLC1", "place_name": "上海", "historical_name": "上海",
        "current_name": "上海", "place_type": "city", "longitude": 121.5, "latitude": 31.2,
        "source_ids": "S1", "confidence": "low"
    }, {
        "place_id": "PLC2", "place_name": "中华艺术大学", "historical_name": "中华艺术大学",
        "current_name": "中华艺术大学", "place_type": "campus", "longitude": 121.4825, "latitude": 31.2603,
        "source_ids": "S2", "confidence": "high"
    }])
    places.to_csv(data_dir / "places.csv", index=False, encoding="utf-8-sig")

    events = pd.DataFrame([{
        "event_id": "E1", "event_name": "左联成立", "event_scope": "collective",
        "canonical_event_key": "zuolian-found", "original_event_names": "",
        "event_date": "1930-03-02", "date_precision": "日", "place_id": "PLC1",
        "historical_location": "上海", "current_address": "上海",
        "longitude": 121.5, "latitude": 31.2, "source_ids": "S1",
        "display_note": "", "correction_reason": "", "confidence": "high", "needs_manual_review": "no"
    }, {
        "event_id": "E2", "event_name": "模糊活动", "event_scope": "entity",
        "canonical_event_key": "", "original_event_names": "",
        "event_date": "1932", "date_precision": "年", "place_id": "",
        "historical_location": "", "current_address": "",
        "longitude": "", "latitude": "", "source_ids": "",
        "display_note": "", "correction_reason": "", "confidence": "low", "needs_manual_review": "yes"
    }])
    events.to_csv(data_dir / "events.csv", index=False, encoding="utf-8-sig")

    participants = pd.DataFrame([{
        "event_participant_id": "EP1", "event_id": "E1", "person_id": "P1",
        "participant_name": "", "participant_role": "", "source_ids": "S1",
        "confidence": "high", "needs_manual_review": "no"
    }])
    participants.to_csv(data_dir / "event_participants.csv", index=False, encoding="utf-8-sig")

    report_path = sandbox_tmp_path / "report.md"
    queue_path = sandbox_tmp_path / "queue.csv"
    summary = audit_events_and_places(data_dir, report_path, queue_path)

    enhanced_places = pd.read_csv(data_dir / "places.csv").fillna("")
    assert "coordinate_precision" in enhanced_places.columns
    city_row = enhanced_places[enhanced_places["place_id"] == "PLC1"].iloc[0]
    campus_row = enhanced_places[enhanced_places["place_id"] == "PLC2"].iloc[0]
    assert city_row["coordinate_precision"] == "city"
    assert campus_row["coordinate_precision"] == "street"

    assert summary["events_total"] == 2
    assert summary["event_review_queue_size"] == 1
    assert summary["place_review_queue_size"] >= 1
    assert report_path.exists()
    assert queue_path.exists()

    queue_df = pd.read_csv(queue_path)
    assert set(queue_df["subject_type"]) == {"event", "place"}
