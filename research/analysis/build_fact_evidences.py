from __future__ import annotations

import argparse
import json
from pathlib import Path

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_DATA_DIR = PROJECT_ROOT / "data" / "processed"
DEFAULT_EVENT_EVIDENCES = DEFAULT_DATA_DIR / "event_evidences.json"

FACT_EVIDENCE_COLUMNS = [
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
]


def _read_csv(path: Path) -> pd.DataFrame:
    return pd.read_csv(path, encoding="utf-8-sig").fillna("")


def _source_level_map(sources: pd.DataFrame) -> dict[str, str]:
    strength_to_level = {"一手": "A", "二手": "B", "转引": "C", "推断": "D"}
    return {
        str(row["source_id"]).strip(): strength_to_level.get(str(row.get("evidence_strength", "")).strip(), "D")
        for _, row in sources.iterrows()
    }


def _generic_support(value: object) -> str:
    mapping = {
        "membership": "support",
        "nonmember": "oppose",
        "related": "lead",
        "lead": "lead",
    }
    return mapping.get(str(value).strip(), "lead")


def _review_status(value: object) -> str:
    status = str(value).strip()
    return status if status in {"pending", "reviewed", "rejected", "machine_extracted"} else "pending"


def build_fact_evidences(data_dir: Path, event_evidence_path: Path | None = None) -> dict[str, int]:
    data_dir = Path(data_dir)
    sources = _read_csv(data_dir / "sources.csv")
    memberships = _read_csv(data_dir / "org_membership_evidences.csv")
    events = _read_csv(data_dir / "events.csv")
    event_names = events.set_index("event_id")["event_name"].to_dict()
    source_levels = _source_level_map(sources)

    rows: list[dict[str, object]] = []
    for _, row in memberships.iterrows():
        origin_id = str(row["evidence_id"]).strip()
        rows.append(
            {
                "evidence_id": f"FE-{origin_id}",
                "subject_type": "person",
                "subject_id": str(row["person_id"]).strip(),
                "predicate": "organization_membership",
                "object_value": str(row["organization_id"]).strip(),
                "source_id": str(row["source_id"]).strip(),
                "locator": str(row.get("locator", "")).strip(),
                "quote": str(row.get("quote", "")).strip(),
                "evidence_support": _generic_support(row.get("evidence_support", "")),
                "source_level": str(row.get("source_level", "")).strip() or source_levels.get(str(row["source_id"]).strip(), "D"),
                "review_status": _review_status(row.get("review_status", "")),
                "reviewer_note": str(row.get("reviewer_note", "")).strip(),
                "origin_evidence_id": origin_id,
            }
        )

    event_evidence_path = event_evidence_path or data_dir / "event_evidences.json"
    event_evidence_count = 0
    skipped_event_evidences = 0
    if event_evidence_path.exists():
        event_evidences = json.loads(event_evidence_path.read_text(encoding="utf-8-sig"))
        for item in event_evidences:
            event_id = str(item.get("event_id", "")).strip()
            if event_id not in event_names:
                skipped_event_evidences += 1
                continue
            origin_id = str(item.get("evidence_id", "")).strip()
            source_id = str(item.get("source_id", "")).strip()
            rows.append(
                {
                    "evidence_id": f"FE-{origin_id}",
                    "subject_type": "event",
                    "subject_id": event_id,
                    "predicate": "event_occurrence",
                    "object_value": str(event_names[event_id]).strip(),
                    "source_id": source_id,
                    "locator": str(item.get("source_loc", "")).strip(),
                    "quote": str(item.get("quote", "")).strip(),
                    "evidence_support": "support",
                    "source_level": source_levels.get(source_id, "D"),
                    "review_status": "machine_extracted",
                    "reviewer_note": f"原事件证据置信度：{item.get('confidence', '')}",
                    "origin_evidence_id": origin_id,
                }
            )
            event_evidence_count += 1

    facts = pd.DataFrame(rows, columns=FACT_EVIDENCE_COLUMNS)
    facts = facts.drop_duplicates(subset=["evidence_id"], keep="first").sort_values("evidence_id")
    facts.to_csv(data_dir / "fact_evidences.csv", index=False, encoding="utf-8-sig")
    return {
        "membership_evidences": len(memberships),
        "event_evidences": event_evidence_count,
        "skipped_event_evidences": skipped_event_evidences,
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="迁移现有领域证据，生成通用事实证据表。")
    parser.add_argument("--data-dir", type=Path, default=DEFAULT_DATA_DIR)
    parser.add_argument("--event-evidences", type=Path, default=DEFAULT_EVENT_EVIDENCES)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    print(build_fact_evidences(args.data_dir, args.event_evidences))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
