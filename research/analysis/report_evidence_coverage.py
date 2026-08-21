from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_DATA_DIR = PROJECT_ROOT / "data" / "processed"
DEFAULT_REPORT = PROJECT_ROOT / "research" / "drafts" / "reports" / "phase2_evidence_coverage_report.md"
DEFAULT_QUEUE = PROJECT_ROOT / "research" / "drafts" / "reports" / "phase2_core_fact_review_queue.csv"


def _read(path: Path) -> pd.DataFrame:
    return pd.read_csv(path, encoding="utf-8-sig").fillna("")


def _coverage(covered: int, total: int) -> dict[str, int | float]:
    return {
        "covered": covered,
        "total": total,
        "rate": round(covered / total, 4) if total else 0.0,
    }


def report_evidence_coverage(data_dir: Path, report_path: Path, queue_path: Path) -> dict[str, dict[str, int | float]]:
    data_dir = Path(data_dir)
    facts = _read(data_dir / "fact_evidences.csv")
    memberships = _read(data_dir / "org_memberships.csv")
    events = _read(data_dir / "events.csv")
    persons = _read(data_dir / "persons.csv")

    membership_keys = {
        (str(row["subject_id"]).strip(), str(row["object_value"]).strip())
        for _, row in facts[facts["predicate"] == "organization_membership"].iterrows()
    }
    covered_memberships = sum(
        (str(row["person_id"]).strip(), str(row["organization_id"]).strip()) in membership_keys
        for _, row in memberships.iterrows()
    )
    event_ids = set(facts.loc[facts["predicate"] == "event_occurrence", "subject_id"].astype(str))
    covered_events = sum(str(event_id) in event_ids for event_id in events["event_id"])

    person_fact_subjects = {
        predicate: set(facts.loc[facts["predicate"] == predicate, "subject_id"].astype(str))
        for predicate in ("birth_year", "death_year", "role")
    }
    summary = {
        "memberships": _coverage(covered_memberships, len(memberships)),
        "events": _coverage(covered_events, len(events)),
        "person_birth_year": _coverage(
            sum(str(person_id) in person_fact_subjects["birth_year"] for person_id in persons["person_id"]),
            len(persons),
        ),
        "person_death_year": _coverage(
            sum(str(person_id) in person_fact_subjects["death_year"] for person_id in persons["person_id"]),
            len(persons),
        ),
        "person_role": _coverage(
            sum(str(person_id) in person_fact_subjects["role"] for person_id in persons["person_id"]),
            len(persons),
        ),
    }

    queue: list[dict[str, str]] = []
    for _, row in memberships.iterrows():
        key = (str(row["person_id"]).strip(), str(row["organization_id"]).strip())
        if key not in membership_keys:
            queue.append(
                {
                    "subject_type": "person",
                    "subject_id": key[0],
                    "predicate": "organization_membership",
                    "object_value": key[1],
                    "reason": "缺少事实级组织身份依据",
                }
            )
    for _, row in events.iterrows():
        event_id = str(row["event_id"]).strip()
        if event_id not in event_ids:
            queue.append(
                {
                    "subject_type": "event",
                    "subject_id": event_id,
                    "predicate": "event_occurrence",
                    "object_value": str(row["event_name"]).strip(),
                    "reason": "缺少可定位事件证据",
                }
            )
    for _, row in persons.iterrows():
        person_id = str(row["person_id"]).strip()
        for predicate, column in (("birth_year", "birth_year"), ("death_year", "death_year"), ("role", "role")):
            if person_id not in person_fact_subjects[predicate] and str(row[column]).strip():
                queue.append(
                    {
                        "subject_type": "person",
                        "subject_id": person_id,
                        "predicate": predicate,
                        "object_value": str(row[column]).strip(),
                        "reason": "实体级来源不能替代事实级证据",
                    }
                )

    queue_path.parent.mkdir(parents=True, exist_ok=True)
    pd.DataFrame(
        queue,
        columns=["subject_type", "subject_id", "predicate", "object_value", "reason"],
    ).to_csv(queue_path, index=False, encoding="utf-8-sig")

    lines = [
        "# Phase 2 事实级证据覆盖率报告",
        "",
        "本报告只统计带有具体来源、定位和摘录的事实证据。实体表中的 `source_ids` 不计为事实级覆盖。",
        "",
        "| 事实类型 | 已覆盖 | 总数 | 覆盖率 |",
        "| --- | ---: | ---: | ---: |",
    ]
    labels = {
        "memberships": "组织身份",
        "events": "事件存在",
        "person_birth_year": "人物出生年",
        "person_death_year": "人物逝世年",
        "person_role": "人物角色",
    }
    for key, label in labels.items():
        item = summary[key]
        lines.append(f"| {label} | {item['covered']} | {item['total']} | {item['rate']:.1%} |")
    lines.extend(
        [
            "",
            f"- 通用事实证据总数：{len(facts)}",
            f"- 待核核心事实数：{len(queue)}",
            f"- 待核队列：`{queue_path.name}`",
            "",
            "## 结论",
            "",
            "当前组织身份已完成事实级证据迁移；事件仅覆盖已有可定位证据的记录；人物生卒年和角色仍需后续补充具体引文。",
        ]
    )
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return summary


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="统计核心事实的事实级证据覆盖率。")
    parser.add_argument("--data-dir", type=Path, default=DEFAULT_DATA_DIR)
    parser.add_argument("--report", type=Path, default=DEFAULT_REPORT)
    parser.add_argument("--queue", type=Path, default=DEFAULT_QUEUE)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    print(report_evidence_coverage(args.data_dir, args.report, args.queue))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
