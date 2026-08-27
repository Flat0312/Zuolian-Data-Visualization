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
    # 三种事件口径，语义互不冒充：
    # 1) 已挂接：存在任一条 event_occurrence 事实证据；
    # 2) 直接支持：至少一条 evidence_support=support 的证据（不论复核/裁决状态）；
    # 3) 已确认：至少一条 reviewed 且（support 或 adjudication_status=resolved_by_event_correction）
    #    ——原始冲突不计入，仅“已由事件级裁决落地”的 conflict 计入。
    event_facts_raw = facts[facts["predicate"] == "event_occurrence"]
    # 三口径均排除 rejected：被拒绝的证据不计入任何覆盖（第四批A裁决）。
    event_facts = event_facts_raw[
        event_facts_raw["review_status"].astype(str).str.strip() != "rejected"
    ]
    adjudication_status = (
        event_facts["adjudication_status"].astype(str).str.strip()
        if "adjudication_status" in event_facts.columns
        else pd.Series("", index=event_facts.index)
    )
    confirmed_frame = event_facts[
        (event_facts["review_status"] == "reviewed")
        & (
            (event_facts["evidence_support"] == "support")
            | (adjudication_status == "resolved_by_event_correction")
        )
    ]
    direct_support_events = set(
        event_facts.loc[event_facts["evidence_support"] == "support", "subject_id"].astype(str).str.strip()
    )
    confirmed_events = set(confirmed_frame["subject_id"].astype(str).str.strip())
    # 已挂接口径同样排除 rejected（第四批A裁决）：仅“非拒绝”的证据构成挂接事实。
    qualifying_event_subjects = set(event_facts["subject_id"].astype(str).str.strip())
    covered_events = sum(
        str(event_id) in qualifying_event_subjects for event_id in events["event_id"]
    )

    total_events = len(events)
    metrics_events = {
        "event_attached_any": _coverage(
            sum(str(event_id) in qualifying_event_subjects for event_id in events["event_id"]),
            total_events,
        ),
        "event_direct_support": _coverage(
            sum(str(event_id) in direct_support_events for event_id in events["event_id"]),
            total_events,
        ),
        "event_confirmed": _coverage(
            sum(str(event_id) in confirmed_events for event_id in events["event_id"]),
            total_events,
        ),
    }

    person_fact_subjects = {
        predicate: set(facts.loc[facts["predicate"] == predicate, "subject_id"].astype(str))
        for predicate in ("birth_year", "death_year", "role")
    }
    summary = {
        "memberships": _coverage(covered_memberships, len(memberships)),
        "events": _coverage(covered_events, len(events)),
        **metrics_events,
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
        if event_id not in qualifying_event_subjects:
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
        "events": "事件存在（口径一：已挂接事实证据）",
        "event_attached_any": "事件存在·已挂接事实证据覆盖（存在任一事实证据即计入）",
        "event_direct_support": "事件存在·直接支持覆盖（evidence_support=support，不论裁决与复核状态）",
        "event_confirmed": (
            "事件存在·已确认覆盖（reviewed 且为 support 或已裁决 conflict"
            "［adjudication_status=resolved_by_event_correction］）"
        ),
        "person_birth_year": "人物出生年",
        "person_death_year": "人物逝世年",
        "person_role": "人物角色",
    }
    ordered_keys = [
        "memberships",
        "event_attached_any",
        "event_direct_support",
        "event_confirmed",
        "person_birth_year",
        "person_death_year",
        "person_role",
    ]
    for key in ordered_keys:
        item = summary[key]
        lines.append(f"| {labels[key]} | {item['covered']} | {item['total']} | {item['rate']:.1%} |")
    attached = metrics_events["event_attached_any"]
    direct = metrics_events["event_direct_support"]
    confirmed = metrics_events["event_confirmed"]
    gap_note = ""
    if confirmed["covered"] < direct["covered"]:
        gap_note = (
            f"已确认覆盖与直接支持覆盖的差额 {direct['covered'] - confirmed['covered']} 个事件，"
            "由仅 machine_extracted 支撑或仅 lead 类证据挂接、尚未经复核/补证的记录构成；"
            "详见 BLOCKED.md B-2。"
        )
    lines.extend(
        [
            "",
            "三种事件覆盖口径含义互不替代：**已挂接≠直接支持≠已确认**——",
            "- 口径一「已挂接」只说明该事件至少被一条事实证据提及，可能是线索（lead）、待裁冲突（conflict）或未复核抽取；",
            "- 口径二「直接支持」要求 evidence_support=support 的支持性证据，但不要求人工复核完成；",
            "- 口径三「已确认」要求复核完成且为支持性证据，或 conflict 已按事件级审核决策落地（结构化状态 resolved_by_event_correction）。",
            "",
            f"- 通用事实证据总数：{len(facts)}",
            f"- 待核核心事实数：{len(queue)}",
            f"- 待核队列：`{queue_path.name}`",
            f"- 三口径速览：已挂接 {attached['covered']}/{attached['total']}；"
            f"直接支持 {direct['covered']}/{direct['total']}；"
            f"已确认 {confirmed['covered']}/{confirmed['total']}。",
        ]
    )
    if gap_note:
        lines.append(f"- {gap_note}")
    lines.extend(
        [
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
