from __future__ import annotations

import csv
import re
from collections import Counter, defaultdict
from collections.abc import Iterable
from dataclasses import dataclass
from pathlib import Path

ORG_ID = "ORG-001"
RELATED_ROLES = {"外围联络人", "相关人士"}
MEMBERSHIP_TYPES = {"confirmed_member", "related_person", "candidate", "disputed"}

EVIDENCE_COLUMNS = [
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
]

MEMBERSHIP_COLUMNS = [
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
]

MEMBERSHIP_PHRASES = (
    "加入左联",
    "加入中国左翼作家联盟",
    "左联盟员",
    "左联成员",
    "左联常委",
    "左联常务委员",
    "左联执行委员",
    "左联书记",
    "担任左联",
    "当选左联",
)
OPPOSE_PHRASES = ("没有加入左联", "未加入左联", "不是左联盟员", "非左联盟员")
ASSOCIATION_PHRASES = ("参加左联活动", "与左联", "左联作家站在同一阵线")


@dataclass(frozen=True)
class MembershipDecision:
    membership_type: str
    membership_role: str
    evidence_status: str
    decision_rule: str
    confidence: str
    needs_manual_review: str


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open(encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def write_csv(path: Path, rows: list[dict[str, str]], columns: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=columns, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def split_ids(value: str) -> list[str]:
    return [item.strip() for item in re.split(r"[;,，；]+", value or "") if item.strip()]


def join_unique(values: Iterable[str]) -> str:
    return ";".join(dict.fromkeys(value for value in values if value))


def classify_source_level(source: dict[str, str]) -> str:
    if source.get("evidence_strength") == "一手":
        return "A"
    if source.get("evidence_type") == "研究论著":
        return "B"
    if source.get("source_kind") in {"web_encyclopedia", "raw_workbook"}:
        return "D"
    return "C"


def classify_membership(
    evidences: list[dict[str, str]],
    fallback_role: str,
) -> MembershipDecision:
    support = [row for row in evidences if row.get("evidence_support") == "membership"]
    oppose = [row for row in evidences if row.get("evidence_support") == "oppose"]

    if support and oppose:
        return MembershipDecision(
            membership_type="disputed",
            membership_role="存在争议",
            evidence_status="conflicting_evidence",
            decision_rule="support_and_oppose_evidence",
            confidence="low",
            needs_manual_review="yes",
        )

    if any(row.get("source_level") == "A" for row in support):
        return MembershipDecision(
            membership_type="confirmed_member",
            membership_role="正式成员",
            evidence_status="evidence_confirmed",
            decision_rule="one_a_level_source",
            confidence="high",
            needs_manual_review="no",
        )

    independent_b = {
        row.get("source_work", "").strip()
        for row in support
        if row.get("source_level") == "B" and row.get("source_work", "").strip()
    }
    if len(independent_b) >= 2:
        return MembershipDecision(
            membership_type="confirmed_member",
            membership_role="正式成员",
            evidence_status="evidence_confirmed",
            decision_rule="two_independent_b_sources",
            confidence="high",
            needs_manual_review="no",
        )

    if support:
        return MembershipDecision(
            membership_type="candidate",
            membership_role="成员身份待核",
            evidence_status="insufficient_member_evidence",
            decision_rule="member_evidence_below_threshold",
            confidence="medium",
            needs_manual_review="yes",
        )

    if fallback_role in RELATED_ROLES or any(
        row.get("evidence_support") == "association" for row in evidences
    ):
        return MembershipDecision(
            membership_type="related_person",
            membership_role="相关人士",
            evidence_status="association_only",
            decision_rule="association_without_member_evidence",
            confidence="medium",
            needs_manual_review="yes",
        )

    return MembershipDecision(
        membership_type="candidate",
        membership_role="成员身份待核",
        evidence_status="lead_only",
        decision_rule="raw_lead_without_member_evidence",
        confidence="low",
        needs_manual_review="yes",
    )


def normalize_ocr(value: str) -> str:
    return re.sub(r"\s+", "", value or "")


def iter_pages(source_text: str) -> Iterable[tuple[int, str]]:
    markers = list(re.finditer(r"第\s*(\d+)\s*页", source_text))
    for index, marker in enumerate(markers):
        start = marker.end()
        end = markers[index + 1].start() if index + 1 < len(markers) else len(source_text)
        yield int(marker.group(1)), source_text[start:end]


def evidence_support_for_person(sentence: str, name: str) -> str:
    escaped_name = re.escape(name)
    direct_oppose = (
        rf"{escaped_name}.{{0,16}}(?:没有|未).{{0,8}}加入(?:中国左翼作家联盟|左联)",
        rf"{escaped_name}.{{0,12}}(?:不是|非)左联盟员",
        rf"{escaped_name}.{{0,250}}等(?:虽)?未加入(?:中国左翼作家联盟|左联)",
    )
    if any(re.search(pattern, sentence) for pattern in direct_oppose):
        return "oppose"

    direct_membership = (
        rf"{escaped_name}.{{0,24}}(?:加入|参加)(?:中国左翼作家联盟|左联)",
        rf"{escaped_name}.{{0,24}}(?:担任|任|当选).{{0,12}}左联.{{0,12}}(?:常委|书记|委员|干事)",
        rf"{escaped_name}.{{0,4}}(?:是|为|系).{{0,2}}(?:左联盟员|左联成员|左联常委|左联常务委员|左联执行委员|左联书记)",
        rf"(?:担任过|担任|任|当选).{{0,12}}左联.{{0,12}}(?:常委|书记|委员|干事).{{0,24}}{escaped_name}",
        rf"(?:左联盟员|左联成员|中国左翼作家联盟成员)[、，,]?(?:(?:作家|诗人|美术家)[、，,]?)?{escaped_name}",
        rf"{escaped_name}.{{0,250}}等(?:左联盟员|左联成员|中国左翼作家联盟成员)(?:作家|诗人|人士)?",
        rf"(?:左联盟员|左联成员|中国左翼作家联盟成员)总数.{{0,50}}主要有[:：]?.{{0,600}}{escaped_name}",
        rf"(?:其中有|名单有|主要有|包括)(?:左联盟员|左联成员|中国左翼作家联盟成员).{{0,300}}{escaped_name}",
    )
    if any(re.search(pattern, sentence) for pattern in direct_membership):
        return "membership"

    direct_association = (
        rf"{escaped_name}.{{0,30}}(?:参加左联活动|与左联|和左联)",
        rf"(?:参加左联活动|与左联|和左联).{{0,30}}{escaped_name}",
    )
    if any(re.search(pattern, sentence) for pattern in direct_association):
        return "association"
    return ""


def extract_person_evidence(
    source_text: str,
    people: list[dict[str, str]],
    source_work: str,
    page_source_ids: dict[int, str],
    source_levels: dict[str, str] | None = None,
    default_source_id: str = "",
) -> list[dict[str, str]]:
    source_levels = source_levels or {}
    extracted: list[dict[str, str]] = []
    seen: set[tuple[str, str, int, str]] = set()

    for page_number, page_text in iter_pages(source_text):
        normalized = normalize_ocr(page_text)
        source_id = page_source_ids.get(page_number, default_source_id)
        if not source_id:
            continue
        sentences = [item for item in re.split(r"[。！？!?]", normalized) if item]
        for sentence in sentences:
            for person in people:
                names = [person.get("standard_name", ""), *split_ids(person.get("aliases", ""))]
                for name in [item for item in names if item and item in sentence]:
                    support = evidence_support_for_person(sentence, name)
                    key = (person["person_id"], source_id, page_number, support)
                    if support and key not in seen:
                        seen.add(key)
                        extracted.append(
                            {
                                "organization_id": ORG_ID,
                                "person_id": person["person_id"],
                                "evidence_support": support,
                                "source_id": source_id,
                                "source_work": source_work,
                                "source_level": source_levels.get(source_id, "B"),
                                "locator": f"第{page_number}页",
                                "quote": sentence[:500],
                                "review_status": "auto_extracted",
                                "reviewer_note": "由明确身份表述自动提取，建议人工抽查。",
                                "extraction_method": "ocr_explicit_phrase",
                            }
                        )
    return extracted


def build_page_source_maps(
    sources: list[dict[str, str]],
) -> tuple[dict[str, dict[int, str]], dict[str, str], dict[str, str]]:
    page_maps: dict[str, dict[int, str]] = defaultdict(dict)
    fallback_ids: dict[str, str] = {}
    source_levels: dict[str, str] = {}
    for source in sources:
        source_id = source.get("source_id", "")
        source_levels[source_id] = classify_source_level(source)
        title = source.get("title", "").strip()
        if not title:
            continue
        fallback_ids.setdefault(title, source_id)
        match = re.search(r"第\s*(\d+)\s*页", source.get("citation", ""))
        if match:
            page_maps[title][int(match.group(1))] = source_id
    return dict(page_maps), fallback_ids, source_levels


def baseline_evidence(
    memberships: list[dict[str, str]],
    people_by_id: dict[str, dict[str, str]],
    sources_by_id: dict[str, dict[str, str]],
) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for membership in memberships:
        if membership.get("organization_id") != ORG_ID:
            continue
        person_id = membership.get("person_id", "")
        person = people_by_id.get(person_id, {})
        source_id = split_ids(membership.get("source_ids", ""))[0] if membership.get("source_ids") else ""
        source = sources_by_id.get(source_id, {})
        role = person.get("role", "")
        rows.append(
            {
                "organization_id": ORG_ID,
                "person_id": person_id,
                "evidence_support": "lead",
                "source_id": source_id,
                "source_work": source.get("title", "原始表格"),
                "source_level": classify_source_level(source) if source else "D",
                "locator": "persons.role",
                "quote": f"原始人物表将该人物标注为“{role}”。",
                "review_status": "pending",
                "reviewer_note": "人物角色不能单独证明正式成员身份。",
                "extraction_method": "legacy_membership_lead",
            }
        )
    return rows


def rebuild_org_memberships(
    data_dir: Path,
    runtime_sources_dir: Path,
) -> dict[str, object]:
    people = read_csv(data_dir / "persons.csv")
    memberships = read_csv(data_dir / "org_memberships.csv")
    sources = read_csv(data_dir / "sources.csv")
    people_by_id = {row["person_id"]: row for row in people}
    sources_by_id = {row["source_id"]: row for row in sources}

    evidence_rows = baseline_evidence(memberships, people_by_id, sources_by_id)
    page_maps, fallback_ids, source_levels = build_page_source_maps(sources)

    for source_path in sorted(runtime_sources_dir.glob("*.txt")):
        source_work = source_path.stem
        if source_work not in page_maps and source_work not in fallback_ids:
            continue
        evidence_rows.extend(
            extract_person_evidence(
                source_text=source_path.read_text(encoding="utf-8"),
                people=people,
                source_work=source_work,
                page_source_ids=page_maps.get(source_work, {}),
                source_levels=source_levels,
                default_source_id=fallback_ids.get(source_work, ""),
            )
        )

    for index, row in enumerate(evidence_rows, start=1):
        row["evidence_id"] = f"OME-{index:05d}"

    evidence_by_person: dict[str, list[dict[str, str]]] = defaultdict(list)
    for row in evidence_rows:
        evidence_by_person[row["person_id"]].append(row)

    conclusion_rows: list[dict[str, str]] = []
    scoped_memberships = [row for row in memberships if row.get("organization_id") == ORG_ID]
    for index, membership in enumerate(scoped_memberships, start=1):
        person_id = membership["person_id"]
        person_evidence = evidence_by_person.get(person_id, [])
        role = people_by_id.get(person_id, {}).get("role", "")
        decision = classify_membership(person_evidence, fallback_role=role)
        conclusion_rows.append(
            {
                "membership_id": f"MEM-{index:05d}",
                "organization_id": ORG_ID,
                "person_id": person_id,
                "membership_role": decision.membership_role,
                "membership_type": decision.membership_type,
                "source_ids": join_unique(row["source_id"] for row in person_evidence),
                "evidence_ids": join_unique(row["evidence_id"] for row in person_evidence),
                "evidence_status": decision.evidence_status,
                "evidence_count": str(len(person_evidence)),
                "decision_rule": decision.decision_rule,
                "confidence": decision.confidence,
                "needs_manual_review": decision.needs_manual_review,
            }
        )

    write_csv(data_dir / "org_membership_evidences.csv", evidence_rows, EVIDENCE_COLUMNS)
    write_csv(data_dir / "org_memberships.csv", conclusion_rows, MEMBERSHIP_COLUMNS)

    return {
        "membership_count": len(conclusion_rows),
        "evidence_count": len(evidence_rows),
        "membership_types": dict(Counter(row["membership_type"] for row in conclusion_rows)),
        "source_levels": dict(Counter(row["source_level"] for row in evidence_rows)),
    }


def main() -> None:
    project_root = Path(__file__).resolve().parents[2]
    summary = rebuild_org_memberships(
        data_dir=project_root / "data" / "processed",
        runtime_sources_dir=project_root / "data" / "processed" / "runtime_sources",
    )
    print(summary)


if __name__ == "__main__":
    main()
