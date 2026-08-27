"""第四批A五项人工批准裁决的幂等执行脚本（生产层落地）。

授权依据：用户原话「5项全部按推荐方案执行」（2026-08-27）。脚本只处理
第四批A审计表列出的 14 条证据与 5 个事件，不搜索、不新增来源、不创建事件。

安全边界：
- 以 30d5f77 的目标行字段和四表计数做前置校验，异常时不写任何文件；
- 所有证据按 evidence_id、事件按 event_id 定点操作；
- reviewer_note 只追加本批裁决出处，quote、locator、source_id 不改；
- 全部后置条件成立时二跑输出“无新增/已完成”并跳过写入。
"""

from __future__ import annotations

import csv
import hashlib
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
DATA = ROOT / "data" / "processed"

PROVENANCE = "2026-08-27 第四批A人工批准（原话「5项全部按推荐方案执行」）落地"

EXPECTED_COUNTS = {
    "sources.csv": (1177, 1177),
    "fact_evidences.csv": (628, 626),
    "events.csv": (148, 147),
    "event_participants.csv": (224, 222),
}

BASELINE_EVENT_STATES = {
    "EVT-00004": {
        "event_name": "鲁迅与柔石会面",
        "canonical_event_key": "鲁迅与柔石会面|ZLH-001",
        "event_date": "1930-03-28",
        "date_precision": "日",
        "source_ids": "SRC-1122;SRC-1143;SRC-1144",
        "confidence": "medium",
        "needs_manual_review": "yes",
    },
    "EVT-00005": {
        "event_name": "鲁迅与柔石会面",
        "canonical_event_key": "鲁迅与柔石会面|ZLH-016",
        "event_date": "1930-03-28",
        "date_precision": "日",
        "source_ids": "SRC-1122;SRC-1143;SRC-1144",
        "confidence": "medium",
        "needs_manual_review": "yes",
    },
    "EVT-00017": {
        "event_name": "鲁迅与内山完造通信",
        "canonical_event_key": "鲁迅与内山完造通信|ZLH-001",
        "event_date": "1929-09-23",
        "date_precision": "日",
        "source_ids": "SRC-0001;SRC-1143;SRC-1144",
        "confidence": "low",
        "needs_manual_review": "yes",
    },
    "EVT-00029": {
        "event_name": "丁玲抵沪就读平民女校",
        "canonical_event_key": "ZLH-021|丁玲抵沪就读平民女校|1922",
        "event_date": "1922",
        "date_precision": "年",
        "source_ids": "SRC-1126;SRC-1154",
        "confidence": "medium",
        "needs_manual_review": "no",
    },
}

POST_EVENT_STATES = {
    "EVT-00004": {
        "event_name": "鲁迅与柔石赴北四川路看屋未成",
        "canonical_event_key": "EVT-00004|鲁迅与柔石赴北四川路看屋未成|1930-03-28",
        "event_date": "1930-03-28",
        "date_precision": "日",
        "source_ids": "SRC-1122;SRC-1143;SRC-1144;SRC-1167",
        "needs_manual_review": "no",
        "display_note": "1930年3月28日，鲁迅与柔石同赴北四川路一带看屋未成；3月30日到寓会面另列为EVT-00006。",
        "correction_reason_append": (
            "2026-08-27 第四批A裁决：按日记将本条收窄为3月28日北四川路看屋未成，"
            "3月30日到寓会面另列为EVT-00006；FE-EVP3-0005改指本条并转为支持。"
        ),
    },
    "EVT-00029": {
        "event_name": "丁玲就读平民女校",
        "canonical_event_key": "EVT-00029|丁玲就读平民女校|1922",
        "event_date": "1922",
        "date_precision": "年",
        "source_ids": "SRC-1126;SRC-1154",
        "needs_manual_review": "no",
        "display_note": "1922年丁玲在平民女校就读；当前条目按年份精度展示，具体入学月份待补证。",
        "correction_reason_append": (
            "2026-08-27 第四批A裁决：按现有证据将事件收窄为“丁玲就读平民女校”，"
            "当前只保留就读与年份事实。"
        ),
    },
}

# evidence_id -> 执行动作。所有目标行的基线状态另由 BASELINE_EVIDENCE_STATES 锚定。
EVIDENCE_ACTIONS = (
    ("FE-EVI-3B84F7AC63", {"kind": "promote_reviewed", "audit": "AUD-B4A-003"}),
    ("FE-EVI-489692805D", {"kind": "promote_reviewed", "audit": "AUD-B4A-004"}),
    ("FE-EVI-58338B3D55", {"kind": "promote_reviewed", "audit": "AUD-B4A-005"}),
    ("FE-EVI-8DB5BD3637", {"kind": "promote_reviewed", "audit": "AUD-B4A-007"}),
    ("FE-EVI-FD8EBDD93E", {"kind": "promote_reviewed", "audit": "AUD-B4A-008"}),
    ("FE-EVI-21007B9057", {"kind": "reject", "audit": "AUD-B4A-002"}),
    ("FE-EVI-6883382B0E", {"kind": "reject", "audit": "AUD-B4A-006"}),
    ("FE-EVI-0528D7CD44", {"kind": "demote_lead", "audit": "AUD-B4A-001"}),
    ("FE-EVI-0357C10A69", {"kind": "reject", "audit": "AUD-B4A-009"}),
    (
        "FE-EVI-7792AD0E80",
        {
            "kind": "remap_support",
            "audit": "AUD-B4A-009",
            "new_subject": "EVT-00004",
            "clear_adjudication": True,
        },
    ),
    ("FE-EVI-04DD852F1C", {"kind": "delete", "audit": "AUD-B4A-010"}),
    ("FE-EVI-43ECB964FE", {"kind": "delete", "audit": "AUD-B4A-011"}),
    (
        "FE-EVI-D016A6A994",
        {
            "kind": "remap_object",
            "audit": "AUD-B4A-012",
            "new_subject": "EVT-00008",
            "new_object_value": (
                "柔石等左联五烈士于1931年2月7日夜或2月8日凌晨在上海龙华警备司令部遇害，"
                "鲁迅约于2月10日获悉"
            ),
        },
    ),
    ("FE-EVI-2EC5596E2B", {"kind": "reject", "audit": "AUD-B4A-013"}),
    ("FE-EVI-CDAFBAFA48", {"kind": "promote_reviewed_from_lead", "audit": "AUD-B4A-014"}),
)

BASELINE_EVIDENCE_STATES = {
    "FE-EVI-3B84F7AC63": {
        "subject_type": "event", "subject_id": "EVT-00001", "predicate": "event_occurrence",
        "object_value": "中国左翼作家联盟成立大会", "source_id": "SRC-0683", "locator": "第256页",
        "evidence_support": "support", "source_level": "B", "review_status": "machine_extracted",
        "origin_evidence_id": "EVI-3B84F7AC63", "adjudication_status": "",
    },
    "FE-EVI-489692805D": {
        "subject_type": "event", "subject_id": "EVT-00001", "predicate": "event_occurrence",
        "object_value": "中国左翼作家联盟成立大会", "source_id": "SRC-0125", "locator": "第175页",
        "evidence_support": "support", "source_level": "B", "review_status": "machine_extracted",
        "origin_evidence_id": "EVI-489692805D", "adjudication_status": "",
    },
    "FE-EVI-58338B3D55": {
        "subject_type": "event", "subject_id": "EVT-00001", "predicate": "event_occurrence",
        "object_value": "中国左翼作家联盟成立大会", "source_id": "SRC-0797", "locator": "第507页",
        "evidence_support": "support", "source_level": "B", "review_status": "machine_extracted",
        "origin_evidence_id": "EVI-58338B3D55", "adjudication_status": "",
    },
    "FE-EVI-8DB5BD3637": {
        "subject_type": "event", "subject_id": "EVT-00001", "predicate": "event_occurrence",
        "object_value": "中国左翼作家联盟成立大会", "source_id": "SRC-0014", "locator": "第128页",
        "evidence_support": "support", "source_level": "B", "review_status": "machine_extracted",
        "origin_evidence_id": "EVI-8DB5BD3637", "adjudication_status": "",
    },
    "FE-EVI-FD8EBDD93E": {
        "subject_type": "event", "subject_id": "EVT-00001", "predicate": "event_occurrence",
        "object_value": "中国左翼作家联盟成立大会", "source_id": "SRC-0892", "locator": "第276页",
        "evidence_support": "support", "source_level": "B", "review_status": "machine_extracted",
        "origin_evidence_id": "EVI-FD8EBDD93E", "adjudication_status": "",
    },
    "FE-EVI-21007B9057": {
        "subject_type": "event", "subject_id": "EVT-00001", "predicate": "event_occurrence",
        "object_value": "中国左翼作家联盟成立大会", "source_id": "SRC-0827", "locator": "第45页",
        "evidence_support": "support", "source_level": "B", "review_status": "machine_extracted",
        "origin_evidence_id": "EVI-21007B9057", "adjudication_status": "",
    },
    "FE-EVI-6883382B0E": {
        "subject_type": "event", "subject_id": "EVT-00001", "predicate": "event_occurrence",
        "object_value": "中国左翼作家联盟成立大会", "source_id": "SRC-0538", "locator": "第354页",
        "evidence_support": "support", "source_level": "B", "review_status": "machine_extracted",
        "origin_evidence_id": "EVI-6883382B0E", "adjudication_status": "",
    },
    "FE-EVI-0528D7CD44": {
        "subject_type": "event", "subject_id": "EVT-00001", "predicate": "event_occurrence",
        "object_value": "中国左翼作家联盟成立大会", "source_id": "SRC-0283", "locator": "第97页",
        "evidence_support": "support", "source_level": "B", "review_status": "machine_extracted",
        "origin_evidence_id": "EVI-0528D7CD44", "adjudication_status": "",
    },
    "FE-EVI-0357C10A69": {
        "subject_type": "event", "subject_id": "EVT-00004", "predicate": "event_occurrence",
        "object_value": "鲁迅与柔石会面", "source_id": "SRC-0117", "locator": "1930年3月30日",
        "evidence_support": "support", "source_level": "A", "review_status": "machine_extracted",
        "origin_evidence_id": "EVI-0357C10A69", "adjudication_status": "",
    },
    "FE-EVI-7792AD0E80": {
        "subject_type": "event", "subject_id": "EVT-00006", "predicate": "event_occurrence",
        "object_value": "北四川路一带看屋实际发生在1930年3月28日而非3月30日", "source_id": "SRC-1167",
        "locator": "日记十九·三月·二十八日条", "evidence_support": "conflict", "source_level": "A",
        "review_status": "reviewed", "origin_evidence_id": "FE-EVP3-0005",
        "adjudication_status": "resolved_by_event_correction",
    },
    "FE-EVI-04DD852F1C": {
        "subject_type": "event", "subject_id": "EVT-00005", "predicate": "event_occurrence",
        "object_value": "鲁迅与柔石会面", "source_id": "SRC-0740", "locator": "第106页",
        "evidence_support": "support", "source_level": "B", "review_status": "machine_extracted",
        "origin_evidence_id": "EVI-04DD852F1C", "adjudication_status": "",
    },
    "FE-EVI-43ECB964FE": {
        "subject_type": "event", "subject_id": "EVT-00005", "predicate": "event_occurrence",
        "object_value": "鲁迅与柔石会面", "source_id": "SRC-0900", "locator": "第526页",
        "evidence_support": "support", "source_level": "B", "review_status": "machine_extracted",
        "origin_evidence_id": "EVI-43ECB964FE", "adjudication_status": "",
    },
    "FE-EVI-D016A6A994": {
        "subject_type": "event", "subject_id": "EVT-00005", "predicate": "event_occurrence",
        "object_value": "鲁迅与柔石会面", "source_id": "SRC-0910", "locator": "第109页",
        "evidence_support": "support", "source_level": "B", "review_status": "machine_extracted",
        "origin_evidence_id": "EVI-D016A6A994", "adjudication_status": "",
    },
    "FE-EVI-2EC5596E2B": {
        "subject_type": "event", "subject_id": "EVT-00017", "predicate": "event_occurrence",
        "object_value": "鲁迅与内山完造通信", "source_id": "SRC-0345", "locator": "1929年2月16日",
        "evidence_support": "support", "source_level": "A", "review_status": "machine_extracted",
        "origin_evidence_id": "EVI-2EC5596E2B", "adjudication_status": "",
    },
    "FE-EVI-CDAFBAFA48": {
        "subject_type": "event", "subject_id": "EVT-00029", "predicate": "event_occurrence",
        "object_value": "丁玲曾在上海平民女校就读（1922年）", "source_id": "SRC-1154",
        "locator": "正文第2段（网页行83-85）", "evidence_support": "lead", "source_level": "B",
        "review_status": "reviewed", "origin_evidence_id": "FE-EVP-0001", "adjudication_status": "",
    },
}

BASELINE_QUOTE_SHA256 = {
    "FE-EVI-3B84F7AC63": "d4726da3d3dc3b3ac9871db5f6f23292b03fd89918d3570de4882f441b0612a1",
    "FE-EVI-489692805D": "c35e911b08477a787b3793823ff6686d84d94aa8e4b2709ed902700195aba8c5",
    "FE-EVI-58338B3D55": "5de2ae7784c1c1c35ada251f80311abc5b037167e788b0cbed6752d7d0218e7f",
    "FE-EVI-8DB5BD3637": "686c276b4e762a18331180f9004783f7634ff9dbd8dfb736fd54b56e4ec1ea42",
    "FE-EVI-FD8EBDD93E": "4b34743b0dae4b116d33543097e7c98c3601115b9bde81b5d4d4d7f14e550280",
    "FE-EVI-21007B9057": "f4b1dfed3f53cc35533cffae1ae4037620d13ccea970ed44f8c48938f0fb437f",
    "FE-EVI-6883382B0E": "68f10f3ac5f01d52eb417a92c7cd0b13c56028d23284c0902b20f166d743c63a",
    "FE-EVI-0528D7CD44": "180878c3b561228fbe84f61dec44423ab6d88b8d42604a7db4842ea147b3edd8",
    "FE-EVI-0357C10A69": "246db688aba5556521af0069dad9e837017c00b78451910c70682ac12258ab76",
    "FE-EVI-7792AD0E80": "a6a2d2ff82a7fb863026445fe0bca2032d815edca1af14c49b47369bdb542cf2",
    "FE-EVI-04DD852F1C": "d6ee7658d72d03d43890ed8be55ae9e4570f786c1b8bc5c3d98a60d1e431c678",
    "FE-EVI-43ECB964FE": "8248c64acbeb77d072d7f98625295998a2ffc74e1b89a88dbebe2a1c6eebd7d1",
    "FE-EVI-D016A6A994": "70debe3e1fea0224a7e862007c0f854cd44ccbc5020b93c7c6999a29144cc075",
    "FE-EVI-2EC5596E2B": "1b75d01bee3e5504cbc373a8511d923e00cc22cc23bf1fcf4f3f10797779c6c6",
    "FE-EVI-CDAFBAFA48": "32e4fce2ec7ab8365af8f51287114589d8bda0d1f65912239359d33c2e828676",
}

BASELINE_PARTICIPANTS = {
    "EVP-00006": {
        "event_id": "EVT-00004", "person_id": "ZLH-001", "participant_name": "鲁迅",
        "participant_role": "直接参与者", "source_ids": "SRC-1122", "confidence": "medium",
        "needs_manual_review": "yes",
    },
    "EVP-00007": {
        "event_id": "EVT-00004", "person_id": "ZLH-016", "participant_name": "柔石",
        "participant_role": "直接参与者", "source_ids": "SRC-1122", "confidence": "medium",
        "needs_manual_review": "yes",
    },
    "EVP-00009": {
        "event_id": "EVT-00005", "person_id": "ZLH-001", "participant_name": "鲁迅",
        "participant_role": "直接参与者", "source_ids": "SRC-1122", "confidence": "medium",
        "needs_manual_review": "yes",
    },
    "EVP-00008": {
        "event_id": "EVT-00005", "person_id": "ZLH-016", "participant_name": "柔石",
        "participant_role": "直接参与者", "source_ids": "SRC-1122", "confidence": "medium",
        "needs_manual_review": "yes",
    },
    "EVP-00043": {
        "event_id": "EVT-00029", "person_id": "ZLH-021", "participant_name": "丁玲",
        "participant_role": "直接参与者", "source_ids": "SRC-1126", "confidence": "medium",
        "needs_manual_review": "yes",
    },
}

DELETE_EVENT_ID = "EVT-00005"


def read_csv(path: Path) -> tuple[list[str], list[dict[str, str]]]:
    with open(path, encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        return list(reader.fieldnames or []), list(reader)


def write_csv(path: Path, fieldnames: list[str], rows: list[dict[str, str]]) -> None:
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def _split_ids(value: str) -> list[str]:
    return [item.strip() for item in value.replace("；", ";").split(";") if item.strip()]


def _join_ids(values: list[str]) -> str:
    seen: list[str] = []
    for value in values:
        if value and value not in seen:
            seen.append(value)
    return ";".join(seen)


def _append_text(row: dict[str, str], field: str, addition: str) -> None:
    if addition in row[field]:
        return
    base = row[field].rstrip()
    separator = "" if not base or base.endswith(("。", "！", "？")) else "。"
    row[field] = f"{base}{separator}{addition}"


def _append_reviewer_note(
    row: dict[str, str], audit: str, before_support: str, before_status: str
) -> None:
    addition = (
        f"{PROVENANCE}（{audit}）：{before_support}/{before_status} 调整为 "
        f"{row['evidence_support']}/{row['review_status']}。"
    )
    _append_text(row, "reviewer_note", addition)


def _quote_hash(row: dict[str, str]) -> str:
    return hashlib.sha256(row.get("quote", "").encode("utf-8")).hexdigest()


def _check_counts(
    sources: list[dict[str, str]],
    facts: list[dict[str, str]],
    events: list[dict[str, str]],
    participants: list[dict[str, str]],
    phase: int,
) -> list[str]:
    rows_by_name = {
        "sources.csv": sources,
        "fact_evidences.csv": facts,
        "events.csv": events,
        "event_participants.csv": participants,
    }
    problems: list[str] = []
    for name, rows in rows_by_name.items():
        expected = EXPECTED_COUNTS[name][phase]
        if len(rows) != expected:
            problems.append(f"{name}: 期望 {expected} 行，实际 {len(rows)} 行")
    return problems


def _check_baseline(
    sources: list[dict[str, str]],
    facts: list[dict[str, str]],
    events: list[dict[str, str]],
    participants: list[dict[str, str]],
) -> list[str]:
    problems = _check_counts(sources, facts, events, participants, 0)
    fact_by_id = {row.get("evidence_id"): row for row in facts}
    event_by_id = {row.get("event_id"): row for row in events}
    participant_by_id = {row.get("event_participant_id"): row for row in participants}

    for evidence_id, expected in BASELINE_EVIDENCE_STATES.items():
        row = fact_by_id.get(evidence_id)
        if row is None:
            problems.append(f"{evidence_id}: 目标证据不存在")
            continue
        for key, want in expected.items():
            if row.get(key, "") != want:
                problems.append(f"{evidence_id}.{key}={row.get(key, '')!r} 与基线 {want!r} 不符")
        if _quote_hash(row) != BASELINE_QUOTE_SHA256[evidence_id]:
            problems.append(f"{evidence_id}.quote 与 30d5f77 基线不符")
        if PROVENANCE in row.get("reviewer_note", ""):
            problems.append(f"{evidence_id}.reviewer_note 已含本批裁决标记，疑似半执行")

    for event_id, expected in BASELINE_EVENT_STATES.items():
        row = event_by_id.get(event_id)
        if row is None:
            problems.append(f"{event_id}: 目标事件不存在")
            continue
        for key, want in expected.items():
            if row.get(key, "") != want:
                problems.append(f"{event_id}.{key}={row.get(key, '')!r} 与基线 {want!r} 不符")

    for participant_id, expected in BASELINE_PARTICIPANTS.items():
        row = participant_by_id.get(participant_id)
        if row is None:
            problems.append(f"{participant_id}: 目标参与者不存在")
            continue
        for key, want in expected.items():
            if row.get(key, "") != want:
                problems.append(f"{participant_id}.{key}={row.get(key, '')!r} 与基线 {want!r} 不符")

    event_ids = set(event_by_id)
    for evidence_id, action in EVIDENCE_ACTIONS:
        new_subject = action.get("new_subject")
        if new_subject and new_subject not in event_ids:
            problems.append(f"{evidence_id}: 改挂目标 {new_subject} 不存在")
    return problems


def _expected_evidence_state(evidence_id: str, action: dict[str, str]) -> dict[str, str]:
    expected = dict(BASELINE_EVIDENCE_STATES[evidence_id])
    kind = action["kind"]
    if kind == "reject":
        expected["review_status"] = "rejected"
    elif kind == "demote_lead":
        expected["evidence_support"] = "lead"
        expected["review_status"] = "pending"
    elif kind == "promote_reviewed":
        expected["review_status"] = "reviewed"
    elif kind == "promote_reviewed_from_lead":
        expected["evidence_support"] = "support"
        expected["review_status"] = "reviewed"
    elif kind == "remap_support":
        expected["subject_id"] = action["new_subject"]
        expected["evidence_support"] = "support"
        expected["review_status"] = "reviewed"
        expected["adjudication_status"] = ""
    elif kind == "remap_object":
        expected["subject_id"] = action["new_subject"]
        expected["object_value"] = action["new_object_value"]
        expected["evidence_support"] = "support"
        expected["review_status"] = "reviewed"
    return expected


def _postconditions_satisfied(
    sources: list[dict[str, str]],
    facts: list[dict[str, str]],
    events: list[dict[str, str]],
    participants: list[dict[str, str]],
) -> tuple[bool, list[str]]:
    drift = _check_counts(sources, facts, events, participants, 1)
    fact_by_id = {row.get("evidence_id"): row for row in facts}
    event_by_id = {row.get("event_id"): row for row in events}
    participant_by_id = {row.get("event_participant_id"): row for row in participants}

    for evidence_id, action in EVIDENCE_ACTIONS:
        row = fact_by_id.get(evidence_id)
        if action["kind"] == "delete":
            if row is not None:
                drift.append(f"{evidence_id}: 删除动作的事实行仍存在")
            continue
        if row is None:
            drift.append(f"{evidence_id}: 执行后目标事实缺失")
            continue
        expected = _expected_evidence_state(evidence_id, action)
        for key, want in expected.items():
            if row.get(key, "") != want:
                drift.append(f"{evidence_id}.{key}={row.get(key, '')!r} 与终值 {want!r} 不符")
        if _quote_hash(row) != BASELINE_QUOTE_SHA256[evidence_id]:
            drift.append(f"{evidence_id}.quote 被改变")
        if f"（{action['audit']}）" not in row.get("reviewer_note", ""):
            drift.append(f"{evidence_id}.reviewer_note 缺少 {action['audit']} 裁决出处")

    for event_id, expected in POST_EVENT_STATES.items():
        row = event_by_id.get(event_id)
        if row is None:
            drift.append(f"{event_id}: 执行后事件缺失")
            continue
        for key in ("event_name", "canonical_event_key", "event_date", "date_precision", "source_ids", "needs_manual_review"):
            if row.get(key, "") != expected[key]:
                drift.append(f"{event_id}.{key}={row.get(key, '')!r} 与终值 {expected[key]!r} 不符")
        if expected["display_note"] not in row.get("display_note", ""):
            drift.append(f"{event_id}.display_note 未落入本批终值")
        if expected["correction_reason_append"] not in row.get("correction_reason", ""):
            drift.append(f"{event_id}.correction_reason 缺少本批裁决出处")

    if DELETE_EVENT_ID in event_by_id:
        drift.append(f"{DELETE_EVENT_ID}: 删除事件仍存在")

    expected_participant_sources = {
        "EVP-00006": "SRC-1122;SRC-1167",
        "EVP-00007": "SRC-1122;SRC-1167",
        "EVP-00043": "SRC-1126;SRC-1154",
    }
    for participant_id, expected in BASELINE_PARTICIPANTS.items():
        row = participant_by_id.get(participant_id)
        if expected["event_id"] == DELETE_EVENT_ID:
            if row is not None:
                drift.append(f"{participant_id}: 删除事件参与者仍存在")
            continue
        if row is None:
            drift.append(f"{participant_id}: 执行后参与者缺失")
            continue
        for key, want in expected.items():
            if key == "source_ids":
                want = expected_participant_sources[participant_id]
            if row.get(key, "") != want:
                drift.append(f"{participant_id}.{key}={row.get(key, '')!r} 与终值 {want!r} 不符")

    event_ids = set(event_by_id)
    for row in facts:
        if row.get("subject_type") == "event" and row.get("subject_id") not in event_ids:
            drift.append(f"{row.get('evidence_id')}: 悬空事件主体 {row.get('subject_id')}")
    for row in participants:
        if row.get("event_id") not in event_ids:
            drift.append(f"{row.get('event_participant_id')}: 悬空事件引用 {row.get('event_id')}")
    return not drift, drift


def main() -> str:
    src_fields, sources = read_csv(DATA / "sources.csv")
    ev_fields, facts = read_csv(DATA / "fact_evidences.csv")
    evt_fields, events = read_csv(DATA / "events.csv")
    part_fields, participants = read_csv(DATA / "event_participants.csv")

    done, _ = _postconditions_satisfied(sources, facts, events, participants)
    if done:
        message = "无新增/已完成：五项第四批A裁决均已落地，跳过写入。"
        print(message)
        return message

    problems = _check_baseline(sources, facts, events, participants)
    if problems:
        raise RuntimeError(
            "apply_batch4a_decisions 前置校验失败（未写入任何文件）：\n- " + "\n- ".join(problems)
        )

    fact_by_id = {row["evidence_id"]: row for row in facts}
    event_by_id = {row["event_id"]: row for row in events}

    # 事件层：EVT-00004 改写，EVT-00029 收窄；说明字段保留历史原因并追加本批出处。
    for event_id, expected in POST_EVENT_STATES.items():
        row = event_by_id[event_id]
        for key in ("event_name", "canonical_event_key", "event_date", "date_precision", "source_ids", "needs_manual_review"):
            row[key] = expected[key]
        row["display_note"] = expected["display_note"]
        _append_text(row, "correction_reason", expected["correction_reason_append"])

    # 参与者：删除 EVT-00005 两行；其余目标行只同步已有来源，不新增参与者。
    removed_participants = sum(1 for row in participants if row["event_id"] == DELETE_EVENT_ID)
    participants = [row for row in participants if row["event_id"] != DELETE_EVENT_ID]
    for row in participants:
        if row["event_id"] == "EVT-00004":
            row["source_ids"] = _join_ids(_split_ids(row["source_ids"]) + ["SRC-1167"])
        elif row["event_id"] == "EVT-00029":
            row["source_ids"] = _join_ids(_split_ids(row["source_ids"]) + ["SRC-1154"])

    # 事件物理删除。
    removed_event = any(row["event_id"] == DELETE_EVENT_ID for row in events)
    events = [row for row in events if row["event_id"] != DELETE_EVENT_ID]

    # 证据层按 evidence_id 定点处置。
    deleted_facts = 0
    for evidence_id, action in EVIDENCE_ACTIONS:
        row = fact_by_id[evidence_id]
        if action["kind"] == "delete":
            facts.remove(row)
            deleted_facts += 1
            continue
        before_support = row["evidence_support"]
        before_status = row["review_status"]
        kind = action["kind"]
        if kind == "promote_reviewed":
            row["review_status"] = "reviewed"
        elif kind == "reject":
            row["review_status"] = "rejected"
        elif kind == "demote_lead":
            row["evidence_support"] = "lead"
            row["review_status"] = "pending"
        elif kind == "promote_reviewed_from_lead":
            row["evidence_support"] = "support"
            row["review_status"] = "reviewed"
        elif kind == "remap_support":
            row["subject_id"] = action["new_subject"]
            row["evidence_support"] = "support"
            row["review_status"] = "reviewed"
            row["adjudication_status"] = ""
        elif kind == "remap_object":
            row["subject_id"] = action["new_subject"]
            row["object_value"] = action["new_object_value"]
            row["evidence_support"] = "support"
            row["review_status"] = "reviewed"
        _append_reviewer_note(row, action["audit"], before_support, before_status)

    write_csv(DATA / "fact_evidences.csv", ev_fields, facts)
    write_csv(DATA / "events.csv", evt_fields, events)
    write_csv(DATA / "event_participants.csv", part_fields, participants)

    message = (
        f"已执行：sources +0、fact_evidences -{deleted_facts}、events -{int(removed_event)}、"
        f"event_participants -{removed_participants}；终值 1177/626/147/222。"
    )
    print(message)
    return message


if __name__ == "__main__":
    main()
