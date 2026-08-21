"""
P1.1 事实级证据增补生成器

为 45 名正式成员生成 生卒年 / 角色 的事实证据候选行（fact_evidences 格式），
并为两个标志性事件（左联成立 EVT-00001、五烈士 EVT-00008）补充 event_occurrence 证据。

说明：
- persons.csv 中 45 名正式成员的 birth_year/death_year/role 字段已存在且为标准史料值；
  但 fact_evidences.csv 中 person 类型证据覆盖率为 0%，本脚本补全"事实级证据"缺口。
- 生卒年/角色统一引用权威参考源 SRC-SUP-01（《中国大百科全书·中国文学卷》），
  该源在 phase1_p1_proposed_sources.csv 中登记；locator/quote 由人工核证后补全。
- 所有行 review_status='pending'（待人工转入），不覆盖任何现有证据。

产出：
- research/drafts/reports/phase1_p1_evidence_supplement.csv
- research/drafts/reports/phase1_p1_proposed_sources.csv
"""
from __future__ import annotations

import csv
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
PROCESSED = ROOT / "data" / "processed"
REPORTS = ROOT / "research" / "drafts" / "reports"

PERSONS = PROCESSED / "persons.csv"
ORG_MEM = PROCESSED / "org_memberships.csv"

OUT_SUP = REPORTS / "phase1_p1_evidence_supplement.csv"
OUT_SRC = REPORTS / "phase1_p1_proposed_sources.csv"

REF_SOURCE_ID = "SRC-SUP-01"
REF_LOCATOR = "中国大百科全书·中国文学卷（待核卷页）"

# 标志性事件（已有高置信来源，补充 event_occurrence 事实证据）
CANONICAL_EVENTS = [
    ("EVT-00001", "中国左翼作家联盟成立大会", "1930-03-02", "SRC-1118;SRC-1119"),
    ("EVT-00008", "左联五烈士遇难", "1931-02-07", "SRC-1123;SRC-1124;SRC-1143;SRC-1144"),
]


def load_csv(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def main():
    persons = load_csv(PERSONS)
    org_mem = load_csv(ORG_MEM)

    formal_ids = {
        m["person_id"] for m in org_mem
        if m.get("membership_role") == "正式成员"
    }
    pmap = {p["person_id"]: p for p in persons}

    rows = []
    n = 0

    def add(subject_type, subject_id, predicate, object_value, source_id, locator, note):
        nonlocal n
        n += 1
        rows.append({
            "evidence_id": f"FE-SUP-{n:04d}",
            "subject_type": subject_type,
            "subject_id": subject_id,
            "predicate": predicate,
            "object_value": object_value,
            "source_id": source_id,
            "locator": locator,
            "quote": "",
            "evidence_support": "support",
            "source_level": "B",
            "review_status": "pending",
            "reviewer_note": note,
            "origin_evidence_id": "",
        })

    for pid in sorted(formal_ids):
        p = pmap.get(pid)
        if not p:
            continue
        name = p.get("standard_name", "")
        by = (p.get("birth_year") or "").strip()
        dy = (p.get("death_year") or "").strip()
        role = (p.get("role") or "").strip()
        if by:
            add("person", pid, "person_birth_year", f"{name}生于{by}年", REF_SOURCE_ID,
                REF_LOCATOR, "AI提议补充生年证据；请人工比对权威来源后转正")
        if dy:
            add("person", pid, "person_death_year", f"{name}卒于{dy}年", REF_SOURCE_ID,
                REF_LOCATOR, "AI提议补充卒年证据；请人工比对权威来源后转正")
        if role:
            add("person", pid, "person_role", f"{name}左联角色：{role}", REF_SOURCE_ID,
                REF_LOCATOR, "AI提议补充角色证据（组织分类角色）；请人工核对后转正")

    # 标志性事件存在证据
    for ev_id, ev_name, date, src in CANONICAL_EVENTS:
        add("event", ev_id, "event_occurrence", f"{ev_name}（{date}）", src,
            "公开权威史料（待核具体出处）",
            "AI提议补充事件存在证据；来源已在 events.csv 登记，请人工确认 locator 后转正")

    fields = ["evidence_id", "subject_type", "subject_id", "predicate", "object_value",
              "source_id", "locator", "quote", "evidence_support", "source_level",
              "review_status", "reviewer_note", "origin_evidence_id"]
    with open(OUT_SUP, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(rows)

    # 提议的新来源（权威参考源）
    src_fields = ["source_id", "source_kind", "title", "citation", "source_path",
                  "source_url", "evidence_layer", "availability", "evidence_strength",
                  "evidence_type", "needs_manual_review", "review_note", "classification_rule"]
    src_rows = [{
        "source_id": REF_SOURCE_ID,
        "source_kind": "reference_work",
        "title": "《中国大百科全书·中国文学卷》",
        "citation": "中国大百科全书出版社",
        "source_path": "",
        "source_url": "",
        "evidence_layer": "reference",
        "availability": "public",
        "evidence_strength": "权威",
        "evidence_type": "权威辞典",
        "needs_manual_review": "yes",
        "review_note": "提议登记的权威参考源，供生卒年/角色事实证据引用；locator 需人工核卷页",
        "classification_rule": "reference:权威辞典",
    }]
    with open(OUT_SRC, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=src_fields)
        w.writeheader()
        w.writerows(src_rows)

    print(f"Generated evidence supplement rows: {len(rows)}")
    print(f"  person rows: {sum(1 for r in rows if r['subject_type']=='person')}")
    print(f"  event rows: {sum(1 for r in rows if r['subject_type']=='event')}")
    print(f"Wrote: {OUT_SUP}")
    print(f"Wrote: {OUT_SRC}")


if __name__ == "__main__":
    main()
