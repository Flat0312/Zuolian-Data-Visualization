"""
AI 辅助关系审查引擎（P0.1 / P1.3 基础）

对 phase5_relation_review_template.csv 的 400 条抽样关系做 *可复现* 的 AI 辅助预审：
- 结合 fact_evidences（事实证据）、org_memberships（组织成员台账）、来源摘录与风险等级，
  为每条关系给出 ai_verdict / ai_confidence / ai_method / ai_note。

重要：本脚本产出的是「AI 辅助草稿」，用于把人工复核工作量聚焦到 needs_human 子集，
不构成最终历史判定。正式结论须由人工复核签字。

输出：
- research/drafts/reports/phase5_relation_review_ai_filled.csv
- 打印分层统计（供 phase5_accuracy_report.md 使用）
"""
from __future__ import annotations

import csv
import json
import re
from collections import Counter, defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
PROCESSED = ROOT / "data" / "processed"
REPORTS = ROOT / "research" / "drafts" / "reports"

TEMPLATE = REPORTS / "phase5_relation_review_template.csv"
FACTS = PROCESSED / "fact_evidences.csv"
ORG_MEM = PROCESSED / "org_memberships.csv"
PERSONS = PROCESSED / "persons.csv"

OUT_FILLED = REPORTS / "phase5_relation_review_ai_filled.csv"

# 关系型事实谓词（用于判断 fact_evidences 是否可佐证一段人物关系）
RELATION_PREDICATES = {
    "relation", "association", "co_membership", "correspondence",
    "collaboration", "coauthorship", "kinship", "mentorship",
    "co_participation", "same_org",
}


def load_csv(path: Path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def build_name_index(persons):
    """person_id -> 候选姓名集合（标准名 + 别名），用于上下文匹配。"""
    idx = {}
    for p in persons:
        pid = p["person_id"]
        names = set()
        if p.get("standard_name"):
            names.add(p["standard_name"].strip())
        if p.get("aliases"):
            for a in re.split(r"[;/、,，]", p["aliases"]):
                a = a.strip()
                if a:
                    names.add(a)
        idx[pid] = names
    return idx


def norm(text: str) -> str:
    return (text or "").replace(" ", "").lower()


def main():
    template = load_csv(TEMPLATE)
    facts = load_csv(FACTS)
    org_mem = load_csv(ORG_MEM)
    persons = load_csv(PERSONS)

    name_idx = build_name_index(persons)
    name_by_id = {p["person_id"]: (p.get("standard_name") or "").strip() for p in persons}

    # 组织成员台账：person_id -> set(org_id)
    org_by_person = defaultdict(set)
    for m in org_mem:
        org_by_person[m["person_id"]].add(m["organization_id"])

    # 事实证据：person_id -> 是否含关系型证据
    fact_rel_person = defaultdict(bool)
    for fe in facts:
        if fe.get("evidence_support") != "support":
            continue
        if fe.get("predicate") in RELATION_PREDICATES:
            fact_rel_person[fe["subject_id"]] = True

    def names_in_context(sp, tp, context):
        c = norm(context)
        sp_names = name_idx.get(sp, set())
        tp_names = name_idx.get(tp, set())
        sp_hit = any(norm(n) in c for n in sp_names if n)
        tp_hit = any(norm(n) in c for n in tp_names if n)
        return sp_hit, tp_hit

    rows_out = []
    stats = {
        "total": 0,
        "verdict_by_risk": defaultdict(lambda: Counter()),
        "verdict_total": Counter(),
        "method_total": Counter(),
        "downgrade_recs": 0,
    }
    downgrade_rows = []

    for r in template:
        stats["total"] += 1
        rid = r["relation_id"]
        sp = r["source_person_id"]
        tp = r["target_person_id"]
        rtype = (r["standard_relation_type"] or "").strip()
        risk = (r["relation_risk_level"] or "").strip()
        context = r.get("context", "")

        sp_hit, tp_hit = names_in_context(sp, tp, context)
        both_names = sp_hit and tp_hit
        same_org = bool(org_by_person[sp] & org_by_person[tp])
        has_fact = fact_rel_person.get(sp) or fact_rel_person.get(tp)

        ai_verdict = None
        ai_confidence = None
        ai_method = None
        ai_note = None

        if rtype == "待核验":
            ai_verdict = "needs_human"
            ai_confidence = "low"
            ai_method = "type_flag"
            ai_note = "关系类型标记为待核验，缺少具体证据，需人工判定或降级"
        elif same_org and rtype == "同属组织":
            ai_verdict = "plausible"
            ai_confidence = "high"
            ai_method = "org_corroborated"
            ai_note = "双方均在组织成员台账中，同属组织关系可佐证"
        elif both_names and risk in ("low", "medium"):
            ai_verdict = "plausible"
            ai_confidence = "medium"
            ai_method = "context_match"
            ai_note = "来源摘录同时出现双方姓名，且风险较低，关系 plausibly 成立"
        elif both_names and risk == "high":
            ai_verdict = "plausible"
            ai_confidence = "medium"
            ai_method = "context_match"
            ai_note = "来源摘录同时出现双方姓名，但风险为 high，建议人工复签"
        elif both_names and risk == "critical":
            ai_verdict = "needs_human"
            ai_confidence = "low"
            ai_method = "context_match_partial"
            ai_note = "摘录含双方姓名但风险 critical，需人工核验后确认"
        elif has_fact:
            ai_verdict = "plausible"
            ai_confidence = "medium"
            ai_method = "fact_corroborated"
            ai_note = "存在相关事实证据可佐证该关系"
        else:
            ai_verdict = "needs_human"
            ai_confidence = "low"
            ai_method = "no_corroboration"
            ai_note = "未找到直接佐证，建议人工核验或降级"
            if risk == "critical":
                ai_note += "（高风险无佐证，建议降级处理）"

        stats["verdict_total"][ai_verdict] += 1
        stats["verdict_by_risk"][risk][ai_verdict] += 1
        stats["method_total"][ai_method] += 1

        # P1.3 降级/修订建议：critical 且无佐证 → 建议降级 + 标记人工
        if ai_verdict == "needs_human" and risk == "critical" and ai_method == "no_corroboration":
            stats["downgrade_recs"] += 1
            downgrade_rows.append({
                "relation_id": rid,
                "source_person_id": sp,
                "target_person_id": tp,
                "standard_relation_type": rtype,
                "current_risk": risk,
                "suggested_risk": "medium",
                "suggested_confidence": "low",
                "needs_manual_review": "yes",
                "correction_reason": "AI预审：critical风险且无任何佐证，建议降级至medium并保留人工复核",
            })

        out = dict(r)
        out["ai_verdict"] = ai_verdict
        out["ai_confidence"] = ai_confidence
        out["ai_method"] = ai_method
        out["ai_note"] = ai_note
        rows_out.append(out)

    # 写出 filled csv
    fields = list(template[0].keys()) + ["ai_verdict", "ai_confidence", "ai_method", "ai_note"]
    with open(OUT_FILLED, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(rows_out)

    # 写出降级建议 csv（P1.3）
    downgrade_path = REPORTS / "phase5_critical_downgrade_recs.csv"
    with open(downgrade_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(downgrade_rows[0].keys()) if downgrade_rows else [])
        w.writeheader()
        w.writerows(downgrade_rows)

    # 打印统计
    print(f"Total reviewed: {stats['total']}")
    print("Verdict total:", dict(stats["verdict_total"]))
    print("Method total:", dict(stats["method_total"]))
    print("Downgrade recs (critical,no_corroboration):", stats["downgrade_recs"])
    print("\nVerdict by risk level:")
    for risk in ("critical", "high", "medium", "low"):
        if risk in stats["verdict_by_risk"]:
            print(f"  {risk}: {dict(stats['verdict_by_risk'][risk])}")
    print(f"\nWrote: {OUT_FILLED}")
    print(f"Wrote: {downgrade_path}")

    # 写出统计 json 供报告生成读取
    stat_json = REPORTS / "phase5_review_stats.json"
    summary = {
        "total": stats["total"],
        "verdict_total": dict(stats["verdict_total"]),
        "method_total": dict(stats["method_total"]),
        "downgrade_recs": stats["downgrade_recs"],
        "verdict_by_risk": {k: dict(v) for k, v in stats["verdict_by_risk"].items()},
    }
    with open(stat_json, "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    print(f"Wrote: {stat_json}")


if __name__ == "__main__":
    main()
