"""
P1.2 / P1.3 收尾队列优先级与修订建议生成器

P1.2: 对 phase4_review_queue.csv（165 条事件/地点）按风险打出优先级并给处置建议。
P1.3: 在 phase5 审查基础上，汇总 critical/high 风险关系的修订/降级建议
      （phase5_critical_downgrade_recs.csv 已含 17 条 critical 无佐证降级建议）。

产出：
- research/drafts/reports/phase4_priority_recs.csv
- （P1.3 直接复用）phase5_critical_downgrade_recs.csv
"""
from __future__ import annotations

import csv
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
REPORTS = ROOT / "research" / "drafts" / "reports"

QUEUE = REPORTS / "phase4_review_queue.csv"
OUT = REPORTS / "phase4_priority_recs.csv"


def load_csv(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def main():
    queue = load_csv(QUEUE)
    out = []
    prio_count = {"高": 0, "中": 0, "低": 0}

    for r in queue:
        conf = (r.get("confidence") or "").strip()
        dp = (r.get("date_precision") or "").strip()
        stype = (r.get("subject_type") or "").strip()
        reason = (r.get("review_reason") or "").strip()
        name = r.get("subject_name", "")

        score = 0
        actions = []
        if conf == "low":
            score += 2
            actions.append("补充权威来源，提升 evidence_strength")
        if dp in ("月", "日"):
            score += 1
            actions.append(f"核实日期精度（当前仅到「{dp}」），补充日级定位")
        if "待核" in reason or "需人工复核" in reason:
            score += 1
            actions.append("人工复核后决定是否转正")
        if not actions:
            actions.append("维持待核标记，定期复核")

        if score >= 3:
            prio = "高"
        elif score == 2:
            prio = "中"
        else:
            prio = "低"
        prio_count[prio] += 1

        out.append({
            "priority": prio,
            "subject_type": stype,
            "subject_id": r.get("subject_id", ""),
            "subject_name": name,
            "confidence": conf,
            "date_precision": dp,
            "recommended_action": "；".join(actions),
        })

    # 高优先级在前
    order = {"高": 0, "中": 1, "低": 2}
    out.sort(key=lambda x: (order[x["priority"]], x["subject_type"]))

    fields = ["priority", "subject_type", "subject_id", "subject_name",
              "confidence", "date_precision", "recommended_action"]
    with open(OUT, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(out)

    print(f"Queue rows processed: {len(out)}")
    print("Priority distribution:", prio_count)
    print(f"Wrote: {OUT}")


if __name__ == "__main__":
    main()
