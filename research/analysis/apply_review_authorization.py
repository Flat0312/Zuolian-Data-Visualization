"""按用户授权把 Phase 5 审核包的 400 条关系按 AI 建议全量落地裁决（签核路径 C）。

输入：``phase5_relation_review_package.csv``（human_verdict 为空的规范审核包）。
输出：
- ``--out``（默认 ``phase5_relation_review_adjudicated.csv``）：全列保留 +
  ``human_verdict``=``ai_suggested_verdict``、wrong_type 行 ``human_note`` 给建议类型、
  每行附 ``authorized_by``/``authorized_at``/``authorization_quote`` 溯源列。
- 同目录 ``phase5_review_adjudication_record.md``：授权登记（逐字引语、路径、分布）。

边界：授权语必须非空且不得为执行者自授（拒绝「AI 自行决定」类占位）；
本脚本只写研究层裁决文件，不改生产层；规范审核包保持空裁决状态，
重跑 ``build_phase5_review_package.py`` 不会覆盖本裁决（产物分文件）。
裁决口径为「授权按建议执行」，非逐条独立人工复核，报告引用时须带此口径。
"""

from __future__ import annotations

import argparse
import csv
from collections import Counter
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_PACKAGE = REPO_ROOT / "research" / "drafts" / "reports" / "phase5_relation_review_package.csv"
DEFAULT_OUT = REPO_ROOT / "research" / "drafts" / "reports" / "phase5_relation_review_adjudicated.csv"

VALID_VERDICTS = ("correct", "wrong_type", "not_supported", "contradicted")
AUTHORIZED_BY = "用户（会话授权）"
AUTHORIZED_AT = "2026-09-20"
PLACEHOLDER_PATTERNS = ("ai 自行决定", "执行者自授", "模型决定")

PROVENANCE_COLUMNS = ["authorized_by", "authorized_at", "authorization_quote"]


def apply(package_path: Path, out_path: Path, authorization_quote: str) -> dict[str, int]:
    quote = authorization_quote.strip()
    if not quote:
        raise ValueError("授权语为空：全量按建议执行裁决必须有用户逐字授权语（签核单路径 C）")
    if any(pattern in quote.lower() for pattern in PLACEHOLDER_PATTERNS):
        raise ValueError(f"授权语疑似执行者自授占位，拒绝执行: {quote!r}")

    with package_path.open(encoding="utf-8-sig", newline="") as handle:
        rows = [dict(row) for row in csv.DictReader(handle)]

    filled = [row["human_verdict"] for row in rows if row.get("human_verdict", "").strip()]
    if filled:
        raise ValueError(f"规范审核包已有 {len(filled)} 条非空 human_verdict，疑似重复落地，拒绝执行")

    counts: Counter[str] = Counter()
    for row in rows:
        verdict = row["ai_suggested_verdict"]
        if verdict not in VALID_VERDICTS:
            raise ValueError(f"{row['relation_id']} 非法 ai_suggested_verdict: {verdict}")
        row["human_verdict"] = verdict
        if verdict == "wrong_type":
            row["human_note"] = f"建议类型：{row['ai_suggested_type']}（授权按AI建议执行）"
        row["authorized_by"] = AUTHORIZED_BY
        row["authorized_at"] = AUTHORIZED_AT
        row["authorization_quote"] = quote
        counts[verdict] += 1

    out_path.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = list(rows[0].keys())
    for column in PROVENANCE_COLUMNS:
        if column not in fieldnames:
            fieldnames.append(column)
    with out_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames, lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)

    _write_record(
        out_path.parent / "phase5_review_adjudication_record.md",
        quote,
        counts,
        out_path.name,
    )
    return {"rows": len(rows), **{verdict: counts[verdict] for verdict in VALID_VERDICTS}}


def _write_record(path: Path, quote: str, counts: Counter[str], out_name: str) -> None:
    lines = [
        "# Phase 5 审核裁决登记（路径 C：全量按 AI 建议执行）",
        "",
        f"- 授权人：{AUTHORIZED_BY}",
        f"- 授权时间：{AUTHORIZED_AT}",
        f"- 授权语（逐字）：「{quote}」",
        "- 授权语境：针对 2026-09-20 P0 收尾交付的三项待裁决清单"
        "（400 条关系裁决、12 条候选裁决含 REL-00523 换引文、T2–T4 整改升级）逐项表示无异议并授权执行者处置。",
        "- 裁决来源：`ai_suggested_verdict`（夜间轮回源核查建议，`phase5_relation_review_package.csv`）。",
        f"- 裁决产物：`{out_name}`（每行带 authorized_by/authorized_at/authorization_quote 溯源列）。",
        "",
        "## 裁决分布",
        "",
        f"- correct {counts['correct']}",
        f"- wrong_type {counts['wrong_type']}（human_note 附建议类型）",
        f"- not_supported {counts['not_supported']}",
        f"- contradicted {counts['contradicted']}",
        f"- 合计 {sum(counts.values())}",
        "",
        "## 口径声明",
        "",
        "本裁决为**授权按建议执行**，非逐条独立人工复核；据此计算的准确率为「授权按建议口径」实测值，",
        "答辩引用时须带此口径。规范审核包 `phase5_relation_review_package.csv` 保持空裁决状态，",
        "如需逐条复核仍可另行进行（以其为准覆盖本裁决需重新授权）。",
        "",
    ]
    path.write_text("\n".join(lines), encoding="utf-8", newline="")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--package", type=Path, default=DEFAULT_PACKAGE)
    parser.add_argument("--out", type=Path, default=DEFAULT_OUT)
    parser.add_argument("--authorization-quote", required=True)
    args = parser.parse_args()
    print(apply(args.package, args.out, args.authorization_quote))


if __name__ == "__main__":
    main()
