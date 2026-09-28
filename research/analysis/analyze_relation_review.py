"""Phase 5 关系审核结果分析：总体与分层准确率、错误率与修订规则。

输入：``phase5_relation_review_package.csv``（含人工填写后的 ``human_verdict`` 列，
允许部分填写）。

输出（写入 ``--out-dir``，默认 ``research/drafts/reports/``）：
- ``phase5_review_accuracy_report.md``：人读报告。
- ``phase5_review_stats.json``：机器可读统计。

口径定义（仅对已裁决行计算，空白行只计入 ``n_pending``）：
- 关系成立率 ``pair_accuracy`` = (correct + wrong_type) / n_adjudicated
  （wrong_type 表示关系成立但类型错，对「关系对是否存在」仍是肯定判定）。
- 类型准确率 ``type_accuracy`` = correct / n_adjudicated。
- ``not_supported_rate`` / ``contradicted_rate`` 分别为对应裁决占比。
- ``ai_agreement_rate`` = human_verdict == ai_suggested_verdict 的占比。

边界：尚无人工裁决时一律输出 ``None`` 与「尚无人工裁决」哨兵文本，
绝不输出预估或模拟的准确率数字；修订规则按已裁决行错误数阈值
（错误 ≥3 且占该分层已裁决数 ≥10%）确定性生成。
"""

from __future__ import annotations

import argparse
import csv
import json
from collections import Counter, defaultdict
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_PACKAGE = REPO_ROOT / "research" / "drafts" / "reports" / "phase5_relation_review_package.csv"
DEFAULT_OUT_DIR = REPO_ROOT / "research" / "drafts" / "reports"

VALID_VERDICTS = ("correct", "wrong_type", "not_supported", "contradicted")
ERROR_VERDICTS = ("wrong_type", "not_supported", "contradicted")
STRATA_DIMENSIONS = (
    ("risk", "relation_risk_level"),
    ("confidence", "confidence"),
    ("support", "overnight_evidence_support"),
    ("type", "standard_relation_type"),
)
RULE_MIN_ERRORS = 3
RULE_MIN_ERROR_RATE = 0.1


def _round(value: float) -> float:
    return round(value, 4)


def _rates(verdict_counts: Counter[str], n: int) -> dict[str, object]:
    if n == 0:
        return {"n": 0, "verdict_counts": {}, "pair_accuracy": None, "type_accuracy": None}
    return {
        "n": n,
        "verdict_counts": {verdict: verdict_counts[verdict] for verdict in VALID_VERDICTS if verdict_counts[verdict]},
        "pair_accuracy": _round((verdict_counts["correct"] + verdict_counts["wrong_type"]) / n),
        "type_accuracy": _round(verdict_counts["correct"] / n),
    }


def analyze(package_path: Path, out_dir: Path) -> dict[str, object]:
    with package_path.open(encoding="utf-8-sig", newline="") as handle:
        rows = [dict(row) for row in csv.DictReader(handle)]

    invalid = [row["relation_id"] for row in rows if row.get("human_verdict", "") and row["human_verdict"] not in VALID_VERDICTS]
    if invalid:
        sample = next(row for row in rows if row["relation_id"] == invalid[0])
        raise ValueError(f"非法 human_verdict：{invalid}，首条 {invalid[0]}={sample['human_verdict']!r}")

    adjudicated = [row for row in rows if row.get("human_verdict", "")]
    pending = [row for row in rows if not row.get("human_verdict", "")]

    overall_counts: Counter[str] = Counter(row["human_verdict"] for row in adjudicated)
    agreement = sum(1 for row in adjudicated if row["human_verdict"] == row["ai_suggested_verdict"])
    n = len(adjudicated)

    stats: dict[str, object] = {
        "n_total": len(rows),
        "n_adjudicated": n,
        "n_pending": len(pending),
        "verdict_counts": {verdict: overall_counts[verdict] for verdict in VALID_VERDICTS},
        "pair_accuracy": _round((overall_counts["correct"] + overall_counts["wrong_type"]) / n) if n else None,
        "type_accuracy": _round(overall_counts["correct"] / n) if n else None,
        "not_supported_rate": _round(overall_counts["not_supported"] / n) if n else None,
        "contradicted_rate": _round(overall_counts["contradicted"] / n) if n else None,
        "ai_agreement_rate": _round(agreement / n) if n else None,
    }

    dimension_values: dict[str, dict[str, list[dict[str, str]]]] = {}
    for dimension, column in STRATA_DIMENSIONS:
        grouped: dict[str, list[dict[str, str]]] = defaultdict(list)
        for row in adjudicated:
            grouped[row[column]].append(row)
        dimension_values[dimension] = dict(grouped)
        stats[f"by_{dimension}"] = {
            value: _rates(Counter(item["human_verdict"] for item in items), len(items))
            for value, items in sorted(grouped.items())
        }

    rules = []
    for dimension, grouped in dimension_values.items():
        for value, items in sorted(grouped.items()):
            counts = Counter(item["human_verdict"] for item in items)
            errors = sum(counts[verdict] for verdict in ERROR_VERDICTS)
            stratum_n = len(items)
            if errors >= RULE_MIN_ERRORS and errors / stratum_n >= RULE_MIN_ERROR_RATE:
                dominant = max(ERROR_VERDICTS, key=lambda verdict: counts[verdict])
                rules.append(
                    {
                        "dimension": dimension,
                        "value": value,
                        "n": stratum_n,
                        "errors": errors,
                        "error_rate": _round(errors / stratum_n),
                        "dominant_error": dominant,
                        "suggestion": f"{value} 层以 {dominant} 为主，建议整层按裁决模式修订（降级待证 / 改类型）后复算",
                    }
                )
    stats["revision_rules"] = rules

    warnings = [
        f"{row['relation_id']} 裁决为 wrong_type 但 human_note 缺建议类型"
        for row in adjudicated
        if row["human_verdict"] == "wrong_type" and not row.get("human_note", "").strip()
    ]
    stats["warnings"] = warnings

    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / "phase5_review_stats.json").write_text(
        json.dumps(stats, ensure_ascii=False, indent=2) + "\n", encoding="utf-8", newline=""
    )
    _write_report(out_dir / "phase5_review_accuracy_report.md", stats, package_path.name)
    return stats


def _fmt_rate(value: object) -> str:
    return "—" if value is None else f"{value:.1%}"


def _write_report(path: Path, stats: dict[str, object], package_name: str) -> None:
    lines = [
        "# Phase 5 · 关系审核准确率报告",
        "",
        "> 生成脚本：``research/analysis/analyze_relation_review.py``（幂等）",
        f"> 输入：``{package_name}`` 的 ``human_verdict`` 人工裁决列",
        "",
        "## 0. 口径声明",
        "",
        "本报告只统计人工已裁决行；AI 建议列不进入任何准确率分子分母。",
        "关系成立率把 wrong_type 计为成立（关系存在、类型需改），类型准确率只认 correct。",
        "",
    ]
    if stats["n_adjudicated"] == 0:
        lines.extend(
            [
                "## 1. 当前状态",
                "",
                f"**尚无人工裁决**（n_pending={stats['n_pending']}）。无法计算准确率，本报告不提供任何预估数字。",
                "请在审核包 CSV 填写 human_verdict 后重跑本脚本。",
                "",
            ]
        )
    else:
        counts: dict[str, int] = stats["verdict_counts"]  # type: ignore[assignment]
        lines.extend(
            [
                "## 1. 总体（仅已裁决行）",
                "",
                f"- 已裁决 {stats['n_adjudicated']} / {stats['n_total']}，待裁决 {stats['n_pending']}",
                f"- 裁决分布：correct {counts['correct']} / wrong_type {counts['wrong_type']} / "
                f"not_supported {counts['not_supported']} / contradicted {counts['contradicted']}",
                f"- 关系成立率（correct+wrong_type）：**{_fmt_rate(stats['pair_accuracy'])}**",
                f"- 类型准确率（correct）：**{_fmt_rate(stats['type_accuracy'])}**",
                f"- not_supported 率：{_fmt_rate(stats['not_supported_rate'])}；contradicted 率：{_fmt_rate(stats['contradicted_rate'])}",
                f"- 人工与 AI 建议一致率：{_fmt_rate(stats['ai_agreement_rate'])}（一致性仅供交叉参考，不替代人工判定）",
                "",
            ]
        )
        for dimension, column_label in (("risk", "风险层"), ("support", "证据分诊层")):
            key = f"by_{dimension}"
            lines.extend([f"## 分层 · {column_label}", "", "| 分层 | 已裁决 | 成立率 | 类型准确率 |", "| --- | ---: | ---: | ---: |"])
            for value, stratum in stats[key].items():  # type: ignore[union-attr]
                lines.append(
                    f"| {value} | {stratum['n']} | {_fmt_rate(stratum['pair_accuracy'])} | {_fmt_rate(stratum['type_accuracy'])} |"
                )
            lines.append("")
        rules: list[dict[str, object]] = stats["revision_rules"]  # type: ignore[assignment]
        lines.extend(["## 修订规则（按错误阈值确定性生成）", ""])
        if rules:
            lines.append("| 维度 | 分层 | 已裁决 | 错误数 | 错误率 | 主导错误 | 建议 |")
            lines.append("| --- | --- | ---: | ---: | ---: | --- | --- |")
            for rule in rules:
                lines.append(
                    f"| {rule['dimension']} | {rule['value']} | {rule['n']} | {rule['errors']} | "
                    f"{_fmt_rate(rule['error_rate'])} | {rule['dominant_error']} | {rule['suggestion']} |"
                )
        else:
            lines.append("当前无满足阈值（错误 ≥3 且 ≥10%）的分层。")
        lines.append("")
    warnings: list[str] = stats["warnings"]  # type: ignore[assignment]
    if warnings:
        lines.extend(["## 警告", ""])
        lines.extend(f"- {warning}" for warning in warnings)
        lines.append("")
    path.write_text("\n".join(lines), encoding="utf-8", newline="")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--package", type=Path, default=DEFAULT_PACKAGE)
    parser.add_argument("--out-dir", type=Path, default=DEFAULT_OUT_DIR)
    args = parser.parse_args()
    stats = analyze(args.package, args.out_dir)
    print(
        {
            key: stats[key]
            for key in ("n_total", "n_adjudicated", "n_pending", "pair_accuracy", "type_accuracy")
        }
    )


if __name__ == "__main__":
    main()
