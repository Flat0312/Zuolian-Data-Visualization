"""构建 Phase 5 关系审核包：把夜间轮回源分诊并入 400 条人工审核模板。

输入：
- ``phase5_relation_review_template.csv``（Phase 5 抽样模板，400 条）
- ``research/content-upgrade-overnight-2026-09-06/relations.jsonl``（夜间轮 411 条
  回源核查记录，取 ``sample400`` 组）

输出（写入 ``--out-dir``，默认 ``research/drafts/reports/``）：
- ``phase5_relation_review_package.csv``：模板列原样保留 + 夜间轮证据列 +
  ``ai_suggested_verdict`` 建议列 + 空 ``human_verdict``/``human_note`` 人工列。
- ``phase5_relation_review_signoff.md``：签核单（裁决词表、分层统计、签核路径）。

边界：本脚本只做打包，不写生产层；``human_verdict``/``human_note`` 必须留空，
人工裁决不得由执行者代填。``ai_suggested_verdict`` 映射规则：
- ``insufficient`` → ``not_supported``（现有证据不足以支持）
- 其余且夜间轮提议类型 ≠ 现类型 → ``wrong_type``（建议类型 = 提议类型）
- 其余 → ``correct``
"""

from __future__ import annotations

import argparse
import csv
import json
from collections import Counter
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_TEMPLATE = REPO_ROOT / "research" / "drafts" / "reports" / "phase5_relation_review_template.csv"
DEFAULT_OVERNIGHT = REPO_ROOT / "research" / "content-upgrade-overnight-2026-09-06" / "relations.jsonl"
DEFAULT_OUT_DIR = REPO_ROOT / "research" / "drafts" / "reports"

PACKAGE_COLUMNS = [
    "relation_id",
    "source_person_id",
    "target_person_id",
    "standard_relation_type",
    "relation_risk_level",
    "confidence",
    "context",
    "overnight_evidence_support",
    "overnight_proposed_type",
    "overnight_reason",
    "overnight_evidence_ids",
    "overnight_search_log_ids",
    "ai_suggested_verdict",
    "ai_suggested_type",
    "human_verdict",
    "human_note",
]

VERDICT_DEFINITIONS = [
    ("correct", "关系成立且 standard_relation_type 恰当，可作为已证关系保留"),
    ("wrong_type", "关系成立但类型应改，human_note 填建议类型"),
    ("not_supported", "现有证据不足以支持，不可作为已证关系保留"),
    ("contradicted", "证据显示关系不成立或错挂，应删除或降级"),
]

VALID_SUPPORT = {"support", "associated", "insufficient"}


def _suggested_verdict(evidence_support: str, proposed_type: str, standard_type: str) -> tuple[str, str]:
    if evidence_support == "insufficient":
        return "not_supported", ""
    if proposed_type != standard_type:
        return "wrong_type", proposed_type
    return "correct", ""


def _load_template(template_path: Path) -> list[dict[str, str]]:
    with template_path.open(encoding="utf-8-sig", newline="") as handle:
        return [dict(row) for row in csv.DictReader(handle)]


def _load_overnight_sample400(overnight_path: Path) -> dict[str, dict]:
    records: dict[str, dict] = {}
    with overnight_path.open(encoding="utf-8") as handle:
        for line in handle:
            line = line.strip()
            if not line:
                continue
            record = json.loads(line)
            if "sample400" not in record.get("groups", []):
                continue
            relation_id = record["relation_id"]
            if relation_id in records:
                raise ValueError(f"relations.jsonl sample400 组存在重复 relation_id: {relation_id}")
            records[relation_id] = record
    return records


def build(template_path: Path, overnight_path: Path, out_dir: Path) -> dict[str, int]:
    template = _load_template(template_path)
    overnight = _load_overnight_sample400(overnight_path)

    missing = [row["relation_id"] for row in template if row["relation_id"] not in overnight]
    if missing:
        raise ValueError(f"夜间轮 sample400 未覆盖模板关系 {len(missing)} 条，首条: {missing[0]}")

    rows: list[dict[str, str]] = []
    support_counts: Counter[str] = Counter()
    verdict_counts: Counter[str] = Counter()
    risk_strata: Counter[tuple[str, str]] = Counter()

    for entry in template:
        record = overnight[entry["relation_id"]]
        evidence_support = record["evidence_support"]
        if evidence_support not in VALID_SUPPORT:
            raise ValueError(f"{entry['relation_id']} 非法 evidence_support: {evidence_support}")
        suggested_verdict, suggested_type = _suggested_verdict(
            evidence_support, record["proposed_relation_type"], entry["standard_relation_type"]
        )
        rows.append(
            {
                "relation_id": entry["relation_id"],
                "source_person_id": entry["source_person_id"],
                "target_person_id": entry["target_person_id"],
                "standard_relation_type": entry["standard_relation_type"],
                "relation_risk_level": entry["relation_risk_level"],
                "confidence": entry["confidence"],
                "context": entry["context"],
                "overnight_evidence_support": evidence_support,
                "overnight_proposed_type": record["proposed_relation_type"],
                "overnight_reason": record["reason"],
                "overnight_evidence_ids": ";".join(record.get("evidence_ids", [])),
                "overnight_search_log_ids": ";".join(record.get("search_log_ids", [])),
                "ai_suggested_verdict": suggested_verdict,
                "ai_suggested_type": suggested_type,
                "human_verdict": "",
                "human_note": "",
            }
        )
        support_counts[evidence_support] += 1
        verdict_counts[suggested_verdict] += 1
        risk_strata[(evidence_support, entry["relation_risk_level"])] += 1

    out_dir.mkdir(parents=True, exist_ok=True)
    package_path = out_dir / "phase5_relation_review_package.csv"
    with package_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=PACKAGE_COLUMNS, lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)

    _write_signoff(out_dir / "phase5_relation_review_signoff.md", support_counts, verdict_counts, risk_strata)

    return {
        "rows": len(rows),
        "support": support_counts["support"],
        "associated": support_counts["associated"],
        "insufficient": support_counts["insufficient"],
        "suggested_correct": verdict_counts["correct"],
        "suggested_wrong_type": verdict_counts["wrong_type"],
        "suggested_not_supported": verdict_counts["not_supported"],
    }


def _write_signoff(
    path: Path,
    support_counts: Counter[str],
    verdict_counts: Counter[str],
    risk_strata: Counter[tuple[str, str]],
) -> None:
    total = sum(support_counts.values())
    risk_levels = ["critical", "high", "medium", "low"]
    lines = [
        "# Phase 5 · 400 条关系审核签核单",
        "",
        "> 生成脚本：``research/analysis/build_phase5_review_package.py``（幂等，重跑字节不变）",
        "> 数据来源：``phase5_relation_review_template.csv`` × 夜间轮 ``relations.jsonl`` sample400 组（2026-09-06 回源核查，全 pending_human_review）",
        "",
        "## 0. 边界声明",
        "",
        "``ai_suggested_verdict`` 是夜间轮回源核查的机器建议，不是人工判定。",
        "本签核单及审核包整体状态为 pending_human_review；``human_verdict``/``human_note``",
        "两列必须由人工评审员填写，执行者不得代填。人工裁决落地前，本包不产生任何准确率结论。",
        "",
        "## 1. 裁决词表（human_verdict 合法值）",
        "",
        "| 值 | 含义 |",
        "| --- | --- |",
    ]
    lines.extend(f"| `{name}` | {meaning} |" for name, meaning in VERDICT_DEFINITIONS)
    lines.extend(
        [
            "",
            "## 2. 分层统计（AI 建议口径）",
            "",
            f"共 {total} 条。夜间轮证据分诊：support {support_counts['support']} / "
            f"associated {support_counts['associated']} / insufficient {support_counts['insufficient']}。",
            f"AI 建议裁决：correct {verdict_counts['correct']} / wrong_type {verdict_counts['wrong_type']} / "
            f"not_supported {verdict_counts['not_supported']} / contradicted 0。",
            "",
            "| 证据分诊 × 风险 | critical | high | medium | low |",
            "| --- | ---: | ---: | ---: | ---: |",
        ]
    )
    for level_name, label in (("support", "support"), ("associated", "associated"), ("insufficient", "insufficient")):
        cells = [str(risk_strata[(level_name, risk)]) for risk in risk_levels]
        lines.append(f"| {label} | " + " | ".join(cells) + " |")
    lines.extend(
        [
            "",
            "## 3. 签核路径",
            "",
            "1. 路径 A（逐条）：按 `phase5_human_review_queue.csv` 顺序或直接在审核包 CSV 中逐条填写 `human_verdict`。",
            "2. 路径 B（分层批量授权）：对某一分层（如 insufficient × critical）整层授权按 `ai_suggested_verdict` 执行，需在授权语中点名分层。",
            "3. 路径 C（全量按建议执行）：对 400 条全部按 `ai_suggested_verdict` 落地，需明确授权语（参照第三批追认模式，落地脚本另行任务书）。",
            "",
            "任何路径下，`wrong_type` 的目标类型与 `not_supported` 的处置（降级/删除）以人工裁决为准；",
            "与 AI 建议不一致的行请在 `human_note` 写明理由。",
            "",
            "## 4. 关联文件",
            "",
            "- 审核包：`research/drafts/reports/phase5_relation_review_package.csv`",
            "- 夜间轮证据与回执：`research/content-upgrade-overnight-2026-09-06/`（relations.jsonl、evidence.jsonl、search_log.jsonl）",
            "- 旧 AI 启发式预审（已被本包取代，保留历史）：`phase5_relation_review_ai_filled.csv`",
            "",
        ]
    )
    path.write_text("\n".join(lines), encoding="utf-8", newline="")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE)
    parser.add_argument("--overnight", type=Path, default=DEFAULT_OVERNIGHT)
    parser.add_argument("--out-dir", type=Path, default=DEFAULT_OUT_DIR)
    args = parser.parse_args()
    summary = build(args.template, args.overnight, args.out_dir)
    print(summary)


if __name__ == "__main__":
    main()
