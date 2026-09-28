from __future__ import annotations

import hashlib
import json
from pathlib import Path

import pandas as pd

from research.analysis.analyze_relation_review import analyze

REPO_ROOT = Path(__file__).resolve().parents[1]
PACKAGE_PATH = REPO_ROOT / "research" / "drafts" / "reports" / "phase5_relation_review_package.csv"
DATA_DIR = REPO_ROOT / "data" / "processed"

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


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _fixture(tmp_path: Path, verdicts: list[tuple[str, str, str, str, str, str]]) -> Path:
    """verdicts: (relation_id, risk, confidence, support, ai_suggested, human_verdict[, note]) 补齐到 6 元。"""
    rows = []
    for index, item in enumerate(verdicts):
        relation_id, risk, confidence, support, suggested, human = item[:6]
        note = item[6] if len(item) > 6 else ""
        rows.append(
            {
                "relation_id": relation_id,
                "source_person_id": "ZLH-001",
                "target_person_id": f"ZLH-{index + 2:03d}",
                "standard_relation_type": "交往",
                "relation_risk_level": risk,
                "confidence": confidence,
                "context": "ctx",
                "overnight_evidence_support": support,
                "overnight_proposed_type": "交往",
                "overnight_reason": "r",
                "overnight_evidence_ids": "EVI-X",
                "overnight_search_log_ids": "SRCH-X",
                "ai_suggested_verdict": suggested,
                "ai_suggested_type": "",
                "human_verdict": human,
                "human_note": note,
            }
        )
    path = tmp_path / "package.csv"
    pd.DataFrame(rows, columns=PACKAGE_COLUMNS).to_csv(path, index=False, encoding="utf-8-sig")
    return path


def test_real_package_without_human_verdicts_reports_pending_not_fake_accuracy(tmp_path: Path) -> None:
    watched = [DATA_DIR / "person_relations.csv"]
    before = {path: _sha256(path) for path in watched}

    stats = analyze(PACKAGE_PATH, tmp_path)

    assert stats["n_total"] == 400
    assert stats["n_adjudicated"] == 0
    assert stats["n_pending"] == 400
    assert stats["pair_accuracy"] is None
    assert stats["type_accuracy"] is None
    report = (tmp_path / "phase5_review_accuracy_report.md").read_text(encoding="utf-8")
    assert "尚无人工裁决" in report
    json.loads((tmp_path / "phase5_review_stats.json").read_text(encoding="utf-8"))
    assert before == {path: _sha256(path) for path in watched}


def test_full_fill_computes_overall_and_stratified_rates(tmp_path: Path) -> None:
    path = _fixture(
        tmp_path,
        [
            # 6 correct, 1 wrong_type, 2 not_supported, 1 contradicted → 10 adjudicated
            ("REL-1", "critical", "high", "support", "correct", "correct"),
            ("REL-2", "critical", "high", "support", "correct", "correct"),
            ("REL-3", "high", "medium", "associated", "correct", "correct"),
            ("REL-4", "low", "low", "associated", "correct", "correct"),
            ("REL-5", "low", "low", "support", "correct", "correct"),
            ("REL-6", "low", "low", "support", "correct", "correct"),
            ("REL-7", "high", "medium", "associated", "wrong_type", "wrong_type", "建议改同属组织"),
            ("REL-8", "critical", "high", "insufficient", "not_supported", "not_supported"),
            ("REL-9", "critical", "high", "insufficient", "not_supported", "not_supported"),
            ("REL-10", "low", "low", "support", "correct", "contradicted"),
        ],
    )

    stats = analyze(path, tmp_path)

    assert stats["n_total"] == 10
    assert stats["n_adjudicated"] == 10
    assert stats["n_pending"] == 0
    assert stats["verdict_counts"] == {"correct": 6, "wrong_type": 1, "not_supported": 2, "contradicted": 1}
    assert stats["pair_accuracy"] == 0.7  # (6+1)/10
    assert stats["type_accuracy"] == 0.6  # 6/10
    assert stats["not_supported_rate"] == 0.2
    assert stats["contradicted_rate"] == 0.1
    assert stats["ai_agreement_rate"] == 0.9  # 9/10 与 AI 建议一致
    critical = stats["by_risk"]["critical"]
    assert critical["n"] == 4
    assert critical["verdict_counts"] == {"correct": 2, "not_supported": 2}
    assert critical["pair_accuracy"] == 0.5
    insufficient = stats["by_support"]["insufficient"]
    assert insufficient["n"] == 2
    assert insufficient["verdict_counts"] == {"not_supported": 2}
    report = (tmp_path / "phase5_review_accuracy_report.md").read_text(encoding="utf-8")
    assert "70.0%" in report  # 关系成立率
    assert "critical" in report


def test_partial_fill_excludes_pending_from_denominators(tmp_path: Path) -> None:
    path = _fixture(
        tmp_path,
        [
            ("REL-1", "critical", "high", "support", "correct", "correct"),
            ("REL-2", "critical", "high", "support", "correct", ""),
            ("REL-3", "low", "low", "insufficient", "not_supported", ""),
        ],
    )

    stats = analyze(path, tmp_path)

    assert stats["n_adjudicated"] == 1
    assert stats["n_pending"] == 2
    assert stats["pair_accuracy"] == 1.0
    assert stats["by_risk"]["critical"]["n"] == 1


def test_invalid_verdict_is_rejected_with_row_ids(tmp_path: Path) -> None:
    path = _fixture(
        tmp_path,
        [
            ("REL-1", "low", "low", "support", "correct", "correct"),
            ("REL-2", "low", "low", "support", "correct", "大概对"),
        ],
    )

    try:
        analyze(path, tmp_path)
    except ValueError as error:
        assert "REL-2" in str(error)
        assert "大概对" in str(error)
    else:
        raise AssertionError("invalid human_verdict must raise")


def test_revision_rules_generated_only_above_threshold(tmp_path: Path) -> None:
    verdicts = [("REL-1", "critical", "high", "insufficient", "not_supported", "not_supported")]
    verdicts += [
        (f"REL-{index}", "critical", "high", "insufficient", "not_supported", "not_supported")
        for index in range(2, 5)
    ]
    # low 层只有 1 条错误，不触发规则
    verdicts.append(("REL-5", "low", "low", "support", "correct", "not_supported"))
    # wrong_type 缺建议类型 → 警告
    verdicts.append(("REL-6", "high", "medium", "associated", "wrong_type", "wrong_type", ""))
    path = _fixture(tmp_path, verdicts)

    stats = analyze(path, tmp_path)

    rules = stats["revision_rules"]
    assert any(rule["dimension"] == "risk" and rule["value"] == "critical" for rule in rules)
    assert not any(rule["dimension"] == "risk" and rule["value"] == "low" for rule in rules)
    warnings = stats["warnings"]
    assert any("REL-6" in warning for warning in warnings)
    report = (tmp_path / "phase5_review_accuracy_report.md").read_text(encoding="utf-8")
    assert "critical" in report


def test_analysis_is_deterministic_on_real_package(tmp_path: Path) -> None:
    first = tmp_path / "a"
    second = tmp_path / "b"
    analyze(PACKAGE_PATH, first)
    analyze(PACKAGE_PATH, second)

    for name in ("phase5_review_accuracy_report.md", "phase5_review_stats.json"):
        assert _sha256(first / name) == _sha256(second / name)
