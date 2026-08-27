"""第四批A候选审核包结构守门测试。

审计对象：5 个事件（EVT-00001/00004/00005/00017/00029）现有 14 条疑似错挂证据。
守门范围：审计表覆盖恰好 14 条证据无重复无遗漏；凭据一一对应；枚举与
pending_human_review 合法；报告含 5 项事件与"待人工决定"签核表；禁止越权结论。
"""

from __future__ import annotations

import csv
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
REPORTS = REPO_ROOT / "research" / "drafts" / "reports"
AUDIT_CSV = REPORTS / "phase2_batch4a_existing_evidence_audit.csv"
RECEIPTS_CSV = REPORTS / "phase2_batch4a_access_receipts.csv"
DECISIONS_MD = REPORTS / "phase2_batch4a_event_decisions.md"

TARGET_EVENT_IDS = {"EVT-00001", "EVT-00004", "EVT-00005", "EVT-00017", "EVT-00029"}

ALLOWED_VERDICTS = {
    "supports_current_event",
    "supports_different_event",
    "wrong_date",
    "irrelevant",
    "insufficient_context",
    "inaccessible",
}
ALLOWED_ACTIONS = {
    "keep_candidate",
    "relink_candidate",
    "correct_event_candidate",
    "remove_candidate",
    "defer",
}
ALLOWED_CONFIDENCE = {"high", "medium", "low"}

FORBIDDEN_CLAIMS = ["已合并", "已删除", "人工审核通过", "人工复核通过", "已转正", "已执行"]

# 第四批A审计做出结论时的基线提交（审计对象在此提交的生产表中恰为 14 条）。
AUDIT_BASELINE_COMMIT = "2a87781"


def _read(path: Path) -> list[dict[str, str]]:
    with open(path, encoding="utf-8-sig", newline="") as fh:
        return list(csv.DictReader(fh))


def _production_target_evidence_ids() -> set[str]:
    """从审计基线提交 2a87781 推导 5 事件当时的证据 ID 集合。

    第四批A执行后，审计对象中两行证据已从生产表物理移除、其余状态已按人工
    批准改变；本审计包的历史快照语义必须锚定到做出这些结论时的提交，
    而非随后的实时生产表。
    """
    import subprocess
    import io

    facts_path = REPO_ROOT / "data" / "processed" / "fact_evidences.csv"
    _ = facts_path
    blob = subprocess.run(
        ["git", "show", f"{AUDIT_BASELINE_COMMIT}:data/processed/fact_evidences.csv"],
        cwd=REPO_ROOT,
        check=True,
        capture_output=True,
    ).stdout.decode("utf-8-sig")
    return {
        row["evidence_id"]
        for row in csv.DictReader(io.StringIO(blob))
        if row["subject_type"] == "event" and row["subject_id"] in TARGET_EVENT_IDS
    }


def _load_audit() -> tuple[list[dict[str, str]], list[str]]:
    assert AUDIT_CSV.exists(), f"缺少审计表：{AUDIT_CSV}"
    with open(AUDIT_CSV, encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        rows = list(reader)
        fieldnames = list(reader.fieldnames or [])
    return rows, fieldnames


def test_audit_covers_exactly_the_14_production_evidence_rows() -> None:
    """审计表恰好覆盖 5 事件的全部现有证据 ID：无重复、无遗漏、无多余。"""
    rows, _ = _load_audit()
    expected = _production_target_evidence_ids()
    audited = [row["evidence_id"] for row in rows]
    assert len(rows) == 14, f"审计行数应为 14，实际 {len(rows)}"
    assert len(set(audited)) == len(audited), "审计表存在重复 evidence_id"
    assert set(audited) == expected, (
        f"覆盖不一致：缺失 {sorted(expected - set(audited))}，多余 {sorted(set(audited) - expected)}"
    )
    assert {row["event_id"] for row in rows} == TARGET_EVENT_IDS


def test_audit_csv_fields_and_enums_are_valid() -> None:
    """必填字段齐备、枚举合法、review_status 恒为 pending_human_review。"""
    rows, fieldnames = _load_audit()
    required = [
        "audit_id",
        "evidence_id",
        "event_id",
        "source_id",
        "original_event_claim",
        "source_locator",
        "context_excerpt",
        "verdict",
        "recommended_action",
        "relink_target",
        "confidence",
        "receipt_id",
        "review_status",
        "review_note",
    ]
    for column in required:
        assert column in fieldnames, f"审计表缺少必需列：{column}"
    audit_ids = [row["audit_id"] for row in rows]
    assert len(set(audit_ids)) == len(audit_ids), "audit_id 重复"

    for row in rows:
        assert row["verdict"] in ALLOWED_VERDICTS, f"{row['audit_id']} verdict 非法：{row['verdict']}"
        assert row["recommended_action"] in ALLOWED_ACTIONS, (
            f"{row['audit_id']} action 非法：{row['recommended_action']}"
        )
        assert row["review_status"] == "pending_human_review", (
            f"{row['audit_id']} review_status 必须为 pending_human_review"
        )
        assert row["confidence"] in ALLOWED_CONFIDENCE, f"{row['audit_id']} confidence 非法"
        assert row["context_excerpt"].strip(), f"{row['audit_id']} context_excerpt 为空"
        assert row["review_note"].strip(), f"{row['audit_id']} review_note 为空"
        assert row["original_event_claim"].strip(), f"{row['audit_id']} original_event_claim 为空"
        # 改挂目标必须是生产表中真实存在的事件 ID，或留空。
        if row["relink_target"].strip():
            events_path = REPO_ROOT / "data" / "processed" / "events.csv"
            with open(events_path, encoding="utf-8-sig", newline="") as fh:
                known = {e["event_id"] for e in csv.DictReader(fh)}
            assert row["relink_target"] in known, (
                f"{row['audit_id']} relink_target 不在生产事件表：{row['relink_target']}"
            )
        elif row["recommended_action"] == "relink_candidate":
            assert "留空" in row["review_note"] or "无对应" in row["review_note"] or "编造" in row["review_note"], (
                f"{row['audit_id']} relink 留空必须在 review_note 说明理由"
            )


def test_receipts_pair_one_to_one_with_audit_rows() -> None:
    """凭据表恰好 14 行、receipt_id 唯一且与审计记录一一对应。"""
    assert RECEIPTS_CSV.exists(), f"缺少访问凭据表：{RECEIPTS_CSV}"
    receipts = _read(RECEIPTS_CSV)
    rows, _ = _load_audit()

    assert len(receipts) == 14, f"凭据应为 14 行，实际 {len(receipts)}"
    receipt_ids = [r["receipt_id"] for r in receipts]
    assert len(set(receipt_ids)) == len(receipt_ids), "receipt_id 重复"
    assert set(receipt_ids) == {row["receipt_id"] for row in rows}, "凭据与审计记录未一一对应"
    by_evidence = {r["evidence_id"]: r for r in receipts}
    assert set(by_evidence) == {row["evidence_id"] for row in rows}

    required = [
        "receipt_id",
        "evidence_id",
        "access_method",
        "target",
        "checked_at",
        "access_status",
        "locator",
        "context_sha256",
        "finding",
    ]
    with open(RECEIPTS_CSV, encoding="utf-8-sig", newline="") as fh:
        fieldnames = list(csv.DictReader(fh).fieldnames or [])
    for column in required:
        assert column in fieldnames, f"凭据表缺少必需列：{column}"

    import re

    for r in receipts:
        assert r["access_method"] in {"local_text", "web_url"}, f"{r['receipt_id']} access_method 非法"
        assert r["access_status"] in {"ok", "inaccessible"}, f"{r['receipt_id']} access_status 非法"
        assert r["checked_at"].strip() and r["target"].strip() and r["locator"].strip()
        assert re.fullmatch(r"[0-9a-f]{64}", r["context_sha256"]), f"{r['receipt_id']} sha256 格式非法"
        if r["access_method"] == "local_text":
            assert Path(r["target"]).is_absolute(), f"{r['receipt_id']} 本地凭据必须记绝对路径"
        else:
            assert r["target"].startswith("http"), f"{r['receipt_id']} 网页凭据必须记原URL"


def test_decisions_report_contains_five_events_and_pending_signoff() -> None:
    """报告必须含 5 项事件小节与 5 行签核表，状态全部「待人工决定」。"""
    assert DECISIONS_MD.exists(), f"缺少事件级建议报告：{DECISIONS_MD}"
    text = DECISIONS_MD.read_text(encoding="utf-8")
    for event_id in sorted(TARGET_EVENT_IDS):
        assert event_id in text, f"报告缺少事件 {event_id}"
    assert text.count("待人工决定") >= 5, "签核表必须至少 5 处『待人工决定』"
    signoff_rows = [ln for ln in text.splitlines() if ln.startswith("| ") and "待人工决定" in ln]
    assert len(signoff_rows) == 5, f"签核表应恰为 5 行，实际 {len(signoff_rows)}"
    for claim in FORBIDDEN_CLAIMS:
        assert claim not in text, f"报告出现越权结论表述：{claim}"


def test_no_overreaching_claims_in_any_batch4a_artifact() -> None:
    """三份交付物均不得声称已合并/已删除/人工复核通过等越权结论。"""
    for path in (AUDIT_CSV, RECEIPTS_CSV, DECISIONS_MD):
        text = path.read_text(encoding="utf-8-sig")
        for claim in FORBIDDEN_CLAIMS:
            assert claim not in text, f"{path.name} 出现越权结论表述：{claim}"


def test_acceptance_guardrails_on_relink_and_support_semantics() -> None:
    """验收整改守门：限定三条易夸大证据的建议语义与一条措辞（2026-08-27 验收）。

    - AUD-B4A-010 为庭审材料、AUD-B4A-011 为资料目录：均不得建议改挂
      EVT-00008“左联五烈士遇难”——只有 AUD-B4A-012 可指向该遇难事实；
    - AUD-B4A-014 仅证明“就读”，不得以 supports_current_event 背书含
      “抵沪”的完整事件名；
    - AUD-B4A-007 来自后出工具书《左联词典》，不得宣称“最直接的同期文献”。
    """
    audit_rows, _ = _load_audit()
    rows = {row["audit_id"]: row for row in audit_rows}

    r10, r11 = rows["AUD-B4A-010"], rows["AUD-B4A-011"]
    assert r10["relink_target"] != "EVT-00008", "庭审材料不得改挂到『左联五烈士遇难』"
    assert r11["relink_target"] != "EVT-00008", "资料目录不得改挂到『左联五烈士遇难』"
    if r10["recommended_action"] == "relink_candidate":
        assert r10["relink_target"] == "" and "留空" in r10["review_note"], "010 改挂目标必须留空并说明"
    assert r11["verdict"] == "irrelevant" and r11["recommended_action"] == "defer", (
        "011 目录条目应判为无关并暂缓，不得建议改挂"
    )

    r12 = rows["AUD-B4A-012"]
    assert r12["relink_target"] == "EVT-00008", "仅 012 直接支持遇难事实，应指向 EVT-00008"

    r14 = rows["AUD-B4A-014"]
    assert r14["verdict"] != "supports_current_event", "014 不得背书含‘抵沪’的完整事件"
    assert r14["verdict"] == "insufficient_context", "014 应判上下文不足（仅证就读，未证抵沪）"
    assert r14["recommended_action"] in {"correct_event_candidate", "defer"}, (
        "014 应建议收窄事件名或暂缓"
    )

    r07 = rows["AUD-B4A-007"]
    assert "最直接" not in r07["review_note"], "007 不得宣称来自后出词典的内容为最直接的同期文献"

    # 报告正文同步不得残留同样夸大表述。
    md = DECISIONS_MD.read_text(encoding="utf-8")
    assert "成立大会最直接的同期文献" not in md
