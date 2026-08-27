"""第三批治理收口守门测试。

覆盖任务书三项硬要求：
1. 授权文档不得声称无凭据的“人工已签核”，必须存在「授权待追认」状态；
2. fact_evidences 必须具备结构化裁决列 adjudication_status（FE-EVP3-* 的
   conflict 行＝resolved_by_event_correction，其余为空），Schema 拒绝非法值；
3. 覆盖率报告必须区分三种事件口径且数值由数据独立复算（不得把“已挂接”
   冒充“直接支持”）。
"""

from __future__ import annotations

import csv
import shutil
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
DECISIONS_MD = REPO_ROOT / "research" / "drafts" / "reports" / "phase2_batch3_event_review_decisions.md"

ALLOWED_ADJUDICATION = {"", "resolved_by_event_correction"}
BATCH3_ORIGIN_PREFIX = "FE-EVP3-"


def _read_facts(path: Path) -> tuple[list[str], list[dict[str, str]]]:
    with open(path, encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        return list(reader.fieldnames or []), list(reader)


# ---------------------------------------------------------------------------
# 任务1：授权记录拆分
# ---------------------------------------------------------------------------


def test_batch3_review_doc_marks_authorization_pending() -> None:
    """审核表必须把技术执行与人工授权拆开，并给出「授权待追认」状态。"""
    text = DECISIONS_MD.read_text(encoding="utf-8")
    assert "授权待追认" in text, "缺少授权待追认状态标记"
    assert "技术上已执行" in text or "技术执行" in text, "缺少技术执行口径说明"
    assert "人工授权凭据待补录" in text, "缺少人工授权凭据待补录说明"
    assert "禁止依据本表开展第四批及后续任何批次的物理合并" in text, "缺少第四批物理合并禁令"


def test_batch3_review_doc_has_no_unattributed_signoff_claim() -> None:
    """禁止无归属的“人工已签核”结论（伪造签核人/时间/勾选）。"""
    text = DECISIONS_MD.read_text(encoding="utf-8")
    forbidden = [
        "# Phase 2 第三批事件级人工审核表（已签核并执行）",
        "**人工已签核并已执行完毕。**",
        "| 签核意见 |",
        "| 签核日期",
        "✅ 已完成：16 项均通过",
    ]
    for pattern in forbidden:
        assert pattern not in text, f"仍存在无归属签核表述：{pattern}"


# ---------------------------------------------------------------------------
# 任务2：结构化裁决语义
# ---------------------------------------------------------------------------


def test_processed_fact_evidences_have_structured_adjudication_status() -> None:
    """第三批 9 条 conflict 必须携带 resolved_by_event_correction，其余行为空。"""
    fields, rows = _read_facts(REPO_ROOT / "data" / "processed" / "fact_evidences.csv")
    assert "adjudication_status" in fields, "fact_evidences.csv 缺少 adjudication_status 列"

    conflicts = [r for r in rows if r.get("evidence_support") == "conflict"]
    resolved = [
        r
        for r in rows
        if r.get("origin_evidence_id", "").startswith(BATCH3_ORIGIN_PREFIX)
        and r["evidence_support"] == "conflict"
    ]
    assert len(resolved) > 0, "未找到任何第三批 conflict 证据"

    for row in rows:
        is_batch3_conflict = (
            row.get("origin_evidence_id", "").startswith(BATCH3_ORIGIN_PREFIX)
            and row["evidence_support"] == "conflict"
        )
        expected = "resolved_by_event_correction" if is_batch3_conflict else ""
        actual = row.get("adjudication_status", "__MISSING__")
        assert actual == expected, f"{row['evidence_id']}: 期望 {expected!r}，实际 {actual!r}"
        assert actual in ALLOWED_ADJUDICATION, f"{row['evidence_id']}: 非法裁决值 {actual!r}"

    assert len({r["evidence_id"] for r in conflicts}) == len(conflicts)
    assert sum(1 for r in rows if r["adjudication_status"] == "resolved_by_event_correction") == len(
        resolved
    ), "非第三批 conflict 行被错误写入裁决值"


def test_kb_schema_rejects_illegal_adjudication_status(tmp_path: Path) -> None:
    """Schema 必须拒绝空值与 resolved_by_event_correction 之外的取值。"""
    publish_like = tmp_path / "processed_copy"
    shutil.copytree(REPO_ROOT / "data" / "processed", publish_like)

    from kb_schema import validate_data_dir

    baseline = validate_data_dir(publish_like)
    baseline_errors = [i.code for i in baseline.errors]
    assert not any("adjudication" in code for code in baseline_errors)

    facts_path = publish_like / "fact_evidences.csv"
    fields, rows = _read_facts(facts_path)
    rows[0]["adjudication_status"] = "manually_signed_off"
    with open(facts_path, "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)

    corrupted = validate_data_dir(publish_like)
    assert any(issue.code == "invalid_fact_adjudication_status" for issue in corrupted.errors), (
        "Schema 未拦截非法裁决值"
    )


# ---------------------------------------------------------------------------
# 任务2/3：覆盖率三口径与发布层一致性
# ---------------------------------------------------------------------------


def _recompute_event_metrics(facts_path: Path, events_path: Path) -> dict[str, int]:
    """与报告实现相互独立的第二算法，用于交叉复算三种事件口径。"""
    _, facts = _read_facts(facts_path)
    with open(events_path, encoding="utf-8-sig") as fh:
        events = [r["event_id"] for r in csv.DictReader(fh)]

    event_ids = set(events)
    attached: set[str] = set()
    direct: set[str] = set()
    confirmed: set[str] = set()
    for row in facts:
        if row["predicate"] != "event_occurrence" or row["subject_id"] not in event_ids:
            continue
        subject = row["subject_id"]
        attached.add(subject)
        if row["evidence_support"] == "support":
            direct.add(subject)
        if row["review_status"] == "reviewed" and (
            row["evidence_support"] == "support"
            or row.get("adjudication_status", "") == "resolved_by_event_correction"
        ):
            confirmed.add(subject)
    return {
        "attached": len(attached),
        "direct": len(direct),
        "confirmed": len(confirmed),
        "total": len(events),
    }


def test_coverage_report_distinguishes_three_event_metrics(sandbox_tmp_path: Path) -> None:
    """三种口径名称互不冒充，数值由独立第二算法交叉复算一致。"""
    from research.analysis.report_evidence_coverage import report_evidence_coverage

    summary = report_evidence_coverage(
        REPO_ROOT / "data" / "processed",
        sandbox_tmp_path / "coverage.md",
        sandbox_tmp_path / "queue.csv",
    )

    for key in ("event_attached_any", "event_direct_support", "event_confirmed"):
        assert key in summary, f"覆盖率报告缺少口径 {key}"

    recomputed = _recompute_event_metrics(
        REPO_ROOT / "data" / "processed" / "fact_evidences.csv",
        REPO_ROOT / "data" / "processed" / "events.csv",
    )
    assert summary["event_attached_any"]["covered"] == recomputed["attached"]
    assert summary["event_direct_support"]["covered"] == recomputed["direct"]
    assert summary["event_confirmed"]["covered"] == recomputed["confirmed"]

    assert summary["event_attached_any"]["total"] == recomputed["total"]
    # 已挂接 ≥ 直接支持：支持族 ⊆ 全部挂接证据。
    # 注意“已确认”不必 ≤ “直接支持”：由 conflict 裁决转正的事件属于确认口径
    # 但不属于 support 族（如第三批 6 个冲突事件中的 5 个）。
    assert summary["event_attached_any"]["covered"] >= summary["event_direct_support"]["covered"]
    assert summary["event_attached_any"]["covered"] >= summary["event_confirmed"]["covered"]


def test_publish_layer_matches_processed_adjudication_semantics() -> None:
    """发布层是过滤后的公开子集：凡发布的行，裁决状态必须与生产层同值；
    全部 9 条 resolved_by_event_correction 必须进入发布层（均为 reviewed）。"""
    _, processed = _read_facts(REPO_ROOT / "data" / "processed" / "fact_evidences.csv")
    _, published = _read_facts(REPO_ROOT / "data" / "publish" / "fact_evidences.csv")

    assert len(published) <= len(processed), "发布层数量不得超过生产层"
    published_by_id = {r["evidence_id"]: r for r in published}
    for p_row in processed:
        b_row = published_by_id.get(p_row["evidence_id"])
        if b_row is None:
            continue
        assert p_row.get("adjudication_status") == b_row.get("adjudication_status"), (
            f"{p_row['evidence_id']} 裁决状态不一致："
            f"{p_row.get('adjudication_status')!r} vs {b_row.get('adjudication_status')!r}"
        )

    resolved_ids = {
        r["evidence_id"]
        for r in processed
        if r.get("adjudication_status") == "resolved_by_event_correction"
    }
    assert len(resolved_ids) == 9
    missing_in_publish = resolved_ids - set(published_by_id)
    assert not missing_in_publish, f"已裁决证据未进发布层：{sorted(missing_in_publish)}"
