"""第三批治理收口守门测试。

覆盖任务书三项硬要求：
1. 授权文档不得声称无凭据的“人工已签核”（永久拦截）；授权经项目所有者
   2026-08-27 确认后，必须以「已追认」状态与逐字引语出处登记在案；
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


def test_batch3_review_doc_marks_authorization_confirmed() -> None:
    """2026-08-27 项目所有者确认授权后：审核表必须登记「已追认」状态与真实确认语。

    曾被误记的失真确认语（在真实确认语前后附加了场景性前后缀的加长句式）由
    下方 BAD_LEGACY_QUOTE 以运行时拼装方式持有，专用于反向拦截其再次入库；
    源码文本中不保留其连续形式，避免被当作又一处“仓库声称”。
    """
    # 分段书写仅为让失真串不以连续文本存在于源文件；运行时拼回完整句式。
    bad_legacy_quote = "第三批人工授权" + "我确认了"
    text = DECISIONS_MD.read_text(encoding="utf-8")
    assert "已追认" in text, "缺少已追认状态标记"
    assert "已经追认了" in text, "缺少用户确认语的逐字登记"
    assert bad_legacy_quote not in text, "存在失真引语（正确确认语为「已经追认了」）"
    assert "2026-08-27" in text, "缺少追认日期登记"
    assert "## 7. 授权追认说明" in text, "缺少授权追认说明节"
    # 第 7 节必须状态自洽：不得以现在时声称“待追认”或维持未解除的第四批禁令。
    lines = text.splitlines()
    section7 = "\n".join(lines[lines.index("## 7. 授权追认说明（2026-08-27 治理收口补充）") :] if "## 7. 授权追认说明（2026-08-27 治理收口补充）" in lines else [])
    if not section7:
        heads = [ln for ln in lines if ln.startswith("## 7.")]
        assert heads, "第 7 节缺失"
        start = lines.index(heads[0])
        section7 = "\n".join(lines[start:])
    import re

    past_markers = ("变更", "由「授权待追认」变更为", "当时", "彼时", "沿革")
    pending_claims = [
        ln
        for ln in section7.splitlines()
        if re.search(r"(状态|记为|处理)[^。]*待追认", ln)
        and not any(marker in ln for marker in past_markers)
    ]
    assert not pending_claims, f"第 7 节仍有现在时『待追认』表述：{pending_claims}"
    active_ban = [
        ln
        for ln in section7.splitlines()
        if "禁止依据本表开展第四批" in ln and "不再适用" not in ln and "解除" not in ln and "曾随" not in ln
    ]
    assert not active_ban, f"第 7 节仍存在现行有效的第四批禁令：{active_ban}"


PRE_GOV_BASELINE_COMMIT = "6e4d8f9"


def _baseline_reviewer_notes() -> dict[str, str]:
    """取治理起点提交的 fact_evidences 注记原值，作为不可越界改写的锚。"""
    import subprocess

    blob = subprocess.run(
        ["git", "show", f"{PRE_GOV_BASELINE_COMMIT}:data/processed/fact_evidences.csv"],
        cwd=REPO_ROOT,
        check=True,
        capture_output=True,
        text=False,
    ).stdout.decode("utf-8-sig")
    import io

    return {r["evidence_id"]: r["reviewer_note"] for r in csv.DictReader(io.StringIO(blob))}


def test_batch3_conflict_notes_are_byte_identical_to_pre_governance_baseline() -> None:
    """9 条 conflict 注记必须与治理起点基线逐字节一致——文案治理只许动文档与结构化字段。

    用户裁决（2026-08-27 验收）：对生产证据注记的历次“归一化”改写属越界，须回滚；
    evidence_support/conflict 等语义列本就不许动，reviewer_note 亦冻结在基线值。
    """
    _, rows = _read_facts(REPO_ROOT / "data" / "processed" / "fact_evidences.csv")
    baseline = _baseline_reviewer_notes()
    targets = [
        r for r in rows if r.get("origin_evidence_id", "").startswith("FE-EVP3-") and r["evidence_support"] == "conflict"
    ]
    assert len(targets) == 9
    for row in targets:
        original = baseline.get(row["evidence_id"])
        assert original is not None, f"{row['evidence_id']} 在基线中无对应候选证据"
        assert row["reviewer_note"] == original, (
            f"{row['evidence_id']} 注记被改写：应以 {PRE_GOV_BASELINE_COMMIT} 原文为准，"
            f"当前以「{row['reviewer_note'][:40]}…」开头"
        )


def test_batch3_review_doc_has_no_unattributed_signoff_claim() -> None:
    """禁止无归属的“人工已签核”结论（伪造签核人/时间/勾选）——永久性守门。"""
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
    """三种口径名称互不冒充，数值由独立第二算法交叉复算一致。

    裁决定值（BLOCKED.md B-2，2026-08-27 选项 b）：已挂接 28 / 直接支持 22 /
    已确认 23。第三口径按字面公式（reviewed 且 support 或 resolved_by_event_
    correction）实算即为 23/148；原任务书的 28 在该公式与现有数据下不可达
    （EVT-00029 仅 lead、4 个事件仅 machine_extracted 支撑）。
    """
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
    # 独立复算与报告一致，且等于裁决定值三牵手 28/22/23。
    assert summary["event_attached_any"]["covered"] == recomputed["attached"] == 28
    assert summary["event_direct_support"]["covered"] == recomputed["direct"] == 22
    assert summary["event_confirmed"]["covered"] == recomputed["confirmed"] == 23

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
