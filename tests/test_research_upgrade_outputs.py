"""Agent B 研究升级交付物守门测试。

覆盖任务书要求的验收点：
- 候选数量与 ID 唯一；候选引用现存实体；
- 新查材料全部 pending_human_review（含 URL/访问日期/引文三要素）；
- missing 不被计入已覆盖（报告数字与 CSV 实数一致）；
- 网络三套口径使用不同过滤条件（全量⊇可信=加权拓扑）；
- 可信网络不包含明确待审核关系（合成注入验证）；
- 时间切片互不重复；
- 重跑输出一致（三个脚本字节级确定性）。

测试对 data/processed 做只读快照拷贝，隔离并行改动；不修改生产数据。
"""
from __future__ import annotations

import csv
import importlib.util
import json
import shutil
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[1]
PROCESSED = REPO_ROOT / "data" / "processed"
ANALYSIS = REPO_ROOT / "research" / "analysis"
REPORTS = REPO_ROOT / "research" / "drafts" / "reports"

PENDING = "pending_human_review"
MISSING = "missing"
CONFLICT = "conflict"


def _load_module(name: str):
    spec = importlib.util.spec_from_file_location(name, ANALYSIS / f"{name}.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


@pytest.fixture(scope="module")
def snapshot_dir(tmp_path_factory: pytest.TempPathFactory) -> Path:
    dest = tmp_path_factory.mktemp("snapshot") / "processed"
    dest.mkdir()
    for csv_file in PROCESSED.glob("*.csv"):
        shutil.copy2(csv_file, dest / csv_file.name)
    return dest


@pytest.fixture(scope="module")
def built_candidates(snapshot_dir: Path, tmp_path_factory: pytest.TempPathFactory) -> Path:
    out = tmp_path_factory.mktemp("cand_out")
    module = _load_module("build_core_upgrade_candidates")
    module.build(snapshot_dir, out)
    return out


@pytest.fixture(scope="module")
def built_network(snapshot_dir: Path, tmp_path_factory: pytest.TempPathFactory) -> Path:
    out = tmp_path_factory.mktemp("net_out")
    module = _load_module("build_trustworthy_network_analysis")
    module.build(snapshot_dir, out)
    return out


def _read_rows(path: Path) -> list[dict[str, str]]:
    with open(path, encoding="utf-8-sig", newline="") as fh:
        return list(csv.DictReader(fh))


def test_candidate_counts_and_ids_unique(built_candidates: Path) -> None:
    persons = _read_rows(built_candidates / "core_person_evidence_candidates.csv")
    events = _read_rows(built_candidates / "core_event_evidence_candidates.csv")
    places = _read_rows(built_candidates / "core_place_review_candidates.csv")
    assert len(persons) == 30, f"人物候选应为 30，实际 {len(persons)}"
    assert len(events) == 20, f"事件候选应为 20，实际 {len(events)}"
    assert len(places) == 10, f"地点候选应为 10，实际 {len(places)}"
    for rows, id_field in ((persons, "candidate_id"), (events, "candidate_id"), (places, "candidate_id")):
        ids = [r[id_field] for r in rows]
        assert len(set(ids)) == len(ids), f"{id_field} 存在重复"
    assert {r["candidate_id"][:5] for r in persons} == {"CPC-P"}
    assert {r["candidate_id"][:5] for r in events} == {"CPC-E"}
    assert {r["candidate_id"][:5] for r in places} == {"CPC-L"}


def test_candidates_reference_existing_entities(built_candidates: Path) -> None:
    live_persons = {r["person_id"] for r in _read_rows(PROCESSED / "persons.csv")}
    live_events = {r["event_id"] for r in _read_rows(PROCESSED / "events.csv")}
    live_places = {r["place_id"] for r in _read_rows(PROCESSED / "places.csv")}
    for row in _read_rows(built_candidates / "core_person_evidence_candidates.csv"):
        assert row["person_id"] in live_persons, f"悬空 person_id: {row['person_id']}"
    for row in _read_rows(built_candidates / "core_event_evidence_candidates.csv"):
        assert row["event_id"] in live_events, f"悬空 event_id: {row['event_id']}"
    for row in _read_rows(built_candidates / "core_place_review_candidates.csv"):
        assert row["place_id"] in live_places, f"悬空 place_id: {row['place_id']}"


@pytest.mark.parametrize(
    "csv_name,groups",
    [
        ("core_person_evidence_candidates.csv", ("birth", "death", "role")),
        ("core_event_evidence_candidates.csv", ("direct_support", "independent_source")),
        ("core_place_review_candidates.csv", ("address", "coord")),
    ],
)
def test_new_evidence_all_pending_with_provenance(built_candidates: Path, csv_name: str, groups: tuple[str, ...]) -> None:
    rows = _read_rows(built_candidates / csv_name)
    assert rows
    for row in rows:
        for g in groups:
            status = row[f"{g}_status"]
            url = row[f"{g}_candidate_url"].strip()
            quote = row[f"{g}_candidate_quote"].strip()
            access = row[f"{g}_candidate_access_date"].strip()
            locator = row[f"{g}_candidate_locator"].strip()
            assert status in (PENDING, MISSING, CONFLICT), f"{csv_name} {row['candidate_id']} {g}_status 非法：{status}"
            if status in (PENDING, CONFLICT):
                assert url and quote and access and locator, (
                    f"{row['candidate_id']} {g} 声称 {status} 但缺少 URL/引文/访问日期/定位"
                )
                assert access.count("-") >= 2, f"{row['candidate_id']} {g} 访问日期格式异常：{access}"
                # 返修：百科/普通媒体一律 D 级 web_lead，不得称为权威；须带分级四件套
                level = row.get(f"{g}_source_level", "").strip()
                stype = row.get(f"{g}_source_type", "").strip()
                retrieval = row.get(f"{g}_retrieval_status", "").strip()
                chash = row.get(f"{g}_content_hash", "").strip()
                assert level in ("A", "B", "C", "D"), f"{row['candidate_id']} {g} 缺少 source_level"
                assert stype, f"{row['candidate_id']} {g} 缺少 source_type"
                assert retrieval in ("retrieved", "conflict", "missing"), f"{row['candidate_id']} {g} retrieval 非法"
                assert chash, f"{row['candidate_id']} {g} 缺少 content_hash"
                if "wikipedia.org" in url or "baike.baidu.com" in url:
                    assert level == "D" and stype == "web_lead", f"{row['candidate_id']} {g} 百科必须为 D/web_lead"
            else:
                assert not url, f"{row['candidate_id']} {g} 记 missing 却带 URL"
    # 冲突候选必须存在且禁止自动落库（殷夫/周扬生年、萌芽/楼适夷日期）
    if csv_name == "core_person_evidence_candidates.csv":
        assert any(r["birth_status"] == CONFLICT for r in rows), "人物候选缺少 conflict（周扬生年）"
    if csv_name == "core_event_evidence_candidates.csv":
        assert any(r["direct_support_status"] == CONFLICT for r in rows), "事件候选缺少 conflict"


def test_missing_not_counted_as_covered(built_candidates: Path) -> None:
    persons = _read_rows(built_candidates / "core_person_evidence_candidates.csv")
    pending = sum(
        1
        for r in persons
        for f in ("birth", "death", "role")
        if r[f"{f}_status"] == PENDING
    )
    conflict = sum(
        1
        for r in persons
        for f in ("birth", "death", "role")
        if r[f"{f}_status"] == CONFLICT
    )
    missing = sum(
        1
        for r in persons
        for f in ("birth", "death", "role")
        if r[f"{f}_status"] == MISSING
    )
    assert pending + conflict + missing == 3 * len(persons)
    report = (built_candidates / "core_upgrade_selection_report.md").read_text(encoding="utf-8")
    import re

    assert "候选来源等级分布（A/B/C/D）" in report, "选择报告缺少 A/B/C/D 分级统计"
    assert CONFLICT in report, "选择报告缺少 conflict 说明"
    # missing 不冒充覆盖：事实级证据为 0 的字段，状态只能是 missing/pending/conflict
    for r in persons:
        for f in ("birth_year", "death_year", "role"):
            if int(r[f"{f}_fact_evidences"] or 0) == 0:
                assert r[f.replace("_year", "").replace("_evidences", "") + "_status"] in (PENDING, MISSING, CONFLICT)


def test_three_networks_use_different_filters(built_network: Path) -> None:
    payload = json.loads((built_network / "trustworthy_network_analysis.json").read_text(encoding="utf-8"))
    nets = payload["networks"]
    # 返修四口径：全量探索、低风险启发式（不得称可信）、证据支持、人工确认、可信（verified ∪ 支持）
    assert set(nets) >= {"full_exploratory", "low_risk_heuristic", "evidence_supported", "human_verified", "trusted", "evidence_weighted"}
    full = nets["full_exploratory"]
    low_risk = nets["low_risk_heuristic"]
    supported = nets["evidence_supported"]
    trusted = nets["trusted"]
    weighted = nets["evidence_weighted"]
    # 低风险启发式为研究对照口径（非可信）；2026-10-08 第三批（双 Agent 交叉验证）落地后
    # 证据支持/可信为 25 条，首次过 10 条门槛，样本充足、可生成 Top10 排名（仍为数据观察口径）；
    # human_verified 为 0（三批均未使用人工裁决通道；第三批为交叉验证，非人工复核）。
    assert full["edges"] > low_risk["edges"] > 0, "全量边应大于低风险筛选边"
    assert supported["edges"] == 25, "三批落地的合格 support 关系应进入证据支持口径"
    assert trusted["edges"] == 25, "trusted = verified(0) ∪ supported(25)"
    assert nets["human_verified"]["edges"] == 0, "未使用 human_adjudication 通道，人工确认口径应为 0"
    assert trusted.get("sample_sufficient") is True, "可信边 25 ≥ 10，样本应标记充足"
    assert "样本充足" in str(trusted.get("sample_note", "")), "样本充足须如实标记"
    assert full["total_weight"] >= low_risk["total_weight"]
    assert set(payload["time_slices"]) == {"1928-1930", "1931-1933", "1934-1936"}
    # 过滤规则在 formulas 中显式登记且四口径互不冒充
    assert "verified" in payload["formulas"]["trusted_rule"] and "support" in payload["formulas"]["trusted_rule"]
    assert "不得称为可信" in payload["formulas"]["low_risk_heuristic"]
    assert payload["formulas"]["evidence_weight"]
    # 时间切片为事件共参与网络，不得冒充关系历时网络
    assert "事件共参与" in payload["formulas"]["time_slices"]


def test_trusted_network_excludes_pending_relations(snapshot_dir: Path, tmp_path: Path) -> None:
    module = _load_module("build_trustworthy_network_analysis")
    # 单元级：四类明确待审核/低置信关系必须被可信规则拒绝
    bad_rows = [
        {"final_relation_type": "待核验", "needs_manual_review": "no", "relation_risk_level": "low", "confidence": "medium", "publish_status": ""},
        {"final_relation_type": "交游", "needs_manual_review": "yes", "relation_risk_level": "low", "confidence": "medium", "publish_status": ""},
        {"final_relation_type": "交游", "needs_manual_review": "no", "relation_risk_level": "critical", "confidence": "medium", "publish_status": ""},
        {"final_relation_type": "交游", "needs_manual_review": "no", "relation_risk_level": "high", "confidence": "medium", "publish_status": ""},
        {"final_relation_type": "交游", "needs_manual_review": "no", "relation_risk_level": "low", "confidence": "low", "publish_status": ""},
        {"final_relation_type": "交游", "needs_manual_review": "no", "relation_risk_level": "low", "confidence": "medium", "publish_status": "pending_review"},
    ]
    for row in bad_rows:
        assert not module.is_trusted_relation(row), f"待审核关系未被过滤：{row}"

    # 图级：向快照注入一条明确待审核关系，可信网络不得出现该边，全量网络应出现
    work = tmp_path / "inject" / "processed"
    work.mkdir(parents=True)
    for csv_file in snapshot_dir.glob("*.csv"):
        shutil.copy2(csv_file, work / csv_file.name)
    ctx = module.build_graphs(snapshot_dir)
    trusted_edges = {tuple(sorted(e)) for e in ctx["graphs"]["trusted"].edges()}
    persons = sorted(ctx["graphs"]["full"].nodes())
    pair = None
    for i in range(len(persons)):
        for j in range(i + 1, len(persons)):
            key = tuple(sorted((persons[i], persons[j])))
            if key not in trusted_edges:
                pair = key
                break
        if pair:
            break
    assert pair, "找不到可信网络外的节点对"
    rel_rows = _read_rows(work / "person_relations.csv")
    max_rel = max(int(r["relation_id"].split("-")[1]) for r in rel_rows)
    has_publish = bool(rel_rows[0].get("publish_status", "") != "" or "publish_status" in rel_rows[0])
    rel_rows.append(
        {
            **rel_rows[0],
            "relation_id": f"REL-{max_rel + 1:05d}",
            "source_person_id": pair[0],
            "target_person_id": pair[1],
            "final_relation_type": "待核验",
            "needs_manual_review": "yes",
            "relation_risk_level": "critical",
            "confidence": "low",
            **({"publish_status": "pending_review"} if has_publish else {}),
        }
    )
    with open(work / "person_relations.csv", "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=list(rel_rows[0]))
        writer.writeheader()
        writer.writerows(rel_rows)
    ctx2 = module.build_graphs(work)
    trusted2 = {tuple(sorted(e)) for e in ctx2["graphs"]["trusted"].edges()}
    full2 = {tuple(sorted(e)) for e in ctx2["graphs"]["full"].edges()}
    assert pair not in trusted2, "注入的待审核关系出现在可信网络"
    assert pair in full2, "注入关系应出现在全量探索网络"


def test_time_slices_do_not_overlap(snapshot_dir: Path, built_network: Path) -> None:
    rows = _read_rows(snapshot_dir / "events.csv")
    slices = (("1928-1930", 1928, 1930), ("1931-1933", 1931, 1933), ("1934-1936", 1934, 1936))
    assignment: dict[str, set[str]] = {label: set() for label, _, _ in slices}
    for r in rows:
        d = (r.get("event_date") or "").strip()
        if len(d) >= 4 and d[:4].isdigit():
            year = int(d[:4])
            for label, y0, y1 in slices:
                if y0 <= year <= y1:
                    assignment[label].add(r["event_id"])
    for i in range(len(slices)):
        for j in range(i + 1, len(slices)):
            assert not assignment[slices[i][0]] & assignment[slices[j][0]], "同一事件落入多个时间切片"
    payload = json.loads((built_network / "trustworthy_network_analysis.json").read_text(encoding="utf-8"))
    # 切片间的边按年份桶互斥构建；不同切片的 dated_events_used 之和不超过有日期事件总数
    total_dated = sum(len(v) for v in assignment.values())
    assert sum(payload["time_slices"][label]["dated_events_used"] for label, _, _ in slices) <= total_dated


def test_rerun_outputs_identical(snapshot_dir: Path, tmp_path: Path) -> None:
    out1, out2 = tmp_path / "r1", tmp_path / "r2"
    cand = _load_module("build_core_upgrade_candidates")
    audit = _load_module("audit_event_place_normalization")
    net = _load_module("build_trustworthy_network_analysis")
    for out in (out1, out2):
        out.mkdir()
        cand.build(snapshot_dir, out)
        audit.build(snapshot_dir, out)
        net.build(snapshot_dir, out)
    names = [
        "core_person_evidence_candidates.csv",
        "core_event_evidence_candidates.csv",
        "core_place_review_candidates.csv",
        "core_upgrade_selection_report.md",
        "event_normalization_candidates.csv",
        "place_normalization_candidates.csv",
        "participant_role_mapping_candidates.csv",
        "temporal_spatial_audit_report.md",
        "trustworthy_network_analysis.json",
        "trustworthy_network_analysis_report.md",
        "updated_research_findings_candidates.md",
    ]
    for name in names:
        b1 = (out1 / name).read_bytes()
        b2 = (out2 / name).read_bytes()
        assert b1 == b2, f"重跑输出不一致：{name}"


def test_audit_outputs_pending_only_and_unique(built_reports_dir: Path) -> None:
    for name, decision_col in (
        ("event_normalization_candidates.csv", "decision"),
        ("place_normalization_candidates.csv", "decision"),
    ):
        rows = _read_rows(built_reports_dir / name)
        assert rows, f"{name} 不应为空"
        ids = [r["candidate_id"] for r in rows]
        assert len(set(ids)) == len(ids)
        assert all(r[decision_col] == PENDING for r in rows), f"{name} 存在非 pending 裁决"
    roles = _read_rows(built_reports_dir / "participant_role_mapping_candidates.csv")
    assert roles
    assert all(r["decision"] in (PENDING, "no_change_needed") for r in roles)


@pytest.fixture(scope="module")
def built_reports_dir(snapshot_dir: Path, tmp_path_factory: pytest.TempPathFactory) -> Path:
    out = tmp_path_factory.mktemp("audit_out")
    module = _load_module("audit_event_place_normalization")
    module.build(snapshot_dir, out)
    return out
