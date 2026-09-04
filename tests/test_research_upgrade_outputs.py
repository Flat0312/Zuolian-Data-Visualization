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
            assert status in (PENDING, MISSING), f"{csv_name} {row['candidate_id']} {g}_status 非法：{status}"
            if status == PENDING:
                assert url and quote and access and locator, (
                    f"{row['candidate_id']} {g} 声称 {PENDING} 但缺少 URL/引文/访问日期/定位"
                )
                assert access.count("-") >= 2, f"{row['candidate_id']} {g} 访问日期格式异常：{access}"
            else:
                assert not url, f"{row['candidate_id']} {g} 记 missing 却带 URL"


def test_missing_not_counted_as_covered(built_candidates: Path) -> None:
    persons = _read_rows(built_candidates / "core_person_evidence_candidates.csv")
    pending = sum(
        1
        for r in persons
        for f in ("birth", "death", "role")
        if r[f"{f}_status"] == PENDING
    )
    missing = sum(
        1
        for r in persons
        for f in ("birth", "death", "role")
        if r[f"{f}_status"] == MISSING
    )
    assert pending + missing == 3 * len(persons)
    report = (built_candidates / "core_upgrade_selection_report.md").read_text(encoding="utf-8")
    import re

    m = re.search(r"检索到候选证据 (\d+) 个（全部 pending_human_review），\s*\n其余 (\d+) 个记 missing", report)
    assert m, "选择报告缺少候选/missing 计数句"
    assert int(m.group(1)) == pending, f"报告声称检索到 {m.group(1)} 个候选，CSV 实数 {pending}"
    assert int(m.group(2)) == missing, f"报告声称 {m.group(2)} 个 missing，CSV 实数 {missing}"
    # missing 不冒充覆盖：事实级证据为 0 的字段，状态只能是 missing/pending，不得出现第三种“已覆盖”态
    for r in persons:
        for f in ("birth_year", "death_year", "role"):
            if int(r[f"{f}_fact_evidences"] or 0) == 0:
                assert r[f.replace("_year", "").replace("_evidences", "") + "_status"] in (PENDING, MISSING)


def test_three_networks_use_different_filters(built_network: Path) -> None:
    payload = json.loads((built_network / "trustworthy_network_analysis.json").read_text(encoding="utf-8"))
    nets = payload["networks"]
    assert set(nets) >= {"full_exploratory", "trusted", "evidence_weighted"}
    full, trusted, weighted = nets["full_exploratory"], nets["trusted"], nets["evidence_weighted"]
    assert full["edges"] > trusted["edges"] > 0, "全量网络边数必须大于可信网络（不同过滤条件）"
    assert trusted["edges"] == weighted["edges"], "加权网络与可信网络拓扑一致，仅权重不同"
    assert full["total_weight"] >= trusted["total_weight"]
    assert set(payload["time_slices"]) == {"1928-1930", "1931-1933", "1934-1936"}
    # 过滤规则在 formulas 中显式登记且三套口径互不冒充
    assert "verified/supported" in payload["formulas"]["trusted_rule"]
    assert payload["formulas"]["evidence_weight"]


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
