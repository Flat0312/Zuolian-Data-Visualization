"""Phase 5 重捕候选第二批（2026-10-08 逐条独立裁决）落地的回归测试。

覆盖：
1. 三条落地引文按候选包 `source_file` + `normalized_start/end`（空白归一）可原样取回，
   且 `quote_sha256` 自洽；
2. REL-01891 判「证据不足」：生产层零痕迹（无本批标记证据行、未写 rejected、状态仍 inferred）；
3. REL-01161 成立但不公开：有本批 support 证据，`publish_status=pending_review`、
   origin 仍 derived、critical 未被反向降险，也未走 human_adjudication；
4. REL-01368 类型已改（standard/final 同步为 签名联署），`correction_reason` 追加不覆盖；
5. 公开层 7 条逐条无 critical/high、无 needs_manual_review、无待核验/推断类型、无 low 置信，
   且每条都有本批或第一批合格 support 引文过双方佐证门；
6. 计数与口径：4238 / 10272（support 23、associated 10249）/ 1178 / 1178 / 65，critical 仍 1974；
7. 幂等：二跑「无新增/已完成」且零写入；
8. 以 ``f011d92`` 为基线的红→绿重放：落地前 supported=5，落地后 supported=7，二跑 no-op。

依赖本地版权全文的测试用 ``conftest.requires_local_texts`` 加跳过守卫（CI 拿不到这些文件）。
"""
from __future__ import annotations

import csv
import hashlib
import importlib.util
import io
import re
import shutil
import subprocess
import sys
import tarfile
from collections import Counter, defaultdict
from pathlib import Path
from types import SimpleNamespace

import pytest
from conftest import PROJECT_ROOT, requires_local_texts

_ANALYSIS = str(PROJECT_ROOT / "research" / "analysis")
if _ANALYSIS not in sys.path:
    sys.path.append(_ANALYSIS)

from quote_attestation import (  # noqa: E402
    name_candidates,
    normalized_bundle,
    quote_attests_pair,
)

DATA = PROJECT_ROOT / "data" / "processed"
REPORTS = PROJECT_ROOT / "research" / "drafts" / "reports"
CANDIDATES = REPORTS / "phase5_quote_recapture_candidates.csv"
ADJUDICATED = REPORTS / "phase5_quote_recapture_adjudicated.csv"
LEDGER = REPORTS / "phase5_recapture_landing_ledger.csv"

RUNTIME_TEXTS = [
    DATA / "runtime_sources" / "左联史.txt",
    DATA / "runtime_sources" / "左联词典.txt",
]
requires_texts = requires_local_texts(*RUNTIME_TEXTS)

BATCH_MARKER = "P5-RECAPTURE-2026-10-08"
PREV_BATCH_MARKER = "P5-LANDING-2026-09-28"
PINNED_PRE_LANDING_COMMIT = "f011d92"

LAND_IDS = {"REL-00622", "REL-01161", "REL-01368"}
INSUFFICIENT_ID = "REL-01891"
EXPECTED_PUBLIC_IDS = {
    "REL-00046", "REL-00059", "REL-00097", "REL-03289", "REL-03518",
    "REL-00622", "REL-01368",
}
AUTHORIZATION_QUOTE = (
    '第一优先REL-00622 和 REL-01368直接过，REL-01891判"证据不足"，REL-01161 成立但不影响公开层'
)
EXPECTED_SOURCE_REUSE = {
    "REL-00622": "SRC-0779",
    "REL-01368": "SRC-0054",
    "REL-01161": "SRC-0739",
}


def _rows(path: Path) -> list[dict[str, str]]:
    with open(path, encoding="utf-8-sig", newline="") as fh:
        return list(csv.DictReader(fh))


def _load_landing_module():
    spec = importlib.util.spec_from_file_location(
        "p5_recapture_landing",
        PROJECT_ROOT / "research" / "analysis" / "apply_phase5_recapture_landing.py",
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


@pytest.fixture(scope="module")
def rel_rows() -> list[dict[str, str]]:
    return _rows(DATA / "person_relations.csv")


@pytest.fixture(scope="module")
def ev_rows() -> list[dict[str, str]]:
    return _rows(DATA / "relation_evidences.csv")


@pytest.fixture(scope="module")
def cands() -> dict[str, dict[str, str]]:
    return {r["relation_id"]: r for r in _rows(CANDIDATES)}


@pytest.fixture(scope="module")
def persons() -> dict[str, dict[str, str]]:
    return {r["person_id"]: r for r in _rows(DATA / "persons.csv")}


def _by_rel(ev_rows: list[dict[str, str]]) -> dict[str, list[dict[str, str]]]:
    out: dict[str, list[dict[str, str]]] = defaultdict(list)
    for row in ev_rows:
        out[row["relation_id"]].append(row)
    return out


def _qualified_support(rows: list[dict[str, str]], marker: str | None = None) -> list[dict[str, str]]:
    out = []
    for e in rows:
        if e["evidence_support"] != "support" or e["review_status"] == "rejected":
            continue
        if not e["locator"].strip() or not (e["quote"].strip() or e["context"].strip()):
            continue
        if marker is not None and marker not in e["reviewer_note"]:
            continue
        out.append(e)
    return out


# ---------------------------------------------------------------- 计数与口径


def test_post_landing_counts_and_critical_unchanged(rel_rows, ev_rows) -> None:
    assert len(rel_rows) == 4238
    assert len(ev_rows) == 10272
    dist = Counter(r["publish_status"] for r in rel_rows)
    assert dist["supported"] == 7
    assert dist["pending_review"] == 2451
    assert dist["inferred"] == 1780
    assert {r["publish_status_origin"] for r in rel_rows} == {"derived"}
    assert int(sum(1 for r in rel_rows if r["relation_risk_level"] == "critical")) == 1974
    support_dist = Counter(r["evidence_support"] for r in ev_rows)
    assert support_dist["support"] == 23
    assert support_dist["associated"] == 10249
    assert sum(1 for r in ev_rows if r["review_status"] == "reviewed") == 23
    assert sum(1 for r in ev_rows if r["review_status"] == "pending") == 10249
    assert len(_rows(DATA / "sources.csv")) == 1178
    assert len(_rows(DATA / "source_passages.csv")) == 1178
    assert len(_rows(DATA / "source_works.csv")) == 65


def test_first_batch_rows_untouched(ev_rows) -> None:
    """第一批 20 条 support 与 10249 条 associated 未被改判。"""
    prev = [e for e in ev_rows if PREV_BATCH_MARKER in e["reviewer_note"]]
    assert len(prev) == 20
    assert {e["evidence_support"] for e in prev} == {"support"}
    assert {e["review_status"] for e in prev} == {"reviewed"}
    batch2 = [e for e in ev_rows if BATCH_MARKER in e["reviewer_note"]]
    assert {e["relation_id"] for e in batch2} == LAND_IDS
    assert len(batch2) == 3


# ---------------------------------------------------------------- REL-01891 零痕迹


def test_rel_01891_zero_trace_in_production(rel_rows, ev_rows, cands) -> None:
    """证据不足 ≠ 人工否定：不新增证据、不写 rejected、保持 inferred。"""
    rel = next(r for r in rel_rows if r["relation_id"] == INSUFFICIENT_ID)
    assert rel["publish_status"] == "inferred"
    assert rel["publish_status_origin"] == "derived"
    assert rel["final_relation_type"] == "交游"
    assert rel["relation_risk_level"] == "low"
    assert rel["confidence"] == "medium"
    assert all(e["review_status"] != "rejected" for e in _by_rel(ev_rows)[INSUFFICIENT_ID])
    assert not [e for e in ev_rows if BATCH_MARKER in e["reviewer_note"] and e["relation_id"] == INSUFFICIENT_ID]
    assert cands[INSUFFICIENT_ID]["secondary_description"] == "yes"


# ---------------------------------------------------------------- REL-01161 成立不公开


def test_rel_01161_support_evidence_but_stays_pending(rel_rows, ev_rows) -> None:
    from relation_publish_status import derive_relation_publish_status_with_origin

    rel = next(r for r in rel_rows if r["relation_id"] == "REL-01161")
    assert rel["publish_status"] == "pending_review"
    assert rel["publish_status_origin"] == "derived"
    assert rel["relation_risk_level"] == "critical"
    assert rel["confidence"] == "low"
    assert rel["final_relation_type"] == "同属组织"
    by_rel = _by_rel(ev_rows)
    batch = _qualified_support(by_rel["REL-01161"], BATCH_MARKER)
    assert len(batch) == 1, "REL-01161 应恰有一条本批合格 support 证据"
    assert batch[0]["source_id"] == EXPECTED_SOURCE_REUSE["REL-01161"]
    assert batch[0]["source_level"] == "B", "二手来源按既有映射应为 B，不得照抄候选包"
    # 门禁自然拦截（critical + low + 推断类型），不是人工放行结果
    assert derive_relation_publish_status_with_origin(rel, by_rel["REL-01161"]) == (
        "pending_review", "derived",
    )


# ---------------------------------------------------------------- REL-01368 类型更正


def test_rel_01368_type_corrected_reason_appended(rel_rows) -> None:
    rel = next(r for r in rel_rows if r["relation_id"] == "REL-01368")
    assert rel["standard_relation_type"] == "签名联署"
    assert rel["final_relation_type"] == "签名联署"
    reason = rel["correction_reason"]
    assert reason.startswith("结构性检测未发现需要自动改写的明确证据。"), "correction_reason 应保留原文（追加不覆盖）"
    assert BATCH_MARKER in reason
    assert "交游" in reason and "签名联署" in reason
    assert rel["publish_status"] == "supported"


# ---------------------------------------------------------------- 公开层语义


def test_public_layer_seven_relations_all_conservative(rel_rows, ev_rows, persons) -> None:
    public = [r for r in rel_rows if r["publish_status"] in ("supported", "verified")]
    assert len(public) == 7
    assert {r["relation_id"] for r in public} == EXPECTED_PUBLIC_IDS
    by_rel = _by_rel(ev_rows)
    for row in public:
        rid = row["relation_id"]
        assert row["publish_status_origin"] == "derived"
        assert row["relation_risk_level"].lower() not in ("critical", "high")
        assert row["needs_manual_review"].lower() != "yes"
        assert row["confidence"].lower() != "low"
        assert row["final_relation_type"] != "待核验"
        assert row["final_relation_type"] not in {"同属组织", "空间共现", "时空共现"}
        qualified = _qualified_support(
            by_rel[rid],
        )
        assert qualified, f"{rid} 公开但无合格 support 证据"
        assert any(
            BATCH_MARKER in e["reviewer_note"] or PREV_BATCH_MARKER in e["reviewer_note"]
            for e in qualified
        ), f"{rid} 合格 support 证据不属于两批落地之一"


@requires_texts
def test_public_support_quotes_attest_both_parties(rel_rows, ev_rows, persons) -> None:
    by_rel = _by_rel(ev_rows)
    for row in rel_rows:
        if row["publish_status"] not in ("supported", "verified"):
            continue
        rid = row["relation_id"]
        qualified = [
            e for e in _qualified_support(by_rel[rid])
            if BATCH_MARKER in e["reviewer_note"] or PREV_BATCH_MARKER in e["reviewer_note"]
        ]
        na = name_candidates(persons[row["source_person_id"]])
        nb = name_candidates(persons[row["target_person_id"]])
        passed = [
            e for e in qualified
            if quote_attests_pair(
                e["quote"], na, nb,
                diary_author_implicit=e["locator"].startswith("鲁迅日记"),
            )[0]
        ]
        assert passed, f"{rid} 公开但合格 support 引文均未过双方佐证门"


# ---------------------------------------------------------------- 逐字回定位 + 哈希


@requires_texts
def test_landed_quotes_relocatable_and_hash_selfconsistent(ev_rows, cands) -> None:
    by_rel = _by_rel(ev_rows)
    checked = 0
    for rid in sorted(LAND_IDS):
        cand = cands[rid]
        rows = [e for e in by_rel[rid] if BATCH_MARKER in e["reviewer_note"]]
        assert len(rows) == 1
        ev = rows[0]
        quote = ev["quote"]
        assert quote == cand["recaptured_quote"], f"{rid} 落地引文与候选包不一致（不得改写）"
        assert ev["locator"] == cand["recorded_locator"]
        assert ev["source_id"] == EXPECTED_SOURCE_REUSE[rid]
        assert hashlib.sha256(quote.encode("utf-8")).hexdigest() == cand["quote_sha256"]
        _, flat, _ = normalized_bundle(DATA / "runtime_sources" / Path(cand["source_file"]).name)
        start, end = int(cand["normalized_start"]), int(cand["normalized_end"])
        assert flat[start:end] == quote, f"{rid} 按偏移未能逐字取回"
        assert f"quote_sha256={cand['quote_sha256']}" in ev["reviewer_note"]
        assert AUTHORIZATION_QUOTE in ev["reviewer_note"], "reviewer_note 必须含本批授权语"
        assert "2026-10-08" in ev["reviewer_note"]
        assert f"phase5_quote_recapture_candidates.csv#relation_id={rid}" in ev["reviewer_note"]
        checked += 1
    assert checked == 3


@requires_texts
def test_rel_01891_quote_is_verbatim_but_human_marked_insufficient(cands) -> None:
    """复核结论与证据等级是两回事：引文真实可取回，但二手著录不足以定为 support。"""
    cand = cands[INSUFFICIENT_ID]
    _, flat, _ = normalized_bundle(DATA / "runtime_sources" / Path(cand["source_file"]).name)
    start, end = int(cand["normalized_start"]), int(cand["normalized_end"])
    assert flat[start:end] == cand["recaptured_quote"]
    assert hashlib.sha256(cand["recaptured_quote"].encode("utf-8")).hexdigest() == cand["quote_sha256"]


# ---------------------------------------------------------------- 裁决文件语义


def test_adjudicated_file_matches_authorization(cands) -> None:
    rows = _rows(ADJUDICATED)
    assert len(rows) == 28
    assert {r["relation_id"] for r in rows} == set(cands)
    landed = {r["relation_id"] for r in rows if r["evidence_landed"] == "yes"}
    assert landed == LAND_IDS
    adjud = {r["relation_id"]: r for r in rows if r["relation_id"] in {*LAND_IDS, INSUFFICIENT_ID}}
    assert len(adjud) == 4
    for item in adjud.values():
        assert item["authorized_at"] == "2026-10-08"
        assert item["authorization_quote"] == AUTHORIZATION_QUOTE
        assert item["batch_marker"] == BATCH_MARKER
        assert item["quote_verbatim_recheck"] == "verbatim_offset_hit_whitespace_normalized"
    assert adjud[INSUFFICIENT_ID]["evidence_landed"] == "no"
    assert "证据不足" in adjud[INSUFFICIENT_ID]["adjudication_2026_10_08"]
    assert adjud["REL-01161"]["publish_status_after"] == "pending_review"
    assert adjud["REL-01368"]["final_relation_type_after"] == "签名联署"
    unadjudicated = [r for r in rows if r["adjudication_2026_10_08"] == "unadjudicated"]
    assert len(unadjudicated) == 24
    assert all(not r["authorized_by"] and not r["authorized_at"] for r in unadjudicated)
    # 候选包与队列文件未被改动来表示裁决
    assert all(c["review_status"] == "pending_human_review" for c in cands.values())


def test_ledger_rows_cover_four_adjudications() -> None:
    rows = _rows(LEDGER)
    assert len(rows) == 4
    by_id = {r["relation_id"]: r for r in rows}
    assert set(by_id) == {*LAND_IDS, INSUFFICIENT_ID}
    for rid in LAND_IDS:
        assert by_id[rid]["evidence_landed"] == "yes"
        assert re.fullmatch(r"RELE-\d{5}", by_id[rid]["new_relation_evidence_id"])
    assert by_id[INSUFFICIENT_ID]["evidence_landed"] == "no"
    assert by_id[INSUFFICIENT_ID]["new_relation_evidence_id"] == ""
    assert by_id["REL-01161"]["publish_status_after"] == "pending_review"
    assert by_id["REL-01368"]["final_relation_type_before"] == "交游"
    assert by_id["REL-01368"]["final_relation_type_after"] == "签名联署"
    for r in rows:
        assert r["source_registered"] == "no", "本批禁止注册新来源"
        assert r["quote_verbatim_recheck"] == "verbatim_offset_hit_whitespace_normalized"
        assert r["authorization_quote"] == AUTHORIZATION_QUOTE


# ---------------------------------------------------------------- 幂等与红→绿重放


def test_landing_is_idempotent_noop() -> None:
    module = _load_landing_module()
    result = module.apply_landing(DATA)
    assert result["status"] == "no-op"
    assert "无新增" in result["message"]


def _materialize(ref: str, dest: Path) -> None:
    out = subprocess.run(
        ["git", "archive", "--format=tar", ref, "data/processed"],
        cwd=PROJECT_ROOT, check=True, capture_output=True,
    )
    with tarfile.open(fileobj=io.BytesIO(out.stdout)) as tar:
        tar.extractall(dest, filter="data")


@requires_texts
def test_replay_from_pinned_pre_landing_commit(tmp_path: Path) -> None:
    """红→绿：f011d92 基线公开 5 条，落地后恰 7 条；二跑零写入。"""
    dest = tmp_path / "baseline"
    _materialize(PINNED_PRE_LANDING_COMMIT, dest)
    data_dir = dest / "data" / "processed"
    module = _load_landing_module()

    rels_before = _rows(data_dir / "person_relations.csv")
    dist_before = Counter(r["publish_status"] for r in rels_before)
    assert len(rels_before) == 4238
    assert dist_before["supported"] == 5, "落地前公开层应为 5（红）"
    assert dist_before["inferred"] == 1782
    evid_before = _rows(data_dir / "relation_evidences.csv")
    assert len(evid_before) == 10269
    assert not [e for e in evid_before if BATCH_MARKER in e["reviewer_note"]]

    result = module.apply_landing(data_dir)
    assert result["status"] == "applied"
    assert result["public_supported"] == 7
    assert result["new_evidence_rows"] == 3
    assert result["type_corrections"] == 1
    assert {e["relation_id"] for e in result["ledger"] if e["evidence_landed"] == "yes"} == LAND_IDS

    rels_after = _rows(data_dir / "person_relations.csv")
    dist_after = Counter(r["publish_status"] for r in rels_after)
    assert len(rels_after) == 4238
    assert dist_after["supported"] == 7
    assert dist_after["pending_review"] == 2451
    assert dist_after["inferred"] == 1780
    assert sum(1 for r in rels_after if r["relation_risk_level"] == "critical") == 1974
    assert len(_rows(data_dir / "relation_evidences.csv")) == 10272
    assert len(_rows(data_dir / "sources.csv")) == 1178
    assert len(_rows(data_dir / "source_passages.csv")) == 1178

    from kb_schema import validate_data_dir

    validation = validate_data_dir(data_dir)
    assert not validation.errors
    assert len(validation.warnings) <= 13

    second = module.apply_landing(data_dir)
    assert second["status"] == "no-op"


# ---------------------------------------------------- 篡改防御：隔离沙盒负向用例
#
# 每个用例都在 `git archive f011d92 data/processed` 还原的落地前基线上跑，
# 候选包与裁决表复制到 tmp_path、模块常量指向沙盒，绝不触碰生产树。
# 断言两件事：①抛 LandingError；②篡改之后、调用之前取三表 sha256，调用后必须一致
# （失败路径零写入），且 supported 仍为 5。


def _file_sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _tree_hashes(data_dir: Path) -> dict[str, str]:
    return {
        name: _file_sha(data_dir / name)
        for name in ("person_relations.csv", "relation_evidences.csv", "sources.csv")
    }


def _edit_csv(path: Path, mutator) -> None:
    with open(path, encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        fields = list(reader.fieldnames or [])
        rows = list(reader)
    rows = mutator(rows)
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)


@pytest.fixture()
def sandbox(tmp_path: Path):
    dest = tmp_path / "baseline"
    _materialize(PINNED_PRE_LANDING_COMMIT, dest)
    data_dir = dest / "data" / "processed"
    reports = tmp_path / "reports"
    reports.mkdir()
    cand = reports / CANDIDATES.name
    adj = reports / ADJUDICATED.name
    shutil.copy2(CANDIDATES, cand)
    shutil.copy2(ADJUDICATED, adj)
    module = _load_landing_module()
    module.DATA = data_dir
    module.REPORTS = reports
    module.CANDIDATES = cand
    module.ADJUDICATED = adj
    module.PERSONS_CSV = data_dir / "persons.csv"
    module.LEDGER = reports / "phase5_recapture_landing_ledger.csv"
    module.REPORT_MD = reports / "phase5_recapture_landing_report.md"
    module.ADJUDICATION_RECORD = reports / "phase5_quote_recapture_adjudication_record.md"
    return SimpleNamespace(module=module, data_dir=data_dir, cand=cand, adj=adj, reports=reports)


def _assert_rejected_and_zero_write(sb: SimpleNamespace, match: str) -> None:
    before = _tree_hashes(sb.data_dir)
    with pytest.raises(sb.module.LandingError, match=match):
        sb.module.apply_landing(sb.data_dir)
    assert _tree_hashes(sb.data_dir) == before, "失败路径留下了写入"
    rels = _rows(sb.data_dir / "person_relations.csv")
    assert Counter(r["publish_status"] for r in rels)["supported"] == 5, "基线公开层被改动"
    assert len(_rows(sb.data_dir / "relation_evidences.csv")) == 10269


def _tamper_adjudicated(sb: SimpleNamespace, rid: str, col: str, value: str) -> None:
    def mutator(rows):
        for row in rows:
            if row["relation_id"] == rid:
                row[col] = value
        return rows

    _edit_csv(sb.adj, mutator)


def test_sandbox_placeholder_authorization_quote_is_rejected(sandbox) -> None:
    """缺陷 1 复现：裁决表授权语改为执行者自授占位 → 必须拒绝且零写入。"""
    def mutator(rows):
        for row in rows:
            if row["batch_marker"] == BATCH_MARKER:
                row["authorization_quote"] = "AI 自行决定按建议执行"
        return rows

    _edit_csv(sandbox.adj, mutator)
    _assert_rejected_and_zero_write(sandbox, "授权语")


def test_sandbox_empty_authorization_quote_is_rejected(sandbox) -> None:
    """缺陷 1：授权语置空 → 授权溯源不完整，必须拒绝。"""
    def mutator(rows):
        for row in rows:
            if row["batch_marker"] == BATCH_MARKER:
                row["authorization_quote"] = ""
        return rows

    _edit_csv(sandbox.adj, mutator)
    _assert_rejected_and_zero_write(sandbox, "授权")


def test_sandbox_evidence_landed_flipped_is_rejected(sandbox) -> None:
    """缺陷 2 复现 A：REL-01891 的 evidence_landed 由 no 改 yes → 必须拒绝。"""
    _tamper_adjudicated(sandbox, INSUFFICIENT_ID, "evidence_landed", "yes")
    _assert_rejected_and_zero_write(sandbox, "evidence_landed")


def test_sandbox_missing_adjudicated_row_is_rejected(sandbox) -> None:
    """缺陷 2 复现 B：删掉 REL-01161 整行 → 必须拒绝。"""
    def mutator(rows):
        return [r for r in rows if r["relation_id"] != "REL-01161"]

    _edit_csv(sandbox.adj, mutator)
    _assert_rejected_and_zero_write(sandbox, "28")


def test_sandbox_adjudicated_quote_sha_mismatch_is_rejected(sandbox) -> None:
    """缺陷 2：裁决表 quote_sha256 与候选包不一致（引文被换）→ 必须拒绝。"""
    _tamper_adjudicated(sandbox, "REL-00622", "quote_sha256", "f" * 64)
    _assert_rejected_and_zero_write(sandbox, "quote_sha256")


def test_sandbox_tampered_candidate_quote_is_rejected(sandbox) -> None:
    """篡改候选包 REL-00622 引文（邵荃麟→邵荃麒）→ sha 不自洽，必须拒绝（固化）。"""
    def mutator(rows):
        for row in rows:
            if row["relation_id"] == "REL-00622":
                row["recaptured_quote"] = row["recaptured_quote"].replace("邵荃麟", "邵荃麒")
        return rows

    _edit_csv(sandbox.cand, mutator)
    _assert_rejected_and_zero_write(sandbox, "quote_sha256|逐字")


def test_sandbox_pre_lowered_risk_is_rejected(sandbox) -> None:
    """预先反向降险（沙盒 REL-01161 critical→low）→ 基线 critical 计数不符，必须拒绝（固化）。"""
    _edit_csv(
        sandbox.data_dir / "person_relations.csv",
        lambda rows: [
            {**r, "relation_risk_level": "low"} if r["relation_id"] == "REL-01161" else r
            for r in rows
        ],
    )
    _assert_rejected_and_zero_write(sandbox, "critical")


@requires_texts
def test_sandbox_clean_baseline_lands_and_idempotent(sandbox) -> None:
    """正向对照：未篡改沙盒必须成功落地 supported=7、evidences=10272，二跑 no-op。"""
    before = _tree_hashes(sandbox.data_dir)
    result = sandbox.module.apply_landing(sandbox.data_dir)
    assert result["status"] == "applied"
    assert result["public_supported"] == 7
    rels = _rows(sandbox.data_dir / "person_relations.csv")
    dist = Counter(r["publish_status"] for r in rels)
    assert dist["supported"] == 7 and dist["pending_review"] == 2451 and dist["inferred"] == 1780
    assert len(_rows(sandbox.data_dir / "relation_evidences.csv")) == 10272
    assert before != _tree_hashes(sandbox.data_dir)
    hash_after = _tree_hashes(sandbox.data_dir)
    second = sandbox.module.apply_landing(sandbox.data_dir)
    assert second["status"] == "no-op"
    assert _tree_hashes(sandbox.data_dir) == hash_after, "二跑发生写入"
