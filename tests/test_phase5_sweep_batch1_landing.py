"""Phase 5 扫掠第 1 批（双 Agent 交叉验证）落地 + 词表归并的回归测试。

覆盖：
1. 活数据计数与口径：4238 / 10290（support 41、reviewed 41）/ 1183 / 1183 / 65，
   公开层恰 25 条（三批口径：5 概括授权 + 2 逐条独立人工裁决 + 18 双 Agent 交叉验证）；
2. 词表归并终值：论战=0、交往=0、文学论战=128、交游=1214；且 `original_relation_type`
   / `raw_relation_type` / `llm_suggested_relation_type` 历史列与基线提交逐字节一致
   （词表归并只许动 standard/final 两列）；
3. 本批 18 条证据行：批次标记、双方佐证门、来源注册（5 条新来源）、类型更正恰 6 条
   且 correction_reason 追加不覆盖；
4. 三批均未使用 human_adjudication / verified / rejected 通道；
5. 台账与授权记录：18 行台账、授权记录逐字含两条授权语；
6. 幂等：二跑「无新增/已完成」且零写入；
7. 红→绿重放：以基线提交 ``e5ff09f``（词表归并前）为真实旧基线，在沙盒重演
   「词表归并 → 第 1 批落地」全链，公开层 7→25，二跑 no-op。

口径红线：本批为**双 Agent 交叉验证，非人工复核**——任何测试不得断言
`publish_status_origin == human_adjudication` 或 `verified`。
"""
from __future__ import annotations

import csv
import hashlib
import importlib.util
import io
import subprocess
import sys
import tarfile
from collections import Counter, defaultdict
from pathlib import Path

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
XVAL = REPORTS / "sweep_batch1_cross_validated.csv"
CANDIDATES = REPORTS / "sweep_candidates.csv"
LEDGER = REPORTS / "phase5_sweep_batch1_landing_ledger.csv"
AUTH_RECORD = REPORTS / "phase5_sweep_batch1_authorization_record.md"
VOCAB_LEDGER = REPORTS / "vocab_merge_2026-10-08_ledger.csv"

RUNTIME_TEXTS = [
    DATA / "runtime_sources" / "左联史.txt",
    DATA / "runtime_sources" / "左联词典.txt",
]
requires_texts = requires_local_texts(*RUNTIME_TEXTS)

BATCH_MARKER = "P5-SWEEP-BATCH1-2026-10-08"
PINNED_PRE_VOCAB_COMMIT = "e5ff09f"

LAND_IDS = {
    "REL-00008", "REL-00011", "REL-00012", "REL-00016", "REL-00022", "REL-00025",
    "REL-00027", "REL-00029", "REL-00032", "REL-00033", "REL-00035", "REL-00036",
    "REL-00038", "REL-00039", "REL-00061", "REL-00082", "REL-00088", "REL-00089",
}
INTERSECTION_IDS = {"REL-00029", "REL-00035", "REL-00089"}
SYNONYM_UNLOCKED_IDS = {"REL-00011", "REL-00038", "REL-00088"}
TYPE_TARGETS = {
    "REL-00011": "交游",
    "REL-00012": "签名联署",
    "REL-00025": "签名联署",
    "REL-00038": "文学论战",
    "REL-00061": "签名联署",
    "REL-00088": "交游",
}
EXPECTED_PUBLIC_IDS = {
    "REL-00046", "REL-00059", "REL-00097", "REL-03289", "REL-03518",
    "REL-00622", "REL-01368",
    *LAND_IDS,
}
EXPECTED_NEW_SOURCES = {
    "REL-00008", "REL-00027", "REL-00033", "REL-00035", "REL-00038",
}
AUTH_XVAL_QUOTE = "不用我裁，你和zcode交叉验证裁决吧"


def _rows(path: Path) -> list[dict[str, str]]:
    with open(path, encoding="utf-8-sig", newline="") as fh:
        return [dict(row) for row in csv.DictReader(fh)]


def _load_module(name: str, filename: str):
    spec = importlib.util.spec_from_file_location(name, PROJECT_ROOT / "research" / "analysis" / filename)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


@pytest.fixture(scope="module")
def rel_rows() -> list[dict[str, str]]:
    return _rows(DATA / "person_relations.csv")


@pytest.fixture(scope="module")
def ev_rows() -> list[dict[str, str]]:
    return _rows(DATA / "relation_evidences.csv")


@pytest.fixture(scope="module")
def persons() -> dict[str, dict[str, str]]:
    return {r["person_id"]: r for r in _rows(DATA / "persons.csv")}


def _by_rel(ev_list: list[dict[str, str]]) -> dict[str, list[dict[str, str]]]:
    out: dict[str, list[dict[str, str]]] = defaultdict(list)
    for row in ev_list:
        out[row["relation_id"]].append(row)
    return out


# ---------------------------------------------------------------- 计数与口径


def test_post_landing_counts(rel_rows, ev_rows) -> None:
    assert len(rel_rows) == 4238
    assert len(ev_rows) == 10290
    dist = Counter(r["publish_status"] for r in rel_rows)
    assert dist == {"supported": 25, "pending_review": 2451, "inferred": 1762}
    assert {r["publish_status_origin"] for r in rel_rows} == {"derived"}
    support_dist = Counter(r["evidence_support"] for r in ev_rows)
    assert support_dist == {"associated": 10249, "support": 41}
    assert sum(1 for r in ev_rows if r["review_status"] == "reviewed") == 41
    assert len(_rows(DATA / "sources.csv")) == 1183
    assert len(_rows(DATA / "source_passages.csv")) == 1183
    assert len(_rows(DATA / "source_works.csv")) == 65
    assert sum(1 for r in rel_rows if r["relation_risk_level"] == "critical") == 1974


def test_vocab_merge_final_distribution(rel_rows) -> None:
    for col in ("standard_relation_type", "final_relation_type"):
        dist = Counter(r[col] for r in rel_rows)
        assert dist.get("论战", 0) == 0, col
        assert dist.get("交往", 0) == 0, col
        assert dist["文学论战"] == 128, col  # 88 + 39(归并) + 1(REL-00038 类型更正)
        assert dist["交游"] == 1214, col     # 1206 + 8(归并) + 2(更正入) - 2(更正出)


def test_vocab_merge_did_not_touch_raw_columns() -> None:
    """词表归并只许改 standard/final 两列：original/raw/llm_suggested 与基线提交逐字节一致。"""
    merged_ids = {r["relation_id"] for r in _rows(VOCAB_LEDGER)}
    assert len(merged_ids) == 47, "词表台账应恰覆盖 47 行（39 + 8）"
    raw = subprocess.run(
        ["git", "show", f"{PINNED_PRE_VOCAB_COMMIT}:data/processed/person_relations.csv"],
        cwd=PROJECT_ROOT, check=True, capture_output=True,
    )
    baseline = {r["relation_id"]: r for r in csv.DictReader(io.StringIO(raw.stdout.decode("utf-8-sig")))}
    live = {r["relation_id"]: r for r in _rows(DATA / "person_relations.csv")}
    for col in ("original_relation_type", "raw_relation_type", "llm_suggested_relation_type"):
        for rid, row in live.items():
            assert row[col] == baseline[rid][col], f"{rid}.{col} 被词表归并改写（禁止）"


def test_public_layer_matches_three_batch_cohorts(rel_rows) -> None:
    public = {r["relation_id"] for r in rel_rows if r["publish_status"] in ("supported", "verified")}
    assert public == EXPECTED_PUBLIC_IDS
    assert not [r for r in rel_rows if r["publish_status"] == "verified"], "三批均未走人工通道"
    assert not [r for r in rel_rows if r["publish_status"] == "rejected"]
    # 本批 18 条的关系行 review_note 应由门禁重算清空（derived 行不带人工痕迹），
    # 交叉验证口径标注在证据行 reviewer_note（见 test_batch_evidence_rows_semantics）。


# ---------------------------------------------------------------- 证据行语义


def test_batch_evidence_rows_semantics(ev_rows) -> None:
    """18 条本批证据行：标记、support/reviewed、来源与佐证门凭据、授权语、溯源。"""
    by_rel = _by_rel(_rows(DATA / "relation_evidences.csv"))
    seen = []
    for rid in sorted(LAND_IDS):
        rows = [e for e in by_rel[rid] if BATCH_MARKER in e["reviewer_note"]]
        assert len(rows) == 1, rid
        ev = rows[0]
        assert ev["evidence_support"] == "support"
        assert ev["review_status"] == "reviewed"
        assert ev["source_level"] == "B", "左联史/左联词典均二手，按既有映射应为 B"
        assert AUTH_XVAL_QUOTE in ev["reviewer_note"], "reviewer_note 必须含逐字授权语"
        assert "双 Agent 交叉验证" in ev["reviewer_note"]
        assert "非人工复核" in ev["reviewer_note"]
        assert f"sweep_candidates.csv#relation_id={rid}" in ev["reviewer_note"]
        assert "quote_sha256=" in ev["reviewer_note"]
        seen.append(rid)
    assert len(seen) == 18


def test_source_registration_matches_expectation(ev_rows) -> None:
    by_rel = _by_rel(ev_rows)
    srcs = {r["source_id"]: r for r in _rows(DATA / "sources.csv")}
    registered = set()
    for rid in sorted(LAND_IDS):
        ev = next(e for e in by_rel[rid] if BATCH_MARKER in e["reviewer_note"])
        sid = ev["source_id"]
        assert sid in srcs, f"{rid} 引用的来源不存在：{sid}"
        if rid in EXPECTED_NEW_SOURCES:
            registered.add(sid)
            assert "P5-SWEEP-BATCH1-2026-10-08" in srcs[sid]["review_note"], rid
            assert srcs[sid]["evidence_strength"] == "二手"
        else:
            assert "P5-SWEEP-BATCH1-2026-10-08" not in srcs[sid]["review_note"], rid
    assert len(registered) == 5
    # 新来源必须有 passage 且 file_hash 指向本地原文
    pas = {r["source_id"]: r for r in _rows(DATA / "source_passages.csv")}
    for sid in registered:
        assert sid in pas, sid
        assert len(pas[sid]["file_hash"]) == 64, f"{sid} passage file_hash 应为 sha256"


def test_type_corrections_appended_not_overwritten(rel_rows) -> None:
    for rid, target in TYPE_TARGETS.items():
        row = next(r for r in rel_rows if r["relation_id"] == rid)
        assert row["standard_relation_type"] == target, rid
        assert row["final_relation_type"] == target, rid
        assert BATCH_MARKER in row["correction_reason"], rid
        assert "双 Agent 交叉验证" in row["correction_reason"], rid
    # 交集行类型必须保持原值
    for rid in INTERSECTION_IDS:
        row = next(r for r in rel_rows if r["relation_id"] == rid)
        assert row["final_relation_type"] == "交游", f"{rid} 交集行按现类型落地，类型不得改"


def test_synonym_unlocked_rows_landed_after_vocab_merge(ev_rows) -> None:
    """3 条 type_synonym 分歧经词表归并解锁：双方提议归一后一致才落地。"""
    xval = {r["relation_id"]: r for r in _rows(XVAL)}
    aliases = {"论战": "文学论战", "交往": "交游"}
    for rid in SYNONYM_UNLOCKED_IDS:
        row = xval[rid]
        ta, tb = row["codex_proposed_type"], row["opus_proposed_type"]
        assert ta != tb, f"{rid} 原始提议应不同（同义词）"
        assert aliases.get(ta, ta) == aliases.get(tb, tb), f"{rid} 归一后应一致"
        assert row["landable"] == "yes", f"{rid} 归并后应为一致可落地"


# ---------------------------------------------------------------- 台账与授权


def test_ledger_covers_all_landings() -> None:
    rows = _rows(LEDGER)
    assert len(rows) == 18
    assert {r["relation_id"] for r in rows} == LAND_IDS
    by_id = {r["relation_id"]: r for r in rows}
    for rid in LAND_IDS:
        entry = by_id[rid]
        assert entry["landed_this_batch"] == "yes"
        assert entry["batch_marker"] == BATCH_MARKER
        assert entry["landing_rule"] == ("保守交集" if rid in INTERSECTION_IDS else "双 Agent 一致")
        assert entry["quote_verbatim_recheck"] == "verbatim_offset_hit_whitespace_normalized"
        assert entry["attestation_basis"], f"{rid} 佐证门凭据缺失"
        assert entry["authorized_at"] == "2026-10-08"
        assert entry["publish_status_after"] == "supported"


def test_authorization_record_contains_verbatim_quotes() -> None:
    text = AUTH_RECORD.read_text(encoding="utf-8")
    assert AUTH_XVAL_QUOTE in text
    assert "采用保守交集" in text
    assert "论战→文学论战，交往→交游" in text
    assert "不是人工复核" in text


# ---------------------------------------------------------------- 独立复核（需本地原文）


@requires_texts
def test_landed_quotes_relocatable_and_attested(rel_rows, ev_rows, persons) -> None:
    """不信任落地脚本自己的结论：18 条引文逐字回定位 + sha256 + 双方佐证门独立复算。"""
    by_rel = _by_rel(ev_rows)
    cands = {r["relation_id"]: r for r in _rows(CANDIDATES)}
    for rid in sorted(LAND_IDS):
        cand = cands[rid]
        ev = next(e for e in by_rel[rid] if BATCH_MARKER in e["reviewer_note"])
        quote = ev["quote"]
        assert quote == cand["quote"], f"{rid} 落地引文与候选池不一致（不得改写）"
        assert hashlib.sha256(quote.encode("utf-8")).hexdigest() == cand["quote_sha256"]
        source_file = Path(cand["source_file"])
        if not source_file.exists():
            source_file = DATA / "runtime_sources" / source_file.name
        _, flat, _ = normalized_bundle(source_file)
        start, end = int(cand["normalized_start"]), int(cand["normalized_end"])
        assert flat[start:end] == quote, f"{rid} 按偏移未能逐字取回"
        row = next(r for r in rel_rows if r["relation_id"] == rid)
        pa = persons[row["source_person_id"]]
        pb = persons[row["target_person_id"]]
        ok, basis = quote_attests_pair(
            quote, name_candidates(pa), name_candidates(pb),
            diary_author_implicit=ev["locator"].startswith("鲁迅日记"),
        )
        assert ok, f"{rid} 独立复算佐证门未通过：{basis}"


# ---------------------------------------------------------------- 幂等与红绿重放


@requires_texts
def test_rerun_is_noop() -> None:
    module = _load_module("p5_sweep_batch1_landing", "apply_phase5_sweep_batch1_landing.py")
    result = module.apply_landing(DATA)
    assert result["status"] == "no-op"


def _materialize(ref: str, dest: Path) -> None:
    out = subprocess.run(
        ["git", "archive", "--format=tar", ref, "data/processed"],
        cwd=PROJECT_ROOT, check=True, capture_output=True,
    )
    with tarfile.open(fileobj=io.BytesIO(out.stdout)) as tar:
        tar.extractall(dest, filter="data")


@requires_texts
def test_replay_from_pinned_pre_vocab_commit(tmp_path: Path) -> None:
    """红→绿：e5ff09f 基线公开 7 条、论战 39/交往 8；归并 + 落地后公开 25 条；二跑零写入。"""
    dest = tmp_path / "baseline"
    _materialize(PINNED_PRE_VOCAB_COMMIT, dest)
    data_dir = dest / "data" / "processed"

    vocab = _load_module("p5_vocab_merge", "merge_relation_type_vocab.py")
    landing = _load_module("p5_sweep_batch1_landing_replay", "apply_phase5_sweep_batch1_landing.py")

    rels_before = _rows(data_dir / "person_relations.csv")
    dist_before = Counter(r["publish_status"] for r in rels_before)
    assert dist_before["supported"] == 7, "词表归并前公开层应为 7（红）"
    assert Counter(r["final_relation_type"] for r in rels_before)["论战"] == 39
    assert Counter(r["final_relation_type"] for r in rels_before)["交往"] == 8

    merged = vocab.apply_merge(data_dir)
    assert merged["status"] == "applied"
    assert merged["affected_rows"] == 47

    result = landing.apply_landing(data_dir)
    assert result["status"] == "applied"
    assert result["new_evidence_rows"] == 18
    assert result["public_supported"] == 25
    assert result["type_corrections"] == 6

    rels_after = _rows(data_dir / "person_relations.csv")
    assert Counter(r["publish_status"] for r in rels_after)["supported"] == 25
    assert Counter(r["final_relation_type"] for r in rels_after)["论战"] == 0
    assert Counter(r["final_relation_type"] for r in rels_after)["交往"] == 0

    assert landing.apply_landing(data_dir)["status"] == "no-op"
    assert vocab.apply_merge(data_dir)["status"] == "no-op"
