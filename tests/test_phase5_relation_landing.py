"""Phase 5 关系裁决落地与「双方佐证门」的回归测试。

覆盖三件事：
1. 佐证门本身的判定语义（含鲁迅日记的日记体例外）；
2. 生产层现状的独立复核——公开关系的支持引文必须真的同时记载双方当事人，
   且逐字命中本地原文（不信任落地脚本自己的结论）；
3. 以落地前提交 ``6c6a12b`` 为真实旧基线做红→绿重放：落地前公开关系为 0，
   落地后恰为 5，二跑「无新增/已完成」。

未通过佐证门的 28 条必须全部留在重捕队列（pending_human_review）且不得进入公开层，
也不得留下任何本批证据痕迹——防止「捕错引文也算已证」的作弊路径。
"""
from __future__ import annotations

import csv
import hashlib
import io
import json
import re
import subprocess
import sys
import tarfile
from collections import Counter
from pathlib import Path

import pytest
from conftest import PROJECT_ROOT, requires_local_texts

# 用 append 而非 insert(0)：避免改变 sys.path 顺序影响其他测试对 app 模块的解析。
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
EVIDENCE_JSONL = PROJECT_ROOT / "research" / "content-upgrade-overnight-2026-09-06" / "evidence.jsonl"
LEDGER = REPORTS / "phase5_relation_landing_ledger.csv"
QUEUE = REPORTS / "phase5_quote_recapture_queue.csv"
ADJUDICATED = REPORTS / "phase5_quote_recapture_adjudicated.csv"

PINNED_PRE_LANDING_COMMIT = "6c6a12b"
BATCH_MARKER = "P5-LANDING-2026-09-28"
RECAPTURE_MARKER = "P5-RECAPTURE-2026-10-08"
EXPECTED_LANDED = 20
# 公开层 5→7：第二批（2026-10-08 逐条独立裁决）追加 REL-00622、REL-01368。
EXPECTED_PUBLIC = 7
# 第一批红→绿重放只重演第一批落地（隔离在 tmp 基线上），公开层恰为 5。
EXPECTED_BATCH1_PUBLIC = 5
EXPECTED_QUEUE = 28


def _required_local_texts() -> list[Path]:
    """本批证据涉及的本地版权全文；按 .gitignore 政策不入库，缺失时相关测试跳过。"""
    paths: set[Path] = set()
    if EVIDENCE_JSONL.exists():
        with open(EVIDENCE_JSONL, encoding="utf-8") as fh:
            for line in fh:
                if not line.strip():
                    continue
                raw = str(json.loads(line).get("local_path") or "").strip()
                if raw:
                    candidate = Path(raw)
                    paths.add(candidate if candidate.is_absolute() else PROJECT_ROOT / raw)
    return sorted(paths)


LOCAL_TEXTS = _required_local_texts()
requires_texts = requires_local_texts(*LOCAL_TEXTS)


def _rows(path: Path) -> list[dict[str, str]]:
    with open(path, encoding="utf-8-sig", newline="") as fh:
        return list(csv.DictReader(fh))


@pytest.fixture(scope="module")
def overnight_evidence() -> dict[str, dict]:
    with open(EVIDENCE_JSONL, encoding="utf-8") as fh:
        return {r["evidence_id"]: r for r in (json.loads(line) for line in fh if line.strip())}



@pytest.fixture(scope="module")
def bundles() -> dict[str, tuple[str, str, list[int]]]:
    out: dict[str, tuple[str, str, list[int]]] = {}
    for path in sorted({*(DATA / "runtime_sources").glob("*.txt"), *(PROJECT_ROOT / "research" / "raw_texts").glob("*.txt")}):
        out[str(path)] = normalized_bundle(path)
    return out


def _bundle_for(local_path: str, bundles: dict) -> tuple[str, str, list[int]]:
    resolved = Path(local_path)
    if not resolved.is_absolute():
        resolved = PROJECT_ROOT / local_path
    key = str(resolved)
    if key not in bundles:
        bundles[key] = normalized_bundle(resolved)
    return bundles[key]


# ---------------------------------------------------------------- 佐证门语义


def test_quote_attests_pair_requires_both_parties() -> None:
    ok, basis = quote_attests_pair("鲁迅与茅盾会晤。", ["鲁迅"], ["茅盾"])
    assert ok and basis == "A=鲁迅;B=茅盾"
    ok, basis = quote_attests_pair("左联成员众多。", ["鲁迅"], ["茅盾"])
    assert not ok and basis == "neither_party_in_quote"
    ok, basis = quote_attests_pair("鲁迅日记某日。", ["鲁迅"], ["茅盾"])
    assert not ok and basis == "only_A_in_quote"
    ok, basis = quote_attests_pair("茅盾来访。", ["鲁迅"], ["茅盾"])
    assert not ok and basis == "only_B_in_quote"


def test_quote_attests_pair_diary_author_implicit() -> None:
    """日记体：作者即甲方，引文不需自名，但必须出现乙方。"""
    # 乙方候选含别名：日记原文多称「雪峰」而非「冯雪峰」
    ok, basis = quote_attests_pair("雪峰来并交稿费。", ["鲁迅"], ["冯雪峰", "雪峰"], diary_author_implicit=True)
    assert ok and basis == "A=diary_author_implicit;B=雪峰"
    ok, _ = quote_attests_pair("晴。上午得丛芜信。", ["鲁迅"], ["冯雪峰", "雪峰"], diary_author_implicit=True)
    assert not ok
    # 日记体例外只免甲方自名，不免乙方
    ok, basis = quote_attests_pair("晴。无事。", ["鲁迅"], ["冯雪峰", "雪峰"], diary_author_implicit=True)
    assert not ok and basis == "neither_party_in_quote"


def test_name_candidates_filters_single_char_aliases() -> None:
    got = name_candidates({"standard_name": "鲁迅", "aliases": "周树人、L.X.、豫"})
    assert got == ["鲁迅", "周树人", "L.X."]


# ------------------------------------------------- 生产层独立复核（不信任脚本）


@requires_texts
def test_public_relations_support_quotes_attest_both_parties(bundles) -> None:
    """公开层每条关系都必须存在合格 support 证据，且其引文过双方佐证门（不限批次标记）。

    2026-10-08 起公开层同时含两批落地行（第一批 5 条 + 第二批 2 条），按单一批次标记过滤
    会漏掉第二批；改为「不限批次、但对每条公开关系都实际校验佐证门」，比原断言更强。
    """
    persons = {r["person_id"]: r for r in _rows(DATA / "persons.csv")}
    rels = {r["relation_id"]: r for r in _rows(DATA / "person_relations.csv")}
    evid = _rows(DATA / "relation_evidences.csv")
    public = [r for r in rels.values() if r["publish_status"] in ("supported", "verified")]
    assert len(public) == EXPECTED_PUBLIC
    for row in public:
        rid = row["relation_id"]
        qualified = [
            e for e in evid
            if e["relation_id"] == rid and e["evidence_support"] == "support"
            and e["review_status"] != "rejected"
            and e["locator"].strip() and (e["quote"].strip() or e["context"].strip())
        ]
        assert qualified, f"{rid} 公开但无合格 support 证据"
        na = name_candidates(persons[row["source_person_id"]])
        nb = name_candidates(persons[row["target_person_id"]])
        passed = [
            e for e in qualified
            if quote_attests_pair(
                e["quote"], na, nb,
                diary_author_implicit="鲁迅日记" in e["locator"],
            )[0]
        ]
        assert passed, f"{rid} 公开关系的合格 support 引文均未同时记载双方"


@requires_texts
def test_landed_support_evidence_is_verbatim_in_local_source(overnight_evidence, bundles) -> None:
    """本批落地的每条 support 证据，引文必须逐字命中本地原文且哈希自洽。"""
    ledger = _rows(LEDGER)
    assert len(ledger) == EXPECTED_LANDED
    checked = 0
    for entry in ledger:
        row = next(e for e in _rows(DATA / "relation_evidences.csv")
                   if e["relation_evidence_id"] == entry["new_relation_evidence_id"])
        src = overnight_evidence[entry["overnight_evidence_id"]]
        assert row["quote"] == src["quote"], entry["relation_id"]
        assert row["locator"] == src["locator"], entry["relation_id"]
        assert row["evidence_support"] == "support"
        assert row["review_status"] == "reviewed"
        _, flat, _ = _bundle_for(src["local_path"], bundles)
        assert re.sub(r"\s+", "", src["quote"]) in flat, f"{entry['relation_id']} 引文未逐字命中原文"
        assert hashlib.sha256(src["quote"].encode("utf-8")).hexdigest() == src["quote_sha256"]
        assert entry["quote_verbatim_recheck"] == "verbatim_whitespace_normalized"
        checked += 1
    assert checked == EXPECTED_LANDED


def test_unattested_relations_stay_out_of_production_and_public() -> None:
    """重捕队列 28 条：扣除已裁决落地的 3 条后，其余 25 条仍不得公开、不得留下任何批次落地证据。

    2026-10-08 第二批把 REL-00622 / REL-01368 / REL-01161 裁决落地（REL-01891 判证据不足、
    生产层零改动，仍在 25 条断言内）；裁决状态只登记在
    ``phase5_quote_recapture_adjudicated.csv``，队列文件本身未被改动。
    """
    queue = _rows(QUEUE)
    assert len(queue) == EXPECTED_QUEUE
    assert {r["review_status"] for r in queue} == {"pending_human_review"}
    rels = {r["relation_id"]: r for r in _rows(DATA / "person_relations.csv")}
    evid = _rows(DATA / "relation_evidences.csv")
    landed = {r["relation_id"] for r in _rows(ADJUDICATED) if r.get("evidence_landed") == "yes"}
    assert landed == {"REL-00622", "REL-01161", "REL-01368"}
    batch2_rids = {e["relation_id"] for e in evid if RECAPTURE_MARKER in e["reviewer_note"]}
    assert batch2_rids == landed, "第二批证据痕迹必须与裁决落地行严格一致"
    batch_rids = {
        e["relation_id"] for e in evid
        if BATCH_MARKER in e["reviewer_note"] or RECAPTURE_MARKER in e["reviewer_note"]
    }
    remaining = [row for row in queue if row["relation_id"] not in landed]
    assert len(remaining) == EXPECTED_QUEUE - 3
    actions = Counter(r["proposed_action"] for r in queue)
    assert set(actions) <= {
        "recapture_quote_then_regrade",
        "cooccurrence_only_keep_associated",
        "no_local_support_mark_insufficient",
    }
    assert actions["no_local_support_mark_insufficient"] >= 1
    for row in remaining:
        rid = row["relation_id"]
        assert rid not in batch_rids, f"{rid} 未落地却留下任一批次落地证据"
        assert rels[rid]["publish_status"] not in ("supported", "verified"), f"{rid} 未过佐证门却进入公开层"
        assert rels[rid]["publish_status_origin"] == "derived"


def test_landing_is_idempotent_on_current_production(capsys) -> None:
    import importlib.util

    spec = importlib.util.spec_from_file_location(
        "p5_landing", PROJECT_ROOT / "research" / "analysis" / "apply_phase5_relation_landing.py"
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    assert module.main.__name__ == "main"
    result = module.apply_landing(DATA)
    assert result["status"] == "no-op"
    assert "无新增" in result["message"]


# ------------------------------------------------------- 红→绿：固定旧提交重放


def _materialize(ref: str, dest: Path) -> None:
    out = subprocess.run(
        ["git", "archive", "--format=tar", ref, "data/processed"],
        cwd=PROJECT_ROOT, check=True, capture_output=True,
    )
    with tarfile.open(fileobj=io.BytesIO(out.stdout)) as tar:
        tar.extractall(dest, filter="data")


def _load_module():
    import importlib.util

    spec = importlib.util.spec_from_file_location(
        "p5_landing_replay", PROJECT_ROOT / "research" / "analysis" / "apply_phase5_relation_landing.py"
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


@requires_texts
def test_replay_from_pinned_pre_landing_commit(tmp_path: Path) -> None:
    """红→绿：旧基线公开关系为 0，落地后恰为 5；二跑零写入。"""
    dest = tmp_path / "baseline"
    _materialize(PINNED_PRE_LANDING_COMMIT, dest)
    data_dir = dest / "data" / "processed"
    module = _load_module()

    before = _rows(data_dir / "person_relations.csv")
    assert len(before) == 4238
    assert not [r for r in before if r["publish_status"] in ("supported", "verified")], "旧基线公开关系必须为 0（红）"
    evid_before = _rows(data_dir / "relation_evidences.csv")
    assert len(evid_before) == 10249
    assert {r["evidence_support"] for r in evid_before} == {"associated"}

    result = module.apply_landing(data_dir)
    assert result["status"] == "applied"
    assert result["attested"] == EXPECTED_LANDED
    assert result["unattested"] == EXPECTED_QUEUE
    assert result["public_supported"] == EXPECTED_BATCH1_PUBLIC
    assert result["type_corrections"] == 3
    assert result["new_evidence_rows"] == EXPECTED_LANDED
    assert len(result["queue"]) == EXPECTED_QUEUE

    after = _rows(data_dir / "person_relations.csv")
    dist = Counter(r["publish_status"] for r in after)
    assert dist["supported"] == EXPECTED_BATCH1_PUBLIC
    assert dist["pending_review"] == 2451
    assert dist["inferred"] == 1782
    assert len(after) == 4238
    assert {r["publish_status_origin"] for r in after} == {"derived"}
    assert len(_rows(data_dir / "relation_evidences.csv")) == 10269
    assert len(_rows(data_dir / "sources.csv")) == 1178
    assert len(_rows(data_dir / "source_passages.csv")) == 1178
    # critical 风险计数未被反向降险改写
    assert sum(1 for r in after if r["relation_risk_level"] == "critical") == 1974

    from kb_schema import validate_data_dir

    validation = validate_data_dir(data_dir)
    assert not validation.errors
    assert len(validation.warnings) <= 13

    second = module.apply_landing(data_dir)
    assert second["status"] == "no-op"
