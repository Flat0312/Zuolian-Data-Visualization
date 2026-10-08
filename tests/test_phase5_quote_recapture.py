"""引文重捕候选包的守门测试：只读、全 pending、重切结果可独立复核。

候选包的价值全在"不越权"：它不得写生产层，不得把任何一条标成已裁决，
且每一条 `recaptured` 都必须能被独立复核——按归一化偏移能在原文里原样取回、
locator 与夜间轮登记同页、并且重切后仍通过双方佐证门。
"""
from __future__ import annotations

import csv
import hashlib
import sys
from pathlib import Path

from conftest import PROJECT_ROOT, RUNTIME_TEXTS, requires_local_texts

_CAND = PROJECT_ROOT / "research" / "drafts" / "reports" / "phase5_quote_recapture_candidates.csv"
_QUEUE = PROJECT_ROOT / "research" / "drafts" / "reports" / "phase5_quote_recapture_queue.csv"
DATA = PROJECT_ROOT / "data" / "processed"

_ANALYSIS = str(PROJECT_ROOT / "research" / "analysis")
if _ANALYSIS not in sys.path:
    sys.path.append(_ANALYSIS)

from build_phase5_quote_recapture_candidates import (  # noqa: E402
    MAX_QUOTE_CHARS,
    _expand_sentence,
    build,
    name_list_pattern,
)
from quote_attestation import name_candidates, normalized_bundle, quote_attests_pair  # noqa: E402

WATCHED = ("person_relations.csv", "relation_evidences.csv", "sources.csv", "source_passages.csv")


def _rows(path: Path) -> list[dict[str, str]]:
    with open(path, encoding="utf-8-sig", newline="") as fh:
        return list(csv.DictReader(fh))


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


# --------------------------------------------------------------- 纯函数单元


def test_name_list_pattern_flags_enumeration_tails() -> None:
    assert name_list_pattern("签名的有田汉、洪深、阿英等共189人。") is True
    assert name_list_pattern("葛琴约了邵荃麟、周起应(周扬)、魏猛克几位友人一同去内山书店拜访鲁迅。") is False


def test_expand_sentence_stops_at_delimiters() -> None:
    flat = "甲来访。乙与丙同往书店。丁未至。"
    lo = flat.index("乙")
    hi = lo + len("丙")
    left, right = _expand_sentence(flat, lo, hi)
    assert flat[left:right] == "乙与丙同往书店。"


# ------------------------------------------------- 候选包整体约束（不需全文）


def test_candidates_cover_queue_and_are_all_pending() -> None:
    queue = _rows(_QUEUE)
    cand = _rows(_CAND)
    assert len(queue) == 28 and len(cand) == 28
    assert {r["relation_id"] for r in cand} == {r["relation_id"] for r in queue}
    assert {r["review_status"] for r in cand} == {"pending_human_review"}
    # 未重捕成功的行不得带 quote_sha256，避免被误当成已核验证据
    for row in cand:
        if row["recapture_status"] != "recaptured":
            assert not row["quote_sha256"], f'{row["relation_id"]} 未重捕成功却带哈希'
        else:
            assert row["locator_agrees"] == "yes", "重捕必须落在夜间轮登记的同一页"
            assert 0 < int(row["quote_char_len"]) <= MAX_QUOTE_CHARS
            assert len(row["quote_sha256"]) == 64


@requires_local_texts(*RUNTIME_TEXTS)
def test_build_writes_nothing_to_production() -> None:
    before = {name: _sha256(DATA / name) for name in WATCHED}
    rows = build(_QUEUE)
    assert len(rows) == 28
    assert before == {name: _sha256(DATA / name) for name in WATCHED}, "候选包生成不得改动生产层"


# ------------------------------------------- 重捕结果的独立复核（需本地全文）


@requires_local_texts(*RUNTIME_TEXTS)
def test_recaptured_quotes_relocate_verbatim_and_attest_both_parties() -> None:
    persons = {r["person_id"]: r for r in _rows(DATA / "persons.csv")}
    bundles: dict[str, tuple[str, str, list[int]]] = {}
    recaptured = [r for r in _rows(_CAND) if r["recapture_status"] == "recaptured"]
    assert recaptured, "本批应至少有一条重捕成功"
    for row in recaptured:
        src = row["source_file"]
        if src not in bundles:
            bundles[src] = normalized_bundle(Path(src))
        _, flat, _ = bundles[src]
        start, end = int(row["normalized_start"]), int(row["normalized_end"])
        # 1. 按偏移能原样取回，且与记录的引文、哈希一致
        assert flat[start:end] == row["recaptured_quote"], row["relation_id"]
        assert hashlib.sha256(row["recaptured_quote"].encode("utf-8")).hexdigest() == row["quote_sha256"]
        # 2. 不跨页标记
        assert "────" not in row["recaptured_quote"]
        # 3. 重切后仍通过双方佐证门
        na = name_candidates(persons[row["person_a_id"]])
        nb = name_candidates(persons[row["person_b_id"]])
        ok, basis = quote_attests_pair(row["recaptured_quote"], na, nb)
        assert ok, f'{row["relation_id"]} 重切引文未同时记载双方'
        assert row["attestation_basis"] == basis
        # 4. 名单句式必须如实标注
        assert row["name_list_pattern"] == ("yes" if name_list_pattern(row["recaptured_quote"]) else "no")


@requires_local_texts(*RUNTIME_TEXTS)
def test_page_mismatch_rejections_are_not_publishable_candidates() -> None:
    """页码不一致的重切一律拒绝：那是全书别处的另一段共现，不是同一处引文的重捕。"""
    cand = _rows(_CAND)
    mismatched = [r for r in cand if r["recapture_status"] == "rejected_locator_page_mismatch"]
    assert mismatched, "本批应存在因跨页被拒的条目"
    for row in mismatched:
        assert row["derived_locator"] != row["recorded_locator"]
        assert not row["quote_sha256"]
        assert row["projected_publish_status"] == ""


# ------------------------------------- 本轮踩到的三个坑，固化为防回归断言


def test_reason_title_wins_over_other_same_page_lists() -> None:
    """reason 点名《篇名》时必须选中含该篇名的同页段落。

    REL-01368 的教训：第132页同时有《为横死之小林遗族募捐启》9 人签署名单与
    营救丁玲、潘梓年的 38 人联名致电，两份都真但不是同一件事。只按"姓名距离最近"
    挑会选到 38 人电报，与 reason 所指不符。
    """
    cand = {r["relation_id"]: r for r in _rows(_CAND)}
    row = cand["REL-01368"]
    assert row["recapture_status"] == "recaptured"
    assert "为横死之小林遗族募捐启" in row["recaptured_quote"]
    assert "营救丁" not in row["recaptured_quote"], "又切回 38 人联名致电了"
    assert row["locator_agrees"] == "yes"


def test_secondary_bibliographic_passages_are_flagged() -> None:
    """书目著录／二手评述必须标注，不得与原始记载混同。

    REL-01891 切到的是《左联词典》第587页的书目条——"…传记小说。李克因作，载《东方纪事》
    1987年…叙述…叶紫同…萧军、萧红夫妇等的交往"。它确实同页、非纯名单、双方都在，
    但语义是"某本传记小说描写了他们的交往"，属二手著录，证据强度低于原始记载。
    """
    cand = {r["relation_id"]: r for r in _rows(_CAND)}
    row = cand["REL-01891"]
    assert row["recapture_status"] == "recaptured"
    assert row["secondary_description"] == "yes", "书目著录未被标注"
    assert row["review_status"] == "pending_human_review"


def test_name_list_rejection_is_relation_type_aware() -> None:
    """名单句式只对 交游/交往/通信 一类否决；签名联署/同属组织 保留并标注。"""
    cand = _rows(_CAND)
    by_id = {r["relation_id"]: r for r in cand}
    # 同属组织 + 名单：保留（左联常委/行政书记任职名单正是该关系的直接记载）
    assert by_id["REL-01161"]["recapture_status"] == "recaptured"
    assert by_id["REL-01161"]["name_list_pattern"] == "yes"
    # 被名单门否决的，其类型不得属于可接受名单的类型
    for row in cand:
        if row["recapture_status"] == "rejected_name_list_only":
            effective = row["proposed_relation_type"] or row["current_final_relation_type"]
            assert effective not in ("签名联署", "同属组织"), row["relation_id"]
