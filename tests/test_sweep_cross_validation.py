"""全量扫掠 + 双 Agent 交叉验证合并器的守门测试（两者均只读，生产层零改动）。

重点钉住三件事：
1. 合并器不得让「只有一方认可」的条目变成可落地——交叉验证的全部意义在于一致才落地；
2. 对两名裁决者一视同仁地校验 key_phrase 必须逐字出自引文，任何一方不合格即判无效；
3. 裁决文件的列串位（理由字段含未转义逗号）按保守规则修复，无法确定归位的一律判无效而非猜测。
"""
from __future__ import annotations

import csv
import hashlib
import sys
from pathlib import Path

import pytest
from conftest import PROJECT_ROOT

_ANALYSIS = str(PROJECT_ROOT / "research" / "analysis")
if _ANALYSIS not in sys.path:
    sys.path.append(_ANALYSIS)

from merge_sweep_cross_validation import (  # noqa: E402
    EXPECTED_COLUMNS,
    GRADES,
    NOT_ASSIGNABLE_TYPES,
    VERDICTS,
    load_verdicts,
    merge,
    validate_one,
)

REPORTS = PROJECT_ROOT / "research" / "drafts" / "reports"
BATCH_INPUT = REPORTS / "sweep_batch1_input.csv"
XVAL = REPORTS / "sweep_batch1_cross_validated.csv"
VERDICT_CODEX = Path(r"D:/1大创/.xval/verdict_codex.csv")
VERDICT_OPUS = PROJECT_ROOT / ".codex_tmp" / "verdict_opus.csv"
PROD = ("person_relations.csv", "relation_evidences.csv", "sources.csv")

pytestmark = pytest.mark.skipif(
    not (VERDICT_CODEX.exists() and VERDICT_OPUS.exists()),
    reason="两名裁决者的输出文件不在本机（交叉验证裁决文件不入库），跳过",
)


def _rows(path: Path) -> list[dict[str, str]]:
    with open(path, encoding="utf-8-sig", newline="") as fh:
        return list(csv.DictReader(fh))


def _write_verdicts(path: Path, rows: list[list[str]]) -> None:
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(EXPECTED_COLUMNS)
        w.writerows(rows)


def test_committed_cross_validated_table_is_consistent() -> None:
    rows = _rows(XVAL)
    assert len(rows) == 50
    for r in rows:
        if r["landable"] == "yes":
            assert r["agreement"] == "agree", r["relation_id"]
            assert r["codex_verdict"] == r["opus_verdict"] in ("成立", "类型需改")
            assert r["codex_grade"] == r["opus_grade"] == "support"
            assert r["current_publish_status"] not in ("supported", "verified")
            eff = r["codex_proposed_type"] or r["current_final_relation_type"]
            assert eff not in NOT_ASSIGNABLE_TYPES
        if r["agreement"] == "disagree":
            assert r["landable"] != "yes", f'{r["relation_id"]} 分歧条目不得可落地'


def test_merge_produces_no_landable_without_unanimity(tmp_path: Path) -> None:
    """把一方裁决全部改成「证据不足」，可落地必须归零——一致才落地是硬规则。"""
    src = _rows(BATCH_INPUT)
    ids = [r["relation_id"] for r in src]
    quotes = {r["relation_id"]: r["quote"] for r in src}

    def mk(verdict: str, grade: str) -> list[list[str]]:
        out = []
        for rid in ids:
            key = "".join(ch for ch in quotes[rid] if not ch.isspace())[:12]
            out.append([rid, verdict, grade, "", key, "t", "x", "2026-10-08"])
        return out

    a = tmp_path / "a.csv"
    b = tmp_path / "b.csv"
    _write_verdicts(a, mk("成立", "support"))
    _write_verdicts(b, mk("证据不足", "associated"))
    result = merge(BATCH_INPUT, a, b, 1)
    assert len(result["rows"]) == 50
    assert not [r for r in result["rows"] if r["landable"] == "yes"]
    assert all(r["agreement"] == "disagree" for r in result["rows"])


def test_key_phrase_must_be_verbatim_substring_for_both_adjudicators(tmp_path: Path) -> None:
    row = _rows(BATCH_INPUT)[0]
    good = {"relation_id": row["relation_id"], "verdict": "成立", "evidence_grade": "support",
            "proposed_type": "", "key_phrase": row["quote"][:10],
            "adjudicator_reason": "r", "adjudicator": "a", "adjudicated_at": "2026-10-08"}
    assert validate_one(row["relation_id"], good, row["quote"], "codex") == []
    bad = dict(good, key_phrase="这段文字绝不可能出现在引文里")
    errs = validate_one(row["relation_id"], bad, row["quote"], "opus")
    assert any("连续子串" in e for e in errs)
    bad2 = dict(good, verdict="证据不足", evidence_grade="support")
    assert any("必须 associated" in e for e in validate_one(row["relation_id"], bad2, row["quote"], "codex"))
    bad3 = dict(good, verdict="类型需改", proposed_type="待核验")
    assert any("机器推断占位标签" in e for e in validate_one(row["relation_id"], bad3, row["quote"], "codex"))


def test_column_shift_repaired_and_unrecoverable_rows_rejected(tmp_path: Path) -> None:
    ids = [r["relation_id"] for r in _rows(BATCH_INPUT)][:3]
    p = tmp_path / "shift.csv"
    with open(p, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(EXPECTED_COLUMNS)
        # 第 1 行正常；第 2 行理由含未转义逗号导致多一列（可修复）；第 3 行字段不足（不可修复）
        w.writerow([ids[0], "成立", "support", "", "k", "理由", "x", "2026-10-08"])
        w.writerow([ids[1], "成立", "support", "", "k", "前半", "后半", "x", "2026-10-08"])
        w.writerow([ids[2], "成立", "support"])
    got, problems = load_verdicts(p, set(ids))
    assert ids[0] in got and ids[1] in got
    assert got[ids[1]]["adjudicator_reason"] == "前半后半"
    assert got[ids[1]]["adjudicator"] == "x"
    assert ids[2] not in got
    assert any("无法确定归位" in x or "字段不足" in x for x in problems)
    assert any("缺少" in x for x in problems)


def test_vocab_constants_match_production_gate() -> None:
    """词表不得与发布门禁漂移。"""
    from relation_publish_status import INFERRED_RELATION_TYPES, PUBLIC_RELATION_STATUSES

    assert "同属组织" in INFERRED_RELATION_TYPES
    assert "同属组织" not in NOT_ASSIGNABLE_TYPES, "同属组织是真实类型，只是不公开，不得禁止人工选用"
    assert set(NOT_ASSIGNABLE_TYPES) & set(INFERRED_RELATION_TYPES) != set()
    assert set(PUBLIC_RELATION_STATUSES) == {"verified", "supported"}
    assert set(VERDICTS) == {"成立", "类型需改", "证据不足"}
    assert set(GRADES) == {"support", "associated"}


def test_merge_writes_nothing_to_production() -> None:
    before = {n: hashlib.sha256((PROJECT_ROOT / "data" / "processed" / n).read_bytes()).hexdigest()
              for n in PROD}
    merge(BATCH_INPUT, VERDICT_CODEX, VERDICT_OPUS, 1)
    after = {n: hashlib.sha256((PROJECT_ROOT / "data" / "processed" / n).read_bytes()).hexdigest()
             for n in PROD}
    assert before == after, "合并器不得改动生产层"
