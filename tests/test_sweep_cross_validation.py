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

# 只有需要读两份**原始**裁决文件的用例才跳过；其余用例只依赖已入库的对照表与批次输入，
# 在 CI 与全新克隆上也必须真跑——否则这套门禁在 CI 里等于不存在。
needs_raw_verdicts = pytest.mark.skipif(
    not (VERDICT_CODEX.exists() and VERDICT_OPUS.exists()),
    reason="两名裁决者的原始输出文件不入库（Codex 的在仓库外、Opus 的在 gitignore 区），跳过",
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
    """入库对照表守门：一致 15 + 保守交集 3 + 仅研究层 1，交集行必须满足交集条件。"""
    rows = _rows(XVAL)
    assert len(rows) == 50
    counts = {"yes": 0, "intersection": 0, "research_only": 0}
    for r in rows:
        if r["landable"] == "yes":
            counts["yes"] += 1
            assert r["agreement"] == "agree", r["relation_id"]
            assert r["codex_verdict"] == r["opus_verdict"] in ("成立", "类型需改")
            assert r["codex_grade"] == r["opus_grade"] == "support"
            assert r["current_publish_status"] not in ("supported", "verified")
            eff = r["codex_proposed_type"] or r["current_final_relation_type"]
            assert eff not in NOT_ASSIGNABLE_TYPES
        elif r["landable"] == "intersection":
            counts["intersection"] += 1
            # 保守交集（2026-10-08 用户决定）：双方 support 且至少一方判成立，按现类型落地
            assert r["agreement"] == "disagree", r["relation_id"]
            assert r["codex_grade"] == r["opus_grade"] == "support"
            assert "成立" in (r["codex_verdict"], r["opus_verdict"])
            assert r["current_publish_status"] not in ("supported", "verified")
            assert r["current_final_relation_type"] not in NOT_ASSIGNABLE_TYPES
        elif r["landable"] == "research_only":
            counts["research_only"] += 1
        if r["agreement"] == "disagree":
            assert r["landable"] != "yes", f'{r["relation_id"]} 分歧条目不得按一致口径落地'
    assert counts == {"yes": 15, "intersection": 3, "research_only": 1}


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


@needs_raw_verdicts
def test_merge_writes_nothing_to_production() -> None:
    before = {n: hashlib.sha256((PROJECT_ROOT / "data" / "processed" / n).read_bytes()).hexdigest()
              for n in PROD}
    merge(BATCH_INPUT, VERDICT_CODEX, VERDICT_OPUS, 1)
    after = {n: hashlib.sha256((PROJECT_ROOT / "data" / "processed" / n).read_bytes()).hexdigest()
             for n in PROD}
    assert before == after, "合并器不得改动生产层"


def test_merge_is_reproducible_from_committed_table(tmp_path: Path) -> None:
    """从已入库的对照表反推双方裁决，重跑合并器必须得到完全相同的一致/分歧/可落地划分。

    这条用例在 CI 上也会跑（只依赖入库文件），保证交叉验证的结论可复现、
    不是只在某台机器上的一次性结果。入库对照表以「保守交集已采用」口径生成，
    故重放同样开启该开关；另设不带开关的对照用例验证交集不开启时分歧一律搁置。
    """
    committed = {r["relation_id"]: r for r in _rows(XVAL)}
    src = _rows(BATCH_INPUT)

    def rebuild(tag: str) -> Path:
        out = tmp_path / f"{tag}.csv"
        with open(out, "w", encoding="utf-8-sig", newline="") as fh:
            w = csv.writer(fh)
            w.writerow(EXPECTED_COLUMNS)
            for row in src:
                c = committed[row["relation_id"]]
                w.writerow([
                    row["relation_id"], c[f"{tag}_verdict"], c[f"{tag}_grade"],
                    c[f"{tag}_proposed_type"], c[f"{tag}_key_phrase"], c[f"{tag}_reason"],
                    tag, "2026-10-08",
                ])
        return out

    codex_csv, opus_csv = rebuild("codex"), rebuild("opus")
    result = merge(BATCH_INPUT, codex_csv, opus_csv, 1, conservative_intersection=True)
    assert not result["errors"], result["errors"][:3]
    again = {r["relation_id"]: r for r in result["rows"]}
    assert set(again) == set(committed)
    for rid, want in committed.items():
        got = again[rid]
        assert got["agreement"] == want["agreement"], rid
        assert got["disagreement_kind"] == want["disagreement_kind"], rid
        if want["landable"] in ("yes", "intersection") and (
                want["current_publish_status"] in ("supported", "verified")
                or got["current_publish_status"] in ("supported", "verified")):
            # 落地后重放：该关系已按本批进入公开层，门禁如实回落为「已在公开层，无需重复落地」。
            # 这是 derived 门禁正确性的表现（不透传旧结论），不属于不可复现。
            assert got["landable"] == "no", rid
            assert "已在公开层" in got["landable_block_reason"], rid
        else:
            assert got["landable"] == want["landable"], rid

    # 对照用例：不开保守交集时，交集行回落为普通分歧（landable=no），一致行不变
    strict = {r["relation_id"]: r for r in merge(
        BATCH_INPUT, codex_csv, opus_csv, 1, conservative_intersection=False)["rows"]}
    for rid, want in committed.items():
        if want["landable"] == "intersection" and strict[rid]["current_publish_status"] not in ("supported", "verified"):
            assert strict[rid]["landable"] == "no", rid
            assert strict[rid]["agreement"] == "disagree", rid
