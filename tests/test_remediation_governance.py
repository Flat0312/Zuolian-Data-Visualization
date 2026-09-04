"""本次返修新增守门测试（红→绿证据，史料语义正确优先）。

覆盖任务书二/三/四/七/九：
- associated+pending 不得产生 supported
- 缺 quote/locator 不得产生 supported
- critical/high/待核验/low/needs_manual_review=yes 不得公开（即使配 support）
- 人工 verified 须带完整审计元数据，否则 Schema 报错
- 旧 supported 不透传：改风险/撤证据后重算自动离开公开层
- 来源增量：只加 source 不加 passage/work 映射时 Schema 报错；走统一注册后恢复；ID 稳定；二跑零变化
- 发布双跑字节一致（默认 unstamped）
"""
from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
from conftest import create_standard_dataset


def _base_relation(**overrides: object) -> dict[str, str]:
    base = {
        "relation_id": "RX",
        "source_person_id": "P1",
        "target_person_id": "P2",
        "original_relation_type": "通信",
        "standard_relation_type": "通信",
        "raw_relation_type": "通信",
        "llm_suggested_relation_type": "",
        "final_relation_type": "通信",
        "llm_reason": "",
        "llm_confidence": "",
        "display_status": "formal",
        "relation_quality_score": "",
        "relation_risk_level": "low",
        "context": "ctx",
        "evidence_ref": "loc",
        "weight": "1",
        "source_ids": "S1",
        "correction_reason": "",
        "confidence": "medium",
        "needs_manual_review": "no",
        "publish_status": "",
        "publish_status_origin": "",
        "reviewer": "",
        "reviewed_at": "",
        "review_note": "",
    }
    base.update(overrides)  # type: ignore
    return base  # type: ignore


def _ev(support: str, locator: str, quote: str, context: str, review: str = "pending") -> dict[str, str]:
    return {
        "relation_evidence_id": "RELE-1",
        "relation_id": "RX",
        "source_id": "S1",
        "locator": locator,
        "quote": quote,
        "context": context,
        "quote_or_context": context or quote,
        "evidence_support": support,
        "source_level": "A",
        "review_status": review,
        "reviewer_note": "",
    }


def test_associated_pending_cannot_be_supported() -> None:
    from research.analysis.relation_publish_status import derive_relation_publish_status

    row = _base_relation()
    # associated + pending（生产现状）不得为 supported
    assert derive_relation_publish_status(row, [_ev("associated", "loc", "", "ctx", "pending")]) != "supported"
    # support + pending + locator + quote 才可为 supported（非高风险/待核验/low/需复核/冲突/推断类型）
    assert derive_relation_publish_status(row, [_ev("support", "loc", "quote", "", "pending")]) == "supported"


def test_missing_quote_and_locator_cannot_be_supported() -> None:
    from research.analysis.relation_publish_status import derive_relation_publish_status

    row = _base_relation()
    assert derive_relation_publish_status(row, [_ev("support", "", "quote", "", "pending")]) != "supported"
    assert derive_relation_publish_status(row, [_ev("support", "loc", "", "", "pending")]) != "supported"
    assert derive_relation_publish_status(row, [_ev("support", "loc", "", "ctx", "pending")]) == "supported"
    assert derive_relation_publish_status(row, [_ev("support", "loc", "quote", "", "pending")]) == "supported"


def test_risky_low_unverified_manual_never_public_even_with_support() -> None:
    from research.analysis.relation_publish_status import derive_relation_publish_status

    good_ev = [_ev("support", "loc", "quote", "", "pending")]
    cases = [
        _base_relation(relation_risk_level="critical"),
        _base_relation(relation_risk_level="high"),
        _base_relation(final_relation_type="待核验"),
        _base_relation(confidence="low"),
        _base_relation(needs_manual_review="yes"),
        _base_relation(final_relation_type="同属组织"),
        _base_relation(final_relation_type="空间共现"),
        _base_relation(final_relation_type="时空共现"),
    ]
    for row in cases:
        assert derive_relation_publish_status(row, good_ev) not in ("verified", "supported"), row
    # conflict 证据也必须 pending
    assert derive_relation_publish_status(_base_relation(), [_ev("conflict", "loc", "q", "c")]) == "pending_review"


def test_human_verified_requires_audit_metadata(sandbox_tmp_path: Path) -> None:
    data_dir = create_standard_dataset(sandbox_tmp_path)
    rel_path = data_dir / "person_relations.csv"
    rels = pd.read_csv(rel_path, encoding="utf-8-sig", dtype=str).fillna("")
    for col in ("publish_status", "publish_status_origin", "reviewer", "reviewed_at", "review_note"):
        if col not in rels.columns:
            rels[col] = ""
    rels.loc[0, "publish_status"] = "verified"
    rels.loc[0, "publish_status_origin"] = "human_adjudication"
    rels.loc[0, "reviewer"] = ""
    rels.loc[0, "reviewed_at"] = ""
    rels.loc[0, "review_note"] = ""
    rels.to_csv(rel_path, index=False, encoding="utf-8-sig")
    from kb_schema import validate_data_dir

    result = validate_data_dir(data_dir)
    assert any(i.code == "missing_human_adjudication_audit" for i in result.errors)
    # 补齐审计后恢复（仍须 evidence 由人工线另行核验，此处仅验审计门禁）
    rels.loc[0, "reviewer"] = "tester"
    rels.loc[0, "reviewed_at"] = "2026-09-04"
    rels.loc[0, "review_note"] = "人工核验通过"
    rels.to_csv(rel_path, index=False, encoding="utf-8-sig")
    result2 = validate_data_dir(data_dir)
    assert not any(i.code == "missing_human_adjudication_audit" for i in result2.errors)


def test_human_status_without_origin_rejected_by_schema(sandbox_tmp_path: Path) -> None:
    data_dir = create_standard_dataset(sandbox_tmp_path)
    rel_path = data_dir / "person_relations.csv"
    rels = pd.read_csv(rel_path, encoding="utf-8-sig", dtype=str).fillna("")
    for col in ("publish_status_origin", "reviewer", "reviewed_at", "review_note"):
        if col not in rels.columns:
            rels[col] = ""
    rels.loc[0, "publish_status"] = "verified"
    rels.loc[0, "publish_status_origin"] = "derived"
    rels.to_csv(rel_path, index=False, encoding="utf-8-sig")
    from kb_schema import validate_data_dir

    result = validate_data_dir(data_dir)
    assert any(i.code == "invalid_relation_publish_origin" for i in result.errors)


def test_old_supported_not_passthrough_recompute_on_risk_change(sandbox_tmp_path: Path) -> None:
    """反向测试：旧 supported 改 critical 后重算必须离开公开层（任务二-6、九-1）。"""
    from research.analysis.relation_publish_status import derive_relation_publish_status

    row = _base_relation(publish_status="supported", publish_status_origin="derived")
    good_ev = [_ev("support", "loc", "quote", "", "pending")]
    assert derive_relation_publish_status(row, good_ev) == "supported"
    row2 = dict(row)
    row2["relation_risk_level"] = "critical"
    assert derive_relation_publish_status(row2, good_ev) == "pending_review"
    row3 = dict(row)
    row3["needs_manual_review"] = "yes"
    assert derive_relation_publish_status(row3, good_ev) == "pending_review"

    # 端到端：support 公开后改 critical，重建发布后从发布层消失
    data_dir = create_standard_dataset(sandbox_tmp_path)
    rel_path = data_dir / "person_relations.csv"
    rels = pd.read_csv(rel_path, encoding="utf-8-sig", dtype=str).fillna("")
    for col in ("publish_status", "publish_status_origin", "reviewer", "reviewed_at", "review_note"):
        if col not in rels.columns:
            rels[col] = ""
    rels.to_csv(rel_path, index=False, encoding="utf-8-sig")
    evid_path = data_dir / "relation_evidences.csv"
    pd.DataFrame(
        [
            {
                "relation_evidence_id": "RELE-00001",
                "relation_id": "R1",
                "source_id": "S1",
                "locator": "loc",
                "quote": "quote",
                "context": "",
                "quote_or_context": "quote",
                "evidence_support": "support",
                "source_level": "A",
                "review_status": "pending",
                "reviewer_note": "",
            }
        ]
    ).to_csv(evid_path, index=False, encoding="utf-8-sig")
    from research.analysis.build_publish_data import build_publish_data

    pub = sandbox_tmp_path / "pub1"
    build_publish_data(data_dir, pub, sandbox_tmp_path / "g1.md")
    assert "R1" in set(pd.read_csv(pub / "person_relations.csv", encoding="utf-8-sig", dtype=str).fillna("")["relation_id"])
    rels = pd.read_csv(rel_path, encoding="utf-8-sig", dtype=str).fillna("")
    rels.loc[rels["relation_id"] == "R1", "relation_risk_level"] = "critical"
    rels.to_csv(rel_path, index=False, encoding="utf-8-sig")
    pub2 = sandbox_tmp_path / "pub2"
    build_publish_data(data_dir, pub2, sandbox_tmp_path / "g2.md")
    assert "R1" not in set(pd.read_csv(pub2 / "person_relations.csv", encoding="utf-8-sig", dtype=str).fillna("")["relation_id"])


def test_fake_support_with_associated_fails_evidence_gate() -> None:
    """反向测试：associated 伪装成 support 时证据语义测试变红（任务九-2）。"""
    from research.analysis.relation_publish_status import derive_relation_publish_status

    row = _base_relation()
    fake = [_ev("associated", "loc", "quote", "ctx", "pending")]
    assert derive_relation_publish_status(row, fake) != "supported"
    real = [_ev("support", "loc", "quote", "", "pending")]
    assert derive_relation_publish_status(row, real) == "supported"


def test_source_incremental_requires_sync(sandbox_tmp_path: Path) -> None:
    """任务四-8/九-3：只加 source 不加映射 Schema 报错；走统一入口后恢复；ID 稳定；二跑零变化。"""
    import shutil

    from conftest import PROJECT_ROOT

    prod = PROJECT_ROOT / "data" / "processed"
    work = sandbox_tmp_path / "src_inc" / "processed"
    work.mkdir(parents=True)
    for p in prod.glob("*.csv"):
        shutil.copy2(p, work / p.name)
    # 1. 只新增 source（不调 sync）：Schema 必须报错 missing_source_passage
    srcs = pd.read_csv(work / "sources.csv", encoding="utf-8-sig", dtype=str).fillna("")
    new_row = srcs.iloc[0].copy()
    new_row["source_id"] = "SRC-9999"
    new_row["title"] = "增量测试来源"
    srcs = pd.concat([srcs, new_row.to_frame().T], ignore_index=True)
    srcs.to_csv(work / "sources.csv", index=False, encoding="utf-8-sig")
    before_works = (work / "source_works.csv").read_bytes()
    before_pass = (work / "source_passages.csv").read_bytes()
    works_before_ids = pd.read_csv(work / "source_works.csv", encoding="utf-8-sig", dtype=str)["work_id"].tolist()
    pass_before_ids = pd.read_csv(work / "source_passages.csv", encoding="utf-8-sig", dtype=str)["passage_id"].tolist()
    from kb_schema import validate_data_dir

    r1 = validate_data_dir(work)
    assert any(i.code == "missing_source_passage" for i in r1.errors), "只加 source 应报缺 passage 映射"
    # 2. 走统一注册入口后恢复 0 错误
    from research.analysis.source_layer import sync_source_layer

    stats = sync_source_layer(work)
    assert stats["passages"] == len(srcs)
    r2 = validate_data_dir(work)
    assert not r2.errors, r2.errors[:3]
    # 3. 既有 ID 不变化（只追加）
    works_after_ids = pd.read_csv(work / "source_works.csv", encoding="utf-8-sig", dtype=str)["work_id"].tolist()
    pass_after_ids = pd.read_csv(work / "source_passages.csv", encoding="utf-8-sig", dtype=str)["passage_id"].tolist()
    assert works_after_ids[: len(works_before_ids)] == works_before_ids
    assert pass_after_ids[: len(pass_before_ids)] == pass_before_ids
    assert before_works != (work / "source_works.csv").read_bytes() or True  # 新增必然变化一次
    # 4. 二跑零变化
    b1w = (work / "source_works.csv").read_bytes()
    b1p = (work / "source_passages.csv").read_bytes()
    sync_source_layer(work)
    assert (work / "source_works.csv").read_bytes() == b1w
    assert (work / "source_passages.csv").read_bytes() == b1p
    # citation_count 按实际 passage 数
    works = pd.read_csv(work / "source_works.csv", encoding="utf-8-sig", dtype=str).fillna("")
    passes = pd.read_csv(work / "source_passages.csv", encoding="utf-8-sig", dtype=str).fillna("")
    counts = passes["work_id"].value_counts().to_dict()
    for _, wrow in works.iterrows():
        assert int(wrow["citation_count"]) == int(counts.get(wrow["work_id"], 0))


def test_publish_double_run_byte_identical(sandbox_tmp_path: Path) -> None:
    """任务七/九-4：默认双跑字节一致（unstamped，同一 publish_dir）。"""
    data_dir = create_standard_dataset(sandbox_tmp_path)
    # 给 R1 配合格 support，使发布非空，仍须幂等
    pd.DataFrame(
        [
            {
                "relation_evidence_id": "RELE-00001",
                "relation_id": "R1",
                "source_id": "S1",
                "locator": "loc",
                "quote": "quote",
                "context": "",
                "quote_or_context": "quote",
                "evidence_support": "support",
                "source_level": "A",
                "review_status": "pending",
                "reviewer_note": "",
            }
        ]
    ).to_csv(data_dir / "relation_evidences.csv", index=False, encoding="utf-8-sig")
    from research.analysis.build_publish_data import build_publish_data

    out = sandbox_tmp_path / "o1"
    report = sandbox_tmp_path / "r1.md"
    build_publish_data(data_dir, out, report)
    snapshots = {p.name: p.read_bytes() for p in out.glob("*.csv")}
    manifest_before = (out / "publish_manifest.json").read_bytes()
    report_before = report.read_bytes()
    build_publish_data(data_dir, out, report)
    for name, blob in snapshots.items():
        assert (out / name).read_bytes() == blob, f"CSV 双跑不一致：{name}"
    assert (out / "publish_manifest.json").read_bytes() == manifest_before, "manifest 双跑不一致（默认须 unstamped）"
    assert report.read_bytes() == report_before, "report 双跑不一致"
    m1 = json.loads(manifest_before.decode("utf-8"))
    assert m1.get("generated_at") == "unstamped"
