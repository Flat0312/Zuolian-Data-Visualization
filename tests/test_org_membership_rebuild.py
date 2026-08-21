from __future__ import annotations

from pathlib import Path

import pandas as pd


def test_classify_membership_confirms_one_a_level_source() -> None:
    from research.analysis.rebuild_org_memberships import classify_membership

    decision = classify_membership(
        [{"evidence_support": "membership", "source_level": "A", "source_work": "组织名录"}],
        fallback_role="普通成员",
    )

    assert decision.membership_type == "confirmed_member"
    assert decision.confidence == "high"
    assert decision.needs_manual_review == "no"


def test_classify_membership_confirms_two_independent_b_sources() -> None:
    from research.analysis.rebuild_org_memberships import classify_membership

    decision = classify_membership(
        [
            {"evidence_support": "membership", "source_level": "B", "source_work": "左联词典"},
            {"evidence_support": "membership", "source_level": "B", "source_work": "左联史"},
        ],
        fallback_role="普通成员",
    )

    assert decision.membership_type == "confirmed_member"
    assert decision.decision_rule == "two_independent_b_sources"


def test_classify_membership_keeps_insufficient_member_evidence_as_candidate() -> None:
    from research.analysis.rebuild_org_memberships import classify_membership

    decision = classify_membership(
        [{"evidence_support": "membership", "source_level": "B", "source_work": "左联词典"}],
        fallback_role="普通成员",
    )

    assert decision.membership_type == "candidate"
    assert decision.needs_manual_review == "yes"


def test_classify_membership_uses_related_fallback_without_member_evidence() -> None:
    from research.analysis.rebuild_org_memberships import classify_membership

    decision = classify_membership(
        [{"evidence_support": "lead", "source_level": "D", "source_work": "原始表格"}],
        fallback_role="外围联络人",
    )

    assert decision.membership_type == "related_person"
    assert decision.confidence == "medium"


def test_classify_membership_marks_conflicting_evidence_as_disputed() -> None:
    from research.analysis.rebuild_org_memberships import classify_membership

    decision = classify_membership(
        [
            {"evidence_support": "membership", "source_level": "A", "source_work": "组织名录"},
            {"evidence_support": "oppose", "source_level": "A", "source_work": "更正名录"},
        ],
        fallback_role="普通成员",
    )

    assert decision.membership_type == "disputed"
    assert decision.needs_manual_review == "yes"


def test_extract_person_evidence_handles_ocr_spaces_and_page_sources() -> None:
    from research.analysis.rebuild_org_memberships import extract_person_evidence

    source_text = """
第 10 页
鲁 迅 是 左 联 常 务 委 员 。
第 11 页
巴 金 没 有 加 入 左 联 ，但与左联作家站在同一阵线。
"""

    evidence = extract_person_evidence(
        source_text=source_text,
        people=[
            {"person_id": "P1", "standard_name": "鲁迅", "aliases": ""},
            {"person_id": "P2", "standard_name": "巴金", "aliases": ""},
        ],
        source_work="左联词典",
        page_source_ids={10: "S10", 11: "S11"},
    )

    luxun = [row for row in evidence if row["person_id"] == "P1"]
    bajin = [row for row in evidence if row["person_id"] == "P2"]

    assert luxun[0]["evidence_support"] == "membership"
    assert luxun[0]["source_id"] == "S10"
    assert luxun[0]["locator"] == "第10页"
    assert bajin[0]["evidence_support"] == "oppose"


def test_extract_person_evidence_does_not_apply_another_persons_opposition() -> None:
    from research.analysis.rebuild_org_memberships import extract_person_evidence

    source_text = """
第 20 页
本书介绍周扬、吴奚如、舒群、白朗等作家，但白朗不是左联盟员，此处有误。
"""

    evidence = extract_person_evidence(
        source_text=source_text,
        people=[
            {"person_id": "P1", "standard_name": "周扬", "aliases": ""},
            {"person_id": "P2", "standard_name": "白朗", "aliases": ""},
        ],
        source_work="左联词典",
        page_source_ids={20: "S20"},
    )

    assert not [row for row in evidence if row["person_id"] == "P1"]
    assert [row for row in evidence if row["person_id"] == "P2"][0]["evidence_support"] == "oppose"


def test_extract_person_evidence_does_not_include_people_after_contrast_marker() -> None:
    from research.analysis.rebuild_org_memberships import extract_person_evidence

    source_text = """
第 21 页
特约撰稿员中有鲁迅、茅盾等左联盟员作家，还有巴金、老舍等知名作家。
"""

    evidence = extract_person_evidence(
        source_text=source_text,
        people=[
            {"person_id": "P1", "standard_name": "鲁迅", "aliases": ""},
            {"person_id": "P2", "standard_name": "巴金", "aliases": ""},
        ],
        source_work="左联史",
        page_source_ids={21: "S21"},
    )

    assert [row for row in evidence if row["person_id"] == "P1"][0]["evidence_support"] == "membership"
    assert not [row for row in evidence if row["person_id"] == "P2"]


def test_extract_person_evidence_does_not_treat_work_author_as_member() -> None:
    from research.analysis.rebuild_org_memberships import extract_person_evidence

    source_text = """
第 22 页
发表的左联成员著译有：《关于创作技巧》（高尔基作，林林译）。
"""

    evidence = extract_person_evidence(
        source_text=source_text,
        people=[{"person_id": "P1", "standard_name": "高尔基", "aliases": ""}],
        source_work="左联史",
        page_source_ids={22: "S22"},
    )

    assert not evidence


def test_extract_person_evidence_recognizes_explicit_nonmember_list() -> None:
    from research.analysis.rebuild_org_memberships import extract_person_evidence

    source_text = """
第 23 页
叶圣陶、巴金、王统照、老舍等虽未加入左联，但与左联作家站在同一阵线。
"""

    evidence = extract_person_evidence(
        source_text=source_text,
        people=[
            {"person_id": "P1", "standard_name": "巴金", "aliases": ""},
            {"person_id": "P2", "standard_name": "王统照", "aliases": ""},
        ],
        source_work="左联词典",
        page_source_ids={23: "S23"},
    )

    assert {row["person_id"] for row in evidence if row["evidence_support"] == "oppose"} == {"P1", "P2"}


def test_build_org_membership_outputs_ledger_and_derived_conclusions(tmp_path: Path) -> None:
    from research.analysis.rebuild_org_memberships import rebuild_org_memberships

    data_dir = tmp_path / "data"
    runtime_dir = data_dir / "runtime_sources"
    runtime_dir.mkdir(parents=True)

    pd.DataFrame(
        [
            {"person_id": "P1", "standard_name": "鲁迅", "aliases": "", "role": "核心领导"},
            {"person_id": "P2", "standard_name": "巴金", "aliases": "", "role": "外围联络人"},
        ]
    ).to_csv(data_dir / "persons.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(
        [
            {
                "membership_id": "M1",
                "organization_id": "ORG-001",
                "person_id": "P1",
                "membership_role": "成员",
                "membership_type": "member",
                "source_ids": "S0",
                "confidence": "high",
                "needs_manual_review": "no",
            },
            {
                "membership_id": "M2",
                "organization_id": "ORG-001",
                "person_id": "P2",
                "membership_role": "成员",
                "membership_type": "member",
                "source_ids": "S0",
                "confidence": "high",
                "needs_manual_review": "no",
            },
        ]
    ).to_csv(data_dir / "org_memberships.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(
        [
            {
                "source_id": "S0",
                "source_kind": "raw_workbook",
                "title": "原始表格",
                "citation": "",
                "source_path": "",
                "source_url": "",
                "evidence_layer": "excel_candidate_fact",
                "availability": "local",
                "evidence_strength": "推断",
                "evidence_type": "档案表格",
                "needs_manual_review": "yes",
                "review_note": "",
                "classification_rule": "fallback",
            },
            {
                "source_id": "S10",
                "source_kind": "local_text_citation",
                "title": "左联词典",
                "citation": "左联词典 第10页",
                "source_path": "左联词典.txt",
                "source_url": "",
                "evidence_layer": "txt_local_evidence",
                "availability": "local",
                "evidence_strength": "一手",
                "evidence_type": "组织名录",
                "needs_manual_review": "no",
                "review_note": "",
                "classification_rule": "test",
            },
        ]
    ).to_csv(data_dir / "sources.csv", index=False, encoding="utf-8-sig")
    (runtime_dir / "左联词典.txt").write_text(
        "第 10 页\n鲁 迅 是 左 联 常 务 委 员 。\n",
        encoding="utf-8",
    )

    summary = rebuild_org_memberships(data_dir=data_dir, runtime_sources_dir=runtime_dir)

    ledger = pd.read_csv(data_dir / "org_membership_evidences.csv")
    memberships = pd.read_csv(data_dir / "org_memberships.csv")

    assert summary["membership_count"] == 2
    assert set(ledger["person_id"]) == {"P1", "P2"}
    assert memberships.loc[memberships["person_id"] == "P1", "membership_type"].iloc[0] == "confirmed_member"
    assert memberships.loc[memberships["person_id"] == "P2", "membership_type"].iloc[0] == "related_person"


def test_standard_pipeline_never_confirms_membership_from_person_role_alone() -> None:
    from research.analysis.build_standard_kb_pipeline import initial_membership_decision_for_role

    direct = initial_membership_decision_for_role("核心领导")
    related = initial_membership_decision_for_role("外围联络人")

    assert direct == ("candidate", "low", "yes")
    assert related == ("related_person", "medium", "yes")
