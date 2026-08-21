from __future__ import annotations


def test_classify_source_row_assigns_expected_strength_and_type() -> None:
    from research.analysis.classify_sources import classify_source_row

    diary = classify_source_row(
        {
            "source_kind": "local_text_citation",
            "title": "鲁迅日记",
            "citation": "鲁迅日记 1933年5月17日",
            "evidence_layer": "txt_local_evidence",
            "availability": "local",
        }
    )
    assert diary["evidence_strength"] == "一手"
    assert diary["evidence_type"] == "日记"
    assert diary["needs_manual_review"] == "no"

    research = classify_source_row(
        {
            "source_kind": "local_text_citation",
            "title": "左联词典",
            "citation": "左联词典 第96页",
            "evidence_layer": "txt_local_evidence",
            "availability": "local",
        }
    )
    assert research["evidence_strength"] == "二手"
    assert research["evidence_type"] == "研究论著"

    ambiguous = classify_source_row(
        {
            "source_kind": "raw_workbook",
            "title": "《左联相关档案资源目录》原始表格",
            "citation": "",
            "evidence_layer": "excel_candidate_fact",
            "availability": "local",
        }
    )
    assert ambiguous["evidence_strength"] == "推断"
    assert ambiguous["evidence_type"] == "档案表格"
    assert ambiguous["needs_manual_review"] == "yes"
