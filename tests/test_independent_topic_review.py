from __future__ import annotations

import unicodedata
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
REVIEW_PATH = (
    REPO_ROOT
    / "research"
    / "content-upgrade-overnight-2026-09-06"
    / "verification"
    / "independent_topic_review_2026-09-20.md"
)
HISTORY = REPO_ROOT / "data" / "processed" / "runtime_sources" / "左联史.txt"
DICTIONARY = REPO_ROOT / "data" / "processed" / "runtime_sources" / "左联词典.txt"


def _normalize(text: str) -> str:
    text = unicodedata.normalize("NFKC", text)
    return text.translate(str.maketrans("", "", " \u3000\t\r\n·"))


def test_review_file_covers_four_topics_with_verdicts() -> None:
    text = REVIEW_PATH.read_text(encoding="utf-8")

    for topic in ("T2", "T3", "T4", "T5"):
        assert f"## {topic}" in text
    assert text.count("issues_found") >= 3
    assert "pass_with_notes" in text
    assert "pending_human_review" in text
    for banned in ("已转正", "人工审核通过", "人工复核通过", "已执行", "已升级"):
        assert banned not in text


def test_review_cites_verifiable_ocr_originals_from_sources() -> None:
    text = REVIEW_PATH.read_text(encoding="utf-8")
    history = _normalize(HISTORY.read_text(encoding="utf-8", errors="ignore"))
    dictionary = _normalize(DICTIONARY.read_text(encoding="utf-8", errors="ignore"))

    # 三个静默校改的原文锚必须真实存在于来源，防止审核报告自身漂移
    anchors = {
        "钱杏邱被推为大会主席团": history,
        "栖石、胡也频等廿三烈士": dictionary,
        "倚重兼迅": dictionary,
    }
    for phrase, source in anchors.items():
        assert phrase in source, f"OCR 原文锚未命中: {phrase}"
        assert phrase in text, f"审核报告未引用原文锚: {phrase}"

    # 报告必须写明校改后的引文形式，构成「原文 vs 引文」对照
    for corrected in ("钱杏邨", "柔石、胡也频", "倚重鲁迅"):
        assert corrected in text
