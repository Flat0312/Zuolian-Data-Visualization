from __future__ import annotations

import argparse
import sys
from collections.abc import Mapping
from pathlib import Path

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from kb_schema import SOURCE_CLASSIFICATION_COLUMNS

DEFAULT_SOURCES_PATH = PROJECT_ROOT / "data" / "processed" / "sources.csv"


def _normalize_text(row: Mapping[str, object]) -> str:
    parts = [
        str(row.get("source_kind", "") or "").strip(),
        str(row.get("title", "") or "").strip(),
        str(row.get("citation", "") or "").strip(),
        str(row.get("evidence_layer", "") or "").strip(),
    ]
    return " ".join(part for part in parts if part).lower()


def classify_source_row(row: Mapping[str, object]) -> dict[str, str]:
    content = _normalize_text(row)
    source_kind = str(row.get("source_kind", "") or "").strip().lower()

    if "日记" in content:
        return {
            "evidence_strength": "一手",
            "evidence_type": "日记",
            "needs_manual_review": "no",
            "review_note": "",
            "classification_rule": "keyword:日记",
        }
    if any(keyword in content for keyword in ("书信", "信件", "函", "致 ")):
        return {
            "evidence_strength": "一手",
            "evidence_type": "信件",
            "needs_manual_review": "no",
            "review_note": "",
            "classification_rule": "keyword:信件",
        }
    if "回忆录" in content:
        return {
            "evidence_strength": "二手",
            "evidence_type": "回忆录",
            "needs_manual_review": "no",
            "review_note": "",
            "classification_rule": "keyword:回忆录",
        }
    if "年谱" in content:
        return {
            "evidence_strength": "二手",
            "evidence_type": "年谱",
            "needs_manual_review": "no",
            "review_note": "",
            "classification_rule": "keyword:年谱",
        }
    if source_kind == "raw_workbook" or any(keyword in content for keyword in ("表格", "目录", "excel_candidate_fact")):
        return {
            "evidence_strength": "推断",
            "evidence_type": "档案表格",
            "needs_manual_review": "yes",
            "review_note": "候选事实来自表格整理结果，建议人工复核。",
            "classification_rule": "fallback:档案表格",
        }
    if any(keyword in content for keyword in ("词典", "研究", "论文", "论著", "专著", "左联史", "史 ")):
        return {
            "evidence_strength": "二手",
            "evidence_type": "研究论著",
            "needs_manual_review": "no",
            "review_note": "",
            "classification_rule": "keyword:研究论著",
        }
    if source_kind == "citation_only":
        return {
            "evidence_strength": "转引",
            "evidence_type": "研究论著",
            "needs_manual_review": "yes",
            "review_note": "仅有引文线索，建议补原始来源后再确认。",
            "classification_rule": "fallback:citation_only",
        }
    if source_kind == "web_url":
        return {
            "evidence_strength": "转引",
            "evidence_type": "研究论著",
            "needs_manual_review": "yes",
            "review_note": "网页来源默认按转引处理，建议人工确认原始出处。",
            "classification_rule": "fallback:web_url",
        }
    return {
        "evidence_strength": "推断",
        "evidence_type": "研究论著",
        "needs_manual_review": "yes",
        "review_note": "规则未能稳定识别来源类型，建议人工复核。",
        "classification_rule": "fallback:unknown",
    }


def apply_source_classification(input_path: Path, output_path: Path | None = None) -> pd.DataFrame:
    frame = pd.read_csv(input_path, encoding="utf-8-sig").fillna("")
    classified = frame.copy()

    for column in SOURCE_CLASSIFICATION_COLUMNS:
        if column not in classified.columns:
            classified[column] = ""

    for index, row in classified.iterrows():
        inferred = classify_source_row(row.to_dict())
        for column in SOURCE_CLASSIFICATION_COLUMNS:
            existing = str(classified.at[index, column]).strip()
            if existing:
                continue
            classified.at[index, column] = inferred[column]

    destination = output_path or input_path
    destination.parent.mkdir(parents=True, exist_ok=True)
    classified.to_csv(destination, index=False, encoding="utf-8-sig")
    return classified


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="为 sources.csv 生成初版证据强度与类型分类。")
    parser.add_argument("--input", type=Path, default=DEFAULT_SOURCES_PATH, help="输入 sources.csv 路径")
    parser.add_argument("--output", type=Path, default=None, help="输出路径；默认原地覆盖输入文件")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    frame = apply_source_classification(args.input, args.output)
    print(f"classified_sources={len(frame)} output={args.output or args.input}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
