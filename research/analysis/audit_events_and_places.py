from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_DATA_DIR = PROJECT_ROOT / "data" / "processed"
DEFAULT_REPORT = PROJECT_ROOT / "research" / "drafts" / "reports" / "phase4_event_place_quality_report.md"
DEFAULT_QUEUE = PROJECT_ROOT / "research" / "drafts" / "reports" / "phase4_review_queue.csv"


def _read(path: Path) -> pd.DataFrame:
    return pd.read_csv(path, encoding="utf-8-sig").fillna("")


def _count(frame: pd.DataFrame, col: str) -> dict[str, int]:
    return frame[col].astype(str).value_counts().to_dict()


def audit_events_and_places(data_dir: Path, report_path: Path, queue_path: Path) -> dict[str, object]:
    data_dir = Path(data_dir)
    events = _read(data_dir / "events.csv")
    places = _read(data_dir / "places.csv")
    participants = _read(data_dir / "event_participants.csv")

    # Event type classification
    event_scope_map = {
        "entity": "entity",
        "collective": "collective",
        "organization": "collective",
    }
    events["event_type"] = events["event_scope"].map(event_scope_map).fillna("unknown")

    # Date precision issues
    year_only = int((events["date_precision"] == "年").sum())
    month_only = int((events["date_precision"] == "月").sum())
    day_precise = int((events["date_precision"] == "日").sum())

    # Review queue: low confidence + needs review
    review_mask = (events["needs_manual_review"] == "yes") | (events["confidence"] == "low")
    event_review = events[review_mask].copy()
    event_review["review_reason"] = ""
    event_review.loc[event_review["needs_manual_review"] == "yes", "review_reason"] += "标记为需人工复核; "
    event_review.loc[event_review["confidence"] == "low", "review_reason"] += "低置信度; "
    event_review.loc[event_review["date_precision"] == "年", "review_reason"] += "仅有年份精度; "

    # Place coordinate precision
    def infer_precision(lon: object, lat: object, place_type: str) -> str:
        try:
            lon_f = float(lon)
            lat_f = float(lat)
        except (ValueError, TypeError):
            return "unknown"
        if place_type in ("historical_place", "memorial_site", "campus"):
            return "street"
        if place_type == "street_site":
            return "district"
        if lon_f == round(lon_f, 1) and lat_f == round(lat_f, 1):
            return "city"
        if lon_f == round(lon_f, 2) and lat_f == round(lat_f, 2):
            return "district"
        return "street"

    places["coordinate_precision"] = places.apply(
        lambda row: infer_precision(row["longitude"], row["latitude"], row["place_type"]),
        axis=1,
    )
    places["geocoding_method"] = "auto_inferred"
    places["coordinate_note"] = places["coordinate_precision"].map(
        {
            "exact": "精确到建筑或门牌",
            "street": "精确到街道级别",
            "district": "精确到区级",
            "city": "城市中心坐标，非精确旧址",
            "unknown": "坐标缺失或无法解析",
        }
    )

    # Place review queue
    place_review = places[places["coordinate_precision"].isin({"city", "unknown"}) | (places["confidence"] == "low")].copy()
    place_review["review_reason"] = ""
    place_review.loc[place_review["coordinate_precision"] == "city", "review_reason"] += "城市级坐标伪装为精确地点; "
    place_review.loc[place_review["coordinate_precision"] == "unknown", "review_reason"] += "坐标无法解析; "
    place_review.loc[place_review["confidence"] == "low", "review_reason"] += "低置信度; "

    # Write enhanced places back
    places.to_csv(data_dir / "places.csv", index=False, encoding="utf-8-sig")

    # Build review queue CSV
    queue_rows = []
    for _, row in event_review.iterrows():
        queue_rows.append({
            "subject_type": "event",
            "subject_id": str(row["event_id"]),
            "subject_name": str(row["event_name"]),
            "review_reason": str(row["review_reason"]).strip(),
            "confidence": str(row["confidence"]),
            "date_precision": str(row["date_precision"]),
            "needs_manual_review": str(row["needs_manual_review"]),
        })
    for _, row in place_review.iterrows():
        queue_rows.append({
            "subject_type": "place",
            "subject_id": str(row["place_id"]),
            "subject_name": str(row["place_name"]),
            "review_reason": str(row["review_reason"]).strip(),
            "confidence": str(row["confidence"]),
            "date_precision": str(row["coordinate_precision"]),
            "needs_manual_review": "",
        })

    queue_df = pd.DataFrame(
        queue_rows,
        columns=["subject_type", "subject_id", "subject_name", "review_reason", "confidence", "date_precision", "needs_manual_review"],
    )
    queue_path.parent.mkdir(parents=True, exist_ok=True)
    queue_df.to_csv(queue_path, index=False, encoding="utf-8-sig")

    # Summary
    summary = {
        "events_total": len(events),
        "event_types": _count(events, "event_type"),
        "date_precision": {"年": year_only, "月": month_only, "日": day_precise},
        "event_review_queue_size": len(event_review),
        "places_total": len(places),
        "coordinate_precision": _count(places, "coordinate_precision"),
        "place_review_queue_size": len(place_review),
        "review_queue_total": len(queue_rows),
    }

    # Write report
    lines = [
        "# Phase 4 事件与地点质量报告",
        "",
        "## 事件质量",
        "",
        f"- 事件总数：{len(events)}",
        f"- 事件类型分布：{dict(_count(events, 'event_type'))}",
        f"- 日期精度分布：年 {year_only}、月 {month_only}、日 {day_precise}",
        f"- 低置信度事件：{int((events['confidence'] == 'low').sum())}",
        f"- 需人工复核事件：{int((events['needs_manual_review'] == 'yes').sum())}",
        f"- 事件审核队列：{len(event_review)} 条",
        "",
        "## 地点质量",
        "",
        f"- 地点总数：{len(places)}",
        f"- 坐标精度分布：{dict(_count(places, 'coordinate_precision'))}",
        f"- 低置信度地点：{int((places['confidence'] == 'low').sum())}",
        f"- 地点审核队列：{len(place_review)} 条",
        "",
        "## 审核队列",
        "",
        f"- 总审核条目：{len(queue_rows)}",
        f"- 队列文件：`{queue_path.name}`",
        "",
        "## 结论",
        "",
        "事件定义已按 scope 分为 entity 和 collective。地点已新增 coordinate_precision 字段，",
        "城市级坐标被明确标记。低置信度和模糊记录已进入审核队列。",
    ]
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text("\n".join(lines) + "\n", encoding="utf-8")

    return summary


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="审计事件与地点数据质量。")
    parser.add_argument("--data-dir", type=Path, default=DEFAULT_DATA_DIR)
    parser.add_argument("--report", type=Path, default=DEFAULT_REPORT)
    parser.add_argument("--queue", type=Path, default=DEFAULT_QUEUE)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    summary = audit_events_and_places(args.data_dir, args.report, args.queue)
    print(summary)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
