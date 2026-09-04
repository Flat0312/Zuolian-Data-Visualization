"""时空与枚举审计脚本（Agent B，只读生产数据，只出建议，不修改任何数据）。

审计内容：
1. events 的 start_date/end_date/date_certainty 需求与缺口；
2. 同名不同日期事件的 canonical_event_key 唯一性；
3. participant_role 枚举混用（直接参与者/关联人物/待核/unclear/不明…）；
4. 地点合并候选（内山书店/内山书店旧址、东方旅社系列、光华书局系列等，含同坐标异名检测）；
5. 泛化挂接到城市级地点（如"上海"，83 个事件）的具体清单；
6. 低置信或未知精度坐标清单。

所有候选 decision 一律 pending_human_review；无法确定的历史沿革不自动合并。

输出（默认 research/drafts/reports/）：
- event_normalization_candidates.csv / place_normalization_candidates.csv
- participant_role_mapping_candidates.csv / temporal_spatial_audit_report.md
"""
from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_DATA_DIR = PROJECT_ROOT / "data" / "processed"
DEFAULT_OUT_DIR = PROJECT_ROOT / "research/drafts" / "reports"

SNAPSHOT_DATE = "2026-09-04"

PENDING = "pending_human_review"

CANONICAL_ROLES = ("直接参与者", "关联人物", "待核", "不明")
LEGACY_ROLE_MAP = {"unclear": "不明", "发起人": "待核", "组织者": "直接参与者", "发起人/组织者": "直接参与者"}

# 规则式地点合并候选（名称归一后视为疑似同一地点；一律待人工确认沿革）
PLACE_MERGE_RULES: list[tuple[tuple[str, ...], str]] = [
    (("内山书店", "内山书店旧址"), "历史店名与旧址名疑似同一地点；注意现址四川北路2050号为1929年迁入，1929年前的店址沿革需人工核对后再合并"),
    (("东方旅社", "上海东方旅社", "上海东方旅社31号房间", "上海东方旅社、中山旅社秘密会议"), "同一旅社的不同表述/房间级记录；建筑级与房间级层级不同，是否合并到建筑级需人工裁决"),
    (("光华书局", "光华书局（四马路）"), "疑似同一书局不同门牌表述，需人工核对地址沿革"),
    (("良友图书公司", "良友图书"), "疑似同一机构简称，需人工确认"),
    (("生活书店", "生活书店（福州路）"), "疑似同一书店不同表述，需人工确认"),
    (("创造社（宝山路）", "创造社（闸北宝山路）"), "疑似同一旧址不同区划表述，需人工确认"),
    (("复旦大学", "复旦大学江湾校区"), "校区与本部关系需人工确认，不自动合并"),
    (("上海大学", "上海大学中国文学系"), "机构与院系层级不同，不建议合并，仅建议建立层级关联"),
]


def _read(data_dir: Path, name: str) -> pd.DataFrame:
    return pd.read_csv(data_dir / name, encoding="utf-8-sig", dtype=str).fillna("")


def _norm_name(name: str) -> str:
    s = str(name).strip()
    for ch in "（）()旧址址":
        s = s.replace(ch, "")
    return s


def build(data_dir: Path, out_dir: Path) -> dict[str, object]:
    events = _read(data_dir, "events.csv")
    places = _read(data_dir, "places.csv")
    parts = _read(data_dir, "event_participants.csv")

    event_cols = set(events.columns)
    place_cols = set(places.columns)

    ev_rows: list[dict[str, str]] = []

    # —— 1. start_date/end_date/date_certainty 需求 ——
    time_fields = ("start_date", "end_date", "date_certainty")
    has_time_cols = all(c in event_cols for c in time_fields)
    if not has_time_cols:
        ev_rows.append(
            {
                "candidate_id": "ENC-T0001",
                "type": "missing_time_columns",
                "event_id": "",
                "detail": "events.csv 缺少 start_date/end_date/date_certainty 列（或部分缺失）",
                "suggestion": "按 event_date/date_precision 派生：start=end=event_date；日→exact、月→approximate_month、年→approximate_year、空→uncertain",
                "decision": PENDING,
            }
        )
    else:
        for i, row in enumerate(events.to_dict("records"), start=1):
            missing = [f for f in time_fields if not str(row.get(f, "")).strip()]
            if missing:
                ev_rows.append(
                    {
                        "candidate_id": f"ENC-T{i:04d}",
                        "type": "missing_time_value",
                        "event_id": row.get("event_id", ""),
                        "detail": f"缺失字段：{'、'.join(missing)}（event_date={row.get('event_date', '')!r}, date_precision={row.get('date_precision', '')!r}）",
                        "suggestion": "按 event_date/date_precision 派生；event_date 为空者 date_certainty=uncertain",
                        "decision": PENDING,
                    }
                )

    # —— 2. canonical_event_key 唯一性（同名不同日期） ——
    if "canonical_event_key" in event_cols:
        key_counts = events["canonical_event_key"].str.strip().value_counts()
        dups = key_counts[key_counts > 1]
        n = 0
        for key, count in dups.items():
            for row in events[events["canonical_event_key"].str.strip() == key].to_dict("records"):
                n += 1
                ev_rows.append(
                    {
                        "candidate_id": f"ENC-K{n:04d}",
                        "type": "duplicate_canonical_event_key",
                        "event_id": row.get("event_id", ""),
                        "detail": f"canonical_event_key={key!r} 被 {count} 个事件共用（本事件日期 {row.get('event_date', '')!r}）",
                        "suggestion": "同名不同日期事件应在 key 末尾追加 |日期 消歧；同日同名事件需人工判断是否为重复条目",
                        "decision": PENDING,
                    }
                )
        # 同名不同日期但 key 恰好不同的，也登记（供人工核对 key 定义完整性）
        name_groups = events.groupby(events["event_name"].str.strip())
        m = 0
        for name, group in name_groups:
            if len(group) < 2:
                continue
            dates = sorted({d for d in group["event_date"].tolist() if d})
            if len(dates) > 1:
                for row in group.to_dict("records"):
                    m += 1
                    ev_rows.append(
                        {
                            "candidate_id": f"ENC-N{m:04d}",
                            "type": "same_name_different_date",
                            "event_id": row.get("event_id", ""),
                            "detail": f"事件名 {name!r} 存在 {len(dates)} 个不同日期（{dates[0]}…），本条日期 {row.get('event_date', '')!r}",
                            "suggestion": "确认 key 已含日期段；若为不同史实建议事件名具体化，若为重复建议走人工合并流程",
                            "decision": PENDING,
                        }
                    )

    # —— 5. 泛化城市级挂接 ——
    generic_place_ids = set(
        places[
            (places["place_type"] == "city") | (places["historical_name"].str.strip() == "上海")
        ]["place_id"]
    )
    g = 0
    for row in events.to_dict("records"):
        if row.get("place_id", "") in generic_place_ids:
            g += 1
            ev_rows.append(
                {
                    "candidate_id": f"ENC-G{g:04d}",
                    "type": "generic_city_attachment",
                    "event_id": row.get("event_id", ""),
                    "detail": f"事件挂接到城市级地点 {row.get('place_id', '')}（上海），historical_location={row.get('historical_location', '')!r}",
                    "suggestion": "依据证据将地点细化到街区/建筑级；无法细化者保留并在展示层标注城市级精度",
                    "decision": PENDING,
                }
            )

    # 空地点挂接
    h = 0
    for row in events.to_dict("records"):
        if not row.get("place_id", "").strip() and row.get("historical_location", "").strip():
            h += 1
            ev_rows.append(
                {
                    "candidate_id": f"ENC-P{h:04d}",
                    "type": "missing_place_id",
                    "event_id": row.get("event_id", ""),
                    "detail": f"无 place_id 但有 historical_location={row.get('historical_location', '')!r}",
                    "suggestion": "人工为该地点建档或挂接到现存地点",
                    "decision": PENDING,
                }
            )

    # —— 3. participant_role 映射候选 ——
    role_rows = []
    if "participant_role" in parts.columns:
        counts = parts["participant_role"].str.strip().value_counts()
        for i, (value, count) in enumerate(counts.items(), start=1):
            if value in CANONICAL_ROLES:
                rec, status = "保留（枚举内）", "no_change_needed"
            elif value in LEGACY_ROLE_MAP:
                rec, status = f"建议映射为 {LEGACY_ROLE_MAP[value]}", PENDING
            else:
                rec, status = "非预设枚举值，人工归类到 直接参与者/关联人物/待核/不明 之一", PENDING
            role_rows.append(
                {
                    "candidate_id": f"PRM-{i:03d}",
                    "observed_value": value,
                    "row_count": int(count),
                    "in_canonical_enum": value in CANONICAL_ROLES,
                    "mapping_recommendation": rec,
                    "decision": status,
                }
            )
    else:
        role_rows.append(
            {
                "candidate_id": "PRM-001",
                "observed_value": "",
                "row_count": 0,
                "in_canonical_enum": False,
                "mapping_recommendation": "event_participants.csv 无 participant_role 列",
                "decision": PENDING,
            }
        )

    # —— 4/6. 地点候选 ——
    plc_rows: list[dict[str, str]] = []
    id_by_name: dict[str, list[str]] = {}
    for row in places.to_dict("records"):
        id_by_name.setdefault(str(row["historical_name"]).strip(), []).append(str(row["place_id"]).strip())
    n = 0
    for group_names, reason in PLACE_MERGE_RULES:
        ids: list[str] = []
        for name in group_names:
            ids.extend(id_by_name.get(name, []))
        ids = sorted(set(ids))
        if len(ids) >= 2:
            n += 1
            plc_rows.append(
                {
                    "candidate_id": f"PNC-M{n:03d}",
                    "type": "merge_candidate",
                    "place_id": ";".join(ids),
                    "place_names": ";".join(group_names),
                    "detail": reason,
                    "suggestion": "仅登记合并建议；需人工确认历史沿革后另行执行，本脚本不自动合并",
                    "decision": PENDING,
                }
            )
    # 同坐标异名
    coord_groups: dict[tuple[str, str], list[str]] = {}
    for row in places.to_dict("records"):
        lon, lat = str(row.get("longitude", "")).strip(), str(row.get("latitude", "")).strip()
        if lon and lat:
            coord_groups.setdefault((lon, lat), []).append(f"{row['place_id']}:{row['historical_name']}")
    m = 0
    for (lon, lat), members in coord_groups.items():
        if len(members) < 2:
            continue
        names = [mem.split(":", 1)[1] for mem in members]
        if len(set(_norm_name(x) for x in names)) <= 1:
            continue  # 已被名称规则覆盖的近似同名不再重复登记
        m += 1
        plc_rows.append(
            {
                "candidate_id": f"PNC-C{m:03d}",
                "type": "same_coordinate_different_name",
                "place_id": ";".join(mem.split(":", 1)[0] for mem in members),
                "place_names": ";".join(names),
                "detail": f"坐标完全相同（{lon},{lat}）但名称归一后不同",
                "suggestion": "人工判断是同一地点多表述、同一建筑多机构，还是概略坐标复用；不得自动合并",
                "decision": PENDING,
            }
        )
    # 坐标质量
    q = 0
    for row in places.to_dict("records"):
        issues = []
        conf = str(row.get("confidence", "")).strip().lower()
        prec = str(row.get("coordinate_precision", "")).strip() or str(row.get("coord_precision", "")).strip()
        lon, lat = str(row.get("longitude", "")).strip(), str(row.get("latitude", "")).strip()
        if conf == "low":
            issues.append("confidence=low")
        if prec.lower() in ("", "unknown", "nan"):
            issues.append("坐标精度未知/未标注")
        if not lon or not lat:
            issues.append("缺坐标")
        if prec.lower() == "city":
            issues.append("城市级概略坐标")
        if issues:
            q += 1
            plc_rows.append(
                {
                    "candidate_id": f"PNC-Q{q:03d}",
                    "type": "coordinate_quality",
                    "place_id": str(row["place_id"]),
                    "place_names": str(row["historical_name"]),
                    "detail": "；".join(issues),
                    "suggestion": "补坐标来源（测绘/文物保护单位页面/地图定位）并人工确认精度等级",
                    "decision": PENDING,
                }
            )

    # —— 写出 ——
    out_dir.mkdir(parents=True, exist_ok=True)
    pd.DataFrame(
        ev_rows,
        columns=["candidate_id", "type", "event_id", "detail", "suggestion", "decision"],
    ).to_csv(out_dir / "event_normalization_candidates.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(
        plc_rows,
        columns=["candidate_id", "type", "place_id", "place_names", "detail", "suggestion", "decision"],
    ).to_csv(out_dir / "place_normalization_candidates.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(
        role_rows,
        columns=["candidate_id", "observed_value", "row_count", "in_canonical_enum", "mapping_recommendation", "decision"],
    ).to_csv(out_dir / "participant_role_mapping_candidates.csv", index=False, encoding="utf-8-sig")

    summary = {
        "time_columns_present": has_time_cols,
        "event_candidates": len(ev_rows),
        "event_types": pd.Series([r["type"] for r in ev_rows]).value_counts().to_dict(),
        "place_candidates": len(plc_rows),
        "place_types": pd.Series([r["type"] for r in plc_rows]).value_counts().to_dict(),
        "role_values": {r["observed_value"]: r["row_count"] for r in role_rows},
        "generic_city_events": g,
        "missing_place_id_events": h,
    }
    _write_report(out_dir, summary, data_dir)
    return summary


def _write_report(out_dir: Path, summary: dict[str, object], data_dir: Path) -> None:
    lines = [
        "# 时空与枚举审计报告（Agent B，只读审计）",
        "",
        f"- 数据快照：`{data_dir}`（读取日期 {SNAPSHOT_DATE}，只读；未修改任何生产数据）",
        f"- events 时间列（start_date/end_date/date_certainty）：{'已存在' if summary['time_columns_present'] else '缺失'}",
        f"- 事件级候选：{summary['event_candidates']} 条；分布：{summary['event_types']}",
        f"- 地点级候选：{summary['place_candidates']} 条；分布：{summary['place_types']}",
        f"- participant_role 现值分布：{summary['role_values']}",
        f"- 泛化挂接到城市级地点（上海）的事件：{summary['generic_city_events']} 个",
        f"- 无 place_id 但有历史地点名的事件：{summary['missing_place_id_events']} 个",
        "",
        "## 审计口径",
        "",
        "1. **时间字段**：缺列或缺值均登记为候选，建议按 event_date/date_precision 派生（日→exact、月→approximate_month、年→approximate_year、空→uncertain）。",
        "2. **canonical_event_key**：重复 key 与『同名不同日期』分别登记；同日同名疑似重复条目只提示人工合并流程，不自动合并。",
        "3. **participant_role**：以 直接参与者/关联人物/待核/不明 为规范枚举；legacy 值（unclear/发起人 等）给出映射建议；非预设值一律人工归类。",
        "4. **地点合并**：仅规则式候选（名称组 + 同坐标异名检测），全部 pending_human_review；涉及历史沿革（如内山书店 1929 年迁入四川北路2050号）须人工核对后再执行。",
        "5. **城市级泛化**：挂接到『上海』等城市级地点的事件逐条列出，建议按证据细化到街区/建筑级，无法细化者保留并在展示层标注精度。",
        "6. **坐标质量**：confidence=low、精度未知、缺坐标、城市级概略坐标逐条列出，建议补坐标来源后人工定级。",
        "",
        "## 证据局限与边界",
        "",
        "- 本审计只描述当前数据的结构与口径问题，**不构成对任何历史事实的判定**。",
        "- 所有 decision=pending_human_review；合并/改名/改期/挂接调整均需人工裁决后另行执行。",
        "- 『上海东方旅社、中山旅社秘密会议』这类复合名地点，其与『东方旅社』『中山旅社』的层级关系（事件名 vs 地点名）需人工先厘清语义再谈合并。",
        "",
    ]
    (out_dir / "temporal_spatial_audit_report.md").write_text("\n".join(lines), encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="时空与枚举审计（只读）")
    parser.add_argument("--data-dir", type=Path, default=DEFAULT_DATA_DIR)
    parser.add_argument("--out-dir", type=Path, default=DEFAULT_OUT_DIR)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    summary = build(args.data_dir, args.out_dir)
    print(f"temporal spatial audit: {summary}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
