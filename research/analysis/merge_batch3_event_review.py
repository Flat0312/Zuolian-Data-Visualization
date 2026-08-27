"""把第三批事件史料候选包按事件级审核决策合并进生产数据。

依据：phase2_batch3_event_review_decisions.md（16 项事件级决策全部落地；
EVT-00006+EVT-00007、EVT-00120+EVT-00119 两组按物理删除合并）。
授权状态：**授权待追认**（决策文档第 7 节）；授权补齐前禁止开展下一批物理合并。

幂等保证：
- 以 fact_evidences.origin_evidence_id（FE-EVP3-*）识别已转正证据；全部已转正时早退不写盘；
- 来源复用映射=已转正证据回推+生产表 URL 查重兜底；SRC-EVP3-010 固定复用生产 SRC-1163；
- 事件字段修正全部为终值赋值或带标记检查的追加，重复执行不产生重复文本或新 ID；
- 物理删除与参与者重定向均带存在性守卫，事件已删时静默跳过。

裁决语义：conflict 证据合并时写入 adjudication_status=resolved_by_event_correction，
其余行为空值——机器因此可区分“原始冲突”“已裁决冲突”与“直接支持”。
"""

from __future__ import annotations

import csv
import hashlib
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
DATA = ROOT / "data" / "processed"
DRAFTS = ROOT / "research" / "drafts" / "reports"

BATCH3_SOURCES = DRAFTS / "phase2_batch3_event_sources.csv"
BATCH3_EVIDENCES = DRAFTS / "phase2_batch3_event_evidences.csv"

ORIGIN_PREFIX = "FE-EVP3-"
ADJUDICATION_RESOLVED = "resolved_by_event_correction"
ADJUDICATION_COLUMN = "adjudication_status"
# SRC-EVP3-010 与生产 SRC-1163 为同一底层页面（纪录小康工程·广东数据库），合并时复用。
REUSED_SOURCES = {"SRC-EVP3-010": "SRC-1163"}
MERGE_NOTE = (
    "2026-08-21 第三批审核决策合并转正（授权待追认，见 phase2_batch3_event_review_decisions.md 第7节）"
)
REMAP_NOTE = "主体原指被合并删除的重复事件，已改指保留条目。"

# 物理合并：被删事件 → 保留事件。
EVENT_MERGES = {"EVT-00007": "EVT-00006", "EVT-00119": "EVT-00120"}

# 16 项签核裁决的字段落地。canonical_event_key 显式给出全值，避免按位拆解出错。
EVENT_DECISIONS: dict[str, dict[str, str]] = {
    "EVT-00019": {
        "confidence": "medium",
        "needs_manual_review": "no",
        "correction_reason_append": "第三批裁决：A级日记证据逐字命中，通过转正。",
    },
    "EVT-00016": {
        "confidence": "medium",
        "needs_manual_review": "no",
        "display_note_append": "证据方向为收信并转寄，地点系通信媒介而非见面地。",
    },
    "EVT-00018": {
        "confidence": "medium",
        "needs_manual_review": "no",
        "correction_reason_append": "第三批裁决：A级日记证据逐字命中，通过转正。",
    },
    "EVT-00256": {
        "confidence": "medium",
        "correction_reason_append": "第三批裁决：两条记录同引一篇《鲁迅日记》，按单一史料族计，不计独立双源。",
    },
    "EVT-00261": {
        "confidence": "high",
        "correction_reason_append": "第三批裁决：中国作家网与中国现代文学馆双独立源一致，置信度升为high。",
    },
    "EVT-00262": {
        "confidence": "medium",
        "display_note_append": "单一学术来源（《新文学史料》网络版），待补青岛本地报刊第二来源。",
    },
    "EVT-00257": {
        "event_name": "东方旅社秘密会议",
        "canonical_event_key": "EVT-00257|东方旅社秘密会议|1931-01-17",
        "confidence": "medium",
        "display_note_append": "来源原始表述为“党内的秘密会议”，“秘密会议”为正式名称依据；参会者以左联作家为主。",
    },
    "EVT-00188": {
        "event_name": "丁玲初次参加左联会议",
        "canonical_event_key": "EVT-00188|丁玲初次参加左联会议|1931-05",
        "event_date": "1931-05",
        "date_precision": "月",
        "historical_location": "北四川路某小学（虹口）",
        "confidence": "medium",
        "display_note_append": "据丁玲回忆（虹口区政府转载），“秘密”系左联组织性质引申、非来源原词；日期精度为月。",
    },
    "EVT-00020": {
        "event_date": "1929-07-06",
        "confidence": "medium",
        "needs_manual_review": "no",
        "correction_reason_append": "第三批裁决：7月7日日记无内山书店记载，按相邻7月6日购书记录改期（纸本日记终核前暂定）。",
    },
    "EVT-00006": {
        "confidence": "medium",
        "needs_manual_review": "no",
        "correction_reason_append": (
            "第三批裁决：与EVT-00007重复条目合并（物理删除）；日记口径为3月30日柔石等人到寓会面，"
            "3月28日北四川路看屋未成、31日海宁路看屋，均非本条日期。"
        ),
    },
    "EVT-00258": {
        "event_date": "1931-09-20",
        "canonical_event_key": "EVT-00258|《北斗》创刊|1931-09-20",
        "confidence": "high",
        "correction_reason_append": "第三批裁决：湖南日报与中国现代文学馆双独立源一致作9月20日，创刊日由9月1日修正。",
    },
    "EVT-00236": {
        "event_date": "1925-10-11",
        "date_precision": "日",
        "canonical_event_key": "EVT-00236|《生活》周刊创刊|1925-10-11",
        "confidence": "medium",
        "correction_reason_append": "第三批裁决：韬奋纪念馆官方数据载1925年10月11日创刊，1929年由误记修正；1929年12月仅为第5卷扩版节点。",
    },
    "EVT-00120": {
        "event_name": "革命文学论争",
        "canonical_event_key": "EVT-00120|革命文学论争|1928",
        "event_date": "1928",
        "confidence": "medium",
        "correction_reason_append": (
            "第三批裁决：与EVT-00119重复条目合并（物理删除）；论争爆发年份通说为1928年，"
            "创造社1928年1月发难写入本条；“闸北宝山路”地点依据不足弃用。"
        ),
    },
    "EVT-00187": {
        "event_name": "文委洛阳书店秘密机关全体党员大会",
        "canonical_event_key": "EVT-00187|文委洛阳书店秘密机关全体党员大会|1931-01-16",
        "event_date": "1931-01-16",
        "date_precision": "日",
        "historical_location": "静安寺路洛阳书店（今南京西路）",
        "place_id": "",
        "longitude": "",
        "latitude": "",
        "current_address": "",
        "confidence": "medium",
        "correction_reason_append": (
            "第三批裁决：“内山书店”系对洛阳书店会议的地点误记，按左联纪念馆馆员文章改写；"
            "人民网版本作“左联全体共产党员大会”，措辞差异待纪念馆正式出版物终核；坐标待补。"
        ),
    },
}


def read_csv(path: Path) -> tuple[list[str], list[dict[str, str]]]:
    with open(path, encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        return list(reader.fieldnames or []), list(reader)


def write_csv(path: Path, fieldnames: list[str], rows: list[dict[str, str]]) -> None:
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def formal_evidence_id(row: dict[str, str], taken: set[str]) -> str:
    basis = "|".join(
        row[key] for key in ("subject_type", "subject_id", "predicate", "object_value", "source_id")
    )
    digest = hashlib.md5(basis.encode("utf-8")).hexdigest()[:10].upper()
    candidate = f"FE-EVI-{digest}"
    while candidate in taken:
        digest = hashlib.md5((candidate + "x").encode("utf-8")).hexdigest()[:10].upper()
        candidate = f"FE-EVI-{digest}"
    return candidate


def main() -> None:
    src_fields, src_rows = read_csv(DATA / "sources.csv")
    ev_fields, ev_rows = read_csv(DATA / "fact_evidences.csv")
    evt_fields, evt_rows = read_csv(DATA / "events.csv")
    part_fields, part_rows = read_csv(DATA / "event_participants.csv")

    _, draft_sources = read_csv(BATCH3_SOURCES)
    _, draft_evidences = read_csv(BATCH3_EVIDENCES)

    origin_index = {r["origin_evidence_id"]: r for r in ev_rows if r["origin_evidence_id"]}
    already_merged = [r for r in draft_evidences if r["evidence_id"] in origin_index]
    pending = [r for r in draft_evidences if r["evidence_id"] not in origin_index]

    if not pending:
        print(f"无新增：{len(already_merged)} 条第三批证据均已转正，跳过写入。")
        return

    # 来源复用映射：固定复用项 → 已转正证据回推 → 生产表 URL 兜底。
    reuse_map: dict[str, str] = {k: v for k, v in REUSED_SOURCES.items()}
    origin_lookup = {r["origin_evidence_id"]: r for r in ev_rows if r["origin_evidence_id"]}
    for row in already_merged:
        merged_row = origin_lookup.get(row["evidence_id"])
        if merged_row is not None and row["source_id"] not in reuse_map:
            reuse_map[row["source_id"]] = merged_row["source_id"]

    url_index = {r["source_url"]: r["source_id"] for r in src_rows if r["source_url"]}
    existing_src_ids = {r["source_id"] for r in src_rows}

    def allocate_source_id() -> str:
        numbers = [int(s.split("-")[1]) for s in existing_src_ids if s.startswith("SRC-")]
        formal_id = f"SRC-{max(numbers) + 1:04d}"
        while formal_id in existing_src_ids:
            numbers.append(max(numbers) + 1)
            formal_id = f"SRC-{max(numbers) + 1:04d}"
        existing_src_ids.add(formal_id)
        return formal_id

    def resolve_source_id(draft_source_id: str, source_url: str) -> str:
        if draft_source_id not in reuse_map:
            reuse_map[draft_source_id] = (
                url_index.get(source_url) if source_url else None
            ) or allocate_source_id()
            if source_url:
                url_index[source_url] = reuse_map[draft_source_id]
        return reuse_map[draft_source_id]

    # 1. 注册尚未入库的候选来源（复用项跳过；URL 已存在则复用正式 ID）。
    registered_sources = 0
    for row in draft_sources:
        if row["source_id"] in reuse_map:
            continue
        if row["source_url"] and row["source_url"] in url_index:
            reuse_map[row["source_id"]] = url_index[row["source_url"]]
            continue
        row["source_id"] = resolve_source_id(row["source_id"], row["source_url"])
        row["needs_manual_review"] = "no"
        if MERGE_NOTE not in row["review_note"]:
            row["review_note"] = f"{row['review_note']}{MERGE_NOTE}。".replace("。。", "。")
        src_rows.append(row)
        registered_sources += 1

    # 2. 合并待转正证据（先解析正式来源 ID，再生成正式证据 ID）。
    taken_ids = {r["evidence_id"] for r in ev_rows}
    event_source_updates: dict[str, set[str]] = {}
    merged_count = 0
    for row in pending:
        row["source_id"] = resolve_source_id(row["source_id"], "")
        original_id = row["evidence_id"]
        if row["subject_id"] in EVENT_MERGES:
            kept = EVENT_MERGES[row["subject_id"]]
            row["subject_id"] = kept
            if REMAP_NOTE not in row["reviewer_note"]:
                row["reviewer_note"] = f"{MERGE_NOTE}（原编号{original_id}）。{REMAP_NOTE}{row['reviewer_note']}"
        else:
            row["reviewer_note"] = f"{MERGE_NOTE}（原编号{original_id}）。{row['reviewer_note']}"
        row["evidence_id"] = formal_evidence_id(row, taken_ids)
        taken_ids.add(row["evidence_id"])
        row["origin_evidence_id"] = original_id
        row["review_status"] = "reviewed"
        # 结构化裁决语义：conflict 合并即“已由事件级裁决落地”，其余行为空值。
        # 只依据 evidence_support 结构判定，不解析 reviewer_note。
        if ADJUDICATION_COLUMN not in row:
            row[ADJUDICATION_COLUMN] = (
                ADJUDICATION_RESOLVED if row["evidence_support"] == "conflict" else ""
            )
        ev_rows.append(row)
        merged_count += 1
        event_source_updates.setdefault(row["subject_id"], set()).add(row["source_id"])

    # 3. 事件表更新：来源挂接 + 签核裁决落地（追加均带标记检查）。
    for row in evt_rows:
        eid = row["event_id"]
        extra_sources = event_source_updates.get(eid)
        if extra_sources:
            current = [s for s in row["source_ids"].split(";") if s]
            row["source_ids"] = ";".join(current + sorted(s for s in extra_sources if s not in current))
        rule = EVENT_DECISIONS.get(eid)
        if not rule:
            continue
        for key in (
            "event_name",
            "canonical_event_key",
            "event_date",
            "date_precision",
            "historical_location",
            "place_id",
            "longitude",
            "latitude",
            "current_address",
            "confidence",
            "needs_manual_review",
        ):
            # 终值赋值（含清空为 "" 的字段），仅 display_note/correction_reason 走追加逻辑。
            if key in rule:
                row[key] = rule[key]
        if rule.get("display_note_append") and rule["display_note_append"] not in row["display_note"]:
            row["display_note"] = f"{row['display_note']}{rule['display_note_append']}"
        if rule.get("correction_reason_append") and rule["correction_reason_append"] not in row["correction_reason"]:
            row["correction_reason"] = f"{row['correction_reason']}{rule['correction_reason_append']}"

    # 4. 物理删除被合并事件及其参与者行（存在性守卫保证幂等）。
    removed_events = {eid for eid in EVENT_MERGES if any(r["event_id"] == eid for r in evt_rows)}
    evt_rows = [r for r in evt_rows if r["event_id"] not in EVENT_MERGES]
    removed_parts = sum(1 for r in part_rows if r["event_id"] in removed_events)
    part_rows = [r for r in part_rows if r["event_id"] not in removed_events]

    if ADJUDICATION_COLUMN not in ev_fields:
        ev_fields = list(ev_fields) + [ADJUDICATION_COLUMN]

    write_csv(DATA / "sources.csv", src_fields, src_rows)
    write_csv(DATA / "fact_evidences.csv", ev_fields, ev_rows)
    write_csv(DATA / "events.csv", evt_fields, evt_rows)
    write_csv(DATA / "event_participants.csv", part_fields, part_rows)

    print(f"sources: +{registered_sources} -> {len(src_rows)}")
    print(f"fact_evidences: +{merged_count} -> {len(ev_rows)}")
    print(f"events: -{len(removed_events)} -> {len(evt_rows)}（删除 {'、'.join(sorted(removed_events)) or '无'}）")
    print(f"event_participants: -{removed_parts} -> {len(part_rows)}")


if __name__ == "__main__":
    main()
