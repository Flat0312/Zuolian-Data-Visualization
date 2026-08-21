from __future__ import annotations

import csv
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path
from typing import Any

import pandas as pd

try:
    from research.analysis.rebuild_org_memberships import rebuild_org_memberships
except ModuleNotFoundError:
    from rebuild_org_memberships import rebuild_org_memberships


PROJECT_ROOT = Path(__file__).resolve().parents[2]
RESEARCH_DIR = PROJECT_ROOT / "research"
RAW_EXCEL_DIR = RESEARCH_DIR / "raw_excel"
RAW_TEXT_DIR = RESEARCH_DIR / "raw_texts"
CLEANED_DIR = RESEARCH_DIR / "intermediate" / "cleaned_data"
KB_DIR = PROJECT_ROOT / "data" / "processed"
LOG_DIR = RESEARCH_DIR / "logs"
REPORT_DIR = RESEARCH_DIR / "drafts" / "reports"
ARCHIVE_DIR = RESEARCH_DIR / "archive" / "legacy" / "old_outputs_旧输出"
APP_DIR = PROJECT_ROOT / "app" / "frontend"

RAW_WORKBOOK = RAW_EXCEL_DIR / "《左联相关档案资源目录》.xlsx"
CLEAN_WORKBOOK = CLEANED_DIR / "《左联相关档案资源目录》.xlsx"
CORRECTED_WORKBOOK = CLEANED_DIR / "《左联相关档案资源目录》_修正版.xlsx"
CORRECTION_LOG_XLSX = CLEANED_DIR / "《左联相关档案资源目录》_修改日志.xlsx"
REVIEW_NEEDED_CSV = CLEANED_DIR / "review_needed.csv"
WORKBOOK_CORRECTED = CLEANED_DIR / "workbook_corrected.xlsx"
CORRECTION_LOG_CSV = LOG_DIR / "correction_log.csv"
REVIEW_QUEUE_CSV = LOG_DIR / "review_queue.csv"
VALIDATION_REPORT = REPORT_DIR / "validation_report.md"
REORGANIZATION_REPORT = REPORT_DIR / "reorganization_report.md"
PIPELINE_LOG = LOG_DIR / "standard_pipeline.log"

CLEAN_SCRIPT = RESEARCH_DIR / "analysis" / "clean_zolian_excel.py"

RELATED_MEMBER_ROLES = {"外围联络人", "相关人士"}

LOCAL_SOURCE_PATHS = {
    "鲁迅日记": RAW_TEXT_DIR / "日记全编：全2册 (鲁迅 著) (Z-Library).txt",
    "左联词典": RAW_TEXT_DIR / "左联词典.txt",
    "左联史": RAW_TEXT_DIR / "左联史.txt",
}

WEB_TITLE_OVERRIDES = {
    "https://www.shhk.gov.cn/xwzx/002008/002008040/20221031/bd8cb3ee-198a-431a-adf7-781e9fc5185d.html": "左联会址纪念馆与成立大会旧址",
    "https://www.shhk.gov.cn/xwzx/002003/20250303/ec139c7d-8fa3-4970-a5dd-3248468989c8.html": "左联成立九十五周年主题活动",
    "https://www.shhk.gov.cn/slh/038001/20260302/5f983b4c-74f7-4a6a-a2b4-930dedf99970.html": "左联五烈士专题纪念",
    "https://www.shhk.gov.cn/xwzx/002006/20210722/07f62353-471f-40d9-880e-ea82be5da936.html": "左联会址纪念馆五烈士介绍",
    "https://cpc.people.com.cn/n1/2022/1209/c443712-32583679.html": "龙华二十四烈士资料",
    "https://www.shhk.gov.cn/xwzx/002006/20210930/96ecb0ec-79e3-49ef-a097-a89c5a5dbc40.html": "内山书店旧址说明",
    "https://www.shhk.gov.cn/xwzx/002008/002008040/20240425/3e9546e4-0e0f-409d-a0e2-91bc115f8f66.html": "内山书店今址活动页",
    "https://al3tai.nenzhu.com/news-id-2373.html": "鲁迅与柔石看屋转录材料",
    "https://m.thepaper.cn/newsDetail_forward_28996791": "澎湃新闻：成为“丁玲”之前，和上海的三次际会",
    "https://museum.shu.edu.cn/info/1034/1373.htm": "上海大学校史馆：历史上的上海大学（1923年）",
    "https://cpc.people.com.cn/BIG5/n1/2024/1006/c443712-40333498.html": "人民网党史频道：上海大学与瞿秋白",
    "https://www.chinawriter.com.cn/n1/2020/0813/c404063-31820381.html": "中国作家网：鲁迅帮助狱中的楼适夷",
}


def text(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_date(value: Any) -> str:
    if pd.isna(value) or value == "":
        return ""
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")
    raw = str(value).strip()
    if raw.endswith(" 00:00:00"):
        raw = raw[:10]
    return raw


def initial_membership_decision_for_role(role: str) -> tuple[str, str, str]:
    if role in RELATED_MEMBER_ROLES:
        return "related_person", "medium", "yes"
    return "candidate", "low", "yes"


def infer_date_precision(value: str) -> str:
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", value):
        return "日"
    if re.fullmatch(r"\d{4}-\d{2}", value):
        return "月"
    if re.fullmatch(r"\d{4}", value):
        return "年"
    return ""


def parse_birth_death(value: str) -> tuple[str, str]:
    matched = re.match(r"^\s*(\d{4}|\?)\s*-\s*(\d{4}|\?)\s*$", value or "")
    if not matched:
        return "", ""
    birth = matched.group(1) if matched.group(1).isdigit() else ""
    death = matched.group(2) if matched.group(2).isdigit() else ""
    return birth, death


def parse_coord_xy(value: str) -> tuple[str, str]:
    if not value:
        return "", ""
    matched = re.search(r"(-?\d+(?:\.\d+)?)\s*,\s*(-?\d+(?:\.\d+)?)", value)
    if not matched:
        return "", ""
    return matched.group(1), matched.group(2)


def unique_nonempty(values: list[str]) -> list[str]:
    seen: set[str] = set()
    ordered: list[str] = []
    for value in values:
        item = text(value)
        if not item or item in seen:
            continue
        seen.add(item)
        ordered.append(item)
    return ordered


def split_multi_value(value: str) -> list[str]:
    if not value:
        return []
    pieces = re.split(r"[;\n]+", value)
    return unique_nonempty(pieces)


def join_values(values: list[str], sep: str = ";") -> str:
    return sep.join(unique_nonempty(values))


def is_generic_activity_event_name(event_name: str) -> bool:
    if not event_name:
        return False
    if any(token in event_name for token in ["成立大会", "遇难", "被捕", "秘密会议", "会面", "论战", "集会"]):
        return False
    return any(
        token in event_name
        for token in ["文学活动", "交往活动", "社交活动", "社会活动", "一般活动", "交流活动", "文学交流", "上海活动", "活动"]
    )


def build_event_key_seed(
    *,
    event_name: str,
    entity_id: str,
    event_scope: str,
    event_date: str,
    explicit_key: str,
) -> str:
    if explicit_key:
        return explicit_key
    if event_scope == "entity" and entity_id:
        if is_generic_activity_event_name(event_name):
            year = event_date[:4] if event_date else "unknown"
            return f"{event_name}|{entity_id}|{year}"
        return f"{event_name}|{entity_id}"
    return event_name


def read_csv_if_exists(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()
    return pd.read_csv(path).fillna("")


def is_legacy_events_file(path: Path) -> bool:
    if not path.exists():
        return False
    header = pd.read_csv(path, nrows=0)
    return {"Entity_ID", "Timestamp", "Hist_Loc", "Current_Loc", "Event"}.issubset(header.columns)


class SourceCatalog:
    def __init__(self) -> None:
        self.rows: list[dict[str, str]] = []
        self.lookup: dict[tuple[str, str, str, str, str], str] = {}

    def register(
        self,
        *,
        source_kind: str,
        title: str,
        citation: str = "",
        source_path: str = "",
        source_url: str = "",
        evidence_layer: str,
        availability: str,
    ) -> str:
        key = (source_kind, title, citation, source_path, source_url)
        existing = self.lookup.get(key)
        if existing:
            return existing

        source_id = f"SRC-{len(self.rows) + 1:04d}"
        self.lookup[key] = source_id
        self.rows.append(
            {
                "source_id": source_id,
                "source_kind": source_kind,
                "title": title,
                "citation": citation,
                "source_path": source_path,
                "source_url": source_url,
                "evidence_layer": evidence_layer,
                "availability": availability,
            }
        )
        return source_id

    def register_raw_workbook(self) -> str:
        return self.register(
            source_kind="raw_workbook",
            title="《左联相关档案资源目录》原始表格",
            source_path=str(RAW_WORKBOOK),
            evidence_layer="excel_candidate_fact",
            availability="local",
        )

    def register_citation(self, citation: str) -> str:
        citation = text(citation)
        if not citation:
            return ""

        for prefix, path in LOCAL_SOURCE_PATHS.items():
            if citation.startswith(prefix):
                return self.register(
                    source_kind="local_text_citation",
                    title=prefix,
                    citation=citation,
                    source_path=str(path),
                    evidence_layer="txt_local_evidence",
                    availability="local",
                )

        title = "左联回忆录" if citation.startswith("左联回忆录") else "未分类引文"
        return self.register(
            source_kind="citation_only",
            title=title,
            citation=citation,
            evidence_layer="citation_reference",
            availability="citation_only",
        )

    def register_url(self, url: str) -> str:
        url = text(url)
        if not url:
            return ""
        title = WEB_TITLE_OVERRIDES.get(url, url)
        return self.register(
            source_kind="web_url",
            title=title,
            source_url=url,
            evidence_layer="web_crosscheck",
            availability="web",
        )

    def attach_sources(self, evidence_ref: str = "", source_url: str = "", fallback_source_id: str = "") -> str:
        source_ids: list[str] = []
        for citation in split_multi_value(evidence_ref):
            source_id = self.register_citation(citation)
            if source_id:
                source_ids.append(source_id)
        for url in split_multi_value(source_url):
            source_id = self.register_url(url)
            if source_id:
                source_ids.append(source_id)
        if not source_ids and fallback_source_id:
            source_ids.append(fallback_source_id)
        return join_values(source_ids)


def ensure_output_dirs() -> None:
    for path in [CLEANED_DIR, KB_DIR, LOG_DIR, REPORT_DIR, ARCHIVE_DIR]:
        path.mkdir(parents=True, exist_ok=True)


def log_message(message: str) -> None:
    PIPELINE_LOG.parent.mkdir(parents=True, exist_ok=True)
    with PIPELINE_LOG.open("a", encoding="utf-8") as handle:
        handle.write(message + "\n")
    print(message)


def run_excel_cleaning() -> None:
    if not RAW_WORKBOOK.exists():
        raise FileNotFoundError(f"缺少原始 Excel：{RAW_WORKBOOK}")
    shutil.copy2(RAW_WORKBOOK, CLEAN_WORKBOOK)
    log_message(f"[clean] sync raw workbook -> {CLEAN_WORKBOOK}")
    result = subprocess.run(
        [sys.executable, str(CLEAN_SCRIPT)],
        cwd=CLEANED_DIR,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    if result.stdout:
        for line in result.stdout.splitlines():
            log_message(f"[clean] {line}")
    if result.stderr:
        for line in result.stderr.splitlines():
            log_message(f"[clean:stderr] {line}")
    if result.returncode != 0:
        raise RuntimeError(f"clean_zolian_excel.py failed: exit={result.returncode}")
    if not CORRECTED_WORKBOOK.exists():
        raise FileNotFoundError(f"缺少清洗结果：{CORRECTED_WORKBOOK}")
    shutil.copy2(CORRECTED_WORKBOOK, WORKBOOK_CORRECTED)
    log_message(f"[clean] workbook corrected -> {WORKBOOK_CORRECTED}")


def load_source_workbook() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    workbook = pd.ExcelFile(CORRECTED_WORKBOOK)
    persons = pd.read_excel(CORRECTED_WORKBOOK, sheet_name="Sheet1").fillna("")
    relations_sheet = "Sheet2_corrected" if "Sheet2_corrected" in workbook.sheet_names else "Sheet2"
    events_sheet = "Sheet3_corrected" if "Sheet3_corrected" in workbook.sheet_names else "Sheet3"
    relations = pd.read_excel(CORRECTED_WORKBOOK, sheet_name=relations_sheet).fillna("")
    events = pd.read_excel(CORRECTED_WORKBOOK, sheet_name=events_sheet).fillna("")
    return persons, relations, events


def load_legacy_inputs() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    legacy_nodes = read_csv_if_exists(KB_DIR / "nodes.csv")
    legacy_edges = read_csv_if_exists(KB_DIR / "edges_audited.csv")
    if legacy_edges.empty:
        legacy_edges = read_csv_if_exists(KB_DIR / "edges.csv")
    legacy_events = read_csv_if_exists(KB_DIR / "events.csv") if is_legacy_events_file(KB_DIR / "events.csv") else pd.DataFrame()
    return legacy_nodes, legacy_edges, legacy_events


def build_legacy_event_coord_index(legacy_events: pd.DataFrame) -> tuple[dict[tuple[str, str, str, str], tuple[str, str]], dict[tuple[str, str, str], tuple[str, str]]]:
    exact: dict[tuple[str, str, str, str], tuple[str, str]] = {}
    loose: dict[tuple[str, str, str], tuple[str, str]] = {}
    if legacy_events.empty:
        return exact, loose

    for _, row in legacy_events.iterrows():
        key_exact = (
            text(row.get("Event")),
            normalize_date(row.get("Timestamp")),
            text(row.get("Hist_Loc")),
            text(row.get("Current_Loc")),
        )
        key_loose = (text(row.get("Event")), text(row.get("Hist_Loc")), text(row.get("Current_Loc")))
        lon = text(row.get("Longitude"))
        lat = text(row.get("Latitude"))
        if lon and lat:
            exact[key_exact] = (lon, lat)
            loose[key_loose] = (lon, lat)
    return exact, loose


def archive_legacy_compatibility() -> list[str]:
    archive_stamp = time.strftime("%Y%m%d_%H%M%S")
    target_dir = ARCHIVE_DIR / f"legacy_kb_compat_{archive_stamp}"
    target_dir.mkdir(parents=True, exist_ok=True)

    moved: list[str] = []
    for filename in ["nodes.csv", "edges.csv", "edges_audited.csv", "merged_events.csv"]:
        source = KB_DIR / filename
        if source.exists():
            target = target_dir / filename
            shutil.move(str(source), str(target))
            moved.append(f"{source.name} -> {target}")

    legacy_events_path = KB_DIR / "events.csv"
    if is_legacy_events_file(legacy_events_path):
        target = target_dir / legacy_events_path.name
        shutil.move(str(legacy_events_path), str(target))
        moved.append(f"{legacy_events_path.name} -> {target}")

    if not moved:
        shutil.rmtree(target_dir, ignore_errors=True)
    return moved


def place_type_for(name: str, current_name: str) -> str:
    probe = f"{name} {current_name}"
    if "书店" in probe:
        return "bookstore"
    if "大学" in probe:
        return "campus"
    if any(token in probe for token in ["陵园", "刑场", "纪念馆"]):
        return "memorial_site"
    if any(token in probe for token in ["路", "弄", "号", "街"]):
        return "street_site"
    if probe in {"上海", "北京", "广州", "杭州"}:
        return "city"
    return "historical_place"


def build_standard_tables(
    persons_sheet: pd.DataFrame,
    relations_sheet: pd.DataFrame,
    events_sheet: pd.DataFrame,
    legacy_nodes: pd.DataFrame,
    legacy_events: pd.DataFrame,
    catalog: SourceCatalog,
    raw_workbook_source_id: str,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, list[dict[str, str]]]:
    legacy_alias_map = legacy_nodes.set_index("Id").to_dict("index") if not legacy_nodes.empty and "Id" in legacy_nodes.columns else {}
    legacy_exact_coords, legacy_loose_coords = build_legacy_event_coord_index(legacy_events)

    persons_rows: list[dict[str, str]] = []
    person_name_map: dict[str, str] = {}
    for _, row in persons_sheet.iterrows():
        person_id = text(row.get("Entity_ID"))
        if not person_id:
            continue
        standard_name = text(row.get("True_Name"))
        birth_death = text(row.get("Birth_Death"))
        birth_year, death_year = parse_birth_death(birth_death)
        legacy_info = legacy_alias_map.get(person_id, {})
        aliases = text(row.get("Alias")) or text(legacy_info.get("Alias", ""))
        reliability = text(row.get("Reliability")) or text(legacy_info.get("Reliability", ""))
        role = text(row.get("Role"))
        persons_rows.append(
            {
                "person_id": person_id,
                "standard_name": standard_name,
                "aliases": aliases,
                "birth_year": birth_year,
                "death_year": death_year,
                "birth_death": birth_death,
                "role": role,
                "reliability": reliability or "0",
                "source_ids": raw_workbook_source_id,
            }
        )
        if standard_name:
            person_name_map[standard_name] = person_id

    persons_df = pd.DataFrame(persons_rows).sort_values("person_id").reset_index(drop=True)
    person_id_name_map = dict(zip(persons_df["person_id"], persons_df["standard_name"]))

    relation_rows: list[dict[str, str]] = []
    for _, row in relations_sheet.iterrows():
        source_person_id = text(row.get("Source_ID"))
        target_person_id = text(row.get("Target_ID"))
        if not source_person_id or not target_person_id:
            continue
        relation_id = f"REL-{len(relation_rows) + 1:05d}"
        original_relation_type = text(row.get("original_relation_type")) or text(row.get("Relation_Type"))
        standard_relation_type = text(row.get("corrected_relation_type")) or original_relation_type
        evidence_ref = text(row.get("Evidence_Ref"))
        source_url = text(row.get("source_url"))
        relation_rows.append(
            {
                "relation_id": relation_id,
                "source_person_id": source_person_id,
                "target_person_id": target_person_id,
                "original_relation_type": original_relation_type,
                "standard_relation_type": standard_relation_type,
                "raw_relation_type": standard_relation_type,
                "llm_suggested_relation_type": "",
                "final_relation_type": standard_relation_type,
                "llm_reason": "",
                "llm_confidence": "",
                "display_status": "formal",
                "relation_quality_score": text(row.get("relation_quality_score")),
                "relation_risk_level": text(row.get("relation_risk_level")),
                "context": text(row.get("Context")),
                "evidence_ref": evidence_ref,
                "weight": text(row.get("Weight")) or "0",
                "source_ids": catalog.attach_sources(evidence_ref=evidence_ref, source_url=source_url, fallback_source_id=raw_workbook_source_id),
                "correction_reason": text(row.get("correction_reason")),
                "confidence": text(row.get("confidence")) or "medium",
                "needs_manual_review": text(row.get("needs_manual_review")) or "no",
            }
        )

    person_relations_df = pd.DataFrame(relation_rows)

    organizations_df = pd.DataFrame(
        [
            {
                "organization_id": "ORG-001",
                "standard_name": "中国左翼作家联盟",
                "aliases": "左联",
                "org_type": "文学团体",
                "start_date": "1930-03-02",
                "end_date": "",
                "source_ids": join_values(
                    [
                        raw_workbook_source_id,
                        catalog.register_url("https://www.shhk.gov.cn/xwzx/002008/002008040/20221031/bd8cb3ee-198a-431a-adf7-781e9fc5185d.html"),
                        catalog.register_url("https://www.shhk.gov.cn/xwzx/002003/20250303/ec139c7d-8fa3-4970-a5dd-3248468989c8.html"),
                    ]
                ),
            },
            {
                "organization_id": "ORG-002",
                "standard_name": "中华艺术大学",
                "aliases": "",
                "org_type": "教育机构",
                "start_date": "",
                "end_date": "",
                "source_ids": catalog.register_url("https://www.shhk.gov.cn/xwzx/002008/002008040/20221031/bd8cb3ee-198a-431a-adf7-781e9fc5185d.html"),
            },
            {
                "organization_id": "ORG-003",
                "standard_name": "内山书店",
                "aliases": "",
                "org_type": "书店",
                "start_date": "",
                "end_date": "",
                "source_ids": join_values(
                    [
                        raw_workbook_source_id,
                        catalog.register_url("https://www.shhk.gov.cn/xwzx/002006/20210930/96ecb0ec-79e3-49ef-a097-a89c5a5dbc40.html"),
                        catalog.register_url("https://www.shhk.gov.cn/xwzx/002008/002008040/20240425/3e9546e4-0e0f-409d-a0e2-91bc115f8f66.html"),
                    ]
                ),
            },
        ]
    )

    membership_rows: list[dict[str, str]] = []
    review_additions: list[dict[str, str]] = []
    for _, row in persons_df.iterrows():
        role = text(row["role"])
        if not role:
            continue
        membership_id = f"MEM-{len(membership_rows) + 1:05d}"
        membership_type, confidence, needs_manual_review = initial_membership_decision_for_role(role)
        review_additions.append(
            {
                "queue_id": "",
                "source_sheet": "persons",
                "record_type": "org_membership",
                "record_key": f"ORG-001|{row['person_id']}",
                "issue_summary": f"角色“{role}”不能自动证明正式成员身份，已按证据待核关系保留。",
                "confidence": confidence,
                "source_ids": raw_workbook_source_id,
                "source_url": "",
                "evidence_ref_used": "",
                "suggested_action": "使用组织成员证据台账确认身份。",
            }
        )
        membership_rows.append(
            {
                "membership_id": membership_id,
                "organization_id": "ORG-001",
                "person_id": row["person_id"],
                "membership_role": role,
                "membership_type": membership_type,
                "source_ids": raw_workbook_source_id,
                "confidence": confidence,
                "needs_manual_review": needs_manual_review,
            }
        )
    org_memberships_df = pd.DataFrame(membership_rows)

    place_lookup: dict[tuple[str, str, str, str], str] = {}
    place_rows: list[dict[str, str]] = []
    event_lookup: dict[tuple[str, str, str, str], str] = {}
    event_rows: list[dict[str, str]] = []
    event_participant_lookup: dict[tuple[str, str], dict[str, str]] = {}

    def ensure_place(historical_location: str, current_address: str, longitude: str, latitude: str, source_ids: str, confidence: str) -> str:
        key = (historical_location, current_address, longitude, latitude)
        existing = place_lookup.get(key)
        if existing:
            return existing
        place_id = f"PLC-{len(place_rows) + 1:05d}"
        place_lookup[key] = place_id
        place_rows.append(
            {
                "place_id": place_id,
                "place_name": historical_location or current_address,
                "historical_name": historical_location,
                "current_name": current_address,
                "place_type": place_type_for(historical_location, current_address),
                "longitude": longitude,
                "latitude": latitude,
                "source_ids": source_ids,
                "confidence": confidence,
            }
        )
        return place_id

    for _, row in events_sheet.iterrows():
        original_event_name = text(row.get("Event"))
        event_name = text(row.get("standard_event_name")) or original_event_name
        entity_role = text(row.get("entity_role_in_event"))
        if entity_role == "冲突":
            continue
        event_date = normalize_date(row.get("corrected_date")) or normalize_date(row.get("original_date")) or normalize_date(row.get("Timestamp"))
        historical_location = text(row.get("historical_location")) or text(row.get("Hist_Loc"))
        current_address = text(row.get("current_address")) or text(row.get("Current_Loc"))
        event_scope = text(row.get("event_scope")) or ("entity" if text(row.get("Entity_ID")) and text(row.get("Entity_ID")) in event_name else "collective")
        canonical_event_key = build_event_key_seed(
            event_name=event_name,
            entity_id=text(row.get("Entity_ID")),
            event_scope=event_scope,
            event_date=event_date,
            explicit_key=text(row.get("canonical_event_key")),
        )
        display_note = text(row.get("display_note")) or text(row.get("correction_reason"))
        source_url = text(row.get("source_url"))
        source_ids = catalog.attach_sources(source_url=source_url, fallback_source_id=raw_workbook_source_id)
        confidence = text(row.get("confidence")) or "medium"
        longitude, latitude = parse_coord_xy(text(row.get("Coord_XY")))
        if not longitude or not latitude:
            coord_exact = legacy_exact_coords.get((original_event_name, normalize_date(row.get("Timestamp")), text(row.get("Hist_Loc")), text(row.get("Current_Loc"))))
            coord_loose = legacy_loose_coords.get((original_event_name, text(row.get("Hist_Loc")), text(row.get("Current_Loc"))))
            if coord_exact:
                longitude, latitude = coord_exact
            elif coord_loose:
                longitude, latitude = coord_loose

        place_id = ensure_place(historical_location, current_address, longitude, latitude, source_ids, confidence)
        event_key = (canonical_event_key, event_date, historical_location, current_address)
        event_id = event_lookup.get(event_key)
        if not event_id:
            event_id = f"EVT-{len(event_rows) + 1:05d}"
            event_lookup[event_key] = event_id
            event_rows.append(
                {
                    "event_id": event_id,
                    "event_name": event_name,
                    "event_scope": event_scope,
                    "canonical_event_key": canonical_event_key,
                    "original_event_names": original_event_name,
                    "event_date": event_date,
                    "date_precision": text(row.get("date_precision")) or infer_date_precision(event_date),
                    "place_id": place_id,
                    "historical_location": historical_location,
                    "current_address": current_address,
                    "longitude": longitude,
                    "latitude": latitude,
                    "source_ids": source_ids,
                    "display_note": display_note,
                    "correction_reason": text(row.get("correction_reason")),
                    "confidence": confidence,
                    "needs_manual_review": text(row.get("needs_manual_review")) or "no",
                }
            )
        else:
            for event_row in event_rows:
                if event_row["event_id"] != event_id:
                    continue
                event_row["source_ids"] = join_values([event_row["source_ids"], source_ids])
                event_row["original_event_names"] = join_values([event_row["original_event_names"], original_event_name], sep=" | ")
                if event_row["needs_manual_review"] != "yes" and text(row.get("needs_manual_review")) == "yes":
                    event_row["needs_manual_review"] = "yes"
                if not event_row["correction_reason"]:
                    event_row["correction_reason"] = text(row.get("correction_reason"))
                if not event_row.get("display_note"):
                    event_row["display_note"] = display_note
                if not event_row.get("canonical_event_key"):
                    event_row["canonical_event_key"] = canonical_event_key
                break

        participant_candidates: list[tuple[str, str, str]] = []
        entity_id = text(row.get("Entity_ID"))
        entity_name = text(row.get("entity_name"))
        entity_role = entity_role or "待核"
        entity_ids = split_multi_value(entity_id)
        if not entity_ids and entity_id:
            entity_ids = [entity_id]
        for single_person_id in entity_ids:
            participant_candidates.append((single_person_id, entity_name or person_id_name_map.get(single_person_id, ""), entity_role))

        corrected_persons = split_multi_value(text(row.get("corrected_persons")).replace("；", ";").replace("、", ";"))
        for person_name in corrected_persons:
            mapped_person_id = person_name_map.get(person_name, "")
            if mapped_person_id:
                participant_candidates.append((mapped_person_id, person_name, "直接参与者"))
            else:
                review_additions.append(
                    {
                        "queue_id": "",
                        "source_sheet": "Sheet3",
                        "record_type": "event_participant",
                        "record_key": f"{event_id}|{person_name}",
                        "issue_summary": f"事件“{event_name}”的修正参与者“{person_name}”未在人物表中找到映射。",
                        "confidence": "medium",
                        "source_ids": source_ids,
                        "source_url": source_url,
                        "evidence_ref_used": "",
                        "suggested_action": "人工补充人物映射或修正参与者姓名。",
                    }
                )

        for person_id, person_name, participant_role in participant_candidates:
            if not person_id:
                continue
            participant_key = (event_id, person_id)
            existing = event_participant_lookup.get(participant_key)
            if existing:
                existing["participant_role"] = join_values([existing["participant_role"], participant_role], sep=" | ")
                existing["source_ids"] = join_values([existing["source_ids"], source_ids])
                if existing["needs_manual_review"] != "yes" and text(row.get("needs_manual_review")) == "yes":
                    existing["needs_manual_review"] = "yes"
                continue
            matched_name = person_name
            if not matched_name:
                matched_name = person_id_name_map.get(person_id, "")
            event_participant_lookup[participant_key] = {
                "event_participant_id": f"EVP-{len(event_participant_lookup) + 1:05d}",
                "event_id": event_id,
                "person_id": person_id,
                "participant_name": matched_name,
                "participant_role": participant_role,
                "source_ids": source_ids,
                "confidence": confidence,
                "needs_manual_review": text(row.get("needs_manual_review")) or "no",
            }

    places_df = pd.DataFrame(place_rows).sort_values("place_id").reset_index(drop=True)
    events_df = pd.DataFrame(event_rows).sort_values(["event_date", "event_name"], na_position="last").reset_index(drop=True)
    event_participants_df = pd.DataFrame(list(event_participant_lookup.values())).sort_values(["event_id", "person_id"]).reset_index(drop=True)

    return (
        persons_df,
        organizations_df,
        places_df,
        events_df,
        person_relations_df,
        org_memberships_df,
        event_participants_df,
        review_additions,
    )


def convert_correction_log(catalog: SourceCatalog, raw_workbook_source_id: str) -> pd.DataFrame:
    if not CORRECTION_LOG_XLSX.exists():
        return pd.DataFrame()
    log_df = pd.read_excel(CORRECTION_LOG_XLSX, sheet_name="modification_log").fillna("")
    if log_df.empty:
        return log_df
    log_df["source_ids"] = log_df.apply(
        lambda row: catalog.attach_sources(
            evidence_ref=text(row.get("evidence_ref_used")),
            source_url=text(row.get("source_url")),
            fallback_source_id=raw_workbook_source_id,
        ),
        axis=1,
    )
    log_df.insert(0, "pipeline_stage", "excel_cleaning")
    return log_df


def build_review_queue(catalog: SourceCatalog, raw_workbook_source_id: str, review_additions: list[dict[str, str]]) -> pd.DataFrame:
    review_rows: list[dict[str, str]] = []
    if REVIEW_NEEDED_CSV.exists():
        current = pd.read_csv(REVIEW_NEEDED_CSV).fillna("")
        for _, row in current.iterrows():
            sheet_name = text(row.get("sheet_name"))
            review_rows.append(
                {
                    "queue_id": "",
                    "source_sheet": sheet_name,
                    "record_type": "person_relation" if sheet_name == "Sheet2" else "event",
                    "record_key": text(row.get("primary_key")),
                    "issue_summary": text(row.get("issue_summary")),
                    "confidence": text(row.get("confidence")) or "low",
                    "source_ids": catalog.attach_sources(
                        evidence_ref=text(row.get("evidence_ref_used")),
                        source_url=text(row.get("source_url")),
                        fallback_source_id=raw_workbook_source_id,
                    ),
                    "source_url": text(row.get("source_url")),
                    "evidence_ref_used": text(row.get("evidence_ref_used")),
                    "suggested_action": "人工复核原始证据与修正规则。",
                }
            )

    review_rows.extend(review_additions)
    for index, item in enumerate(review_rows, start=1):
        item["queue_id"] = f"RVW-{index:05d}"

    if not review_rows:
        return pd.DataFrame(
            columns=[
                "queue_id",
                "source_sheet",
                "record_type",
                "record_key",
                "issue_summary",
                "confidence",
                "source_ids",
                "source_url",
                "evidence_ref_used",
                "suggested_action",
            ]
        )
    return pd.DataFrame(review_rows)


def write_csv(df: pd.DataFrame, path: Path) -> None:
    df.to_csv(path, index=False, encoding="utf-8-sig", quoting=csv.QUOTE_MINIMAL)
    log_message(f"[write] {path}")


def validate_referential_integrity(
    persons_df: pd.DataFrame,
    places_df: pd.DataFrame,
    events_df: pd.DataFrame,
    person_relations_df: pd.DataFrame,
    org_memberships_df: pd.DataFrame,
    event_participants_df: pd.DataFrame,
    sources_df: pd.DataFrame,
    review_queue_df: pd.DataFrame,
) -> dict[str, Any]:
    person_ids = set(persons_df["person_id"].astype(str))
    place_ids = set(places_df["place_id"].astype(str))
    event_ids = set(events_df["event_id"].astype(str))
    source_ids = set(sources_df["source_id"].astype(str))

    missing_relation_people = int(
        (~person_relations_df["source_person_id"].astype(str).isin(person_ids)).sum()
        + (~person_relations_df["target_person_id"].astype(str).isin(person_ids)).sum()
    )
    missing_membership_people = int((~org_memberships_df["person_id"].astype(str).isin(person_ids)).sum())
    missing_event_places = int((~events_df["place_id"].astype(str).isin(place_ids)).sum())
    missing_event_participants = int(
        (~event_participants_df["event_id"].astype(str).isin(event_ids)).sum()
        + (~event_participants_df["person_id"].astype(str).isin(person_ids)).sum()
    )

    def count_missing_source_refs(df: pd.DataFrame, column: str) -> int:
        total = 0
        for value in df[column].astype(str):
            for source_id in split_multi_value(value):
                if source_id and source_id not in source_ids:
                    total += 1
        return total

    return {
        "duplicate_person_ids": int(persons_df["person_id"].duplicated().sum()),
        "duplicate_place_ids": int(places_df["place_id"].duplicated().sum()),
        "duplicate_event_ids": int(events_df["event_id"].duplicated().sum()),
        "duplicate_relation_ids": int(person_relations_df["relation_id"].duplicated().sum()),
        "duplicate_membership_ids": int(org_memberships_df["membership_id"].duplicated().sum()),
        "duplicate_event_participant_ids": int(event_participants_df["event_participant_id"].duplicated().sum()),
        "missing_relation_people": missing_relation_people,
        "missing_membership_people": missing_membership_people,
        "missing_event_places": missing_event_places,
        "missing_event_participants": missing_event_participants,
        "missing_relation_sources": count_missing_source_refs(person_relations_df, "source_ids"),
        "missing_event_sources": count_missing_source_refs(events_df, "source_ids"),
        "missing_review_sources": count_missing_source_refs(review_queue_df, "source_ids") if not review_queue_df.empty else 0,
    }


def validate_app_runtime() -> dict[str, str]:
    command = [
        sys.executable,
        "-m",
        "streamlit",
        "run",
        "app.py",
        "--server.headless",
        "true",
        "--server.address",
        "127.0.0.1",
        "--server.port",
        "8521",
    ]
    captured: list[str] = []
    success = False
    process = subprocess.Popen(
        command,
        cwd=APP_DIR,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        bufsize=1,
    )
    try:
        deadline = time.time() + 30
        while time.time() < deadline:
            line = process.stdout.readline() if process.stdout else ""
            if line:
                captured.append(line.rstrip())
                if "You can now view your Streamlit app in your browser" in line or "Local URL:" in line:
                    success = True
                    break
            elif process.poll() is not None:
                break
            else:
                time.sleep(0.2)
    finally:
        if process.poll() is None:
            process.terminate()
            try:
                process.wait(timeout=10)
            except subprocess.TimeoutExpired:
                process.kill()
        if process.stdout:
            remainder = process.stdout.read()
            if remainder:
                captured.extend(remainder.splitlines())

    return {
        "success": "yes" if success else "no",
        "details": "\n".join(captured[-30:]),
    }


def build_output_inventory() -> list[str]:
    inventory: list[str] = []
    for base_dir in [KB_DIR, LOG_DIR, REPORT_DIR, ARCHIVE_DIR]:
        if not base_dir.exists():
            continue
        for path in sorted(base_dir.rglob("*")):
            if path.is_file():
                inventory.append(str(path.relative_to(PROJECT_ROOT)))
    return inventory


def write_reports(
    *,
    moved_files: list[str],
    validation: dict[str, Any],
    app_validation: dict[str, str],
    inventory: list[str],
    row_counts: dict[str, int],
) -> None:
    validation_lines = [
        "# validation_report",
        "",
        "## 标准表产出",
        f"- persons.csv: {row_counts['persons']}",
        f"- organizations.csv: {row_counts['organizations']}",
        f"- places.csv: {row_counts['places']}",
        f"- events.csv: {row_counts['events']}",
        f"- person_relations.csv: {row_counts['person_relations']}",
        f"- org_memberships.csv: {row_counts['org_memberships']}",
        f"- event_participants.csv: {row_counts['event_participants']}",
        f"- sources.csv: {row_counts['sources']}",
        f"- review_queue.csv: {row_counts['review_queue']}",
        f"- correction_log.csv: {row_counts['correction_log']}",
        "",
        "## 校验结果",
        f"- duplicate_person_ids: {validation['duplicate_person_ids']}",
        f"- duplicate_place_ids: {validation['duplicate_place_ids']}",
        f"- duplicate_event_ids: {validation['duplicate_event_ids']}",
        f"- duplicate_relation_ids: {validation['duplicate_relation_ids']}",
        f"- duplicate_membership_ids: {validation['duplicate_membership_ids']}",
        f"- duplicate_event_participant_ids: {validation['duplicate_event_participant_ids']}",
        f"- missing_relation_people: {validation['missing_relation_people']}",
        f"- missing_membership_people: {validation['missing_membership_people']}",
        f"- missing_event_places: {validation['missing_event_places']}",
        f"- missing_event_participants: {validation['missing_event_participants']}",
        f"- missing_relation_sources: {validation['missing_relation_sources']}",
        f"- missing_event_sources: {validation['missing_event_sources']}",
        f"- missing_review_sources: {validation['missing_review_sources']}",
        "",
        "## 高风险联网核验",
        "- 中国左翼作家联盟成立大会：1930-03-02，中华艺术大学教室（今址上海市虹口区多伦路201弄2号）。",
        "- 左联五烈士遇难：1931-02-07，上海龙华淞沪警备司令部刑场。",
        "- 内山书店秘密会议：保留为 1931 年级别，地点统一为内山书店旧址 / 四川北路2050号，具体日期继续待核。",
        "",
        "## app.py 验证",
        f"- headless_streamlit_run: {app_validation['success']}",
        "- validation_log:",
        "```text",
        app_validation["details"] or "(no output)",
        "```",
        "",
        "## 兼容层归档",
    ]
    if moved_files:
        validation_lines.extend([f"- {item}" for item in moved_files])
    else:
        validation_lines.append("- 本次运行前未检测到旧版兼容 CSV。")

    validation_lines.extend(["", "## 产物文件清单"])
    validation_lines.extend([f"- {item}" for item in inventory])
    VALIDATION_REPORT.write_text("\n".join(validation_lines), encoding="utf-8")

    report_lines = [
        "# reorganization_report",
        "",
        "## 当前工程结构分析",
        "- 原始输入已集中到 research/raw_excel 与 research/raw_texts；清洗与验证脚本集中到 research/analysis；中间结果集中到 research/intermediate；唯一生产数据目录锁定为 data/processed/。",
        "- 现有最可复用的数据资产是 research/intermediate/cleaned_data/《左联相关档案资源目录》_修正版.xlsx，其中 Sheet2_corrected / Sheet3_corrected 已包含关系纠偏、事件校核与待复核标记。",
        "- 旧版 knowledge base 兼容 CSV（nodes/edges/events）仅作为本次标准化输入参考，不再作为最终生产数据源。",
        "",
        "## 识别出的主数据流程",
        "1. research/raw_excel/《左联相关档案资源目录》.xlsx -> clean_zolian_excel.py",
        "2. clean_zolian_excel.py -> research/intermediate/cleaned_data/《左联相关档案资源目录》_修正版.xlsx",
        "3. build_standard_kb_pipeline.py -> data/processed/*.csv",
        "4. app/frontend/app.py 从标准表优先加载，并在内存中构造 UI 兼容视图。",
        "",
        "## 被复用的脚本列表",
        "- research/analysis/clean_zolian_excel.py",
        "- research/analysis/build_standard_kb_pipeline.py",
        "- app/frontend/data_paths.py",
        "- app/frontend/app.py",
        "",
        "## 被废弃/归档的脚本",
        "- 已在 research/archive/legacy/scripts_废弃脚本/ 中归档的旧脚本继续保持归档状态。",
        "- 不再纳入主 pipeline 的旧链路脚本：process_sheet2.py、expand_sheet3.py、verify_evidence.py、verify_with_llm.py。",
        "",
        "## 数据去重说明",
        "- 原有 nodes.csv / edges.csv / edges_audited.csv / merged_events.csv 以及旧 schema 的 events.csv 已从生产数据目录移出，归档到 research/archive/legacy/old_outputs_旧输出/。",
        "- 最终知识库数据只保留标准表：persons / organizations / places / events / person_relations / org_memberships / event_participants / sources。",
        "",
        "## 路径修复说明",
        "- app/frontend/data_paths.py 只解析 data/processed/ 一个目录。",
        "- app/frontend/app.py 改为优先读取标准表，不再要求旧版 nodes/edges/events 作为入口。",
        "",
        "## app.py 运行结果",
        f"- headless_streamlit_run: {app_validation['success']}",
        "- 详情见 validation_report.md。",
    ]
    REORGANIZATION_REPORT.write_text("\n".join(report_lines), encoding="utf-8")


def main() -> None:
    ensure_output_dirs()
    PIPELINE_LOG.write_text("", encoding="utf-8")
    log_message("[pipeline] start")
    run_excel_cleaning()

    persons_sheet, relations_sheet, events_sheet = load_source_workbook()
    legacy_nodes, _, legacy_events = load_legacy_inputs()
    catalog = SourceCatalog()
    raw_workbook_source_id = catalog.register_raw_workbook()

    (
        persons_df,
        organizations_df,
        places_df,
        events_df,
        person_relations_df,
        org_memberships_df,
        event_participants_df,
        review_additions,
    ) = build_standard_tables(
        persons_sheet,
        relations_sheet,
        events_sheet,
        legacy_nodes,
        legacy_events,
        catalog,
        raw_workbook_source_id,
    )

    correction_log_df = convert_correction_log(catalog, raw_workbook_source_id)
    review_queue_df = build_review_queue(catalog, raw_workbook_source_id, review_additions)

    moved_files = archive_legacy_compatibility()

    write_csv(persons_df, KB_DIR / "persons.csv")
    write_csv(organizations_df, KB_DIR / "organizations.csv")
    write_csv(places_df, KB_DIR / "places.csv")
    write_csv(events_df, KB_DIR / "events.csv")
    write_csv(person_relations_df, KB_DIR / "person_relations.csv")
    write_csv(org_memberships_df, KB_DIR / "org_memberships.csv")
    write_csv(event_participants_df, KB_DIR / "event_participants.csv")
    sources_df = pd.DataFrame(catalog.rows).sort_values("source_id").reset_index(drop=True)
    write_csv(sources_df, KB_DIR / "sources.csv")
    rebuild_org_memberships(KB_DIR, KB_DIR / "runtime_sources")
    write_csv(review_queue_df, REVIEW_QUEUE_CSV)
    write_csv(correction_log_df, CORRECTION_LOG_CSV)

    validation = validate_referential_integrity(
        persons_df,
        places_df,
        events_df,
        person_relations_df,
        org_memberships_df,
        event_participants_df,
        sources_df,
        review_queue_df,
    )
    app_validation = validate_app_runtime()
    inventory = build_output_inventory()
    row_counts = {
        "persons": len(persons_df),
        "organizations": len(organizations_df),
        "places": len(places_df),
        "events": len(events_df),
        "person_relations": len(person_relations_df),
        "org_memberships": len(org_memberships_df),
        "event_participants": len(event_participants_df),
        "sources": len(sources_df),
        "review_queue": len(review_queue_df),
        "correction_log": len(correction_log_df),
    }
    write_reports(
        moved_files=moved_files,
        validation=validation,
        app_validation=app_validation,
        inventory=inventory,
        row_counts=row_counts,
    )

    log_message("[pipeline] completed")


if __name__ == "__main__":
    main()
