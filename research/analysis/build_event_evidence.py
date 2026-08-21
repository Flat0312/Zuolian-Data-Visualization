from __future__ import annotations

import argparse
import hashlib
import json
import re
from collections import defaultdict
from pathlib import Path
from typing import Any

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[2]
KB_DIR = PROJECT_ROOT / "data" / "processed"
OUTPUT_FILE = KB_DIR / "event_evidences.json"

EVENT_TITLE_SUFFIXES = (
    "成立大会",
    "成立",
    "秘密会议",
    "会议",
    "会面",
    "遇难",
    "被捕事件",
    "被捕",
    "拜访",
    "通信",
    "文学活动",
    "社交活动",
    "交往活动",
    "一般活动",
)

GENERIC_LOCATION_TOKENS = {
    "上海",
    "中国",
    "虹口区",
    "闸北区",
    "徐汇区",
    "黄浦区",
    "静安区",
    "宝山区",
    "虹口",
    "闸北",
    "徐汇",
    "黄浦",
    "地址",
    "附近",
    "一带",
}

MATCH_RULE_NOTE = {
    "relation_exact_date": "关系行+原始引文：精确日期与人物双重命中",
    "relation_collective": "关系行+原始引文：集体事件命中多人、年份与地点",
    "direct_title_date": "原始引文：事件标题与日期直接命中",
    "direct_specific_visit": "原始引文：精确日期与地点/参与人命中",
    "direct_collective": "原始引文：集体事件命中标题/多人/地点",
}


def text(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def clean_ocr_text(value: str) -> str:
    raw = text(value).replace("\u3000", " ")
    if not raw:
        return ""
    raw = raw.replace("\r", "\n")
    raw = re.sub(r"(?<=[\u4e00-\u9fff])\s+(?=[\u4e00-\u9fff])", "", raw)
    raw = re.sub(r"(?<=[\u4e00-\u9fff])\s+(?=[，。！？；：、“”‘’（）《》])", "", raw)
    raw = re.sub(r"(?<=[，。！？；：、“”‘’（）《》])\s+(?=[\u4e00-\u9fff])", "", raw)
    raw = re.sub(r"[ \t]+", " ", raw)
    raw = re.sub(r"\n{3,}", "\n\n", raw)
    return raw.strip()


def normalize_text(value: str) -> str:
    cleaned = clean_ocr_text(value)
    cleaned = cleaned.lower()
    return re.sub(r"[\s\u3000,，。！？；：、“”‘’（）()《》〈〉【】\\[\\]·•/\\-]+", "", cleaned)


def split_ids(value: Any) -> list[str]:
    items = [item.strip() for item in text(value).split(";")]
    return [item for item in items if item]


def parse_year(value: str) -> int | None:
    matched = re.search(r"((?:18|19|20)\d{2})", text(value))
    return int(matched.group(1)) if matched else None


def parse_exact_date(value: str) -> str:
    matched = re.search(r"((?:18|19|20)\d{2})年(\d{1,2})月(\d{1,2})日", text(value))
    if not matched:
        return ""
    year, month, day = matched.groups()
    return f"{int(year):04d}-{int(month):02d}-{int(day):02d}"


def parse_unique_exact_date_from_text(value: str) -> str:
    matches = {
        f"{int(year):04d}-{int(month):02d}-{int(day):02d}"
        for year, month, day in re.findall(r"((?:18|19|20)\d{2})年(\d{1,2})月(\d{1,2})日", text(value))
    }
    if len(matches) == 1:
        return next(iter(matches))
    return ""


def chinese_numeral_to_int(value: str) -> int | None:
    digits = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}
    raw = value.strip()
    if raw == "十":
        return 10
    if raw == "廿":
        return 20
    if raw == "卅":
        return 30
    if raw.startswith("十"):
        return 10 + digits.get(raw[1], 0)
    if raw.startswith("廿"):
        return 20 + digits.get(raw[1], 0)
    if raw.startswith("卅"):
        return 30 + digits.get(raw[1], 0)
    if raw.endswith("十"):
        return digits.get(raw[0], 0) * 10
    if len(raw) == 2 and raw[0] in digits and raw[1] in digits:
        return digits[raw[0]] * 10 + digits[raw[1]]
    return digits.get(raw)


def build_diary_entries(path: Path) -> dict[str, str]:
    lines = path.read_text(encoding="utf-8").splitlines()
    entries: dict[str, str] = {}
    current_year: int | None = None
    current_month: int | None = None
    current_date: str = ""
    buffer: list[str] = []

    def flush() -> None:
        nonlocal current_date, buffer
        if current_date and buffer:
            entries[current_date] = clean_ocr_text("\n".join(buffer))
        current_date = ""
        buffer = []

    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            if buffer:
                buffer.append("")
            continue

        year_match = re.match(r"日记.+?（(\d{4})年）", line)
        if year_match:
            flush()
            current_year = int(year_match.group(1))
            current_month = None
            continue

        month_match = re.fullmatch(r"([一二三四五六七八九十廿卅]+)月", line)
        if month_match:
            flush()
            current_month = chinese_numeral_to_int(month_match.group(1))
            continue

        day_match = re.match(r"([一二三四五六七八九十廿卅]+)日[　\\s]*(.*)", line)
        if current_year and current_month and day_match:
            flush()
            day = chinese_numeral_to_int(day_match.group(1))
            if day is None:
                continue
            current_date = f"{current_year:04d}-{current_month:02d}-{day:02d}"
            buffer = [clean_ocr_text(line)]
            continue

        if current_date:
            buffer.append(clean_ocr_text(line))

    flush()
    return entries


def build_paged_text(path: Path) -> dict[int, str]:
    lines = path.read_text(encoding="utf-8").splitlines()
    pages: dict[int, list[str]] = {}
    current_page: int | None = None
    for raw_line in lines:
        line = raw_line.rstrip()
        page_match = re.match(r"第\s+(\d+)\s+页", line.strip())
        if page_match:
            current_page = int(page_match.group(1))
            pages.setdefault(current_page, [])
            continue
        if current_page is not None:
            pages[current_page].append(line)
    return {page: clean_ocr_text("\n".join(content)) for page, content in pages.items()}


def build_source_lookup(sources_df: pd.DataFrame) -> dict[str, dict[str, Any]]:
    diary_cache: dict[str, dict[str, str]] = {}
    page_cache: dict[str, dict[int, str]] = {}
    source_lookup: dict[str, dict[str, Any]] = {}

    for _, row in sources_df.iterrows():
        if text(row.get("source_kind")) != "local_text_citation":
            continue
        source_id = text(row.get("source_id"))
        source_path = Path(text(row.get("source_path")))
        citation = text(row.get("citation"))
        source_title = text(row.get("title"))
        if not source_id or not source_path.exists():
            continue

        quote = ""
        source_loc = ""
        exact_date = parse_exact_date(citation)
        year = parse_year(citation)

        if "鲁迅日记" in citation:
            cache_key = str(source_path)
            diary_entries = diary_cache.setdefault(cache_key, build_diary_entries(source_path))
            if exact_date:
                quote = diary_entries.get(exact_date, "")
                source_loc = citation.replace("鲁迅日记", "", 1).strip() or citation
        else:
            page_match = re.search(r"第(\d+)页", citation)
            if page_match:
                cache_key = str(source_path)
                page_lookup = page_cache.setdefault(cache_key, build_paged_text(source_path))
                page_num = int(page_match.group(1))
                quote = page_lookup.get(page_num, "")
                source_loc = f"第{page_num}页"

        quote = clean_ocr_text(quote)
        if not quote:
            continue
        if not exact_date:
            exact_date = parse_unique_exact_date_from_text(quote)
        if not year:
            year = parse_year(quote)

        source_lookup[source_id] = {
            "source_id": source_id,
            "source": source_title or source_path.stem,
            "source_file": source_path.relative_to(PROJECT_ROOT).as_posix(),
            "source_loc": source_loc or citation,
            "citation": citation,
            "quote": quote,
            "quote_norm": normalize_text(quote),
            "year": year or parse_year(quote),
            "exact_date": exact_date,
        }
    return source_lookup


def unique_ordered(values: list[str]) -> list[str]:
    seen: set[str] = set()
    items: list[str] = []
    for value in values:
        item = text(value)
        if not item or item in seen:
            continue
        seen.add(item)
        items.append(item)
    return items


def build_title_variants(title: str) -> list[str]:
    cleaned = text(title)
    variants = {normalize_text(cleaned)}
    for suffix in EVENT_TITLE_SUFFIXES:
        if cleaned.endswith(suffix):
            base = cleaned[: -len(suffix)]
            if len(base) >= 2 and not base.endswith(("联盟", "左联")):
                variants.add(normalize_text(base))
    variants.add(normalize_text(cleaned.replace("中国", "")))
    return sorted([item for item in variants if len(item) >= 4], key=len, reverse=True)


def strong_collective_title_hit(title: str, candidate_norm: str) -> bool:
    cleaned = text(title)
    phrases = {normalize_text(cleaned)}
    if cleaned.endswith("成立大会"):
        phrases.add(normalize_text(cleaned.replace("成立大会", "成立")))
        phrases.add(normalize_text(cleaned.replace("中国左翼作家联盟", "左联")))
        phrases.add(normalize_text("左联成立大会"))
    if cleaned.endswith("遇难"):
        phrases.add(normalize_text(cleaned.replace("遇难", "")))
    return any(phrase and phrase in candidate_norm for phrase in phrases if len(phrase) >= 4)


def build_location_tokens(*values: str) -> list[str]:
    tokens: list[str] = []
    for value in values:
        cleaned = clean_ocr_text(value)
        if not cleaned:
            continue
        tokens.append(cleaned)
        tokens.extend(re.findall(r"[\u4e00-\u9fffA-Za-z0-9·]{2,}", cleaned))
    ordered = unique_ordered(tokens)
    return [
        item
        for item in ordered
        if len(item) >= 2
        and item not in GENERIC_LOCATION_TOKENS
        and not item.endswith("号")
    ]


def build_event_meta(events_df: pd.DataFrame, participants_df: pd.DataFrame) -> dict[str, dict[str, Any]]:
    participant_lookup: dict[str, list[dict[str, str]]] = defaultdict(list)
    for _, row in participants_df.iterrows():
        event_id = text(row.get("event_id"))
        if not event_id:
            continue
        participant_lookup[event_id].append(
            {
                "person_id": text(row.get("person_id")),
                "name": text(row.get("participant_name")),
                "relation": text(row.get("participant_role")),
            }
        )

    event_meta: dict[str, dict[str, Any]] = {}
    for _, row in events_df.iterrows():
        event_id = text(row.get("event_id"))
        if not event_id:
            continue
        title = text(row.get("event_name"))
        event_date = text(row.get("event_date"))
        participants = participant_lookup.get(event_id, [])
        participant_names = unique_ordered([item["name"] for item in participants if item["name"]])
        title_variants = build_title_variants(title)
        event_meta[event_id] = {
            "event_id": event_id,
            "title": title,
            "title_variants": title_variants,
            "event_date": event_date,
            "year": parse_year(event_date),
            "participants": participants,
            "participant_names": participant_names,
            "participant_name_set": set(participant_names),
            "location_tokens": build_location_tokens(text(row.get("historical_location")), text(row.get("current_address"))),
            "is_pair_event": len(participant_names) == 2,
            "is_single_event": len(participant_names) == 1,
            "is_collective_event": len(participant_names) >= 4,
            "is_generic_activity": any(keyword in title for keyword in ("文学活动", "一般活动", "社交活动", "交往活动")),
            "is_collective_anchor": any(keyword in title for keyword in ("成立", "遇难", "被捕", "秘密会议", "会议")),
        }
    return event_meta


def choose_quote_excerpt(source_quote: str, keywords: list[str], limit: int = 220) -> str:
    body = clean_ocr_text(source_quote)
    if not body:
        return ""
    sentences = [item.strip() for item in re.split(r"(?<=[。！？；])", body) if item.strip()]
    normalized_keywords = [normalize_text(item) for item in keywords if item]
    for index, sentence in enumerate(sentences):
        sentence_norm = normalize_text(sentence)
        if any(keyword and keyword in sentence_norm for keyword in normalized_keywords):
            start = max(index - 1, 0)
            end = min(index + 2, len(sentences))
            snippet = "".join(sentences[start:end]).strip()
            return snippet[:limit].rstrip("，；、 ") if len(snippet) > limit else snippet
    return body[:limit].rstrip("，；、 ") if len(body) > limit else body


def participant_hits_in_text(names: list[str], candidate_norm: str) -> int:
    return sum(1 for name in names if normalize_text(name) and normalize_text(name) in candidate_norm)


def location_hit(tokens: list[str], candidate_norm: str) -> bool:
    return any(normalize_text(token) in candidate_norm for token in tokens if len(normalize_text(token)) >= 2)


def title_hit(variants: list[str], candidate_norm: str) -> bool:
    return any(variant in candidate_norm for variant in variants if variant)


def make_evidence_id(event_id: str, source_id: str, match_rule: str) -> str:
    digest = hashlib.sha1(f"{event_id}|{source_id}|{match_rule}".encode()).hexdigest()[:10]
    return f"EVI-{digest.upper()}"


def append_evidence(
    bucket: dict[str, dict[str, Any]],
    event: dict[str, Any],
    source: dict[str, Any],
    confidence: float,
    match_rule: str,
    keywords: list[str],
) -> None:
    dedupe_key = f"{event['event_id']}::{source['source_id']}"
    existing = bucket.get(dedupe_key)
    if existing and float(existing["confidence"]) >= confidence:
        return
    bucket[dedupe_key] = {
        "evidence_id": make_evidence_id(event["event_id"], source["source_id"], match_rule),
        "event_id": event["event_id"],
        "source": source["source"],
        "source_file": source["source_file"],
        "source_loc": source["source_loc"],
        "quote": choose_quote_excerpt(source["quote"], keywords=keywords),
        "confidence": round(confidence, 2),
        "source_id": source["source_id"],
        "match_rule": MATCH_RULE_NOTE.get(match_rule, match_rule),
    }


def match_from_relation_row(
    event_meta: dict[str, dict[str, Any]],
    source_lookup: dict[str, dict[str, Any]],
    relations_df: pd.DataFrame,
    person_name_map: dict[str, str],
) -> dict[str, dict[str, Any]]:
    evidences: dict[str, dict[str, Any]] = {}

    for _, row in relations_df.iterrows():
        source_ids = [source_id for source_id in split_ids(row.get("source_ids")) if source_id in source_lookup]
        if not source_ids:
            continue

        relation_names = unique_ordered(
            [
                person_name_map.get(text(row.get("source_person_id")), text(row.get("source_person_id"))),
                person_name_map.get(text(row.get("target_person_id")), text(row.get("target_person_id"))),
            ]
        )
        context = clean_ocr_text(text(row.get("context")))
        evidence_ref = text(row.get("evidence_ref"))

        for source_id in source_ids:
            source = source_lookup[source_id]
            candidate_norm = normalize_text(" ".join([source["quote"], context, evidence_ref]))
            matches: list[tuple[str, float, list[str]]] = []

            for event in event_meta.values():
                overlap = len(set(relation_names) & event["participant_name_set"])
                title_match = title_hit(event["title_variants"], candidate_norm)
                loc_match = location_hit(event["location_tokens"], candidate_norm)
                name_hits = participant_hits_in_text(event["participant_names"], candidate_norm)
                exact_date_match = bool(source["exact_date"] and source["exact_date"] == event["event_date"])
                year_match = bool(source["year"] and source["year"] == event["year"])

                if exact_date_match and overlap >= 2 and (loc_match or title_match or event["is_pair_event"]):
                    confidence = 0.84 + min(overlap, 2) * 0.03 + (0.03 if loc_match else 0) + (0.02 if title_match else 0)
                    keywords = relation_names + event["participant_names"] + event["location_tokens"] + [event["title"]]
                    matches.append((event["event_id"], min(confidence, 0.95), keywords))
                    continue

                if (
                    year_match
                    and event["is_collective_event"]
                    and event["is_collective_anchor"]
                    and name_hits >= 4
                    and loc_match
                    and strong_collective_title_hit(event["title"], candidate_norm)
                ):
                    confidence = 0.83 + min(name_hits, 5) * 0.02 + (0.03 if title_match else 0)
                    keywords = relation_names + event["participant_names"] + event["location_tokens"] + [event["title"]]
                    matches.append((event["event_id"], min(confidence, 0.92), keywords))

            matches.sort(key=lambda item: item[1], reverse=True)
            if not matches:
                continue
            if len(matches) > 1 and matches[0][1] - matches[1][1] < 0.05:
                continue

            best_event_id, confidence, keywords = matches[0]
            append_evidence(
                evidences,
                event_meta[best_event_id],
                source,
                confidence=confidence,
                match_rule="relation_exact_date" if source["exact_date"] else "relation_collective",
                keywords=keywords,
            )

    return evidences


def match_from_source_only(
    event_meta: dict[str, dict[str, Any]],
    source_lookup: dict[str, dict[str, Any]],
) -> dict[str, dict[str, Any]]:
    evidences: dict[str, dict[str, Any]] = {}

    for source in source_lookup.values():
        candidate_norm = source["quote_norm"]
        matches: list[tuple[str, float, str, list[str]]] = []

        for event in event_meta.values():
            title_match_value = title_hit(event["title_variants"], candidate_norm)
            location_match_value = location_hit(event["location_tokens"], candidate_norm)
            name_hits = participant_hits_in_text(event["participant_names"], candidate_norm)
            exact_date_match = bool(source["exact_date"] and source["exact_date"] == event["event_date"])
            year_match = bool(source["year"] and source["year"] == event["year"])

            if title_match_value and exact_date_match:
                confidence = 0.9 + min(name_hits, 2) * 0.02
                keywords = event["participant_names"] + event["location_tokens"] + [event["title"]]
                matches.append((event["event_id"], min(confidence, 0.97), "direct_title_date", keywords))
                continue

            if (
                exact_date_match
                and location_match_value
                and not event["is_generic_activity"]
                and (
                    ("通信" not in event["title"] and (name_hits >= 1 or event["is_single_event"]))
                    or ("通信" in event["title"] and (title_match_value or name_hits >= 2))
                )
            ):
                confidence = 0.83 + min(name_hits, 2) * 0.02 + (0.03 if title_match_value else 0)
                keywords = event["participant_names"] + event["location_tokens"] + [event["title"]]
                matches.append((event["event_id"], min(confidence, 0.9), "direct_specific_visit", keywords))
                continue

            if (
                event["is_collective_event"]
                and event["is_collective_anchor"]
                and (
                    (year_match and strong_collective_title_hit(event["title"], candidate_norm) and ("成立" in event["title"] or "大会" in event["title"]))
                    or (year_match and ("遇难" in event["title"] or "被捕" in event["title"]) and name_hits >= 4 and (location_match_value or "龙华" in candidate_norm))
                    or (exact_date_match and location_match_value and name_hits >= 3)
                )
            ):
                confidence = 0.84 + min(name_hits, 4) * 0.02 + (0.02 if title_match_value else 0)
                keywords = event["participant_names"] + event["location_tokens"] + [event["title"]]
                matches.append((event["event_id"], min(confidence, 0.91), "direct_collective", keywords))

        matches.sort(key=lambda item: item[1], reverse=True)
        if not matches:
            continue
        if len(matches) > 1 and matches[0][1] - matches[1][1] < 0.05:
            continue

        best_event_id, confidence, match_rule, keywords = matches[0]
        append_evidence(
            evidences,
            event_meta[best_event_id],
            source,
            confidence=confidence,
            match_rule=match_rule,
            keywords=keywords,
        )

    return evidences


def load_tables(kb_dir: Path) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    events_df = pd.read_csv(kb_dir / "events.csv").fillna("")
    participants_df = pd.read_csv(kb_dir / "event_participants.csv").fillna("")
    relations_df = pd.read_csv(kb_dir / "person_relations.csv").fillna("")
    sources_df = pd.read_csv(kb_dir / "sources.csv").fillna("")
    return events_df, participants_df, relations_df, sources_df


def build_event_evidences(kb_dir: Path) -> list[dict[str, Any]]:
    events_df, participants_df, relations_df, sources_df = load_tables(kb_dir)
    person_rows = pd.read_csv(kb_dir / "persons.csv").fillna("")
    person_name_map = person_rows.set_index("person_id")["standard_name"].to_dict()

    event_meta = build_event_meta(events_df, participants_df)
    source_lookup = build_source_lookup(sources_df)
    relation_matches = match_from_relation_row(event_meta, source_lookup, relations_df, person_name_map)
    direct_matches = match_from_source_only(event_meta, source_lookup)

    merged = relation_matches
    for dedupe_key, evidence in direct_matches.items():
        existing = merged.get(dedupe_key)
        if existing is None or float(existing["confidence"]) < float(evidence["confidence"]):
            merged[dedupe_key] = evidence

    grouped: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for evidence in merged.values():
        grouped[evidence["event_id"]].append(evidence)

    ordered: list[dict[str, Any]] = []
    for event_id in sorted(grouped.keys()):
        items = sorted(
            grouped[event_id],
            key=lambda item: (-float(item["confidence"]), item["source"], item["source_loc"], item["evidence_id"]),
        )[:8]
        ordered.extend(items)
    return ordered


def main() -> None:
    parser = argparse.ArgumentParser(description="Build structured event evidences from local raw texts and relation traces.")
    parser.add_argument("--kb-dir", type=Path, default=KB_DIR, help="知识库标准表目录，默认 data/processed")
    parser.add_argument("--output", type=Path, default=OUTPUT_FILE, help="结构化事件证据输出文件")
    args = parser.parse_args()

    evidences = build_event_evidences(args.kb_dir)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(evidences, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"WROTE\t{args.output}")
    print(f"EVENT_EVIDENCES\t{len(evidences)}")
    print(f"EVENTS_WITH_EVIDENCE\t{len({item['event_id'] for item in evidences})}")


if __name__ == "__main__":
    main()
