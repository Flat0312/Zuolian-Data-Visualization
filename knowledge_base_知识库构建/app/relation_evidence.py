from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path
from typing import Any

import pandas as pd

DATE_CN_RE = re.compile(r"((?:18|19|20)\d{2})年(\d{1,2})月(\d{1,2})日")
DATE_ISO_RE = re.compile(r"((?:18|19|20)\d{2})-(\d{1,2})-(\d{1,2})(?:\s+\d{1,2}:\d{1,2}:\d{1,2})?")
YEAR_MONTH_CN_RE = re.compile(r"((?:18|19|20)\d{2})年(\d{1,2})月")
YEAR_CN_RE = re.compile(r"((?:18|19|20)\d{2})年")
YEAR_RANGE_RE = re.compile(r"((?:18|19|20)\d{2})\s*[-—–~～至到]+\s*((?:18|19|20)\d{2})年?")
PAGE_HEADER_RE = re.compile(r"第\s*(\d+)\s*页")
PAGE_CITATION_RE = re.compile(r"第\s*(\d+)\s*页")
LEFT_WING_SOURCE_TITLES = {"左联词典", "左联史", "左联回忆录"}
LEFT_WING_PERIOD_LABEL = "左联时期（1930-03-02 至 1936年初）"
LEFT_WING_PERIOD_START = "1930-03-02"
LEFT_WING_PERIOD_END_DISPLAY = "1936年初"
LEFT_WING_PERIOD_END_SORT = "1936-02-01"
HISTORICAL_YEAR_MIN = 1925
HISTORICAL_YEAR_MAX = 1937


def canonical_pair_key(person_a_id: str, person_b_id: str) -> str:
    return "__".join(sorted((str(person_a_id), str(person_b_id))))


def _clean(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return " ".join(str(value).replace("\r", " ").split())


def _unique(values: list[str]) -> list[str]:
    items: list[str] = []
    for value in values:
        text = _clean(value)
        if text and text not in items:
            items.append(text)
    return items


def _ordered_people(person_a_id: str, person_a_name: str, person_b_id: str, person_b_name: str) -> tuple[str, str, str, str]:
    ordered = sorted(
        ((str(person_a_name), str(person_a_id)), (str(person_b_name), str(person_b_id))),
        key=lambda item: (item[0], item[1]),
    )
    return ordered[0][1], ordered[0][0], ordered[1][1], ordered[1][0]


def _extract_date_token(*texts: str) -> tuple[str, str]:
    for text in texts:
        content = _clean(text)
        if not content:
            continue
        match_cn = DATE_CN_RE.search(content)
        if match_cn:
            year, month, day = match_cn.groups()
            return match_cn.group(0), f"{int(year):04d}-{int(month):02d}-{int(day):02d}"
        match_iso = DATE_ISO_RE.search(content)
        if match_iso:
            year, month, day = match_iso.groups()
            return match_iso.group(0), f"{int(year):04d}-{int(month):02d}-{int(day):02d}"
    return "", ""


@dataclass(slots=True)
class TimeHint:
    start_display: str = ""
    end_display: str = ""
    start_sort: str = ""
    end_sort: str = ""
    label: str = ""
    precision: str = ""

    @property
    def is_empty(self) -> bool:
        return not (self.start_display or self.end_display or self.label)


def _month_last_day(year: int, month: int) -> int:
    if month == 2:
        if (year % 400 == 0) or (year % 4 == 0 and year % 100 != 0):
            return 29
        return 28
    if month in {4, 6, 9, 11}:
        return 30
    return 31


def _time_hint(
    start_display: str,
    start_sort: str,
    end_display: str = "",
    end_sort: str = "",
    precision: str = "",
    label: str = "",
) -> TimeHint:
    start_display = _clean(start_display)
    end_display = _clean(end_display) or start_display
    start_sort = _clean(start_sort)
    end_sort = _clean(end_sort) or start_sort
    computed_label = _clean(label)
    if not computed_label and start_display and end_display:
        computed_label = start_display if start_display == end_display else f"{start_display} 至 {end_display}"
    return TimeHint(
        start_display=start_display,
        end_display=end_display,
        start_sort=start_sort,
        end_sort=end_sort,
        label=computed_label,
        precision=precision,
    )


def _left_wing_period_hint() -> TimeHint:
    return _time_hint(
        start_display=LEFT_WING_PERIOD_START,
        start_sort=LEFT_WING_PERIOD_START,
        end_display=LEFT_WING_PERIOD_END_DISPLAY,
        end_sort=LEFT_WING_PERIOD_END_SORT,
        precision="fallback",
        label=LEFT_WING_PERIOD_LABEL,
    )


def _extract_year_month_tokens(*texts: str) -> list[tuple[int, int]]:
    tokens: set[tuple[int, int]] = set()
    for text in texts:
        content = _clean(text)
        if not content:
            continue
        for year_text, month_text in YEAR_MONTH_CN_RE.findall(content):
            year = int(year_text)
            month = int(month_text)
            if HISTORICAL_YEAR_MIN <= year <= HISTORICAL_YEAR_MAX and 1 <= month <= 12:
                tokens.add((year, month))
    return sorted(tokens)


def _extract_year_tokens(*texts: str) -> list[int]:
    years: set[int] = set()
    for text in texts:
        content = _clean(text)
        if not content:
            continue
        for start_text, end_text in YEAR_RANGE_RE.findall(content):
            start_year = int(start_text)
            end_year = int(end_text)
            for year in (start_year, end_year):
                if HISTORICAL_YEAR_MIN <= year <= HISTORICAL_YEAR_MAX:
                    years.add(year)
        for year_text in YEAR_CN_RE.findall(content):
            year = int(year_text)
            if HISTORICAL_YEAR_MIN <= year <= HISTORICAL_YEAR_MAX:
                years.add(year)
    return sorted(years)


def _extract_time_hint(*texts: str) -> TimeHint:
    exact_dates: list[str] = []
    for text in texts:
        content = _clean(text)
        if not content:
            continue
        for _, iso_date in [_extract_date_token(content)]:
            if iso_date:
                exact_dates.append(iso_date)
        for year_text, month_text, day_text in DATE_CN_RE.findall(content):
            exact_dates.append(f"{int(year_text):04d}-{int(month_text):02d}-{int(day_text):02d}")
        for year_text, month_text, day_text in DATE_ISO_RE.findall(content):
            exact_dates.append(f"{int(year_text):04d}-{int(month_text):02d}-{int(day_text):02d}")

    exact_dates = sorted(set(exact_dates))
    if exact_dates:
        return _time_hint(
            start_display=exact_dates[0],
            start_sort=exact_dates[0],
            end_display=exact_dates[-1],
            end_sort=exact_dates[-1],
            precision="day",
        )

    year_months = _extract_year_month_tokens(*texts)
    if year_months:
        start_year, start_month = year_months[0]
        end_year, end_month = year_months[-1]
        start_display = f"{start_year:04d}-{start_month:02d}"
        end_display = f"{end_year:04d}-{end_month:02d}"
        return _time_hint(
            start_display=start_display,
            start_sort=f"{start_year:04d}-{start_month:02d}-01",
            end_display=end_display,
            end_sort=f"{end_year:04d}-{end_month:02d}-{_month_last_day(end_year, end_month):02d}",
            precision="month",
        )

    years = _extract_year_tokens(*texts)
    if years:
        return _time_hint(
            start_display=str(years[0]),
            start_sort=f"{years[0]:04d}-01-01",
            end_display=str(years[-1]),
            end_sort=f"{years[-1]:04d}-12-31",
            precision="year",
        )
    return TimeHint()


def _parse_paged_txt(path: Path) -> dict[int, str]:
    pages: dict[int, list[str]] = {}
    current_page: int | None = None
    if not path.exists():
        return {}
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        match = PAGE_HEADER_RE.search(raw_line)
        if match:
            current_page = int(match.group(1))
            pages.setdefault(current_page, [])
            continue
        if current_page is None:
            continue
        pages[current_page].append(raw_line)
    return {page: "\n".join(lines) for page, lines in pages.items()}


def _parse_paged_json(path: Path) -> dict[int, str]:
    if not path.exists():
        return {}
    payload = json.loads(path.read_text(encoding="utf-8"))
    result: dict[int, str] = {}
    for key, value in payload.items():
        if str(key).isdigit():
            result[int(key)] = _clean(value)
    return result


@lru_cache(maxsize=1)
def _load_source_page_index(base_dir_text: str) -> dict[str, dict[int, str]]:
    base_dir = Path(base_dir_text)
    project_root = base_dir.parents[1]
    return {
        "左联词典": _parse_paged_txt(project_root / "input_输入" / "raw_texts_原始文本" / "左联词典.txt"),
        "左联史": _parse_paged_txt(project_root / "input_输入" / "raw_texts_原始文本" / "左联史.txt"),
        "左联回忆录": _parse_paged_json(project_root / "work_处理中间数据" / "extracted_抽取结果" / "左联回忆录_ocr_text.json"),
    }


def _extract_time_hint_from_citation_pages(base_dir: Path, citation_ref: str) -> TimeHint:
    page_index = _load_source_page_index(str(base_dir.resolve()))
    page_texts: list[str] = []
    for fragment in re.split(r"[;；]+", citation_ref):
        citation = _clean(fragment)
        if not citation:
            continue
        for source_title, pages in page_index.items():
            if source_title not in citation:
                continue
            page_match = PAGE_CITATION_RE.search(citation)
            if not page_match:
                continue
            page_number = int(page_match.group(1))
            for nearby_page in (page_number - 1, page_number, page_number + 1):
                page_text = _clean(pages.get(nearby_page))
                if page_text:
                    page_texts.append(page_text)
    if not page_texts:
        return TimeHint()
    return _extract_time_hint(*page_texts)


def _fallback_time_hint(*texts: str) -> TimeHint:
    for text in texts:
        content = _clean(text)
        if any(source_title in content for source_title in LEFT_WING_SOURCE_TITLES):
            return _left_wing_period_hint()
    return TimeHint()


def _infer_source_title(citation_ref: str) -> str:
    if not citation_ref:
        return "来源待补"
    first = _clean(citation_ref).split(";")[0]
    title_match = re.match(r"^(.*?)(?=(?:\s*(?:18|19|20)\d{2}[年-])|\s*第\d|$)", first)
    title = _clean(title_match.group(1) if title_match else first)
    return title or first or "来源待补"


def _support_note(person_a_name: str, person_b_name: str, relation_type: str) -> str:
    relation = _clean(relation_type)
    if "通信" in relation:
        return f"该条材料记录了 {person_a_name} 与 {person_b_name} 的书信或消息往来，可作为通信关系的直接证据。"
    if "亲属" in relation:
        return f"该条材料呈现了 {person_a_name} 与 {person_b_name} 的家庭往来，可作为亲属关系的支撑材料。"
    if "组织" in relation:
        return f"该条材料把 {person_a_name} 与 {person_b_name} 置于共同的左联活动或组织语境中，可支持组织关联判断。"
    if "时空共现" in relation:
        return f"该条材料显示 {person_a_name} 与 {person_b_name} 在同一时间或场景中的并置，可作为时空共现证据。"
    return f"该条材料可用来说明 {person_a_name} 与 {person_b_name} 之间存在“{relation or '未标注'}”关联。"


@dataclass(slots=True)
class EvidenceRecord:
    evidence_id: str
    relation_type: str
    source_title: str
    source_date: str
    excerpt: str
    citation_ref: str
    support_note: str
    provenance: str = "derived"
    sort_date: str = ""
    sort_end_date: str = ""
    display_start: str = ""
    display_end: str = ""
    time_precision: str = ""

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "EvidenceRecord":
        return cls(
            evidence_id=_clean(data.get("evidence_id")),
            relation_type=_clean(data.get("relation_type")),
            source_title=_clean(data.get("source_title")) or "来源待补",
            source_date=_clean(data.get("source_date")) or "时间待补",
            excerpt=_clean(data.get("excerpt")),
            citation_ref=_clean(data.get("citation_ref")),
            support_note=_clean(data.get("support_note")),
            provenance=_clean(data.get("provenance")) or "mock",
            sort_date=_clean(data.get("sort_date")),
            sort_end_date=_clean(data.get("sort_end_date")),
            display_start=_clean(data.get("display_start")),
            display_end=_clean(data.get("display_end")),
            time_precision=_clean(data.get("time_precision")),
        )


@dataclass(slots=True)
class RelationDetail:
    relation_id: str
    pair_key: str
    person_a_id: str
    person_a_name: str
    person_b_id: str
    person_b_name: str
    relation_types: list[str] = field(default_factory=list)
    strength: float = 0.0
    first_seen: str = "时间待补"
    last_seen: str = "时间待补"
    time_range_label: str = "时间待补"
    evidence_samples: list[EvidenceRecord] = field(default_factory=list)
    provenance: str = "derived"

    @property
    def evidence_count(self) -> int:
        return len(self.evidence_samples)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "RelationDetail":
        relation = cls(
            relation_id=_clean(data.get("relation_id")) or canonical_pair_key(data.get("person_a_id", ""), data.get("person_b_id", "")),
            pair_key=_clean(data.get("pair_key")) or canonical_pair_key(data.get("person_a_id", ""), data.get("person_b_id", "")),
            person_a_id=_clean(data.get("person_a_id")),
            person_a_name=_clean(data.get("person_a_name")),
            person_b_id=_clean(data.get("person_b_id")),
            person_b_name=_clean(data.get("person_b_name")),
            relation_types=_unique([str(item) for item in data.get("relation_types", [])]),
            strength=float(data.get("strength") or 0.0),
            first_seen=_clean(data.get("first_seen")) or "时间待补",
            last_seen=_clean(data.get("last_seen")) or "时间待补",
            time_range_label=_clean(data.get("time_range_label")) or "时间待补",
            evidence_samples=[EvidenceRecord.from_dict(item) for item in data.get("evidence_samples", [])],
            provenance=_clean(data.get("provenance")) or "mock",
        )
        relation.finalize()
        return relation

    def finalize(self) -> None:
        self.relation_types = _unique(self.relation_types)
        dated = [item for item in self.evidence_samples if item.sort_date]
        if dated:
            earliest = min(dated, key=lambda item: (item.sort_date, item.display_start or item.source_date, item.source_title))
            latest = max(
                dated,
                key=lambda item: (
                    item.sort_end_date or item.sort_date,
                    item.display_end or item.source_date,
                    item.source_title,
                ),
            )
            if not self.first_seen or self.first_seen == "时间待补":
                self.first_seen = earliest.display_start or earliest.source_date or earliest.sort_date
            if not self.last_seen or self.last_seen == "时间待补":
                self.last_seen = latest.display_end or latest.source_date or latest.sort_end_date or latest.sort_date
        if not self.first_seen:
            self.first_seen = "时间待补"
        if not self.last_seen:
            self.last_seen = "时间待补"
        if not self.time_range_label or self.time_range_label == "时间待补":
            if self.first_seen == self.last_seen and self.first_seen != "时间待补":
                self.time_range_label = self.first_seen
            elif self.first_seen != "时间待补" or self.last_seen != "时间待补":
                self.time_range_label = f"{self.first_seen} 至 {self.last_seen}"
            else:
                self.time_range_label = "时间待补"


def _row_to_evidence(
    base_dir: Path,
    row: pd.Series,
    pair_key: str,
    person_a_name: str,
    person_b_name: str,
    index: int,
) -> EvidenceRecord | None:
    citation_ref = _clean(row.get("Evidence_Ref"))
    excerpt = _clean(row.get("Context"))
    if not citation_ref and not excerpt:
        return None

    relation_type = _clean(row.get("Relation_Family")) or _clean(row.get("Relation_Type")) or "未标注"
    time_hint = _extract_time_hint(citation_ref, excerpt)
    if time_hint.is_empty and citation_ref:
        time_hint = _extract_time_hint_from_citation_pages(base_dir, citation_ref)
    if time_hint.is_empty:
        time_hint = _fallback_time_hint(citation_ref, excerpt, _infer_source_title(citation_ref))
    return EvidenceRecord(
        evidence_id=f"{pair_key}-ev-{index:03d}",
        relation_type=relation_type,
        source_title=_infer_source_title(citation_ref),
        source_date=time_hint.label or "时间待补",
        excerpt=excerpt,
        citation_ref=citation_ref,
        support_note=_support_note(person_a_name, person_b_name, relation_type),
        provenance="derived",
        sort_date=time_hint.start_sort,
        sort_end_date=time_hint.end_sort,
        display_start=time_hint.start_display,
        display_end=time_hint.end_display,
        time_precision=time_hint.precision,
    )


def _merge_override(actual: RelationDetail | None, override: RelationDetail) -> RelationDetail:
    if actual is None:
        override.finalize()
        return override

    merged = RelationDetail(
        relation_id=actual.relation_id or override.relation_id,
        pair_key=actual.pair_key or override.pair_key,
        person_a_id=override.person_a_id or actual.person_a_id,
        person_a_name=override.person_a_name or actual.person_a_name,
        person_b_id=override.person_b_id or actual.person_b_id,
        person_b_name=override.person_b_name or actual.person_b_name,
        relation_types=_unique(actual.relation_types + override.relation_types),
        strength=override.strength or actual.strength,
        first_seen=override.first_seen or actual.first_seen,
        last_seen=override.last_seen or actual.last_seen,
        time_range_label=override.time_range_label or actual.time_range_label,
        evidence_samples=override.evidence_samples or actual.evidence_samples,
        provenance=override.provenance or actual.provenance,
    )
    merged.finalize()
    return merged


def load_mock_relation_details(base_dir: Path) -> dict[str, RelationDetail]:
    mock_path = Path(base_dir) / "relation_evidence_mock.json"
    if not mock_path.exists():
        return {}
    payload = json.loads(mock_path.read_text(encoding="utf-8"))
    details = [RelationDetail.from_dict(item) for item in payload]
    return {detail.pair_key: detail for detail in details}


def build_relation_detail_index(base_dir: Path, nodes_df: pd.DataFrame, edges_df: pd.DataFrame) -> dict[str, RelationDetail]:
    name_map = nodes_df.set_index("Id")["Label"].to_dict()
    working = edges_df.copy()
    if "Relation_Family" not in working.columns:
        working["Relation_Family"] = working["Relation_Type"].astype(str).str.replace(r"^(强关联-|弱关联-)", "", regex=True)
    working["pair_key"] = working.apply(lambda row: canonical_pair_key(row["Source"], row["Target"]), axis=1)

    details: dict[str, RelationDetail] = {}
    for pair_key, group in working.groupby("pair_key", sort=False):
        source_id = str(group.iloc[0]["Source"])
        target_id = str(group.iloc[0]["Target"])
        source_name = str(name_map.get(source_id, source_id))
        target_name = str(name_map.get(target_id, target_id))
        person_a_id, person_a_name, person_b_id, person_b_name = _ordered_people(
            source_id, source_name, target_id, target_name
        )
        relation_types = _unique(group["Relation_Family"].astype(str).tolist())

        # The current edge table does not yet have a first-class evidence schema,
        # so we derive a minimum viable evidence chain from each edge row.
        evidence_samples: list[EvidenceRecord] = []
        seen_keys: set[tuple[str, str, str]] = set()
        for index, (_, row) in enumerate(group.iterrows(), start=1):
            evidence = _row_to_evidence(base_dir, row, pair_key, person_a_name, person_b_name, index)
            if evidence is None:
                continue
            dedupe_key = (evidence.relation_type, evidence.citation_ref, evidence.excerpt)
            if dedupe_key in seen_keys:
                continue
            seen_keys.add(dedupe_key)
            evidence_samples.append(evidence)

        dated_starts = [item for item in evidence_samples if item.sort_date]
        detail = RelationDetail(
            relation_id=pair_key,
            pair_key=pair_key,
            person_a_id=person_a_id,
            person_a_name=person_a_name,
            person_b_id=person_b_id,
            person_b_name=person_b_name,
            relation_types=relation_types,
            strength=float(pd.to_numeric(group["Weight"], errors="coerce").fillna(0).max()),
            first_seen=(
                min(dated_starts, key=lambda item: item.sort_date).display_start
                if dated_starts
                else "时间待补"
            ),
            last_seen=(
                max(dated_starts, key=lambda item: item.sort_end_date or item.sort_date).display_end
                if dated_starts
                else "时间待补"
            ),
            evidence_samples=evidence_samples,
            provenance="derived",
        )
        detail.finalize()
        details[pair_key] = detail

    # Mock JSON can override selected pairs with richer archival fields until a real source system arrives.
    for pair_key, override in load_mock_relation_details(base_dir).items():
        details[pair_key] = _merge_override(details.get(pair_key), override)
    return details
