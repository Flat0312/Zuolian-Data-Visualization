from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import pandas as pd

from utils import clean_text as _clean, split_ids as _split_ids


def _to_float(value: object) -> float | None:
    if value is None or value == "":
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    if pd.isna(number):
        return None
    return number


def _parse_year(value: object) -> int | None:
    text = _clean(value)
    if not text:
        return None
    match = pd.Series([text]).str.extract(r"((?:18|19|20)\d{2})")[0].iloc[0]
    if pd.isna(match):
        return None
    return int(match)


def _summary(location_name: str, category: str, people: list[str]) -> str:
    people_text = "、".join(people[:4]) if people else "相关人物待补"
    return f"{location_name}发生“{category}”事件，涉及 {people_text}，可作为左联时空活动重构的基础线索。"


def _significance(category: str, people: list[str], location_name: str) -> str:
    if "成立" in category or "会议" in category:
        return f"该事件把 {location_name} 标定为左联组织化的重要空间节点，可用于重建网络形成的历史现场。"
    if "逮捕" in category or "牺牲" in category:
        return f"该事件揭示了左联网络遭遇压迫的关键时刻，有助于理解人物关系从文化合作转向政治风险的轨迹。"
    if "通信" in category or "交往" in category:
        return f"该事件体现人物之间的联络与流动，可补充左联关系网络在日常层面的连接机制。"
    if people:
        return f"该事件把 {location_name} 与 {len(people)} 位相关人物联系起来，有助于分析左联网络在空间上的聚合方式。"
    return f"该事件为 {location_name} 的左联活动提供了时空锚点。"


@dataclass(slots=True)
class EventEvidence:
    evidence_id: str
    event_id: str
    source: str
    source_file: str
    source_loc: str
    quote: str
    confidence: float | None = None
    source_id: str = ""
    match_rule: str = ""

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "EventEvidence":
        return cls(
            evidence_id=_clean(data.get("evidence_id")) or _clean(data.get("id")),
            event_id=_clean(data.get("event_id")),
            source=_clean(data.get("source")) or _clean(data.get("source_title")) or "来源待补",
            source_file=_clean(data.get("source_file")),
            source_loc=_clean(data.get("source_loc")) or _clean(data.get("citation_ref")) or _clean(data.get("date")),
            quote=_clean(data.get("quote")) or _clean(data.get("excerpt")),
            confidence=_to_float(data.get("confidence")),
            source_id=_clean(data.get("source_id")),
            match_rule=_clean(data.get("match_rule")),
        )

    @property
    def source_title(self) -> str:
        return self.source

    @property
    def date(self) -> str:
        return self.source_loc or "位置待补"

    @property
    def excerpt(self) -> str:
        return self.quote

    @property
    def citation_ref(self) -> str:
        return self.source_loc

    @property
    def note(self) -> str:
        if self.confidence is None:
            return ""
        return f"自动映射置信度 {self.confidence:.2f}"


@dataclass(slots=True)
class RelatedPerson:
    person_id: str
    name: str
    relation: str = ""

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "RelatedPerson":
        return cls(
            person_id=_clean(data.get("person_id")),
            name=_clean(data.get("name")) or _clean(data.get("participant_name")),
            relation=_clean(data.get("relation")) or _clean(data.get("participant_role")),
        )


@dataclass(slots=True)
class HistoricalEvent:
    # This schema is intentionally close to a JSON/document payload so the current
    # mock layer can later be replaced by a database row or a DataV/API response.
    event_id: str
    title: str
    date: str
    year: int | None
    location_name: str
    region_code: str
    map_region: str
    latitude: float | None
    longitude: float | None
    related_persons: list[RelatedPerson] = field(default_factory=list)
    category: str = ""
    summary: str = ""
    evidences: list[EventEvidence] = field(default_factory=list)
    significance: str = ""
    source_excerpt: str = ""
    source_ids: list[str] = field(default_factory=list)
    region_match_status: str = "matched"

    @property
    def id(self) -> str:
        return self.event_id

    @property
    def people(self) -> list[str]:
        return [person.name for person in self.related_persons if person.name]

    @property
    def evidence(self) -> list[EventEvidence]:
        return self.evidences


@dataclass(slots=True)
class MapRegion:
    region_code: str
    name: str
    display_name: str
    keywords: list[str]
    polygons: list[list[list[float]]]
    centroid_lon: float
    centroid_lat: float


def _centroid(points: list[list[float]]) -> tuple[float, float]:
    xs = [point[0] for point in points]
    ys = [point[1] for point in points]
    return sum(xs) / len(xs), sum(ys) / len(ys)


def _flatten_polygons(feature: dict[str, Any]) -> list[list[list[float]]]:
    geometry = feature.get("geometry", {})
    geom_type = geometry.get("type")
    coords = geometry.get("coordinates", [])
    if geom_type == "Polygon":
        return [coords[0]]
    if geom_type == "MultiPolygon":
        return [polygon[0] for polygon in coords]
    return []


def _point_in_polygon(lon: float, lat: float, polygon: list[list[float]]) -> bool:
    inside = False
    j = len(polygon) - 1
    for i in range(len(polygon)):
        xi, yi = polygon[i]
        xj, yj = polygon[j]
        intersects = ((yi > lat) != (yj > lat)) and (
            lon < (xj - xi) * (lat - yi) / ((yj - yi) or 1e-12) + xi
        )
        if intersects:
            inside = not inside
        j = i
    return inside


def load_map_regions(base_dir: Path) -> tuple[list[MapRegion], dict[str, Any]]:
    # The current mock GeoJSON stands in for a formal historical boundary file.
    # Replacing the file with official GeoJSON should not require UI changes.
    geojson_path = Path(base_dir) / "shanghai_datav_mock.geojson"
    payload = json.loads(geojson_path.read_text(encoding="utf-8"))
    regions: list[MapRegion] = []
    for feature in payload.get("features", []):
        properties = feature.get("properties", {})
        polygons = _flatten_polygons(feature)
        all_points = [point for polygon in polygons for point in polygon]
        centroid_lon, centroid_lat = _centroid(all_points)
        regions.append(
            MapRegion(
                region_code=_clean(properties.get("region_code")),
                name=_clean(properties.get("name")),
                display_name=_clean(properties.get("display_name")) or _clean(properties.get("name")),
                keywords=[_clean(item) for item in properties.get("keywords", []) if _clean(item)],
                polygons=polygons,
                centroid_lon=centroid_lon,
                centroid_lat=centroid_lat,
            )
        )
    return regions, payload


def _infer_region(
    lon: float | None,
    lat: float | None,
    current_loc: str,
    hist_loc: str,
    regions: list[MapRegion],
) -> tuple[str, str, str]:
    if lon is not None and lat is not None:
        for region in regions:
            if any(_point_in_polygon(lon, lat, polygon) for polygon in region.polygons):
                return region.region_code, region.display_name, "matched"

    location_text = f"{current_loc} {hist_loc}"
    for region in regions:
        if any(keyword and keyword in location_text for keyword in region.keywords):
            return region.region_code, region.display_name, "keyword"

    if lon is not None and lat is not None and regions:
        nearest = min(
            regions,
            key=lambda region: (region.centroid_lon - lon) ** 2 + (region.centroid_lat - lat) ** 2,
        )
        return nearest.region_code, nearest.display_name, "nearest"

    return "unmatched", "区域待考", "unmatched"


def load_event_overrides(base_dir: Path) -> dict[str, dict[str, Any]]:
    override_path = Path(base_dir) / "historical_event_overrides.json"
    if not override_path.exists():
        return {}
    payload = json.loads(override_path.read_text(encoding="utf-8"))
    return {str(item["id"]): item for item in payload}


def load_event_evidence_index(data_dir: Path | None) -> dict[str, list[EventEvidence]]:
    if data_dir is None:
        return {}
    evidence_path = Path(data_dir) / "event_evidences.json"
    if not evidence_path.exists():
        return {}
    payload = json.loads(evidence_path.read_text(encoding="utf-8"))
    evidence_index: dict[str, list[EventEvidence]] = {}
    for item in payload:
        evidence = EventEvidence.from_dict(item)
        if not evidence.event_id:
            continue
        evidence_index.setdefault(evidence.event_id, []).append(evidence)
    for event_id, items in evidence_index.items():
        evidence_index[event_id] = sorted(
            items,
            key=lambda evidence: (
                -(evidence.confidence if evidence.confidence is not None else -1.0),
                evidence.source,
                evidence.source_loc,
                evidence.evidence_id,
            ),
        )
    return evidence_index


def _coerce_related_persons(value: object) -> list[RelatedPerson]:
    if value is None or value == "" or (isinstance(value, float) and pd.isna(value)):
        return []
    raw_items: list[dict[str, Any]]
    if isinstance(value, list):
        raw_items = [item for item in value if isinstance(item, dict)]
    else:
        text = str(value).strip()
        if not text:
            return []
        try:
            parsed = json.loads(text)
        except json.JSONDecodeError:
            return []
        raw_items = [item for item in parsed if isinstance(item, dict)]
    people: list[RelatedPerson] = []
    for item in raw_items:
        person = RelatedPerson.from_dict(item)
        if person.person_id or person.name:
            people.append(person)
    return people


def build_historical_events(
    base_dir: Path,
    data_dir: Path | None,
    nodes_df: pd.DataFrame,
    events_df: pd.DataFrame,
) -> tuple[list[HistoricalEvent], dict[str, Any]]:
    name_map = nodes_df.set_index("Id")["Label"].to_dict()
    regions, geojson = load_map_regions(base_dir)
    overrides = load_event_overrides(base_dir)
    evidence_index = load_event_evidence_index(data_dir)

    events: list[HistoricalEvent] = []
    for index, row in events_df.reset_index(drop=True).iterrows():
        # Assemble a stable event payload for the UI while preserving a clear
        # seam where real archive evidence or database-backed metadata can enter.
        event_id = _clean(row.get("Event_ID")) or f"EVT-{index:04d}"
        related_persons = _coerce_related_persons(row.get("Related_Persons"))
        if not related_persons:
            related_persons = [
                RelatedPerson(person_id=item, name=name_map.get(item, item), relation="")
                for item in _split_ids(row.get("Entity_ID"))
            ]
        title = _clean(row.get("Event")) or "未命名事件"
        date = _clean(row.get("Timestamp")) or "时间待补"
        year = row.get("Year")
        year_value = int(year) if pd.notna(year) and str(year) != "" else _parse_year(date)
        location_name = _clean(row.get("Current_Loc")) or _clean(row.get("Hist_Loc")) or "地点待考"
        lon = pd.to_numeric(pd.Series([row.get("Longitude")]), errors="coerce").iloc[0]
        lat = pd.to_numeric(pd.Series([row.get("Latitude")]), errors="coerce").iloc[0]
        longitude = None if pd.isna(lon) else float(lon)
        latitude = None if pd.isna(lat) else float(lat)
        region_code, map_region, match_status = _infer_region(
            longitude,
            latitude,
            _clean(row.get("Current_Loc")),
            _clean(row.get("Hist_Loc")),
            regions,
        )

        event = HistoricalEvent(
            event_id=event_id,
            title=title,
            date=date,
            year=year_value,
            location_name=location_name,
            region_code=region_code,
            map_region=map_region,
            latitude=latitude,
            longitude=longitude,
            related_persons=related_persons,
            category=title,
            summary=_summary(location_name, title, [person.name for person in related_persons if person.name]),
            evidences=evidence_index.get(event_id, []),
            significance=_significance(title, [person.name for person in related_persons if person.name], location_name),
            source_excerpt="",
            source_ids=_split_ids(row.get("Source_IDs")),
            region_match_status=match_status,
        )
        if event.evidences:
            event.source_excerpt = event.evidences[0].quote

        override = overrides.get(event_id)
        if override:
            event.title = _clean(override.get("title")) or event.title
            event.date = _clean(override.get("date")) or event.date
            event.year = int(override["year"]) if override.get("year") else event.year
            event.location_name = _clean(override.get("location_name")) or event.location_name
            event.region_code = _clean(override.get("region_code")) or event.region_code
            event.map_region = _clean(override.get("map_region")) or event.map_region
            event.category = _clean(override.get("category")) or event.category
            event.summary = _clean(override.get("summary")) or event.summary
            event.significance = _clean(override.get("significance")) or event.significance
            if override.get("people"):
                event.related_persons = [
                    RelatedPerson(person_id="", name=_clean(item), relation="")
                    for item in override.get("people", [])
                    if _clean(item)
                ]
            if override.get("evidence"):
                event.evidences = [EventEvidence.from_dict(item) for item in override.get("evidence", [])]
            if event.evidences:
                event.source_excerpt = event.evidences[0].quote
        events.append(event)

    return events, geojson


def events_to_frame(events: list[HistoricalEvent]) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    for event in events:
        rows.append(
            {
                "id": event.id,
                "title": event.title,
                "date": event.date,
                "year": event.year,
                "location_name": event.location_name,
                "region_code": event.region_code,
                "map_region": event.map_region,
                "latitude": event.latitude,
                "longitude": event.longitude,
                "people": event.people,
                "people_label": "、".join(event.people) if event.people else "人物待补",
                "category": event.category,
                "summary": event.summary,
                "significance": event.significance,
                "evidence_count": len(event.evidences),
                "source_excerpt": event.source_excerpt,
                "region_match_status": event.region_match_status,
                "event": event,
            }
        )
    return pd.DataFrame(rows)
