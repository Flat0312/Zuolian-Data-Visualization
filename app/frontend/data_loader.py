from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pandas as pd
import streamlit as st

from data_paths import candidate_data_dirs, format_candidate_paths, resolve_data_dir
from utils import clean_text, split_ids


@dataclass(frozen=True, slots=True)
class LoadedData:
    nodes: pd.DataFrame
    edges: pd.DataFrame
    events: pd.DataFrame
    data_dir: str


def match(series: pd.Series, query: str) -> pd.Series:
    return series.fillna("").astype(str).str.contains(query, case=False, regex=False)


def show(value: object, fallback: str = "未标注") -> str:
    if value is None:
        return fallback
    if isinstance(value, float) and pd.isna(value):
        return fallback
    text = str(value).strip()
    return text or fallback


@st.cache_data(show_spinner=False)
def standardized_data_exists(data_dir: Path) -> bool:
    required = [
        "persons.csv",
        "organizations.csv",
        "places.csv",
        "events.csv",
        "person_relations.csv",
        "org_memberships.csv",
        "event_participants.csv",
        "sources.csv",
    ]
    return all((data_dir / filename).exists() for filename in required)


def _fill_from_place(df: pd.DataFrame, col: str, place_map: dict) -> pd.Series:
    return df[col].where(df[col].astype(str).str.strip() != "", df["place_id"].map(place_map).fillna(""))


def load_standardized_views(data_dir: Path) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    persons = pd.read_csv(data_dir / "persons.csv").fillna("")
    relations = pd.read_csv(data_dir / "person_relations.csv").fillna("")
    standard_events = pd.read_csv(data_dir / "events.csv").fillna("")
    participants = pd.read_csv(data_dir / "event_participants.csv").fillna("")
    places = pd.read_csv(data_dir / "places.csv").fillna("")

    nodes = persons.rename(
        columns={
            "person_id": "Id",
            "standard_name": "Label",
            "aliases": "Alias",
            "reliability": "Reliability",
            "birth_death": "Birth_Death",
            "role": "Role",
        }
    ).copy()
    for column in ["Id", "Label", "Alias", "Reliability", "Birth_Death", "Role"]:
        if column not in nodes.columns:
            nodes[column] = ""
    nodes = nodes[["Id", "Label", "Alias", "Reliability", "Birth_Death", "Role"]]

    edges = relations.rename(
        columns={
            "source_person_id": "Source",
            "target_person_id": "Target",
            "standard_relation_type": "Relation_Type",
            "context": "Context",
            "evidence_ref": "Evidence_Ref",
            "weight": "Weight",
        }
    ).copy()
    if "Relation_Type" not in edges.columns and "original_relation_type" in edges.columns:
        edges["Relation_Type"] = edges["original_relation_type"]
    for column in ["Source", "Target", "Relation_Type", "Context", "Evidence_Ref", "Weight"]:
        if column not in edges.columns:
            edges[column] = ""

    if "place_id" in places.columns:
        place_hist_map = places.set_index("place_id")["historical_name"].to_dict()
        place_current_map = places.set_index("place_id")["current_name"].to_dict()
        place_lon_map = places.set_index("place_id")["longitude"].to_dict()
        place_lat_map = places.set_index("place_id")["latitude"].to_dict()
    else:
        place_hist_map = {}
        place_current_map = {}
        place_lon_map = {}
        place_lat_map = {}

    if not participants.empty and {"event_id", "person_id"}.issubset(participants.columns):
        participant_ids = (
            participants.groupby("event_id")["person_id"]
            .apply(lambda values: ";".join(dict.fromkeys([str(v).strip() for v in values if str(v).strip()])))
            .to_dict()
        )
        participant_details: dict[str, list[dict[str, str]]] = {}
        for event_id, group in participants.groupby("event_id", sort=False):
            participant_details[str(event_id)] = [
                {
                    "person_id": str(row.get("person_id", "")).strip(),
                    "name": str(row.get("participant_name", "")).strip(),
                    "relation": str(row.get("participant_role", "")).strip(),
                }
                for _, row in group.iterrows()
                if str(row.get("person_id", "")).strip() or str(row.get("participant_name", "")).strip()
            ]
    else:
        participant_ids = {}
        participant_details = {}

    events = standard_events.copy()
    event_ids = events.get("event_id", pd.Series([""] * len(events), index=events.index))
    events["Event_ID"] = event_ids
    events["Source_IDs"] = events.get("source_ids", "")
    events["Related_Persons"] = event_ids.map(participant_details).apply(
        lambda value: value if isinstance(value, list) else []
    )
    events["Entity_ID"] = event_ids.map(participant_ids).fillna("")
    events["Timestamp"] = events.get("event_date", "")
    events["Hist_Loc"] = events.get("historical_location", "")
    events["Current_Loc"] = events.get("current_address", "")
    events["Longitude"] = events.get("longitude", "")
    events["Latitude"] = events.get("latitude", "")
    if "place_id" in events.columns:
        for col, place_map in [
            ("Hist_Loc", place_hist_map),
            ("Current_Loc", place_current_map),
            ("Longitude", place_lon_map),
            ("Latitude", place_lat_map),
        ]:
            events[col] = _fill_from_place(events, col, place_map)
    events["Event"] = events.get("event_name", "")

    required_columns = [
        "Event_ID",
        "Source_IDs",
        "Related_Persons",
        "Entity_ID",
        "Timestamp",
        "Hist_Loc",
        "Current_Loc",
        "Longitude",
        "Latitude",
        "Event",
    ]
    for column in required_columns:
        if column not in events.columns:
            events[column] = ""
    return nodes, edges, events[required_columns]


def load_legacy_views(data_dir: Path, base_dir: Path) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    nodes_path = data_dir / "nodes.csv"
    events_path = data_dir / "events.csv"
    edges_path = data_dir / "edges_audited.csv"
    if not edges_path.exists():
        edges_path = data_dir / "edges.csv"

    required = [nodes_path, edges_path, events_path]
    if not all(path.exists() for path in required):
        st.error("找不到核心数据文件。已搜索目录：\n" + format_candidate_paths(candidate_data_dirs(base_dir)))
        st.stop()

    nodes = pd.read_csv(nodes_path).fillna("")
    edges = pd.read_csv(edges_path).fillna("")
    events = pd.read_csv(events_path).fillna("")
    return nodes, edges, events


@st.cache_data(show_spinner=False)
def load_data(base_dir: Path | None = None) -> LoadedData:
    resolved_base_dir = (base_dir or Path(__file__).resolve().parent).resolve()
    data_dir = resolve_data_dir(resolved_base_dir)
    if standardized_data_exists(data_dir):
        nodes, edges, events = load_standardized_views(data_dir)
    else:
        nodes, edges, events = load_legacy_views(data_dir, resolved_base_dir)

    nodes = nodes.copy()
    edges = edges.copy()
    events = events.copy()

    nodes["search_text"] = (
        nodes["Label"].astype(str)
        + " "
        + nodes["Alias"].astype(str)
        + " "
        + nodes["Role"].astype(str)
    )
    nodes["Reliability"] = pd.to_numeric(nodes["Reliability"], errors="coerce").fillna(0).astype(int)

    name_map = nodes.set_index("Id")["Label"].to_dict()
    role_map = nodes.set_index("Id")["Role"].to_dict()

    if "raw_relation_type" not in edges.columns:
        if "original_relation_type" in edges.columns:
            edges["raw_relation_type"] = edges["original_relation_type"]
        else:
            edges["raw_relation_type"] = edges["Relation_Type"]
    if "llm_suggested_relation_type" not in edges.columns:
        edges["llm_suggested_relation_type"] = ""
    if "final_relation_type" not in edges.columns:
        edges["final_relation_type"] = edges["Relation_Type"]
    if "llm_reason" not in edges.columns:
        edges["llm_reason"] = ""
    if "llm_confidence" not in edges.columns:
        edges["llm_confidence"] = ""
    if "display_status" not in edges.columns:
        edges["display_status"] = "formal"

    edges["Weight"] = pd.to_numeric(edges["Weight"], errors="coerce").fillna(0)
    if "original_relation_type" in edges.columns:
        edges["raw_relation_type"] = edges["raw_relation_type"].where(
            edges["raw_relation_type"].astype(str).str.strip() != "",
            edges["original_relation_type"],
        )
    edges["raw_relation_type"] = edges["raw_relation_type"].where(
        edges["raw_relation_type"].astype(str).str.strip() != "",
        edges["Relation_Type"],
    )
    edges["final_relation_type"] = edges["final_relation_type"].where(
        edges["final_relation_type"].astype(str).str.strip() != "",
        edges["Relation_Type"],
    )
    edges["Relation_Type"] = edges["final_relation_type"].where(
        edges["final_relation_type"].astype(str).str.strip() != "",
        "未标注",
    )
    edges["LLM_Confidence"] = pd.to_numeric(edges["llm_confidence"], errors="coerce")
    edges["Display_Status"] = edges["display_status"].replace("", "formal")
    edges["Relation_Family"] = edges["Relation_Type"].astype(str).str.replace(r"^(强关联-|弱关联-)", "", regex=True)
    edges["Context_Preview"] = edges["Context"].apply(clean_text)
    edges["Evidence_Preview"] = edges["Evidence_Ref"].apply(lambda value: clean_text(value, 70))
    edges["Source_Name"] = edges["Source"].map(name_map).fillna(edges["Source"])
    edges["Target_Name"] = edges["Target"].map(name_map).fillna(edges["Target"])
    edges["Source_Role"] = edges["Source"].map(role_map).fillna("")
    edges["Target_Role"] = edges["Target"].map(role_map).fillna("")

    events["Longitude"] = pd.to_numeric(events["Longitude"], errors="coerce")
    events["Latitude"] = pd.to_numeric(events["Latitude"], errors="coerce")
    generated_event_ids = [f"EVT-{index:04d}" for index in range(len(events))]
    events["Event_ID"] = events["Event_ID"].where(
        events["Event_ID"].astype(str).str.strip() != "",
        generated_event_ids,
    )
    events["Datetime"] = pd.to_datetime(events["Timestamp"], errors="coerce", format="mixed")
    events["Year"] = events["Datetime"].dt.year
    events["Subjects"] = events["Entity_ID"].apply(lambda value: "、".join(name_map.get(item, item) for item in split_ids(value)))
    events["search_text"] = (
        events["Timestamp"].astype(str)
        + " "
        + events["Event"].astype(str)
        + " "
        + events["Hist_Loc"].astype(str)
        + " "
        + events["Current_Loc"].astype(str)
        + " "
        + events["Subjects"].astype(str)
    )

    return LoadedData(nodes=nodes, edges=edges, events=events, data_dir=str(data_dir))
