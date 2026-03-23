from __future__ import annotations

import math
from html import escape
from pathlib import Path

import folium
import pandas as pd
import streamlit as st
from streamlit_folium import st_folium

from historical_map import HistoricalEvent, build_historical_events, events_to_frame
from styles import ACCENT, BORDER, CHART_FONT, INK, MUTED, PAPER_LIGHT, PRIMARY


@st.cache_data(show_spinner=False)
def load_historical_map_bundle(
    data_dir: Path,
    nodes_df: pd.DataFrame,
    events_df: pd.DataFrame,
) -> tuple[pd.DataFrame, dict[str, HistoricalEvent], dict[str, object]]:
    events, geojson = build_historical_events(Path(__file__).resolve().parent, data_dir, nodes_df, events_df)
    frame = events_to_frame(events)
    index = {event.id: event for event in events}
    return frame, index, geojson


def event_stage_options() -> dict[str, tuple[int, int] | None]:
    return {
        "全部阶段": None,
        "萌芽与汇流（1922-1927）": (1922, 1927),
        "组织化推进（1928-1931）": (1928, 1931),
        "扩散与转折（1932-1934）": (1932, 1934),
        "战时前夕（1935-1940）": (1935, 1940),
    }


def filter_map_events(
    event_frame: pd.DataFrame,
    person_name: str,
    category: str,
    stage_label: str,
    year_choice: str,
    keyword: str,
) -> pd.DataFrame:
    display = event_frame.copy()
    if person_name != "全部人物":
        display = display[display["people"].apply(lambda values: person_name in values)]
    if category != "全部事件":
        display = display[display["category"] == category]

    stage_range = event_stage_options().get(stage_label)
    if stage_range:
        display = display[display["year"].fillna(stage_range[0]).between(stage_range[0], stage_range[1])]
    if year_choice != "全部年份":
        display = display[display["year"] == int(year_choice)]
    if keyword:
        display = display[
            display["title"].astype(str).str.contains(keyword, case=False, regex=False)
            | display["location_name"].astype(str).str.contains(keyword, case=False, regex=False)
            | display["people_label"].astype(str).str.contains(keyword, case=False, regex=False)
            | display["summary"].astype(str).str.contains(keyword, case=False, regex=False)
        ]
    return display.sort_values(["year", "date", "title"], na_position="last")


def _prepare_map_points(map_df: pd.DataFrame) -> pd.DataFrame:
    mapped = map_df.dropna(subset=["longitude", "latitude"]).copy()
    if mapped.empty:
        return mapped

    mapped = mapped.sort_values(["year", "date", "id"], na_position="last").reset_index(drop=True)
    mapped["map_lat"] = mapped["latitude"].astype(float)
    mapped["map_lon"] = mapped["longitude"].astype(float)

    for _, group in mapped.groupby(["latitude", "longitude"], sort=False):
        if len(group) <= 1:
            continue
        total = len(group)
        for position, row_index in enumerate(group.index):
            angle = (2 * math.pi * position) / total
            mapped.at[row_index, "map_lon"] = float(mapped.at[row_index, "longitude"]) + math.cos(angle) * 0.0013
            mapped.at[row_index, "map_lat"] = float(mapped.at[row_index, "latitude"]) + math.sin(angle) * 0.0010
    return mapped


def build_historical_event_leaflet_map(
    map_df: pd.DataFrame,
    selected_event_id: str | None,
    zoom_start: int = 12,
) -> tuple[folium.Map, pd.DataFrame]:
    mapped = _prepare_map_points(map_df)
    center_lat = float(mapped["latitude"].mean()) if not mapped.empty else 31.2304
    center_lon = float(mapped["longitude"].mean()) if not mapped.empty else 121.4737

    event_map = folium.Map(
        location=[center_lat, center_lon],
        zoom_start=zoom_start,
        control_scale=True,
        tiles=None,
        prefer_canvas=True,
    )
    folium.TileLayer(
        tiles="https://webrd0{s}.is.autonavi.com/appmaptile?lang=zh_cn&size=1&scale=1&style=8&x={x}&y={y}&z={z}",
        attr="高德地图",
        name="街道底图",
        subdomains=["1", "2", "3", "4"],
        max_zoom=18,
    ).add_to(event_map)

    if not mapped.empty:
        south = float(mapped["latitude"].min()) - 0.012
        north = float(mapped["latitude"].max()) + 0.012
        west = float(mapped["longitude"].min()) - 0.016
        east = float(mapped["longitude"].max()) + 0.016
        event_map.fit_bounds([[south, west], [north, east]])

        if len(mapped) > 1:
            folium.PolyLine(
                locations=mapped[["latitude", "longitude"]].values.tolist(),
                color=PRIMARY,
                weight=2.2,
                opacity=0.55,
                dash_array="6 7",
            ).add_to(event_map)

        for _, row in mapped.iterrows():
            marker_color = PRIMARY if row["evidence_count"] else ACCENT
            radius = 8 + min(int(row["evidence_count"]), 3)
            tooltip = (
                f"{row['title']}<br>"
                f"时间：{row['date']}<br>"
                f"地点：{row['location_name']}<br>"
                f"相关人物：{row['people_label']}"
            )
            popup = folium.Popup(
                (
                    f"<div style='font-family:{CHART_FONT};line-height:1.6;'>"
                    f"<strong>{escape(str(row['title']))}</strong><br>"
                    f"{escape(str(row['date']))}<br>"
                    f"{escape(str(row['location_name']))}</div>"
                ),
                max_width=280,
            )
            folium.CircleMarker(
                location=[float(row["map_lat"]), float(row["map_lon"])],
                radius=radius + 2 if row["id"] == selected_event_id else radius,
                color=PAPER_LIGHT,
                weight=1.2,
                fill=True,
                fill_color=marker_color,
                fill_opacity=0.92,
                tooltip=tooltip,
                popup=popup,
            ).add_to(event_map)

            if row["id"] == selected_event_id:
                folium.CircleMarker(
                    location=[float(row["map_lat"]), float(row["map_lon"])],
                    radius=radius + 8,
                    color=PRIMARY,
                    weight=1.2,
                    fill=True,
                    fill_color=PRIMARY,
                    fill_opacity=0.08,
                ).add_to(event_map)
                folium.Marker(
                    location=[float(row["map_lat"]) + 0.0012, float(row["map_lon"])],
                    icon=folium.DivIcon(
                        html=(
                            f"<div style=\"font-family:{CHART_FONT};font-size:12px;color:{PRIMARY};"
                            "font-weight:600;white-space:nowrap;text-shadow:0 0 2px #f7f0e2;\">"
                            f"{escape(str(row['title']))}</div>"
                        )
                    ),
                ).add_to(event_map)

    folium.Element(
        f"""
        <div style="
            position:absolute; left:14px; top:14px; z-index:9999;
            padding:10px 12px; border:1px solid {BORDER};
            background:rgba(247,240,226,.9); box-shadow:none;
            font:13px/1.55 {CHART_FONT}; color:{INK};">
            <div style="color:{PRIMARY}; font-weight:600; margin-bottom:2px;">上海活动区历史事件地图</div>
            <div style="color:{MUTED};">作为关系网络的空间背景层，保留点位、时序与史料入口。</div>
        </div>
        """
    ).add_to(event_map.get_root().html)
    return event_map, mapped


def render_event_map(
    map_df: pd.DataFrame,
    geojson: dict[str, object],
    key_prefix: str,
    height: int = 560,
) -> str | None:
    _ = geojson
    if map_df.empty:
        st.info("当前条件下没有可展示的事件地图记录。")
        return None

    selected_key = f"{key_prefix}_selected_event"
    available_ids = map_df["id"].tolist()
    if st.session_state.get(selected_key) not in available_ids:
        st.session_state[selected_key] = available_ids[0]
    selected_event_id = st.session_state.get(selected_key)

    coord_df = map_df.dropna(subset=["longitude", "latitude"]).copy()
    unmatched_count = int((map_df["region_match_status"] == "unmatched").sum())
    if coord_df.empty:
        st.info("当前条件下暂无可定位坐标，无法绘制地图。")
        return selected_event_id
    if unmatched_count:
        st.caption(f"{unmatched_count} 条事件尚未匹配到底图区域，当前采用坐标或近邻方式落点。")

    event_map, plotted_df = build_historical_event_leaflet_map(coord_df, selected_event_id)
    map_state = st_folium(
        event_map,
        key=f"{key_prefix}_folium",
        height=height,
        use_container_width=True,
        returned_objects=["last_object_clicked"],
    )
    clicked_object = map_state.get("last_object_clicked") if map_state else None
    if clicked_object and not plotted_df.empty:
        lat = float(clicked_object.get("lat", 0))
        lon = float(clicked_object.get("lng", 0))
        distances = (plotted_df["map_lat"] - lat).pow(2) + (plotted_df["map_lon"] - lon).pow(2)
        nearest_index = distances.idxmin()
        clicked_id = str(plotted_df.loc[nearest_index, "id"]) if float(distances.loc[nearest_index]) < 0.00002 else None
        if clicked_id and clicked_id in available_ids and clicked_id != st.session_state.get(selected_key):
            st.session_state[selected_key] = clicked_id
            st.rerun()
    return st.session_state.get(selected_key)


def render_event_detail(event: HistoricalEvent | None, title: str = "事件详情") -> None:
    st.markdown(f"#### {title}")
    if event is None:
        st.info("请先在地图中点击一个事件点位。")
        return

    st.markdown(
        f"""
        <div class="relation-detail-head">
            <div class="relation-detail-title">{escape(event.title)}</div>
            <div class="relation-detail-subtitle">
                时间：{escape(event.date)}<br>
                地点：{escape(event.location_name)}｜区域：{escape(event.map_region)}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    meta_a, meta_b = st.columns(2)
    related_persons = event.related_persons or []
    evidences = event.evidences or []
    meta_a.metric("相关人物", len(related_persons))
    meta_b.metric("证据条数", len(evidences))

    st.markdown("**相关人物**")
    st.write("、".join(person.name for person in related_persons if person.name) if related_persons else "人物待补")
    st.markdown("**事件类型**")
    st.write(event.category or "未标注")
    st.markdown("**简要说明**")
    st.write(event.summary or "暂无说明")
    st.markdown("**该事件在左联网络中的意义**")
    st.write(event.significance or "暂无说明")

    st.markdown("**史料摘录 / 证据摘要**")
    if not evidences:
        st.info("暂无可回溯证据")
        return

    for index, evidence in enumerate(evidences, start=1):
        loc_text = evidence.source_loc or "位置待补"
        with st.expander(f"证据 {index}｜{evidence.source or '来源待补'}｜{loc_text}", expanded=index == 1):
            confidence_line = ""
            if evidence.confidence is not None:
                confidence_line = f"<br>置信度：{evidence.confidence:.2f}"
            st.markdown(
                f"""
                <div class="evidence-summary">
                    <div class="evidence-meta">
                        来源：{escape(evidence.source or "来源待补")}<br>
                        位置：{escape(loc_text)}<br>
                        文件：{escape(evidence.source_file or "待补")}{confidence_line}
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            if evidence.quote:
                st.markdown(
                    f"<div class='excerpt-block'>{escape(evidence.quote).replace(chr(10), '<br>')}</div>",
                    unsafe_allow_html=True,
                )
            else:
                st.info("暂无史料摘录")


def render_events(
    nodes_df: pd.DataFrame,
    historical_event_frame: pd.DataFrame,
    historical_event_index: dict[str, HistoricalEvent],
    historical_geojson: dict[str, object],
) -> None:
    st.markdown(
        '<div class="page-note">地图页作为关系网络的空间延展层，重点回答左联人物活动发生在哪里、如何迁移，以及哪些空间节点承载了关系形成与事件爆发。</div>',
        unsafe_allow_html=True,
    )
    if historical_event_frame.empty:
        st.info("当前没有可用于地图展示的事件数据。")
        return

    year_values = historical_event_frame["year"].dropna().astype(int)
    timeline_options = ["全部年份"] + [str(year) for year in sorted(year_values.unique())]
    category_options = ["全部事件"] + sorted(
        historical_event_frame["category"].fillna("未标注").astype(str).unique().tolist()
    )

    pcol, ccol, scol = st.columns([1.0, 1.0, 1.0])
    with pcol:
        person = st.selectbox("按人物筛选", ["全部人物"] + sorted(nodes_df["Label"].tolist()))
    with ccol:
        event_type = st.selectbox("按事件类型筛选", category_options)
    with scol:
        stage_label = st.selectbox("历史阶段", list(event_stage_options().keys()))

    year_choice = st.select_slider("时间轴", options=timeline_options, value="全部年份")
    keyword = st.text_input("检索事件", placeholder="按事件、地点、人物或史料摘要搜索")

    filtered_map_df = filter_map_events(
        historical_event_frame,
        person_name=person,
        category=event_type,
        stage_label=stage_label,
        year_choice=year_choice,
        keyword=keyword,
    )

    if filtered_map_df.empty:
        st.info("当前筛选条件下暂无事件记录。")
        return

    related_people = {person_name for values in filtered_map_df["people"] for person_name in values}
    summary_a, summary_b, summary_c, summary_d = st.columns(4)
    summary_a.metric("可见事件", len(filtered_map_df))
    summary_b.metric("覆盖区域", int(filtered_map_df["map_region"].nunique()))
    summary_c.metric("涉及人物", len(related_people))
    summary_d.metric("史料条数", int(filtered_map_df["evidence_count"].sum()))

    st.caption("切换年份、阶段或人物后，地图点位与右侧证据详情会同步刷新。")

    left, right = st.columns([1.2, 0.8])
    with left:
        selected_event_id = render_event_map(
            filtered_map_df,
            historical_geojson,
            key_prefix="events_map",
            height=620,
        )
        st.dataframe(
            filtered_map_df[
                ["date", "title", "location_name", "map_region", "people_label", "category", "evidence_count"]
            ].rename(
                columns={
                    "date": "时间",
                    "title": "事件标题",
                    "location_name": "地点",
                    "map_region": "区域",
                    "people_label": "相关人物",
                    "category": "事件类型",
                    "evidence_count": "证据条数",
                }
            ),
            width="stretch",
            hide_index=True,
        )
    with right:
        render_event_detail(historical_event_index.get(selected_event_id), title="事件详情")
        unmatched = filtered_map_df[filtered_map_df["region_match_status"] == "unmatched"]
        if not unmatched.empty:
            st.caption(f"区域匹配失败 {len(unmatched)} 条：当前保留点位展示，后续可补正式边界或历史地名映射。")
        no_coord = filtered_map_df["longitude"].isna() | filtered_map_df["latitude"].isna()
        if int(no_coord.sum()) > 0:
            st.caption(f"无坐标事件 {int(no_coord.sum())} 条：补充经纬度后可进入底图展示。")
