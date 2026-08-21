from __future__ import annotations

import argparse
import json
import shutil
from collections import Counter, defaultdict
from html import escape
from pathlib import Path
from typing import Iterable

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parent
DATA_DIR = PROJECT_ROOT / "data" / "publish"
DOCS_DIR = PROJECT_ROOT / "docs"
APP_ASSETS_DIR = PROJECT_ROOT / "app" / "frontend" / "assets"
STATIC_SITE_DIR = PROJECT_ROOT / "static_site"

ASSET_FILES = (
    "banner.png",
    "historical_archive_bg.svg",
    "paper_texture.png",
    "stamp.png",
    "zuolian_relationship_bg.svg",
    "site.css",
    "site.js",
)


def text(value: object, fallback: str = "", limit: int = 0) -> str:
    if value is None:
        return fallback
    if isinstance(value, float) and pd.isna(value):
        return fallback
    cleaned = " ".join(str(value).split())
    if not cleaned:
        return fallback
    if limit and len(cleaned) > limit:
        return f"{cleaned[:limit].rstrip()}..."
    return cleaned


def split_ids(value: object) -> list[str]:
    raw = text(value)
    if not raw:
        return []
    normalized = raw.replace("；", ";").replace("、", ";")
    return [item.strip() for item in normalized.split(";") if item.strip()]


def unique_ordered(values: Iterable[str]) -> list[str]:
    seen: set[str] = set()
    ordered: list[str] = []
    for value in values:
        item = text(value)
        if not item or item in seen:
            continue
        seen.add(item)
        ordered.append(item)
    return ordered


def excerpt(value: object, limit: int = 120, fallback: str = "暂无摘录") -> str:
    cleaned = text(value)
    if not cleaned:
        return fallback
    if len(cleaned) <= limit:
        return cleaned
    return f"{cleaned[:limit].rstrip()}..."


def extract_year(value: object) -> int | None:
    cleaned = text(value)
    for chunk in cleaned.replace("年", "-").replace("/", "-").split("-"):
        piece = chunk.strip()
        if len(piece) == 4 and piece.isdigit():
            return int(piece)
    return None


def select_featured_events(events: list[dict[str, object]], limit: int = 6) -> list[dict[str, object]]:
    buckets = [
        lambda event: text(event.get("date_precision")) in {"日", "月"} and text(event.get("confidence")) != "low",
        lambda event: text(event.get("date_precision")) in {"日", "月"},
        lambda event: text(event.get("needs_manual_review")) != "yes",
        lambda event: True,
    ]
    featured: list[dict[str, object]] = []
    seen: set[str] = set()
    for matcher in buckets:
        for event in events:
            event_id = text(event.get("id"))
            if not event_id or event_id in seen or not matcher(event):
                continue
            featured.append(event)
            seen.add(event_id)
            if len(featured) >= limit:
                return featured
    return featured


def pair_key(person_a_id: str, person_b_id: str) -> str:
    left, right = sorted([person_a_id, person_b_id])
    return f"{left}|{right}"


def pair_anchor(person_a_id: str, person_b_id: str) -> str:
    left, right = sorted([person_a_id.lower(), person_b_id.lower()])
    return f"pair-{left}-{right}"


def status_label(status_counts: Counter[str]) -> str:
    parts: list[str] = []
    mapping = {
        "formal": "正式证据",
        "review": "待复核",
        "hidden": "隐藏记录",
    }
    for key in ("formal", "review", "hidden"):
        count = int(status_counts.get(key, 0))
        if count:
            parts.append(f"{mapping[key]} {count}")
    for key, count in status_counts.items():
        if key in mapping or not count:
            continue
        parts.append(f"{key} {count}")
    return " / ".join(parts) if parts else "状态未标注"


def metric_card(label: str, value: str, note: str) -> str:
    return (
        '<article class="metric-card">'
        f'<div class="metric-card__label">{escape(label)}</div>'
        f'<div class="metric-card__value">{escape(value)}</div>'
        f'<p class="metric-card__note">{escape(note)}</p>'
        "</article>"
    )


def badge_items(items: Iterable[str]) -> str:
    badges = [f'<span class="badge">{escape(item)}</span>' for item in unique_ordered(items) if text(item)]
    return "".join(badges) or '<span class="badge">信息待补</span>'


def source_reference_items(source_ids: list[str], source_map: dict[str, dict[str, str]], limit: int = 14) -> list[dict[str, str]]:
    references: list[dict[str, str]] = []
    seen: set[str] = set()
    for source_id in source_ids:
        if source_id in seen:
            continue
        seen.add(source_id)
        record = source_map.get(source_id)
        if not record:
            continue
        references.append(record)
        if len(references) >= limit:
            break
    return references


def source_reference_list(source_ids: list[str], source_map: dict[str, dict[str, str]], limit: int = 14) -> str:
    references = source_reference_items(source_ids, source_map, limit=limit)
    if not references:
        return '<p class="empty-state">当前页面没有可公开展示的来源卡片。</p>'
    items: list[str] = []
    for record in references:
        title = escape(record["title"])
        citation = escape(record["citation"])
        kind = escape(record["kind"])
        source_id = escape(record["id"])
        if record["url"]:
            title_html = f'<a href="{escape(record["url"], quote=True)}" target="_blank" rel="noreferrer">{title}</a>'
        else:
            title_html = title
        citation_html = f'<p class="source-list__citation">{citation}</p>' if citation else ""
        items.append(
            '<li class="source-list__item">'
            f'<div class="source-list__title">{title_html}</div>'
            f'{citation_html}'
            f'<div class="source-list__meta">{source_id} · {kind}</div>'
            "</li>"
        )
    return f'<ul class="source-list">{"".join(items)}</ul>'


def page_shell(
    *,
    title: str,
    description: str,
    active_nav: str,
    body: str,
    depth: int = 0,
) -> str:
    prefix = "../" * depth
    nav_items = [
        ("home", "首页", f"{prefix}index.html"),
        ("people", "人物档案", f"{prefix}people/index.html"),
        ("events", "事件索引", f"{prefix}events/index.html"),
        ("relations", "关系索引", f"{prefix}relations/index.html"),
        ("search", "全文搜索", f"{prefix}search/index.html"),
        ("graph", "关系图谱", f"{prefix}graph/index.html"),
        ("timeline", "事件时间轴", f"{prefix}timeline/index.html"),
    ]
    nav_html = "".join(
        (
            f'<a class="site-nav__link{" is-active" if key == active_nav else ""}" '
            f'href="{escape(href, quote=True)}">{escape(label)}</a>'
        )
        for key, label, href in nav_items
    )
    return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{escape(title)}</title>
  <meta name="description" content="{escape(description, quote=True)}">
  <link rel="stylesheet" href="{prefix}assets/site.css">
  <script defer src="{prefix}assets/site.js"></script>
</head>
<body>
  <a class="skip-link" href="#main">跳到正文</a>
  <header class="site-header">
    <div class="site-header__inner">
      <a class="brand" href="{prefix}index.html">
        <span class="brand__eyebrow">STATIC READER</span>
        <span class="brand__title">左联知识库</span>
      </a>
      <nav class="site-nav" aria-label="主导航">{nav_html}</nav>
      <button class="site-nav-toggle" aria-label="展开导航" onclick="this.closest('.site-header__inner').querySelector('.site-nav').classList.toggle('is-open')">☰</button>
    </div>
  </header>
  <main id="main" class="page-main">
    <div class="page-shell">{body}</div>
  </main>
  <footer class="site-footer">
    <div class="site-footer__inner">
      <p>静态阅读版基于仓库内标准知识库数据自动生成，适合 GitHub Pages 发布。</p>
      <p>复杂分析与交互地图已内置于本站，支持关系图谱与事件时间轴在线浏览。</p>
    </div>
  </footer>
</body>
</html>
"""


def person_card(person: dict[str, object], href: str) -> str:
    aliases = text(person["aliases"], "别名待补")
    years = text(person["birth_death"], "生卒待补")
    role = text(person["role"], "角色待补")
    relation_total = f"{int(person['pair_total'])} 位关联人物"
    event_total = f"{int(person['event_total'])} 条关联事件"
    badges = badge_items([role, years, f"可靠度 {person['reliability']}"])
    return (
        f'<article class="index-card" data-search="{escape(person["search_blob"], quote=True)}">'
        f'<a class="index-card__link" href="{escape(href, quote=True)}">'
        f'<div class="index-card__header"><h3>{escape(text(person["name"]))}</h3><span>{escape(role)}</span></div>'
        f'<p class="index-card__summary">{escape(aliases)}</p>'
        f'<div class="index-card__badges">{badges}</div>'
        f'<div class="index-card__meta">{escape(relation_total)} · {escape(event_total)}</div>'
        "</a></article>"
    )


def event_card(event: dict[str, object], href: str) -> str:
    participants = "、".join([item["name"] for item in event["participants"][:4]]) or "参与人物待补"
    location = text(event["historical_location"] or event["current_address"], "地点待补")
    badges = badge_items(
        [
            text(event["date"], "时间待补"),
            text(event["date_precision"], "精度待补"),
            f"{len(event['participants'])} 位相关人物",
        ]
    )
    return (
        f'<article class="index-card" data-search="{escape(event["search_blob"], quote=True)}">'
        f'<a class="index-card__link" href="{escape(href, quote=True)}">'
        f'<div class="index-card__header"><h3>{escape(text(event["name"]))}</h3><span>{escape(text(event["date"], "时间待补"))}</span></div>'
        f'<p class="index-card__summary">{escape(location)} · {escape(participants)}</p>'
        f'<div class="index-card__badges">{badges}</div>'
        f'<div class="index-card__meta">{escape(excerpt(event["note"], 78, "暂无备注"))}</div>'
        "</a></article>"
    )


def relation_card(relation: dict[str, object], person_prefix: str) -> str:
    left_href = f"{person_prefix}{relation['person_a_id']}.html"
    right_href = f"{person_prefix}{relation['person_b_id']}.html"
    title = (
        f'<a href="{escape(left_href, quote=True)}">{escape(relation["person_a_name"])}</a>'
        " × "
        f'<a href="{escape(right_href, quote=True)}">{escape(relation["person_b_name"])}</a>'
    )
    meta = f"{relation['count']} 条关系记录 · {relation['status_text']}"
    summary = " / ".join(relation["types"][:4])
    details = relation["evidences"][:2]
    detail_text = "；".join(details) if details else relation["context_preview"]
    return (
        f'<article class="relation-card" id="{escape(relation["anchor"], quote=True)}" '
        f'data-search="{escape(relation["search_blob"], quote=True)}">'
        f'<h3 class="relation-card__title">{title}</h3>'
        f'<p class="relation-card__summary">{escape(summary)}</p>'
        f'<div class="relation-card__meta">{escape(meta)}</div>'
        f'<p class="relation-card__excerpt">{escape(detail_text)}</p>'
        "</article>"
    )


def timeline_items(events: list[dict[str, object]], href_prefix: str) -> str:
    if not events:
        return '<p class="empty-state">当前人物尚未关联到可展示的事件时间线。</p>'
    items = []
    for event in events:
        href = f"{href_prefix}{event['id']}.html"
        location = text(event["historical_location"] or event["current_address"], "地点待补")
        people = "、".join(participant["name"] for participant in event["participants"][:4]) or "参与人物待补"
        items.append(
            '<li class="timeline__item">'
            f'<div class="timeline__date">{escape(text(event["date"], "时间待补"))}</div>'
            '<div class="timeline__body">'
            f'<h3><a href="{escape(href, quote=True)}">{escape(text(event["name"]))}</a></h3>'
            f'<p>{escape(location)} · {escape(people)}</p>'
            "</div></li>"
        )
    return f'<ol class="timeline">{"".join(items)}</ol>'


def write_text(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")


def load_frame(filename: str) -> pd.DataFrame:
    return pd.read_csv(DATA_DIR / filename, encoding="utf-8-sig").fillna("")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="构建左联知识库静态阅读站点。")
    parser.add_argument("--data-dir", type=Path, default=DATA_DIR, help="标准数据目录")
    parser.add_argument("--output-dir", type=Path, default=DOCS_DIR, help="静态站输出目录")
    return parser.parse_args()


def cleanup_output() -> None:
    DOCS_DIR.mkdir(exist_ok=True)
    generated_paths = [
        DOCS_DIR / "assets",
        DOCS_DIR / "people",
        DOCS_DIR / "events",
        DOCS_DIR / "relations",
        DOCS_DIR / "search",
        DOCS_DIR / "index.html",
        DOCS_DIR / ".nojekyll",
    ]
    for path in generated_paths:
        if path.is_dir():
            shutil.rmtree(path)
        elif path.exists():
            path.unlink()


def copy_assets() -> None:
    assets_dir = DOCS_DIR / "assets"
    assets_dir.mkdir(parents=True, exist_ok=True)
    asset_sources = {
        "banner.png": APP_ASSETS_DIR / "banner.png",
        "historical_archive_bg.svg": APP_ASSETS_DIR / "historical_archive_bg.svg",
        "paper_texture.png": APP_ASSETS_DIR / "paper_texture.png",
        "stamp.png": APP_ASSETS_DIR / "stamp.png",
        "zuolian_relationship_bg.svg": APP_ASSETS_DIR / "zuolian_relationship_bg.svg",
        "site.css": STATIC_SITE_DIR / "site.css",
        "site.js": STATIC_SITE_DIR / "site.js",
    }
    for name in ASSET_FILES:
        shutil.copy2(asset_sources[name], assets_dir / name)
    write_text(DOCS_DIR / ".nojekyll", "")


def build_sources_lookup(sources_df: pd.DataFrame) -> dict[str, dict[str, str]]:
    source_map: dict[str, dict[str, str]] = {}
    for row in sources_df.to_dict("records"):
        source_id = text(row.get("source_id"))
        if not source_id:
            continue
        url = text(row.get("source_url"))
        if not url.startswith(("http://", "https://")):
            url = ""
        citation = text(row.get("citation"))
        title = text(row.get("title")) or citation or source_id
        source_map[source_id] = {
            "id": source_id,
            "title": title,
            "citation": citation,
            "url": url,
            "kind": text(row.get("source_kind"), "source"),
        }
    return source_map


def build_relation_profiles(relations_df: pd.DataFrame, name_map: dict[str, str]) -> list[dict[str, object]]:
    profiles: dict[str, dict[str, object]] = {}
    for row in relations_df.to_dict("records"):
        source_id = text(row.get("source_person_id"))
        target_id = text(row.get("target_person_id"))
        if not source_id or not target_id:
            continue
        key = pair_key(source_id, target_id)
        profile = profiles.setdefault(
            key,
            {
                "person_a_id": min(source_id, target_id),
                "person_b_id": max(source_id, target_id),
                "person_a_name": name_map.get(min(source_id, target_id), min(source_id, target_id)),
                "person_b_name": name_map.get(max(source_id, target_id), max(source_id, target_id)),
                "count": 0,
                "total_weight": 0.0,
                "types": [],
                "contexts": [],
                "evidences": [],
                "source_ids": [],
                "status_counts": Counter(),
            },
        )
        profile["count"] += 1
        profile["total_weight"] += float(row.get("weight") or 0)
        profile["types"].append(text(row.get("final_relation_type") or row.get("standard_relation_type") or row.get("original_relation_type"), "未标注"))
        context_text = excerpt(row.get("context"), 140, "")
        if context_text:
            profile["contexts"].append(context_text)
        profile["evidences"].extend(split_ids(row.get("evidence_ref")))
        profile["source_ids"].extend(split_ids(row.get("source_ids")))
        profile["status_counts"][text(row.get("display_status"), "formal")] += 1

    relation_profiles: list[dict[str, object]] = []
    for profile in profiles.values():
        profile["types"] = unique_ordered(profile["types"])
        profile["contexts"] = unique_ordered(profile["contexts"])
        profile["evidences"] = unique_ordered(profile["evidences"])
        profile["source_ids"] = unique_ordered(profile["source_ids"])
        profile["status_text"] = status_label(profile["status_counts"])
        profile["context_preview"] = profile["contexts"][0] if profile["contexts"] else "暂无语境摘录"
        profile["anchor"] = pair_anchor(profile["person_a_id"], profile["person_b_id"])
        profile["search_blob"] = " ".join(
            [
                text(profile["person_a_name"]),
                text(profile["person_b_name"]),
                " ".join(profile["types"]),
                " ".join(profile["evidences"][:4]),
                profile["context_preview"],
            ]
        ).lower()
        relation_profiles.append(profile)
    relation_profiles.sort(
        key=lambda item: (
            -int(item["count"]),
            -float(item["total_weight"]),
            text(item["person_a_name"]),
            text(item["person_b_name"]),
        )
    )
    return relation_profiles


def build_event_records(
    events_df: pd.DataFrame,
    participants_df: pd.DataFrame,
    name_map: dict[str, str],
) -> list[dict[str, object]]:
    participants_by_event: dict[str, list[dict[str, object]]] = defaultdict(list)
    for row in participants_df.to_dict("records"):
        event_id = text(row.get("event_id"))
        if not event_id:
            continue
        person_id = text(row.get("person_id"))
        participant_name = text(row.get("participant_name")) or name_map.get(person_id, person_id) or "未标注人物"
        participants_by_event[event_id].append(
            {
                "person_id": person_id,
                "name": participant_name,
                "role": text(row.get("participant_role"), "待核"),
                "source_ids": split_ids(row.get("source_ids")),
            }
        )

    event_records: list[dict[str, object]] = []
    for row in events_df.to_dict("records"):
        event_id = text(row.get("event_id"))
        if not event_id:
            continue
        participants = participants_by_event.get(event_id, [])
        historical_location = text(row.get("historical_location"))
        current_address = text(row.get("current_address"))
        event = {
            "id": event_id,
            "name": text(row.get("event_name"), event_id),
            "original_names": text(row.get("original_event_names")),
            "date": text(row.get("event_date"), "时间待补"),
            "date_precision": text(row.get("date_precision"), "精度待补"),
            "event_scope": text(row.get("event_scope"), ""),
            "canonical_event_key": text(row.get("canonical_event_key"), ""),
            "historical_location": historical_location,
            "current_address": current_address,
            "year": extract_year(row.get("event_date")),
            "participants": participants,
            "source_ids": unique_ordered(split_ids(row.get("source_ids"))),
            "note": text(row.get("display_note")) or text(row.get("correction_reason"), "暂无备注"),
            "internal_note": text(row.get("correction_reason"), ""),
            "confidence": text(row.get("confidence"), "未标注"),
            "needs_manual_review": text(row.get("needs_manual_review"), "no"),
        }
        event["search_blob"] = " ".join(
            [
                event["name"],
                event["date"],
                historical_location,
                current_address,
                " ".join(item["name"] for item in participants),
                event["note"],
            ]
        ).lower()
        event_records.append(event)
    event_records.sort(key=lambda item: (item["year"] or 9999, item["date"], item["name"]))
    return event_records


def build_person_records(
    persons_df: pd.DataFrame,
    relation_profiles: list[dict[str, object]],
    event_records: list[dict[str, object]],
    source_map: dict[str, dict[str, str]],
) -> list[dict[str, object]]:
    relation_map: dict[str, list[dict[str, object]]] = defaultdict(list)
    for relation in relation_profiles:
        relation_map[relation["person_a_id"]].append(relation)
        relation_map[relation["person_b_id"]].append(relation)

    event_map: dict[str, list[dict[str, object]]] = defaultdict(list)
    for event in event_records:
        for participant in event["participants"]:
            person_id = text(participant.get("person_id"))
            if person_id:
                event_map[person_id].append(event)

    people: list[dict[str, object]] = []
    for row in persons_df.to_dict("records"):
        person_id = text(row.get("person_id"))
        if not person_id:
            continue
        related_pairs = relation_map.get(person_id, [])
        related_events = event_map.get(person_id, [])
        primary_source_ids = split_ids(row.get("source_ids"))
        source_ids: list[str] = []
        source_ids.extend(primary_source_ids)
        for relation in related_pairs[:10]:
            source_ids.extend(relation["source_ids"][:4])
        for event in related_events[:10]:
            source_ids.extend(event["source_ids"][:4])
        person = {
            "id": person_id,
            "name": text(row.get("standard_name"), person_id),
            "aliases": text(row.get("aliases"), "别名待补"),
            "birth_death": text(row.get("birth_death"), "生卒待补"),
            "role": text(row.get("role"), "角色待补"),
            "reliability": text(row.get("reliability"), "0"),
            "pair_total": len(related_pairs),
            "relation_total": sum(int(item["count"]) for item in related_pairs),
            "event_total": len(related_events),
            "related_pairs": related_pairs[:20],
            "related_events": related_events[:12],
            "source_ids": unique_ordered(source_ids),
        }
        person["search_blob"] = " ".join(
            [
                person["name"],
                person["aliases"],
                person["birth_death"],
                person["role"],
                " ".join(
                    text(item["person_b_name"] if item["person_a_id"] == person_id else item["person_a_name"])
                    for item in related_pairs[:8]
                ),
            ]
        ).lower()
        person["source_list_html"] = source_reference_list(person["source_ids"], source_map, limit=16)
        people.append(person)
    people.sort(key=lambda item: (-int(item["pair_total"]), item["name"]))
    return people


def render_home_page(
    people: list[dict[str, object]],
    relations: list[dict[str, object]],
    events: list[dict[str, object]],
    sources_count: int,
) -> str:
    top_people = "".join(person_card(person, f"people/{person['id']}.html") for person in people[:6])
    top_relations = "".join(relation_card(relation, "people/") for relation in relations[:8])
    featured_events = "".join(event_card(event, f"events/{event['id']}.html") for event in select_featured_events(events))

    role_counter = Counter(text(person["role"], "角色待补") for person in people)
    relation_type_counter = Counter()
    for relation in relations:
        for relation_type in relation["types"]:
            relation_type_counter[relation_type] += int(relation["count"])
    role_rank = "".join(
        f'<li><span>{escape(name)}</span><strong>{count}</strong></li>'
        for name, count in role_counter.most_common(5)
    )
    relation_rank = "".join(
        f'<li><span>{escape(name)}</span><strong>{count}</strong></li>'
        for name, count in relation_type_counter.most_common(6)
    )
    body = f"""
    <section class="hero hero--home">
      <div class="hero__content">
        <p class="eyebrow">GITHUB PAGES 静态阅读版</p>
        <h1>把左联人物、关系与事件整理成可浏览的档案站</h1>
        <p class="hero__lead">这一版不依赖 Python 运行环境，直接把标准化 CSV 生成静态页面，适合在线阅读、引用、搜索与 GitHub Pages 发布。</p>
        <div class="hero__actions">
          <a class="button button--primary" href="people/index.html">进入人物档案</a>
          <a class="button" href="graph/index.html">关系图谱</a>
          <a class="button" href="timeline/index.html">事件时间轴</a>
          <a class="button" href="search/index.html">全文搜索</a>
        </div>
      </div>
      <div class="hero__media">
        <img src="assets/banner.png" alt="左联知识库静态阅读版横幅">
      </div>
    </section>

    <section class="metric-grid">
      {metric_card("人物档案", f"{len(people):,}", "可直接进入单人页面阅读")}
      {metric_card("关系对", f"{len(relations):,}", "聚合为人物对关系卡片")}
      {metric_card("历史事件", f"{len(events):,}", "保留时间、地点与参与者")}
      {metric_card("证据来源", f"{sources_count:,}", "仅展示可公开说明的来源信息")}
    </section>

    <section class="section-grid section-grid--two">
      <article class="section-panel">
        <div class="section-panel__eyebrow">概览</div>
        <h2>知识库的静态阅读入口</h2>
        <p>这一版支持人物档案、关系图谱、事件时间轴与全文搜索，直接在浏览器中运行，无需 Python 环境。</p>
        <ul class="rank-list">{role_rank}</ul>
      </article>
      <article class="section-panel">
        <div class="section-panel__eyebrow">关系类型</div>
        <h2>目前最常见的关系标签</h2>
        <p>静态站会优先把高频关系、人际共现与史料证据组织成阅读友好的关系卡片。</p>
        <ul class="rank-list">{relation_rank}</ul>
      </article>
    </section>

    <section class="page-section">
      <div class="section-heading">
        <div>
          <p class="eyebrow">人物入口</p>
          <h2>先看连接最密集的人物</h2>
        </div>
        <a class="section-link" href="people/index.html">查看全部人物</a>
      </div>
      <div class="card-grid">{top_people}</div>
    </section>

    <section class="page-section">
      <div class="section-heading">
        <div>
          <p class="eyebrow">关系索引</p>
          <h2>最值得先读的一组关系卡片</h2>
        </div>
        <a class="section-link" href="relations/index.html">查看全部关系</a>
      </div>
      <div class="relation-grid">{top_relations}</div>
    </section>

    <section class="page-section">
      <div class="section-heading">
        <div>
          <p class="eyebrow">事件时间线</p>
          <h2>按时间进入左联活动现场</h2>
        </div>
        <a class="section-link" href="events/index.html">查看全部事件</a>
      </div>
      <div class="card-grid">{featured_events}</div>
    </section>
    """
    return page_shell(
        title="左联知识库静态阅读版",
        description="适合 GitHub Pages 发布的左联知识库静态阅读站。",
        active_nav="home",
        body=body,
    )


def render_people_index(people: list[dict[str, object]]) -> str:
    cards = "".join(person_card(person, f"./{person['id']}.html") for person in people)
    body = f"""
    <section class="page-hero">
      <p class="eyebrow">人物档案</p>
      <h1>按人物阅读左联知识网络</h1>
      <p class="page-hero__lead">人物页会聚合生卒、角色、关联人物、相关事件与来源卡片，适合做单人档案阅读。</p>
      <div class="filter-box">
        <label for="people-filter">按姓名、别名、角色筛选</label>
        <input id="people-filter" type="search" placeholder="例如：鲁迅、核心领导、上海" data-list-filter="people-list">
        <p class="filter-box__meta" data-count-for="people-list">共 {len(people)} 条人物记录</p>
      </div>
    </section>
    <section class="page-section">
      <div id="people-list" class="card-grid">{cards}</div>
      <p class="empty-state" data-empty-for="people-list" hidden>没有匹配到人物，请换一个关键词。</p>
    </section>
    """
    return page_shell(
        title="人物档案 - 左联知识库静态阅读版",
        description="浏览左联知识库中的人物档案索引。",
        active_nav="people",
        body=body,
        depth=1,
    )


def render_events_index(events: list[dict[str, object]]) -> str:
    cards = "".join(event_card(event, f"./{event['id']}.html") for event in events)
    body = f"""
    <section class="page-hero">
      <p class="eyebrow">事件索引</p>
      <h1>按时间与地点阅读历史事件</h1>
      <p class="page-hero__lead">每条事件保留日期、地点、参与人物与来源说明，适合做课堂展示或史料线索浏览。</p>
      <div class="filter-box">
        <label for="events-filter">按时间、地点、人物筛选</label>
        <input id="events-filter" type="search" placeholder="例如：1930、上海、鲁迅" data-list-filter="events-list">
        <p class="filter-box__meta" data-count-for="events-list">共 {len(events)} 条事件记录</p>
      </div>
    </section>
    <section class="page-section">
      <div id="events-list" class="card-grid">{cards}</div>
      <p class="empty-state" data-empty-for="events-list" hidden>没有匹配到事件，请换一个关键词。</p>
    </section>
    """
    return page_shell(
        title="事件索引 - 左联知识库静态阅读版",
        description="浏览左联知识库中的历史事件索引。",
        active_nav="events",
        body=body,
        depth=1,
    )


def render_relations_index(relations: list[dict[str, object]]) -> str:
    cards = "".join(relation_card(relation, "../people/") for relation in relations)
    body = f"""
    <section class="page-hero">
      <p class="eyebrow">关系索引</p>
      <h1>把人物对关系压缩成可读卡片</h1>
      <p class="page-hero__lead">这里展示的是按人物对聚合后的关系摘要，适合快速定位谁和谁之间有何种关联、证据来自哪里。</p>
      <div class="filter-box">
        <label for="relations-filter">按人名、关系类型、证据关键词筛选</label>
        <input id="relations-filter" type="search" placeholder="例如：鲁迅 通信 上海" data-list-filter="relations-list">
        <p class="filter-box__meta" data-count-for="relations-list">共 {len(relations)} 组人物关系</p>
      </div>
    </section>
    <section class="page-section">
      <div id="relations-list" class="relation-grid">{cards}</div>
      <p class="empty-state" data-empty-for="relations-list" hidden>没有匹配到关系卡片，请换一个关键词。</p>
    </section>
    """
    return page_shell(
        title="关系索引 - 左联知识库静态阅读版",
        description="浏览左联知识库中的人物关系索引。",
        active_nav="relations",
        body=body,
        depth=1,
    )


def render_search_page(total_records: int) -> str:
    body = f"""
    <section class="page-hero">
      <p class="eyebrow">全文搜索</p>
      <h1>统一搜索人物、事件与关系卡片</h1>
      <p class="page-hero__lead">搜索索引是构建时从 CSV 预先生成的 JSON，不依赖后端服务，适合在 GitHub Pages 直接运行。</p>
    </section>

    <section class="search-section">
      <div class="search-app" data-search-index="../assets/search-index.json">
        <label class="search-app__label" for="global-search-input">输入关键词</label>
        <input id="global-search-input" type="search" placeholder="例如：鲁迅、同属组织、1931、上海" data-search-input>
        <p class="search-app__meta" data-search-meta>搜索索引共 {total_records} 条记录。</p>
        <div class="search-results" data-search-results>
          <article class="search-result search-result--placeholder">
            <h2>输入关键词后开始检索</h2>
            <p>支持姓名、别名、角色、关系类型、事件时间、地点与来源摘录等文本。</p>
          </article>
        </div>
      </div>
    </section>
    """
    return page_shell(
        title="全文搜索 - 左联知识库静态阅读版",
        description="在人物、事件与关系之间进行统一全文搜索。",
        active_nav="search",
        body=body,
        depth=1,
    )


def render_person_detail(person: dict[str, object], source_map: dict[str, dict[str, str]]) -> str:
    relation_blocks: list[str] = []
    for relation in person["related_pairs"]:
        is_left = relation["person_a_id"] == person["id"]
        counterpart_id = relation["person_b_id"] if is_left else relation["person_a_id"]
        counterpart_name = relation["person_b_name"] if is_left else relation["person_a_name"]
        relation_blocks.append(
            '<article class="detail-card detail-card--compact">'
            f'<h3><a href="./{escape(counterpart_id, quote=True)}.html">{escape(counterpart_name)}</a></h3>'
            f'<p>{escape(" / ".join(relation["types"][:4]))}</p>'
            f'<div class="detail-card__meta">{relation["count"]} 条记录 · {escape(relation["status_text"])}</div>'
            f'<p>{escape(relation["context_preview"])}</p>'
            "</article>"
        )

    source_html = source_reference_list(person["source_ids"], source_map, limit=16)
    body = f"""
    <section class="detail-hero">
      <div class="breadcrumb"><a href="../index.html">首页</a> / <a href="./index.html">人物档案</a> / {escape(person["name"])}</div>
      <p class="eyebrow">人物档案</p>
      <h1>{escape(person["name"])}</h1>
      <p class="detail-hero__lead">{escape(person["aliases"])}</p>
      <div class="hero-badges">{badge_items([person["role"], person["birth_death"], f"可靠度 {person['reliability']}"])}</div>
    </section>

    <section class="metric-grid">
      {metric_card("关联人物", str(person["pair_total"]), "按人物对聚合后统计")}
      {metric_card("关系记录", str(person["relation_total"]), "包含多条证据与上下文")}
      {metric_card("相关事件", str(person["event_total"]), "通过参与者表回连")}
      {metric_card("来源线索", str(len(person["source_ids"])), "人物、关系、事件的合并来源")}
    </section>

    <section class="section-grid section-grid--detail">
      <article class="section-panel">
        <div class="section-panel__eyebrow">身份概况</div>
        <h2>可读档案摘要</h2>
        <p>该人物在静态站中的信息来自标准人物表、人物关系表和事件参与表。页面更适合做阅读、引用和跳转，不承担复杂计算。</p>
        <ul class="fact-list">
          <li><span>人物 ID</span><strong>{escape(person["id"])}</strong></li>
          <li><span>角色</span><strong>{escape(person["role"])}</strong></li>
          <li><span>生卒</span><strong>{escape(person["birth_death"])}</strong></li>
          <li><span>别名</span><strong>{escape(person["aliases"])}</strong></li>
        </ul>
      </article>
      <article class="section-panel">
        <div class="section-panel__eyebrow">来源卡片</div>
        <h2>当前人物页引用到的证据</h2>
        {source_html}
      </article>
    </section>

    <section class="page-section">
      <div class="section-heading">
        <div>
          <p class="eyebrow">人物关系</p>
          <h2>最值得先读的关联人物</h2>
        </div>
      </div>
      <div class="detail-grid">{''.join(relation_blocks) or '<p class="empty-state">暂无可展示的关系卡片。</p>'}</div>
    </section>

    <section class="page-section">
      <div class="section-heading">
        <div>
          <p class="eyebrow">事件时间线</p>
          <h2>与该人物相关的事件</h2>
        </div>
      </div>
      {timeline_items(person["related_events"], "../events/")}
    </section>
    """
    return page_shell(
        title=f"{person['name']} - 左联知识库静态阅读版",
        description=f"{person['name']}的人物档案与相关关系、事件索引。",
        active_nav="people",
        body=body,
        depth=1,
    )


def render_event_detail(event: dict[str, object], source_map: dict[str, dict[str, str]]) -> str:
    participant_cards = []
    for participant in event["participants"]:
        if participant["person_id"]:
            person_link = f'../people/{escape(participant["person_id"], quote=True)}.html'
            name_html = f'<a href="{person_link}">{escape(participant["name"])}</a>'
        else:
            name_html = escape(participant["name"])
        participant_cards.append(
            '<article class="detail-card detail-card--compact">'
            f"<h3>{name_html}</h3>"
            f'<p>{escape(participant["role"])}</p>'
            "</article>"
        )

    source_html = source_reference_list(event["source_ids"], source_map, limit=12)
    body = f"""
    <section class="detail-hero">
      <div class="breadcrumb"><a href="../index.html">首页</a> / <a href="./index.html">事件索引</a> / {escape(event["name"])}</div>
      <p class="eyebrow">历史事件</p>
      <h1>{escape(event["name"])}</h1>
      <p class="detail-hero__lead">{escape(text(event["historical_location"] or event["current_address"], "地点待补"))}</p>
      <div class="hero-badges">{badge_items([event["date"], event["date_precision"], f"置信度 {event['confidence']}"])}</div>
    </section>

    <section class="section-grid section-grid--detail">
      <article class="section-panel">
        <div class="section-panel__eyebrow">事件摘要</div>
        <h2>时间、地点与备注</h2>
        <ul class="fact-list">
          <li><span>事件 ID</span><strong>{escape(event["id"])}</strong></li>
          <li><span>时间</span><strong>{escape(event["date"])}</strong></li>
          <li><span>历史地点</span><strong>{escape(text(event["historical_location"], "地点待补"))}</strong></li>
          <li><span>现址</span><strong>{escape(text(event["current_address"], "现址待补"))}</strong></li>
        </ul>
        <p class="section-panel__note">{escape(event["note"])}</p>
      </article>
      <article class="section-panel">
        <div class="section-panel__eyebrow">来源卡片</div>
        <h2>当前事件页引用到的证据</h2>
        {source_html}
      </article>
    </section>

    <section class="page-section">
      <div class="section-heading">
        <div>
          <p class="eyebrow">参与人物</p>
          <h2>与事件直接相连的人物</h2>
        </div>
      </div>
      <div class="detail-grid">{''.join(participant_cards) or '<p class="empty-state">暂无参与人物。</p>'}</div>
    </section>
    """
    return page_shell(
        title=f"{event['name']} - 左联知识库静态阅读版",
        description=f"{event['name']}的事件详情与参与人物。",
        active_nav="events",
        body=body,
        depth=1,
    )


def build_search_index(
    people: list[dict[str, object]],
    events: list[dict[str, object]],
    relations: list[dict[str, object]],
) -> list[dict[str, str]]:
    records: list[dict[str, str]] = []
    for person in people:
        records.append(
            {
                "type": "人物",
                "title": text(person["name"]),
                "subtitle": f"{text(person['role'])} · {text(person['birth_death'])}",
                "url": f"../people/{person['id']}.html",
                "text": person["search_blob"],
            }
        )
    for event in events:
        location = text(event["historical_location"] or event["current_address"], "地点待补")
        records.append(
            {
                "type": "事件",
                "title": text(event["name"]),
                "subtitle": f"{text(event['date'])} · {location}",
                "url": f"../events/{event['id']}.html",
                "text": event["search_blob"],
            }
        )
    for relation in relations:
        records.append(
            {
                "type": "关系",
                "title": f"{relation['person_a_name']} × {relation['person_b_name']}",
                "subtitle": " / ".join(relation["types"][:4]),
                "url": f"../relations/index.html#{relation['anchor']}",
                "text": relation["search_blob"],
            }
        )
    return records


def build_timeline_data(event_records: list[dict]) -> list[dict]:
    items = []
    for ev in event_records:
        date = text(ev.get("date", ""))
        if not date:
            continue
        year = extract_year(date)
        if not year:
            continue
        items.append({
            "id": ev["id"],
            "name": ev["name"],
            "date": date,
            "year": year,
            "location": text(ev.get("historical_location") or ev.get("location", ""), "地点待补"),
            "participants": [p["name"] for p in ev.get("participants", [])[:4]],
            "confidence": text(ev.get("confidence", "")),
        })
    items.sort(key=lambda x: x["date"])
    return items


def render_timeline_page() -> str:
    body = """
<div class="page-hero">
  <h1 class="page-hero__title">事件时间轴</h1>
  <p class="page-hero__lead">左联历史事件按时间排列，拖动滑块筛选年份范围，点击事件查看详情。</p>
</div>
<div style="display:flex;gap:1rem;flex-wrap:wrap;align-items:center;margin-bottom:1rem">
  <label style="font-size:.9rem">年份范围：
    <input id="year-min" type="number" value="1920" min="1900" max="1940"
      style="width:5rem;padding:.25rem .5rem;border-radius:4px;border:1px solid var(--line)">
    —
    <input id="year-max" type="number" value="1940" min="1900" max="1940"
      style="width:5rem;padding:.25rem .5rem;border-radius:4px;border:1px solid var(--line)">
  </label>
  <label style="font-size:.9rem">关键词：
    <input id="kw-filter" type="text" placeholder="人名/地点/事件名"
      style="padding:.25rem .5rem;border-radius:4px;border:1px solid var(--line);width:12rem">
  </label>
  <span id="tl-stats" style="font-size:.85rem;color:var(--muted)"></span>
</div>
<div id="tl-container" style="width:100%;height:65vh;border:1px solid var(--line);border-radius:8px;background:#faf6ef"></div>
<script src="https://cdn.jsdelivr.net/npm/echarts@5/dist/echarts.min.js"></script>
<script>
(function(){
  var chart = echarts.init(document.getElementById('tl-container'));
  var allItems = [];

  function buildOption(items) {
    var years = items.map(function(d){ return d.year; });
    var minY = Math.min.apply(null, years) || 1920;
    var maxY = Math.max.apply(null, years) || 1940;
    var data = items.map(function(d, i){
      return {
        value: [d.date, i % 8, d.name],
        id: d.id,
        name: d.name,
        location: d.location,
        participants: d.participants,
        confidence: d.confidence,
      };
    });
    return {
      tooltip: {
        formatter: function(p) {
          var d = p.data;
          var parts = [
            '<b>' + d.name + '</b>',
            '日期：' + d.value[0],
            '地点：' + d.location,
          ];
          if (d.participants && d.participants.length)
            parts.push('参与：' + d.participants.join('、'));
          return parts.join('<br>');
        }
      },
      grid: { left: 60, right: 40, top: 20, bottom: 60 },
      xAxis: {
        type: 'category',
        data: Array.from({length: maxY - minY + 1}, function(_, i){ return String(minY + i); }),
        axisLabel: { rotate: 45, fontSize: 11 },
        name: '年份',
      },
      yAxis: { show: false, min: -1, max: 9 },
      dataZoom: [
        { type: 'slider', xAxisIndex: 0, bottom: 5, height: 20 },
        { type: 'inside', xAxisIndex: 0 }
      ],
      series: [{
        type: 'scatter',
        data: data,
        symbolSize: 10,
        itemStyle: { color: '#8e2528', opacity: 0.75 },
        emphasis: { itemStyle: { color: '#536779', opacity: 1 }, scale: 1.5 },
      }]
    };
  }

  function applyFilter() {
    var minY = parseInt(document.getElementById('year-min').value) || 1920;
    var maxY = parseInt(document.getElementById('year-max').value) || 1940;
    var kw = (document.getElementById('kw-filter').value || '').trim().toLowerCase();
    var filtered = allItems.filter(function(d) {
      if (d.year < minY || d.year > maxY) return false;
      if (kw && d.name.toLowerCase().indexOf(kw) < 0 &&
          d.location.toLowerCase().indexOf(kw) < 0 &&
          d.participants.join('').toLowerCase().indexOf(kw) < 0) return false;
      return true;
    });
    document.getElementById('tl-stats').textContent = '显示 ' + filtered.length + ' 条事件';
    chart.setOption(buildOption(filtered), true);
  }

  fetch('../assets/timeline-data.json').then(function(r){ return r.json(); }).then(function(data){
    allItems = data;
    ['year-min','year-max','kw-filter'].forEach(function(id){
      document.getElementById(id).addEventListener('input', applyFilter);
    });
    applyFilter();
  });

  chart.on('click', function(p){
    if (p.data && p.data.id) window.location.href = '../events/' + p.data.id + '.html';
  });
  window.addEventListener('resize', function(){ chart.resize(); });
})();
</script>
"""
    return page_shell(
        title="事件时间轴 · 左联知识库",
        description="左联历史事件时间轴",
        active_nav="timeline",
        body=body,
        depth=1,
    )


def build_graph_data(people: list[dict], relation_profiles: list[dict]) -> dict:
    nodes = [
        {
            "id": p["id"],
            "name": p["name"],
            "value": int(p["pair_total"]),
            "role": p.get("role", ""),
        }
        for p in people
    ]
    edges = []
    seen: set[str] = set()
    for rel in relation_profiles:
        a, b = rel["person_a_id"], rel["person_b_id"]
        key = f"{min(a,b)}|{max(a,b)}"
        if key in seen:
            continue
        seen.add(key)
        edges.append({
            "source": a,
            "target": b,
            "types": rel["types"][:2],
            "weight": int(rel.get("weight", 1)),
        })
    return {"nodes": nodes, "edges": edges}


def render_graph_page() -> str:
    body = """
<div class="page-hero">
  <h1 class="page-hero__title">关系图谱</h1>
  <p class="page-hero__lead">左联成员人物关系交互网络，点击节点查看人物详情，滚轮缩放，拖拽平移。</p>
</div>
<div style="margin-bottom:1rem;display:flex;gap:.75rem;flex-wrap:wrap;align-items:center">
  <label style="font-size:.9rem">筛选关系类型：
    <select id="rel-filter" style="margin-left:.4rem;padding:.25rem .5rem;border-radius:4px;border:1px solid var(--line)">
      <option value="">全部</option>
      <option>交游</option><option>合作</option><option>同属组织</option>
      <option>签名联署</option><option>纪念悼念</option><option>待核验</option>
    </select>
  </label>
  <label style="font-size:.9rem">最少关联人数：
    <input id="min-degree" type="number" min="0" value="0"
      style="width:4rem;margin-left:.4rem;padding:.25rem .5rem;border-radius:4px;border:1px solid var(--line)">
  </label>
  <span id="graph-stats" style="font-size:.85rem;color:var(--muted)"></span>
</div>
<div id="graph-container" style="width:100%;height:70vh;border:1px solid var(--line);border-radius:8px;background:#faf6ef"></div>
<script src="https://cdn.jsdelivr.net/npm/echarts@5/dist/echarts.min.js"></script>
<script>
(function(){
  var chart = echarts.init(document.getElementById('graph-container'));
  var allNodes = [], allEdges = [];

  function buildOption(nodes, edges) {
    return {
      tooltip: {
        formatter: function(p) {
          if (p.dataType === 'node') return p.data.name + (p.data.role ? ' · ' + p.data.role : '');
          return (p.data.sourceNode||'') + ' — ' + (p.data.targetNode||'') + '<br>' + (p.data.types||[]).join(' / ');
        }
      },
      series: [{
        type: 'graph', layout: 'force',
        roam: true, draggable: true,
        force: { repulsion: 120, edgeLength: [60, 200], gravity: 0.1 },
        label: { show: true, fontSize: 11, color: '#2d251d' },
        lineStyle: { color: '#b09070', opacity: 0.5, width: 1 },
        itemStyle: { color: '#8e2528', borderColor: '#fff', borderWidth: 1.5 },
        emphasis: { focus: 'adjacency', lineStyle: { width: 3 } },
        nodes: nodes,
        edges: edges,
        symbolSize: function(val) { return Math.max(14, Math.min(50, 14 + val * 0.25)); }
      }]
    };
  }

  function applyFilter() {
    var relType = document.getElementById('rel-filter').value;
    var minDeg = parseInt(document.getElementById('min-degree').value) || 0;
    var activeEdges = allEdges.filter(function(e) {
      return !relType || (e.types && e.types.indexOf(relType) >= 0);
    });
    var activeIds = new Set();
    activeEdges.forEach(function(e){ activeIds.add(e.source); activeIds.add(e.target); });
    var activeNodes = allNodes.filter(function(n){
      return n.value >= minDeg && (minDeg > 0 ? activeIds.has(n.id) : true);
    });
    var nodeIds = new Set(activeNodes.map(function(n){ return n.id; }));
    var filteredEdges = activeEdges.filter(function(e){
      return nodeIds.has(e.source) && nodeIds.has(e.target);
    }).map(function(e){
      return Object.assign({}, e, {
        sourceNode: (allNodes.find(function(n){return n.id===e.source;})||{}).name,
        targetNode: (allNodes.find(function(n){return n.id===e.target;})||{}).name
      });
    });
    document.getElementById('graph-stats').textContent =
      '显示 ' + activeNodes.length + ' 人 / ' + filteredEdges.length + ' 条关系';
    chart.setOption(buildOption(activeNodes, filteredEdges), true);
  }

  fetch('../assets/graph-data.json').then(function(r){ return r.json(); }).then(function(data){
    allNodes = data.nodes;
    allEdges = data.edges;
    document.getElementById('rel-filter').addEventListener('change', applyFilter);
    document.getElementById('min-degree').addEventListener('input', applyFilter);
    applyFilter();
  });

  chart.on('click', function(p){
    if (p.dataType === 'node') window.location.href = '../people/' + p.data.id + '.html';
  });
  window.addEventListener('resize', function(){ chart.resize(); });
})();
</script>
"""
    return page_shell(
        title="关系图谱 · 左联知识库",
        description="左联成员人物关系交互网络图谱",
        active_nav="graph",
        body=body,
        depth=1,
    )


def main() -> None:
    global DATA_DIR, DOCS_DIR
    args = parse_args()
    DATA_DIR = args.data_dir.resolve()
    DOCS_DIR = args.output_dir.resolve()

    required = [
        "persons.csv",
        "person_relations.csv",
        "events.csv",
        "event_participants.csv",
        "sources.csv",
    ]
    missing = [filename for filename in required if not (DATA_DIR / filename).exists()]
    if missing:
        raise FileNotFoundError(f"静态站生成失败，缺少数据文件：{', '.join(missing)}")

    persons_df = load_frame("persons.csv")
    relations_df = load_frame("person_relations.csv")
    events_df = load_frame("events.csv")
    participants_df = load_frame("event_participants.csv")
    sources_df = load_frame("sources.csv")

    source_map = build_sources_lookup(sources_df)
    name_map = {
        text(row.get("person_id")): text(row.get("standard_name"), text(row.get("person_id")))
        for row in persons_df.to_dict("records")
        if text(row.get("person_id"))
    }
    relation_profiles = build_relation_profiles(relations_df, name_map)
    event_records = build_event_records(events_df, participants_df, name_map)
    people = build_person_records(persons_df, relation_profiles, event_records, source_map)
    search_records = build_search_index(people, event_records, relation_profiles)

    cleanup_output()
    copy_assets()

    write_text(DOCS_DIR / "index.html", render_home_page(people, relation_profiles, event_records, len(source_map)))
    write_text(DOCS_DIR / "people" / "index.html", render_people_index(people))
    write_text(DOCS_DIR / "events" / "index.html", render_events_index(event_records))
    write_text(DOCS_DIR / "relations" / "index.html", render_relations_index(relation_profiles))
    write_text(DOCS_DIR / "search" / "index.html", render_search_page(len(search_records)))
    write_text(DOCS_DIR / "assets" / "search-index.json", json.dumps(search_records, ensure_ascii=False, indent=2))

    graph_data = build_graph_data(people, relation_profiles)
    write_text(DOCS_DIR / "assets" / "graph-data.json", json.dumps(graph_data, ensure_ascii=False))
    write_text(DOCS_DIR / "graph" / "index.html", render_graph_page())

    timeline_data = build_timeline_data(event_records)
    write_text(DOCS_DIR / "assets" / "timeline-data.json", json.dumps(timeline_data, ensure_ascii=False))
    write_text(DOCS_DIR / "timeline" / "index.html", render_timeline_page())

    for person in people:
        write_text(DOCS_DIR / "people" / f"{person['id']}.html", render_person_detail(person, source_map))

    for event in event_records:
        write_text(DOCS_DIR / "events" / f"{event['id']}.html", render_event_detail(event, source_map))

    print(
        "Static site generated:",
        f"{len(people)} people,",
        f"{len(relation_profiles)} relation cards,",
        f"{len(event_records)} events.",
    )


if __name__ == "__main__":
    main()
