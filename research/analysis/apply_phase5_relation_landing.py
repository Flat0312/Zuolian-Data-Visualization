"""Phase 5 四百条关系裁决的生产层落地（Tier 1：证据派生 supported，幂等）。

授权依据
--------
用户 2026-09-20 会话授权，签核路径 C（全量按 AI 建议执行），逐字授权语
「我都没问题，你自己看着办吧」，登记于 ``phase5_review_adjudication_record.md``；
裁决结果落在 ``phase5_relation_review_adjudicated.csv``（400 行，每行带授权溯源列）。

落地范围（保守）
----------------
只处理同时满足三条的关系：夜间轮回源核查判定 ``evidence_support=support``、人工裁决为
``correct`` 或 ``wrong_type``、且引文能在本地原文中通过运行时逐字复核。
共 48 条关系 / 23 条去重证据；预期公开层 derived ``supported`` 12 条。

明确不做的事
------------
- 不把既有 10249 条机器迁移生成的 ``associated`` 证据行改判为 ``support``：这些行的
  reviewer_note 写明「仅表示来源关联，不断言支持强度」，改判等于夸大证据等级。
  本脚本改为**新增**经逐字复核的 support 证据行，旧行原样保留。
- 不触碰 associated 层（113 条人工成立但证据仅关联级）与 insufficient 层（239 条）。
- 不把 ``not_supported`` 映射为 ``rejected``：五态中 ``rejected`` 语义是「人工否定」
  （对应 contradicted，本批 0 条），``not_supported`` 只表示证据不足，映射会夸大结论。
- 不用 ``human_adjudication`` 通道把 critical/high 风险关系推入公开层：那会违反
  ``relation_publish_status`` 的保守原则与既有治理门禁测试语义。公开层只接受 derived supported。
- 不改写 ``relation_risk_level``（禁止反向降险）。

运行时校验（任一失败即整体退出、不写任何文件）
----------------------------------------------
1. 裁决表 400 行、human_verdict 全非空且合法、授权三列全非空、授权语非执行者自授占位；
2. 裁决表 overnight_evidence_support 与 relations.jsonl 的 evidence_support 逐条一致（防漂移）；
3. 每条待落地引文在 local_path 原文中空白归一后逐字命中，且 quote_sha256、原文 content_hash 一致；
4. 生产表基线计数与字段状态符合预期；
5. 落地后计数、发布状态分布、Schema 0 错误全部符合预期。
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import os
import re
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
DATA = ROOT / "data" / "processed"
REPORTS = ROOT / "research" / "drafts" / "reports"
PERSONS = ROOT / "data" / "processed" / "persons.csv"
RUNTIME_SOURCES = ROOT / "data" / "processed" / "runtime_sources"
RAW_TEXTS = ROOT / "research" / "raw_texts"
OVERNIGHT = ROOT / "research" / "content-upgrade-overnight-2026-09-06"

ADJUDICATED = REPORTS / "phase5_relation_review_adjudicated.csv"
RELATIONS_JSONL = OVERNIGHT / "relations.jsonl"
EVIDENCE_JSONL = OVERNIGHT / "evidence.jsonl"
LEDGER = REPORTS / "phase5_relation_landing_ledger.csv"
REPORT_MD = REPORTS / "phase5_relation_landing_report.md"
RECAPTURE_QUEUE = REPORTS / "phase5_quote_recapture_queue.csv"

# 导入不得污染 sys.path 顺序：把仓库根插到最前会让根目录的 app.py 垫片遮蔽
# app/frontend/app.py（tests/test_utils_and_views.py 依赖后者）。因此优先按包路径导入，
# 只有以脚本方式直接运行时才回退，且用 append 追加脚本自身目录。
_ANALYSIS_DIR = Path(__file__).resolve().parent
try:
    from research.analysis.quote_attestation import (
        find_cooccurrence,
        looks_like_name_list,
        name_candidates,
        normalized_bundle,
        quote_attests_pair,
    )
    from research.analysis.relation_publish_status import (
        INFERRED_RELATION_TYPES,
        PUBLIC_RELATION_STATUSES,
        assign_relation_status_columns,
    )
    from research.analysis.source_layer import sync_source_layer
except ImportError:
    if str(_ANALYSIS_DIR) not in sys.path:
        sys.path.append(str(_ANALYSIS_DIR))
    from quote_attestation import (  # noqa: E402
        find_cooccurrence,
        looks_like_name_list,
        name_candidates,
        normalized_bundle,
        quote_attests_pair,
    )
    from relation_publish_status import (  # noqa: E402
        INFERRED_RELATION_TYPES,
        PUBLIC_RELATION_STATUSES,
        assign_relation_status_columns,
    )
    from source_layer import sync_source_layer  # noqa: E402

BATCH_MARKER = "P5-LANDING-2026-09-28"
AUTHORIZED_BY = "用户（会话授权）"
AUTHORIZED_AT = "2026-09-20"
AUTHORIZATION_QUOTE = "我都没问题，你自己看着办吧"
LANDED_AT = "2026-09-28"

EXPECTED_BASELINE = {
    "person_relations.csv": 4238,
    "relation_evidences.csv": 10249,
    "sources.csv": 1177,
    "source_passages.csv": 1177,
}
EXPECTED_POST = {
    "person_relations.csv": 4238,
    "relation_evidences.csv": 10269,
    "sources.csv": 1178,
    "source_passages.csv": 1178,
}
EXPECTED_TARGET_RELATIONS = 48
EXPECTED_DISTINCT_EVIDENCES = 23
EXPECTED_ATTESTED_RELATIONS = 20
EXPECTED_UNATTESTED_RELATIONS = 28
EXPECTED_NEW_EVIDENCE_ROWS = 20
EXPECTED_TYPE_CORRECTIONS = 3
EXPECTED_PUBLIC_SUPPORTED = 5
EXPECTED_POST_PUBLISH_STATUS = {"supported": 5, "pending_review": 2451, "inferred": 1782}
NEW_SOURCE_LOCATOR = "鲁迅日记 1928年7月1日"
DIARY_TEMPLATE_SOURCE_ID = "SRC-0497"
STRENGTH_TO_LEVEL = {"一手": "A", "二手": "B", "转引": "C", "参考": "D", "推断": "D"}
VALID_VERDICTS = ("correct", "wrong_type", "not_supported", "contradicted")
UPHELD_VERDICTS = ("correct", "wrong_type")
SELF_AUTH_PLACEHOLDERS = ("ai 自行决定", "执行者自授", "模型决定")
FAMILY_MAP = {"鲁迅日记": "luxun_diary", "左联词典": "zuolian_cidian", "左联史": "zuolian_shi"}
LEDGER_COLUMNS = [
    "relation_id", "source_person_id", "target_person_id", "person_a_name", "person_b_name",
    "batch_marker", "attestation_basis",
    "human_verdict", "final_relation_type_before", "final_relation_type_after",
    "publish_status_before", "publish_status_after", "publish_status_origin",
    "new_relation_evidence_id", "overnight_evidence_id", "source_id", "source_registered",
    "locator", "quote_sha256", "search_log_id", "quote_verbatim_recheck",
    "authorized_by", "authorized_at", "authorization_quote",
]

def _read(path: Path) -> tuple[list[str], list[dict[str, str]]]:
    with open(path, encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        return list(reader.fieldnames or []), [dict(row) for row in reader]


def _write(path: Path, columns: list[str], rows: list[dict[str, str]]) -> None:
    tmp = path.with_suffix(path.suffix + ".tmp")
    with open(tmp, "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=columns, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)
    os.replace(tmp, path)


def _norm(text: object) -> str:
    return re.sub(r"\s+", "", str(text or ""))


def _sha256_text(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _read_jsonl(path: Path) -> list[dict]:
    with open(path, encoding="utf-8") as fh:
        return [json.loads(line) for line in fh if line.strip()]


def _next_rele_id(existing_ids: list[str]) -> int:
    max_n = 0
    for value in existing_ids:
        m = re.fullmatch(r"RELE-(\d+)", str(value).strip())
        if m:
            max_n = max(max_n, int(m.group(1)))
    return max_n + 1


class LandingError(RuntimeError):
    """任何前置/后置校验失败都抛出本异常，调用方保证不写任何文件。"""


def load_adjudication(path: Path) -> dict[str, dict[str, str]]:
    cols, rows = _read(path)
    if len(rows) != 400:
        raise LandingError(f"裁决表应为 400 行，实际 {len(rows)}")
    required = ("human_verdict", "overnight_evidence_support", "ai_suggested_type",
                "authorized_by", "authorized_at", "authorization_quote")
    missing = [c for c in required if c not in cols]
    if missing:
        raise LandingError(f"裁决表缺少必需列：{missing}")
    out: dict[str, dict[str, str]] = {}
    for row in rows:
        rid = row["relation_id"].strip()
        verdict = row["human_verdict"].strip()
        if verdict not in VALID_VERDICTS:
            raise LandingError(f"{rid} human_verdict 非法或为空：{verdict!r}")
        quote = row["authorization_quote"].strip()
        if not quote or not row["authorized_by"].strip() or not row["authorized_at"].strip():
            raise LandingError(f"{rid} 授权溯源列不完整，拒绝落地")
        if any(p in quote.lower() for p in SELF_AUTH_PLACEHOLDERS):
            raise LandingError(f"{rid} 授权语疑似执行者自授占位：{quote!r}")
        if quote != AUTHORIZATION_QUOTE:
            raise LandingError(f"{rid} 授权语与本批登记的逐字授权语不一致：{quote!r}")
        if row["authorized_at"].strip() != AUTHORIZED_AT:
            raise LandingError(f"{rid} 授权日期与本批登记不一致：{row['authorized_at']!r}")
        out[rid] = row
    return out


def load_overnight(adj: dict[str, dict[str, str]]) -> tuple[list[dict], dict[str, dict]]:
    relations = [r for r in _read_jsonl(RELATIONS_JSONL) if "sample400" in (r.get("groups") or [])]
    evidences = {r["evidence_id"]: r for r in _read_jsonl(EVIDENCE_JSONL)}
    for rec in relations:
        rid = rec["relation_id"]
        if rid not in adj:
            raise LandingError(f"relations.jsonl 的 {rid} 不在裁决表中，疑似样本漂移")
        if str(rec.get("evidence_support", "")).strip() != adj[rid]["overnight_evidence_support"].strip():
            raise LandingError(
                f"{rid} 裁决表 overnight_evidence_support 与 relations.jsonl 不一致，拒绝落地"
            )
    return relations, evidences


def build_target_set(relations: list[dict], adj: dict[str, dict[str, str]]) -> list[dict]:
    target = [
        rec for rec in relations
        if rec.get("evidence_support") == "support"
        and adj[rec["relation_id"]]["human_verdict"].strip() in UPHELD_VERDICTS
    ]
    if len(target) != EXPECTED_TARGET_RELATIONS:
        raise LandingError(f"目标关系应为 {EXPECTED_TARGET_RELATIONS} 条，实际 {len(target)}")
    distinct = {eid for rec in target for eid in rec["evidence_ids"]}
    if len(distinct) != EXPECTED_DISTINCT_EVIDENCES:
        raise LandingError(f"去重证据应为 {EXPECTED_DISTINCT_EVIDENCES} 条，实际 {len(distinct)}")
    for rec in target:
        if len(rec["evidence_ids"]) != 1:
            raise LandingError(f"{rec['relation_id']} 证据数不是 1，本批只处理单证据关系：{rec['evidence_ids']}")
        if rec.get("review_status") != "pending_human_review":
            raise LandingError(f"{rec['relation_id']} 夜间轮 review_status 异常：{rec.get('review_status')!r}")
    return target


def verify_quotes(target: list[dict], evidences: dict[str, dict]) -> dict[str, str]:
    """逐字复核：返回 evidence_id -> 'verbatim_whitespace_normalized'。失败即抛异常。"""
    cache: dict[str, str] = {}
    result: dict[str, str] = {}
    for rec in target:
        for eid in rec["evidence_ids"]:
            if eid in result:
                continue
            ev = evidences.get(eid)
            if ev is None:
                raise LandingError(f"{eid} 不在 evidence.jsonl 中")
            local = Path(str(ev["local_path"]))
            if not local.exists():
                raise LandingError(f"{eid} 本地原文不存在，无法逐字复核：{local}")
            key = str(local)
            if key not in cache:
                cache[key] = local.read_text(encoding="utf-8", errors="replace")
                actual_hash = _sha256_file(local)
                if actual_hash != ev.get("content_hash"):
                    raise LandingError(
                        f"{eid} 原文 content_hash 不一致（文件已变动）：期望 {ev.get('content_hash')}，实际 {actual_hash}"
                    )
            quote = str(ev.get("quote", ""))
            if not quote.strip():
                raise LandingError(f"{eid} 引文为空")
            if _sha256_text(quote) != ev.get("quote_sha256"):
                raise LandingError(f"{eid} quote_sha256 不一致，引文已被改动")
            if _norm(quote) not in _norm(cache[key]):
                raise LandingError(f"{eid} 引文未在本地原文中逐字命中（空白归一后仍不命中）：{ev.get('locator')}")
            if not str(ev.get("locator", "")).strip():
                raise LandingError(f"{eid} locator 为空")
            if ev.get("retrieval_status") != "local_checked":
                raise LandingError(f"{eid} retrieval_status 非 local_checked：{ev.get('retrieval_status')!r}")
            result[eid] = "verbatim_whitespace_normalized"
    return result

def _next_source_id(existing: list[str]) -> str:
    max_n = 0
    for value in existing:
        m = re.fullmatch(r"SRC-(\d+)", str(value).strip())
        if m:
            max_n = max(max_n, int(m.group(1)))
    return f"SRC-{max_n + 1:04d}"


def resolve_sources(
    source_cols: list[str],
    sources: list[dict[str, str]],
    evidences: dict[str, dict],
    needed_eids: list[str],
) -> tuple[dict[str, str], list[dict[str, str]], list[str]]:
    """把每条待落地证据映射到 sources.csv 的 source_id；缺失的按同族模板注册一条新来源。

    返回 (evidence_id -> source_id, 新增来源行, 新注册的 locator 列表)。
    """
    by_locator: dict[tuple[str, str], str] = {}
    for row in sources:
        fam = (row.get("source_family") or "").strip()
        cit = (row.get("citation") or "").strip()
        if fam and cit:
            by_locator.setdefault((fam, cit), row["source_id"].strip())
    template = next((r for r in sources if r["source_id"].strip() == DIARY_TEMPLATE_SOURCE_ID), None)
    if template is None:
        raise LandingError(f"同族模板来源 {DIARY_TEMPLATE_SOURCE_ID} 不存在，无法注册新来源")

    mapping: dict[str, str] = {}
    added: list[dict[str, str]] = []
    registered: list[str] = []
    existing_ids = [r["source_id"].strip() for r in sources]
    for eid in needed_eids:
        ev = evidences[eid]
        family = FAMILY_MAP.get(str(ev.get("source_family", "")).strip())
        if not family:
            raise LandingError(f"{eid} 未知 source_family：{ev.get('source_family')!r}")
        locator = str(ev.get("locator", "")).strip()
        sid = by_locator.get((family, locator))
        if sid:
            mapping[eid] = sid
            continue
        if locator != NEW_SOURCE_LOCATOR:
            raise LandingError(
                f"{eid}  locator {locator!r} 在 sources.csv 中无对应来源，且不在本批允许注册的白名单内"
            )
        if family != "luxun_diary":
            raise LandingError(f"{eid} 新注册来源仅限 luxun_diary 族，当前 {family}")
        new_id = _next_source_id(existing_ids + [r["source_id"] for r in added])
        row = {col: template.get(col, "") for col in source_cols}
        row["source_id"] = new_id
        row["citation"] = locator
        row["title"] = template.get("title", "鲁迅日记")
        row["source_family"] = family
        row["review_note"] = f"{BATCH_MARKER} 按夜间轮逐字复核证据注册；引文见 relation_evidences。"
        added.append(row)
        by_locator[(family, locator)] = new_id
        mapping[eid] = new_id
        registered.append(locator)
    return mapping, added, registered


def build_evidence_row(
    cols: list[str],
    rele_id: str,
    relation_id: str,
    source_id: str,
    ev: dict,
    rec: dict,
    source_level: str,
    recheck: str,
    newly_registered: bool,
) -> dict[str, str]:
    search_logs = ";".join(rec.get("search_log_ids") or [])
    note = (
        f"{BATCH_MARKER}：夜间轮回源核查证据 {ev['evidence_id']} 落地（{LANDED_AT}）。"
        f"引文逐字复核={recheck}；quote_sha256={ev.get('quote_sha256', '')}；"
        f"检索凭据={search_logs or '无'}；retrieval_status={ev.get('retrieval_status', '')}；"
        f"核录时间={ev.get('accessed_at', '')}；夜间轮介质分级={str(ev.get('source_level', ''))[:1]}；"
        f"本行 source_level 按 sources.evidence_strength 既有映射取值。"
        f"分诊理由：{rec.get('reason', '')}"
        f"｜授权：{AUTHORIZED_BY} {AUTHORIZED_AT} 签核路径C 逐字「{AUTHORIZATION_QUOTE}」。"
        + ("｜来源为本批新注册。" if newly_registered else "")
    )
    row = {col: "" for col in cols}
    row.update({
        "relation_evidence_id": rele_id,
        "relation_id": relation_id,
        "source_id": source_id,
        "locator": str(ev.get("locator", "")).strip(),
        "quote": str(ev.get("quote", "")),
        "context": str(rec.get("reason", "")),
        "quote_or_context": str(ev.get("quote", "")),
        "evidence_support": "support",
        "source_level": source_level,
        "review_status": "reviewed",
        "reviewer_note": note,
    })
    return row


def already_landed(evid_rows: list[dict[str, str]]) -> bool:
    return any(BATCH_MARKER in (r.get("reviewer_note") or "") for r in evid_rows)

def load_persons(path: Path) -> dict[str, dict[str, str]]:
    _, rows = _read(path)
    return {r["person_id"].strip(): r for r in rows}


def gate_attestation(
    target: list[dict], evidences: dict[str, dict], persons: dict[str, dict[str, str]]
) -> tuple[list[tuple[dict, str]], list[tuple[dict, str]]]:
    """双方佐证门：引文须同时记载关系两端（鲁迅日记按日记体免甲方自名）。

    返回 (通过列表[(rec, basis)], 未通过列表[(rec, basis)])。未通过者不落地、不改类型，
    只进入重捕队列——引文捕错段落不等于关系不成立，也不等于可以公开。
    """
    passed: list[tuple[dict, str]] = []
    failed: list[tuple[dict, str]] = []
    for rec in target:
        ev = evidences[rec["evidence_ids"][0]]
        pa = persons.get(rec["source_person_id"])
        pb = persons.get(rec["target_person_id"])
        if pa is None or pb is None:
            raise LandingError(f"{rec['relation_id']} 的人物 ID 不在 persons.csv 中")
        diary = str(ev.get("source_family", "")).strip() == "鲁迅日记"
        ok, basis = quote_attests_pair(
            ev.get("quote", ""), name_candidates(pa), name_candidates(pb),
            diary_author_implicit=diary,
        )
        (passed if ok else failed).append((rec, basis))
    if len(passed) != EXPECTED_ATTESTED_RELATIONS:
        raise LandingError(f"通过双方佐证门的关系应为 {EXPECTED_ATTESTED_RELATIONS} 条，实际 {len(passed)}")
    if len(failed) != EXPECTED_UNATTESTED_RELATIONS:
        raise LandingError(f"未通过双方佐证门的关系应为 {EXPECTED_UNATTESTED_RELATIONS} 条，实际 {len(failed)}")
    return passed, failed


QUEUE_COLUMNS = [
    "relation_id", "person_a_id", "person_a_name", "person_b_id", "person_b_name",
    "final_relation_type", "human_verdict", "overnight_evidence_support",
    "recorded_locator", "recorded_evidence_id", "attestation_basis",
    "quote_verbatim_recheck", "local_file", "cooccurrence_found", "best_distance",
    "best_excerpt", "best_list_like", "proposed_action", "overnight_reason",
    "search_log_ids", "review_status",
]


def build_recapture_queue(
    failed: list[tuple[dict, str]],
    evidences: dict[str, dict],
    persons: dict[str, dict[str, str]],
    recheck: dict[str, str],
) -> list[dict[str, str]]:
    """为引文未佐证双方的关系生成重捕队列（只读研究层产物，不动生产数据）。"""
    bundles: dict[str, tuple[str, str, list[int]]] = {}
    rows: list[dict[str, str]] = []
    for rec, basis in sorted(failed, key=lambda item: item[0]["relation_id"]):
        ev = evidences[rec["evidence_ids"][0]]
        pa = persons[rec["source_person_id"]]
        pb = persons[rec["target_person_id"]]
        local = str(ev.get("local_path", ""))
        resolved = Path(local)
        if not resolved.is_absolute():
            resolved = ROOT / local
        if not resolved.exists():
            raise LandingError(f"{rec['relation_id']} 本地原文不存在，无法评估重捕可能：{resolved}")
        key = str(resolved)
        if key not in bundles:
            bundles[key] = normalized_bundle(resolved)
        hits = find_cooccurrence(bundles[key], name_candidates(pa), name_candidates(pb))
        best = hits[0] if hits else None
        list_like = bool(best) and looks_like_name_list(str(best.get("excerpt", "")))
        if best is None:
            action = "no_local_support_mark_insufficient"
        elif list_like:
            action = "cooccurrence_only_keep_associated"
        else:
            action = "recapture_quote_then_regrade"
        rows.append({
            "relation_id": rec["relation_id"],
            "person_a_id": rec["source_person_id"],
            "person_a_name": (pa.get("standard_name") or "").strip(),
            "person_b_id": rec["target_person_id"],
            "person_b_name": (pb.get("standard_name") or "").strip(),
            "final_relation_type": str(rec.get("proposed_relation_type") or "").strip(),
            "human_verdict": "",
            "overnight_evidence_support": str(rec.get("evidence_support") or "").strip(),
            "recorded_locator": str(ev.get("locator") or "").strip(),
            "recorded_evidence_id": str(ev.get("evidence_id") or "").strip(),
            "attestation_basis": basis,
            "quote_verbatim_recheck": recheck.get(ev.get("evidence_id"), ""),
            "local_file": str(resolved),
            "cooccurrence_found": "yes" if best else "no",
            "best_distance": str(best.get("distance", "")) if best else "",
            "best_excerpt": str(best.get("excerpt", ""))[:400] if best else "",
            "best_list_like": "yes" if list_like else ("no" if best else ""),
            "proposed_action": action,
            "overnight_reason": str(rec.get("reason") or "")[:400],
            "search_log_ids": ";".join(rec.get("search_log_ids") or []),
            "review_status": "pending_human_review",
        })
    return rows


def write_recapture_queue(rows: list[dict[str, str]], path: Path, adj: dict[str, dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=QUEUE_COLUMNS, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            item = dict(row)
            verdict = adj.get(item["relation_id"], {})
            item["human_verdict"] = verdict.get("human_verdict", "")
            writer.writerow(item)


def _check_counts(rows: dict[str, int], expected: dict[str, int], stage: str) -> None:
    for name, want in expected.items():
        got = rows.get(name)
        if got != want:
            raise LandingError(f"{stage}：{name} 应为 {want}，实际 {got}")


def _assert_public_semantics(rel_rows: list[dict[str, str]], ev_rows: list[dict[str, str]]) -> None:
    """公开层语义守门：每条公开关系必须有合格 support 证据，且不得带任何保守拦截标记。"""
    by_rel: dict[str, list[dict[str, str]]] = {}
    for row in ev_rows:
        by_rel.setdefault(row["relation_id"].strip(), []).append(row)
    public = [r for r in rel_rows if (r.get("publish_status") or "").strip() in PUBLIC_RELATION_STATUSES]
    if not public:
        raise LandingError("落地后公开层仍为空，未达成目标")
    for row in public:
        rid = row["relation_id"].strip()
        if (row.get("publish_status") or "").strip() != "supported":
            raise LandingError(f"{rid} 本批只允许 derived supported，实际 {row.get('publish_status')!r}")
        if (row.get("publish_status_origin") or "").strip() != "derived":
            raise LandingError(f"{rid} origin 必须为 derived，实际 {row.get('publish_status_origin')!r}")
        if (row.get("relation_risk_level") or "").strip().lower() in ("critical", "high"):
            raise LandingError(f"{rid} 公开层不得含 critical/high 风险关系")
        if (row.get("needs_manual_review") or "").strip().lower() == "yes":
            raise LandingError(f"{rid} 公开层不得含 needs_manual_review=yes")
        if (row.get("confidence") or "").strip().lower() == "low":
            raise LandingError(f"{rid} 公开层不得含 low 置信关系")
        ftype = (row.get("final_relation_type") or "").strip()
        if ftype == "待核验" or ftype in INFERRED_RELATION_TYPES:
            raise LandingError(f"{rid} 公开层不得含待核验或推断类型（{ftype}）")
        qualified = [
            e for e in by_rel.get(rid, [])
            if e["evidence_support"] == "support" and e["review_status"] != "rejected"
            and (e.get("locator") or "").strip()
            and ((e.get("quote") or "").strip() or (e.get("context") or "").strip())
        ]
        if not qualified:
            raise LandingError(f"{rid} 公开层关系缺少合格 support 证据")
        if not any(BATCH_MARKER in (e.get("reviewer_note") or "") for e in qualified):
            raise LandingError(f"{rid} 合格 support 证据不来自本批落地，疑似既有数据漂移")


def apply_landing(data_dir: Path, dry_run: bool = False) -> dict:
    data_dir = Path(data_dir)
    rel_cols, rel_rows = _read(data_dir / "person_relations.csv")
    ev_cols, ev_rows = _read(data_dir / "relation_evidences.csv")
    src_cols, src_rows = _read(data_dir / "sources.csv")
    pas_cols, pas_rows = _read(data_dir / "source_passages.csv")

    if already_landed(ev_rows):
        _check_counts({
            "person_relations.csv": len(rel_rows),
            "relation_evidences.csv": len(ev_rows),
            "sources.csv": len(src_rows),
            "source_passages.csv": len(pas_rows),
        }, EXPECTED_POST, "二跑幂等校验")
        return {"status": "no-op",
                "message": "无新增/已完成：本批落地痕迹已存在且计数符合预期，跳过写入。"}

    _check_counts({
        "person_relations.csv": len(rel_rows),
        "relation_evidences.csv": len(ev_rows),
        "sources.csv": len(src_rows),
        "source_passages.csv": len(pas_rows),
    }, EXPECTED_BASELINE, "落地前基线")

    if {r["evidence_support"].strip() for r in ev_rows} != {"associated"}:
        raise LandingError("基线漂移：relation_evidences.evidence_support 应全为 associated")
    if {r["review_status"].strip() for r in ev_rows} != {"pending"}:
        raise LandingError("基线漂移：relation_evidences.review_status 应全为 pending")
    risk_before = sum(1 for r in rel_rows if (r.get("relation_risk_level") or "").strip() == "critical")

    adj = load_adjudication(ADJUDICATED)
    overnight_relations, evidences = load_overnight(adj)
    target = build_target_set(overnight_relations, adj)
    recheck = verify_quotes(target, evidences)

    persons = load_persons(PERSONS)
    attested, unattested = gate_attestation(target, evidences, persons)
    queue = build_recapture_queue(unattested, evidences, persons, recheck)

    needed = sorted({eid for rec, _ in attested for eid in rec["evidence_ids"]})
    src_by_id = {r["source_id"].strip(): r for r in src_rows}
    mapping, added_sources, registered = resolve_sources(src_cols, src_rows, evidences, needed)
    for row in added_sources:
        strength = (row.get("evidence_strength") or "").strip()
        if strength not in STRENGTH_TO_LEVEL:
            raise LandingError(f"新来源 evidence_strength 无法映射 source_level：{strength!r}")

    rel_by_id = {r["relation_id"].strip(): r for r in rel_rows}
    next_n = _next_rele_id([r["relation_evidence_id"] for r in ev_rows])
    ledger: list[dict[str, str]] = []
    new_evidence: list[dict[str, str]] = []
    touched: list[str] = []

    for rec, basis in sorted(attested, key=lambda item: item[0]["relation_id"]):
        rid = rec["relation_id"]
        row = rel_by_id.get(rid)
        if row is None:
            raise LandingError(f"{rid} 不在 person_relations.csv 中")
        verdict = adj[rid]["human_verdict"].strip()
        type_before = (row.get("final_relation_type") or "").strip()
        status_before = (row.get("publish_status") or "").strip()
        if verdict == "wrong_type":
            suggested = (adj[rid].get("ai_suggested_type") or "").strip()
            if not suggested:
                raise LandingError(f"{rid} 裁决为 wrong_type 但缺少建议类型")
            row["standard_relation_type"] = suggested
            row["final_relation_type"] = suggested
            reason = (row.get("correction_reason") or "").strip()
            addition = (
                f"{BATCH_MARKER}：人工裁决 wrong_type，类型由「{type_before}」改为「{suggested}」"
                f"（{AUTHORIZED_BY} {AUTHORIZED_AT} 授权按建议执行）。"
            )
            row["correction_reason"] = f"{reason}｜{addition}" if reason else addition
        eid = rec["evidence_ids"][0]
        ev = evidences[eid]
        sid = mapping[eid]
        strength = (src_by_id.get(sid, {}).get("evidence_strength") or "").strip()
        if not strength:
            strength = next(
                (r.get("evidence_strength", "") for r in added_sources if r["source_id"] == sid), ""
            )
        strength = (strength or "").strip()
        if strength not in STRENGTH_TO_LEVEL:
            raise LandingError(f"{eid} 来源 {sid} evidence_strength 无法映射 source_level：{strength!r}")
        newly_registered = any(a["source_id"] == sid for a in added_sources)
        rele_id = f"RELE-{next_n:05d}"
        next_n += 1
        new_evidence.append(build_evidence_row(
            ev_cols, rele_id, rid, sid, ev, rec,
            STRENGTH_TO_LEVEL[strength], recheck[eid], newly_registered,
        ))
        touched.append(rid)
        pa = persons[rec["source_person_id"]]
        pb = persons[rec["target_person_id"]]
        ledger.append({
            "relation_id": rid,
            "source_person_id": rec["source_person_id"],
            "target_person_id": rec["target_person_id"],
            "person_a_name": (pa.get("standard_name") or "").strip(),
            "person_b_name": (pb.get("standard_name") or "").strip(),
            "batch_marker": BATCH_MARKER,
            "attestation_basis": basis,
            "human_verdict": verdict,
            "final_relation_type_before": type_before,
            "final_relation_type_after": (row.get("final_relation_type") or "").strip(),
            "publish_status_before": status_before,
            "publish_status_after": "",
            "publish_status_origin": "",
            "new_relation_evidence_id": rele_id,
            "overnight_evidence_id": eid,
            "source_id": sid,
            "source_registered": "yes" if newly_registered else "no",
            "locator": str(ev.get("locator", "")).strip(),
            "quote_sha256": str(ev.get("quote_sha256", "")),
            "search_log_id": ";".join(rec.get("search_log_ids") or []),
            "quote_verbatim_recheck": recheck[eid],
            "authorized_by": AUTHORIZED_BY,
            "authorized_at": AUTHORIZED_AT,
            "authorization_quote": AUTHORIZATION_QUOTE,
        })

    if len(new_evidence) != EXPECTED_NEW_EVIDENCE_ROWS:
        raise LandingError(f"新增证据行应为 {EXPECTED_NEW_EVIDENCE_ROWS}，实际 {len(new_evidence)}")
    ids = [r["relation_evidence_id"] for r in ev_rows] + [r["relation_evidence_id"] for r in new_evidence]
    if len(set(ids)) != len(ids):
        raise LandingError("relation_evidence_id 出现重复")
    ev_rows.extend(new_evidence)
    if added_sources:
        src_rows.extend(added_sources)
        src_by_id.update({r["source_id"].strip(): r for r in added_sources})

    by_rel: dict[str, list[dict[str, str]]] = {}
    for row in ev_rows:
        by_rel.setdefault(row["relation_id"].strip(), []).append(row)
    for rid in touched:
        row = rel_by_id[rid]
        cols = assign_relation_status_columns(row, by_rel.get(rid, []), existing={
            "reviewer": row.get("reviewer", ""), "reviewed_at": row.get("reviewed_at", ""),
            "review_note": row.get("review_note", ""),
        })
        row["publish_status"] = cols["publish_status"]
        row["publish_status_origin"] = cols["publish_status_origin"]
        row["reviewer"] = cols["reviewer"]
        row["reviewed_at"] = cols["reviewed_at"]
        row["review_note"] = cols["review_note"]
    corrections = 0
    for entry in ledger:
        row = rel_by_id[entry["relation_id"]]
        entry["publish_status_after"] = row["publish_status"]
        entry["publish_status_origin"] = row["publish_status_origin"]
        if entry["final_relation_type_before"] != entry["final_relation_type_after"]:
            corrections += 1

    _check_counts({
        "person_relations.csv": len(rel_rows),
        "relation_evidences.csv": len(ev_rows),
        "sources.csv": len(src_rows),
        "source_passages.csv": len(pas_rows) + len(added_sources),
    }, EXPECTED_POST, "落地后计数")
    actual_dist = {k: v for k, v in Counter(
        (r.get("publish_status") or "").strip() for r in rel_rows).items() if v}
    if actual_dist != EXPECTED_POST_PUBLISH_STATUS:
        raise LandingError(f"落地后发布状态分布不符：期望 {EXPECTED_POST_PUBLISH_STATUS}，实际 {actual_dist}")
    if corrections != EXPECTED_TYPE_CORRECTIONS:
        raise LandingError(f"类型更正应为 {EXPECTED_TYPE_CORRECTIONS} 条，实际 {corrections}")
    risk_after = sum(1 for r in rel_rows if (r.get("relation_risk_level") or "").strip() == "critical")
    if risk_after != risk_before:
        raise LandingError(f"relation_risk_level 被改写：critical {risk_before} -> {risk_after}（禁止反向降险）")
    if {r.get("publish_status_origin", "").strip() for r in rel_rows} != {"derived"}:
        raise LandingError("本批只允许 derived origin，出现 human_adjudication")
    _assert_public_semantics(rel_rows, ev_rows)
    public_n = sum(1 for r in rel_rows if (r.get("publish_status") or "").strip() in PUBLIC_RELATION_STATUSES)
    if public_n != EXPECTED_PUBLIC_SUPPORTED:
        raise LandingError(f"公开层关系应为 {EXPECTED_PUBLIC_SUPPORTED} 条，实际 {public_n}")

    summary = {
        "attested": len(attested), "unattested": len(unattested),
        "public_supported": public_n, "type_corrections": corrections,
        "new_evidence_rows": len(new_evidence), "registered_sources": registered,
        "queue": queue, "ledger": ledger,
    }
    if dry_run:
        return {"status": "dry-run", "message": "校验全部通过（未写入）。", **summary}

    _write(data_dir / "person_relations.csv", rel_cols, rel_rows)
    _write(data_dir / "relation_evidences.csv", ev_cols, ev_rows)
    if added_sources:
        _write(data_dir / "sources.csv", src_cols, src_rows)
        sync_source_layer(data_dir)
    return {"status": "applied", "message": "落地完成。", **summary}


def write_ledger(ledger: list[dict[str, str]], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=LEDGER_COLUMNS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(ledger)


def write_report(result: dict, ledger: list[dict[str, str]], path: Path) -> None:
    public = [e for e in ledger if e["publish_status_after"] in PUBLIC_RELATION_STATUSES]
    corrections = [e for e in ledger if e["final_relation_type_before"] != e["final_relation_type_after"]]
    dist = Counter(e["publish_status_after"] for e in ledger)
    lines = [
        "# Phase 5 关系裁决生产层落地报告（Tier 1：证据派生 supported）",
        "",
        "> 执行脚本：`research/analysis/apply_phase5_relation_landing.py`（幂等）  ",
        f"> 批次标记：`{BATCH_MARKER}`  落地日期：{LANDED_AT}",
        "",
        "## 0. 授权与口径声明",
        "",
        f"- 授权人：{AUTHORIZED_BY}；授权时间：{AUTHORIZED_AT}；签核路径：C（全量按 AI 建议执行）。",
        f"- 授权语（逐字）：「{AUTHORIZATION_QUOTE}」，登记于 `phase5_review_adjudication_record.md`。",
        "- 裁决口径为**授权按建议执行**，非逐条独立人工复核；公开层展示与答辩引用均须带此口径。",
        "- 本批公开层只接受 **derived `supported`**（证据派生），未使用 `human_adjudication` 通道，"
        "未把 critical/high 风险关系推入公开层，未改写 `relation_risk_level`。",
        "",
        "## 1. 落地范围",
        "",
        f"- 候选关系：{EXPECTED_TARGET_RELATIONS} 条（夜间轮判定 `evidence_support=support` ∧ 人工裁决 correct/wrong_type）。",
        f"- **双方佐证门通过：{result.get('attested', 0)} 条**（引文空白归一后同时含两端姓名/别名；"
        "鲁迅日记按日记体免甲方自名）。只有通过的才落地。",
        f"- 双方佐证门未通过：{result.get('unattested', 0)} 条 → 全部转入 `phase5_quote_recapture_queue.csv`"
        "（`pending_human_review`，生产层零改动）。",
        f"- 引文逐字复核：{EXPECTED_DISTINCT_EVIDENCES}/{EXPECTED_DISTINCT_EVIDENCES} 通过"
        "（quote_sha256 与原文 content_hash 一致；OCR 全文字间带空格，须空白归一后比对）。",
        f"- 新增关系证据行：{result.get('new_evidence_rows', 0)} 条。",
        f"- 新注册来源：{len(result.get('registered_sources') or [])} 条"
        + (f"（{', '.join(result.get('registered_sources') or [])}）" if result.get("registered_sources") else "") + "。",
        f"- 关系类型更正：{len(corrections)} 条。",
        "",
        "## 2. 结果分布（仅本批 48 条）",
        "",
        "| 落地后 publish_status | 条数 |",
        "| --- | ---: |",
    ]
    for key in ("supported", "pending_review", "inferred"):
        lines.append(f"| {key} | {dist.get(key, 0)} |")
    lines += [
        "",
        f"全表发布状态：supported {EXPECTED_POST_PUBLISH_STATUS['supported']} / "
        f"pending_review {EXPECTED_POST_PUBLISH_STATUS['pending_review']} / "
        f"inferred {EXPECTED_POST_PUBLISH_STATUS['inferred']}（合计 4238，origin 全为 derived）。",
        "",
        "## 3. 进入公开层的关系",
        "",
        "| relation_id | 人物对 | 类型（落地后） | 证据 locator | 新证据行 |",
        "| --- | --- | --- | --- | --- |",
    ]
    for e in sorted(public, key=lambda x: x["relation_id"]):
        lines.append(
            f"| {e['relation_id']} | {e['source_person_id']} → {e['target_person_id']} "
            f"| {e['final_relation_type_after']} | {e['locator']} | {e['new_relation_evidence_id']} |"
        )
    lines += [
        "",
        "## 4. 类型更正明细",
        "",
        "| relation_id | 更正前 | 更正后 | 落地后状态 |",
        "| --- | --- | --- | --- |",
    ]
    for e in sorted(corrections, key=lambda x: x["relation_id"]):
        lines.append(
            f"| {e['relation_id']} | {e['final_relation_type_before']} "
            f"| {e['final_relation_type_after']} | {e['publish_status_after']} |"
        )
    lines += [
        "",
        "## 5. 本批发现的系统性缺陷：引文窗口捕错",
        "",
        "夜间轮记录的 `quote` 多数是按页码定位截取的固定窗口，而不是真正记载该关系的那一句，",
        "导致 `reason` 字段描述的史实与 `quote` 内容不一致。48 条候选中：",
        "",
        f"- {result.get('attested', 0)} 条引文确实同时记载双方当事人 → 已落地；",
        f"- {result.get('unattested', 0)} 条引文里找不到当事人（16 条两端都缺、11 条只缺乙方、1 条只缺甲方）。",
        "",
        "未通过的 28 条按下述三类进入重捕队列，**生产层零改动**：",
        "",
        "| proposed_action | 含义 | 条数 |",
        "| --- | --- | ---: |",
    ]
    qdist = Counter(e["proposed_action"] for e in (result.get("queue") or []))
    for key, desc in (
        ("recapture_quote_then_regrade", "同一本地原文中存在同时记载两人的叙述性段落，可重捕引文后再评级"),
        ("cooccurrence_only_keep_associated", "同窗共现仅为顿号人名罗列，属关联级，不得升为 support"),
        ("no_local_support_mark_insufficient", "本地原文中找不到任何同窗共现，应判证据不足"),
    ):
        lines.append(f"| `{key}` | {desc} | {qdist.get(key, 0)} |")
    lines += [
        "",
        "结论：`evidence_support=support` 不能单独作为公开依据，必须再过双方佐证门。",
        "该门已固化在 `research/analysis/quote_attestation.py`，可复用于后续所有关系补证批次。",
        "",
        "## 6. 明确未做的事（避免夸大结论）",
        "",
        "- **未改判既有 10249 条 `associated` 证据**：其 reviewer_note 写明「仅表示来源关联，不断言支持强度」，"
        "改判为 support 等于凭空提升证据等级。本批改为新增经逐字复核的 support 行，旧行原样保留。",
        "- **未落地 associated 层 113 条**（人工裁决成立但证据仅为关联级）：证据等级不足以进入公开层，保留在研究层。",
        "- **未把 239 条 `not_supported` 写为 `rejected`**：五态中 `rejected` 语义是「人工否定」（对应 contradicted，本批 0 条），"
        "`not_supported` 只表示证据不足；映射会把「未证实」夸大成「已证伪」。这些关系仍为 pending_review/inferred，不进入公开层。",
        "- **未使用 `human_adjudication` 通道**：通过佐证门但被机器保守标记挡住的关系仍留在 pending_review，"
        "要公开它们需要逐条独立人工复核，不能由一次概括授权代行。",
        "- **未对未通过佐证门的关系做任何类型更正或降级**：引文捕错不等于关系不成立，"
        "改类型或判不足都需要重捕引文后另行裁决。",
        "- **未改写 `relation_risk_level`**：critical 计数保持 1974。",
        "",
        "## 7. 复核方式",
        "",
        "```powershell",
        "python research/analysis/apply_phase5_relation_landing.py --dry-run   # 只校验不写",
        "python research/analysis/apply_phase5_relation_landing.py             # 落地（二跑输出「无新增/已完成」）",
        "python research/analysis/build_publish_data.py",
        "python build_static_site.py",
        "python -m pytest -q",
        "```",
        "",
        f"逐条痕迹见 `phase5_relation_landing_ledger.csv`（{len(ledger)} 行，每行带 quote_sha256、检索凭据与授权溯源）。",
        "",
    ]
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines), encoding="utf-8")


def verify_post_state(data_dir: Path) -> dict[str, int]:
    from kb_schema import validate_data_dir

    counts = {}
    for name in EXPECTED_POST:
        _, rows = _read(data_dir / name)
        counts[name] = len(rows)
    _check_counts(counts, EXPECTED_POST, "写后复核")
    result = validate_data_dir(data_dir)
    if result.errors:
        raise LandingError(f"写后 Schema 校验出现 {len(result.errors)} 个错误：{result.errors[:3]}")
    return {"counts": counts, "schema_errors": len(result.errors), "schema_warnings": len(result.warnings)}


def main() -> int:
    parser = argparse.ArgumentParser(description="Phase 5 关系裁决生产层落地（Tier 1，幂等）。")
    parser.add_argument("--data-dir", type=Path, default=DATA, help="生产数据目录")
    parser.add_argument("--dry-run", action="store_true", help="只执行全部校验，不写任何文件")
    parser.add_argument("--skip-report", action="store_true", help="不写台账与报告（沙盒测试用）")
    args = parser.parse_args()

    try:
        result = apply_landing(args.data_dir, dry_run=args.dry_run)
    except LandingError as exc:
        print(f"落地失败，未写入任何文件：{exc}")
        return 1

    print(result["message"])
    if result["status"] == "no-op":
        return 0
    print(f"候选 {EXPECTED_TARGET_RELATIONS} 条 → 双方佐证门通过 {result['attested']} 条 / 未通过 {result['unattested']} 条")
    print(f"公开层 derived supported：{result['public_supported']} 条")
    print(f"新增关系证据行：{result['new_evidence_rows']} 条")
    print(f"关系类型更正：{result.get('type_corrections', 0)} 条")
    print(f"新注册来源：{result.get('registered_sources') or '无'}")

    if result["status"] == "applied":
        post = verify_post_state(args.data_dir)
        print(f"写后复核：计数 {post['counts']}；Schema {post['schema_errors']} errors / {post['schema_warnings']} warnings")
    if not args.skip_report and not args.dry_run:
        adj_for_queue = load_adjudication(ADJUDICATED)
        write_recapture_queue(result["queue"], RECAPTURE_QUEUE, adj_for_queue)
        print(f"重捕队列：{RECAPTURE_QUEUE.relative_to(ROOT)}（{len(result['queue'])} 条，pending_human_review）")
        write_ledger(result["ledger"], LEDGER)
        write_report(result, result["ledger"], REPORT_MD)
        print(f"台账：{LEDGER.relative_to(ROOT)}")
        print(f"报告：{REPORT_MD.relative_to(ROOT)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
