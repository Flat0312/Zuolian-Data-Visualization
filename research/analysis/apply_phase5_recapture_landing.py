"""Phase 5 重捕候选第二批生产层落地（2026-10-08 逐条独立人工裁决，幂等）。

授权依据
--------
用户 2026-10-08 会话逐条裁决，逐字授权语：
「第一优先REL-00622 和 REL-01368直接过，REL-01891判"证据不足"，REL-01161 成立但不影响公开层」。
这是**逐条独立裁决**，口径强于第一批（2026-09-20「授权按建议执行」的概括授权，
签核路径 C）；台账、裁决记录与站点文案必须区分两种口径，不得混同。

四条裁决的落地语义（照裁决逐条实现，不自行发挥）
------------------------------------------------
- REL-00622（周扬—邵荃麟）：过。新增一条 support 证据行；类型不改；
  预期 derived ``supported`` 进入公开层。
- REL-01368（郁达夫—陈望道）：过。新增 support 证据行；类型 交游→签名联署
  （``standard_relation_type`` 与 ``final_relation_type`` 同步改，
  ``correction_reason`` 追加不覆盖）；预期 derived ``supported``。
- REL-01161（阳翰笙—林淡秋）：成立但不进公开层。新增 support 证据行；类型不改；
  由发布门禁自然挡在 ``pending_review``（critical 风险 + low 置信 + 推断类型三重拦截）。
  **不使用** ``human_adjudication`` 通道放行。
- REL-01891（叶紫—萧军）：证据不足。**不新增任何证据行**，生产表零改动，保持
  ``inferred``。重切段落是书目著录（"…传记小说。李克因作，载《东方纪事》1987年…
  叙述…叶紫同…萧军…的交往"），属二手著录，不足以定为 support。
  也不写 ``rejected``：五态中 rejected 语义是人工否定，映射会把「未证实」夸大成「已证伪」。

明确不做的事
------------
- 不改门禁：``relation_publish_status.py`` 的判定顺序、``INFERRED_RELATION_TYPES``、
  ``PUBLIC_RELATION_STATUSES`` 一律不动；公开层只接受 derived ``supported``。
- 不改判既有行：10249 条 ``associated`` 与第一批 20 条 ``support`` 原样保留。
- 不注册新来源：四条 locator 全部复用既有 ``source_id``（SRC-0779 / SRC-0054 /
  SRC-0739；SRC-0946 本批用不到）。
- 不改 ``relation_risk_level`` / ``needs_manual_review`` / ``confidence``。
- 不改候选包与队列文件：裁决另出 ``phase5_quote_recapture_adjudicated.csv``。

运行时校验（任一失败即整体退出、不写任何文件；全部前置到任何写盘动作之前）
------------------------------------------------------------------------
1. 候选包 28 行、全部 ``pending_human_review``；四条裁决行存在且字段自洽
   （recapture_status/quote_form/locator_agrees/quote_char_len）；
2. **登记裁决表（``phase5_quote_recapture_adjudicated.csv``）逐行三列授权校验**
   （2026-10-08 返工，缺陷 1）：四条已裁决行的 ``authorized_by``/``authorized_at``
   非空且与登记一致、``authorization_quote`` 非空、不含执行者自授占位
   （大小写不敏感）、且逐字等于模块常量 ``AUTHORIZATION_QUOTE``；
   24 条未裁决行不得携带任何授权列内容（防后门授权）；
3. **裁决表与内置 ``ADJUDICATION`` 双向交叉校验**（返工，缺陷 2）：行数与 relation_id
   集合恰与候选包一致；携带批次标记的裁决行集合恰为 ``ADJUDICATION`` 的 4 条
   （多、少、改都拒绝）；每行 ``adjudication_2026_10_08``/``evidence_landed``/
   ``final_relation_type_after``/``publish_status_after`` 与内置裁决语义一致；
   每行 ``quote_sha256``/``recorded_locator`` 与候选包同 relation_id 行一致；
   落地行的 ``new_relation_evidence_id`` 与本次计算的证据号一致；
4. 三条落地引文（含 REL-01891 留档）按候选包 ``source_file`` + ``normalized_start/end``
   在空白归一后的本地原文中**逐字回定位取回**，``quote_sha256`` 自洽；
5. 三条落地引文过双方佐证门（复用 ``quote_attestation``，不放宽）；
6. 每个 locator 在 sources.csv 按 citation 唯一命中既有 source_id 且与预期一致；
7. 落地前基线、落地后终值全部为硬后置条件（含公开层 7、critical 1974 不变、
   REL-01891 零痕迹、REL-01161 不公开、类型更正恰 1）；
8. 落地后全表按门禁重算零漂移、Schema 0 错误。

幂等：存在本批标记行时只校验「本批标记行恰 3 条」即输出「无新增/已完成」并零写入，
对后续批次免疫。原子写（tmp + ``os.replace``）；CSV 为 utf-8-sig + CRLF + QUOTE_MINIMAL。
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import os
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
DATA = ROOT / "data" / "processed"
REPORTS = ROOT / "research" / "drafts" / "reports"
PERSONS_CSV = DATA / "persons.csv"
RUNTIME_SOURCES = DATA / "runtime_sources"

CANDIDATES = REPORTS / "phase5_quote_recapture_candidates.csv"
ADJUDICATED = REPORTS / "phase5_quote_recapture_adjudicated.csv"
ADJUDICATION_RECORD = REPORTS / "phase5_quote_recapture_adjudication_record.md"
LEDGER = REPORTS / "phase5_recapture_landing_ledger.csv"
REPORT_MD = REPORTS / "phase5_recapture_landing_report.md"

# 导入不得污染 sys.path 顺序（与第一批脚本同一约束）：优先包路径，脚本直跑才回退 append。
_ANALYSIS_DIR = Path(__file__).resolve().parent
try:
    from research.analysis.quote_attestation import (
        name_candidates,
        normalized_bundle,
        quote_attests_pair,
    )
    from research.analysis.relation_publish_status import (
        INFERRED_RELATION_TYPES,
        PUBLIC_RELATION_STATUSES,
        assign_relation_status_columns,
        derive_relation_publish_status_with_origin,
    )
except ImportError:
    if str(_ANALYSIS_DIR) not in sys.path:
        sys.path.append(str(_ANALYSIS_DIR))
    from quote_attestation import (  # noqa: E402
        name_candidates,
        normalized_bundle,
        quote_attests_pair,
    )
    from relation_publish_status import (  # noqa: E402
        INFERRED_RELATION_TYPES,
        PUBLIC_RELATION_STATUSES,
        assign_relation_status_columns,
        derive_relation_publish_status_with_origin,
    )

BATCH_MARKER = "P5-RECAPTURE-2026-10-08"
PREV_BATCH_MARKER = "P5-LANDING-2026-09-28"
AUTHORIZED_BY = "用户（逐条独立裁决）"
AUTHORIZED_AT = "2026-10-08"
AUTHORIZATION_QUOTE = (
    '第一优先REL-00622 和 REL-01368直接过，REL-01891判"证据不足"，REL-01161 成立但不影响公开层'
)
LANDED_AT = "2026-10-08"

# 四条裁决的精确语义。land=是否新增 support 证据行；type_after=类型更正目标（None 不改）。
ADJUDICATION = {
    "REL-00622": {
        "verdict": "过（直接进公开层）",
        "land": True,
        "type_after": None,
        "expect_publish": "supported",
        "note": "叙述性同往内山书店记载，非名单罗列。",
    },
    "REL-01368": {
        "verdict": '过（直接进公开层）；类型更正 交游→签名联署',
        "land": True,
        "type_after": "签名联署",
        "expect_publish": "supported",
        "note": "《为横死之小林遗族募捐启》9 人签署名单为共同联署直接记载。",
    },
    "REL-01161": {
        "verdict": "成立但不影响公开层",
        "land": True,
        "type_after": None,
        "expect_publish": "pending_review",
        "note": "门禁自然拦截（critical 风险 + low 置信 + 同属组织推断类型）；"
                "未使用 human_adjudication 通道放行。",
    },
    "REL-01891": {
        "verdict": "证据不足（生产层零改动）",
        "land": False,
        "type_after": None,
        "expect_publish": "inferred",
        "note": "重切段落为书目著录（二手著录=yes），不足以定为 support；"
                "不写 rejected（rejected 语义是人工否定，本条只是未证实）。",
    },
}
LAND_IDS = sorted(rid for rid, spec in ADJUDICATION.items() if spec["land"])
INSUFFICIENT_IDS = sorted(rid for rid, spec in ADJUDICATION.items() if not spec["land"])

STRENGTH_TO_LEVEL = {"一手": "A", "二手": "B", "转引": "C", "参考": "D", "推断": "D"}
EXPECTED_SOURCE_REUSE = {
    "REL-00622": "SRC-0779",
    "REL-01368": "SRC-0054",
    "REL-01161": "SRC-0739",
    "REL-01891": "SRC-0946",  # 本批用不到，仅留档
}
# 执行者自授占位（大小写不敏感）：授权语命中任一子串即拒绝，与第一批 load_adjudication 同一防线。
SELF_AUTH_PLACEHOLDERS = ("ai 自行决定", "执行者自授", "模型决定")

# 落地前基线（2026-10-08 实测）
EXPECTED_BASELINE_COUNTS = {
    "person_relations.csv": 4238,
    "relation_evidences.csv": 10269,
    "sources.csv": 1178,
    "source_passages.csv": 1178,
    "source_works.csv": 65,
}
EXPECTED_BASELINE_STATUS = {"supported": 5, "pending_review": 2451, "inferred": 1782}
EXPECTED_BASELINE_SUPPORT = {"associated": 10249, "support": 20}
EXPECTED_BASELINE_REVIEWED = 20
EXPECTED_BASELINE_CRITICAL = 1974

# 落地后硬后置条件（任务书第 3 节，任一不符即整体退出、不写任何文件）
EXPECTED_LANDED_ROWS = 3
EXPECTED_POST_COUNTS = {
    "person_relations.csv": 4238,
    "relation_evidences.csv": 10272,
    "sources.csv": 1178,
    "source_passages.csv": 1178,
    "source_works.csv": 65,
}
EXPECTED_POST_STATUS = {"supported": 7, "pending_review": 2451, "inferred": 1780}
EXPECTED_POST_SUPPORT = {"associated": 10249, "support": 23}
EXPECTED_POST_REVIEWED = 23
EXPECTED_PUBLIC_IDS = {
    "REL-00046", "REL-00059", "REL-00097", "REL-03289", "REL-03518",  # 第一批
    "REL-00622", "REL-01368",  # 本批新增
}
EXPECTED_TYPE_CORRECTIONS = 1
EXPECTED_CANDIDATE_ROWS = 28

LEDGER_COLUMNS = [
    "relation_id", "person_a_id", "person_a_name", "person_b_id", "person_b_name",
    "batch_marker", "adjudication", "landed_this_batch", "evidence_landed",
    "new_relation_evidence_id", "candidate_human_verdict",
    "final_relation_type_before", "final_relation_type_after",
    "publish_status_before", "publish_status_after", "publish_status_origin",
    "reused_source_id", "source_registered", "locator",
    "normalized_start", "normalized_end", "quote_sha256",
    "quote_verbatim_recheck", "attestation_basis",
    "authorized_by", "authorized_at", "authorization_quote",
    "candidate_row_source", "note",
]
ADJUDICATED_COLUMNS = [
    "relation_id", "person_a_id", "person_a_name", "person_b_id", "person_b_name",
    "current_final_relation_type", "proposed_relation_type", "recapture_status",
    "quote_sha256", "recorded_locator", "candidate_human_verdict",
    "adjudication_2026_10_08", "evidence_landed", "new_relation_evidence_id",
    "reused_source_id", "final_relation_type_after", "publish_status_after",
    "quote_verbatim_recheck", "batch_marker",
    "authorized_by", "authorized_at", "authorization_quote",
    "note",
]


class LandingError(RuntimeError):
    """任何前置/后置校验失败都抛出本异常，调用方保证不写任何文件。"""


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


def _write_md(path: Path, text: str) -> None:
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(text, encoding="utf-8")
    os.replace(tmp, path)


def _sha256_text(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _next_rele_id(existing_ids: list[str]) -> int:
    max_n = 0
    for value in existing_ids:
        m = re.fullmatch(r"RELE-(\d+)", str(value).strip())
        if m:
            max_n = max(max_n, int(m.group(1)))
    return max_n + 1


def _check_counts(rows: dict[str, int], expected: dict[str, int], stage: str) -> None:
    for name, want in expected.items():
        got = rows.get(name)
        if got != want:
            raise LandingError(f"{stage}：{name} 应为 {want}，实际 {got}")


def load_candidates(path: Path) -> dict[str, dict[str, str]]:
    cols, rows = _read(path)
    if len(rows) != EXPECTED_CANDIDATE_ROWS:
        raise LandingError(f"候选包应为 {EXPECTED_CANDIDATE_ROWS} 行，实际 {len(rows)}")
    for col in ("relation_id", "recaptured_quote", "quote_sha256", "normalized_start",
                "normalized_end", "source_file", "recorded_locator", "review_status",
                "human_verdict", "proposed_relation_type"):
        if col not in cols:
            raise LandingError(f"候选包缺列 {col}")
    by_id: dict[str, dict[str, str]] = {}
    for row in rows:
        rid = row["relation_id"].strip()
        if rid in by_id:
            raise LandingError(f"候选包 relation_id 重复：{rid}")
        if row["review_status"].strip() != "pending_human_review":
            raise LandingError(f"{rid} 候选行 review_status 非 pending_human_review，疑似已被改动")
        by_id[rid] = row
    missing = [rid for rid in ADJUDICATION if rid not in by_id]
    if missing:
        raise LandingError(f"候选包缺少已裁决关系：{missing}")
    return by_id


def validate_adjudicated(cands: dict[str, dict[str, str]]) -> dict[str, dict[str, str]]:
    """对登记裁决表做前置校验（2026-10-08 返工），任一不符抛 LandingError、调用方零写入。

    裁决表不只是产物：它先作为**输入**被校验，与内置 ``ADJUDICATION``、候选包双向交叉，
    保证「授权凭据」与「代码行为」不可能静默不一致。

    第一道（缺陷 1，逐行授权三列）：已裁决行的 ``authorized_by``/``authorized_at`` 非空且
    与登记一致、``authorization_quote`` 非空、不含执行者自授占位、且逐字等于
    ``AUTHORIZATION_QUOTE``；未裁决行不得携带任何授权内容。
    第二道（缺陷 2，与内置裁决交叉）：行数与 relation_id 集合恰与候选包一致；
    裁决行集合恰为 ``ADJUDICATION`` 四条；每行裁决语义/落地标志/落地后类型与状态/
    quote_sha256/locator 与内置字典及候选包一致。

    返回 relation_id -> 裁决行（四条），供后续证据号一致性复核。
    """
    if not ADJUDICATED.exists():
        raise LandingError(f"登记裁决表不存在，无法完成授权溯源前置校验：{ADJUDICATED}")
    try:
        cols, rows = _read(ADJUDICATED)
    except OSError as exc:
        raise LandingError(f"登记裁决表不可读：{ADJUDICATED}") from exc
    required_cols = {
        "relation_id", "batch_marker", "adjudication_2026_10_08", "evidence_landed",
        "new_relation_evidence_id", "quote_sha256", "recorded_locator",
        "final_relation_type_after", "publish_status_after",
        "authorized_by", "authorized_at", "authorization_quote",
    }
    missing = sorted(required_cols - set(cols))
    if missing:
        raise LandingError(f"登记裁决表缺少必需列：{missing}")
    if len(rows) != len(cands):
        raise LandingError(
            f"登记裁决表应恰为候选包的 {len(cands)} 行（28 行全量处置），实际 {len(rows)} 行"
        )
    rids = [r["relation_id"].strip() for r in rows]
    if len(set(rids)) != len(rids):
        raise LandingError("登记裁决表 relation_id 出现重复")
    if set(rids) != set(cands):
        raise LandingError(
            f"登记裁决表 relation_id 集合与候选包不一致：多出 {sorted(set(rids) - set(cands))}，"
            f"缺少 {sorted(set(cands) - set(rids))}"
        )
    by_rid = {r["relation_id"].strip(): r for r in rows}

    def _is_adjudicated(row: dict[str, str]) -> bool:
        return bool((row.get("batch_marker") or "").strip()) or (
            (row.get("adjudication_2026_10_08") or "").strip() not in ("", "unadjudicated")
        )

    adjudicated_rids = {rid for rid, row in by_rid.items() if _is_adjudicated(row)}
    if adjudicated_rids != set(ADJUDICATION):
        raise LandingError(
            "登记裁决表的已裁决行集合必须恰为内置 ADJUDICATION 的 4 条："
            f"期望 {sorted(ADJUDICATION)}，实际 {sorted(adjudicated_rids)}"
            "（多、少、改都拒绝——防止审计链被静默篡改）"
        )

    # 第一道 + 第二道：逐条对四个已裁决行做全字段交叉
    for rid in sorted(ADJUDICATION):
        spec = ADJUDICATION[rid]
        row = by_rid[rid]
        cand = cands[rid]
        authorized_by = (row["authorized_by"] or "").strip()
        authorized_at = (row["authorized_at"] or "").strip()
        quote = (row["authorization_quote"] or "").strip()
        if not authorized_by or not authorized_at or not quote:
            raise LandingError(f"{rid} 授权溯源列不完整（authorized_by/authorized_at/authorization_quote 须非空）")
        if any(p in quote.lower() for p in SELF_AUTH_PLACEHOLDERS):
            raise LandingError(f"{rid} 授权语疑似执行者自授占位，拒绝落地：{quote!r}")
        if quote != AUTHORIZATION_QUOTE:
            raise LandingError(f"{rid} 授权语与登记逐字授权语不一致：{quote!r}")
        if authorized_at != AUTHORIZED_AT:
            raise LandingError(f"{rid} 授权日期与本批登记不一致：{authorized_at!r}")
        if authorized_by != AUTHORIZED_BY:
            raise LandingError(f"{rid} 授权人与本批登记不一致：{authorized_by!r}")
        if (row["batch_marker"] or "").strip() != BATCH_MARKER:
            raise LandingError(f"{rid} 批次标记异常：{(row['batch_marker'] or '').strip()!r}")
        if (row["adjudication_2026_10_08"] or "").strip() != spec["verdict"]:
            raise LandingError(
                f"{rid} 裁决语义与内置 ADJUDICATION 不一致：期望 {spec['verdict']!r}，"
                f"实际 {(row['adjudication_2026_10_08'] or '').strip()!r}"
            )
        want_landed = "yes" if spec["land"] else "no"
        if (row["evidence_landed"] or "").strip() != want_landed:
            raise LandingError(
                f"{rid} evidence_landed 与内置裁决不一致：期望 {want_landed!r}，"
                f"实际 {(row['evidence_landed'] or '').strip()!r}（审计链被篡改，拒绝）"
            )
        if not spec["land"] and (row["new_relation_evidence_id"] or "").strip():
            raise LandingError(f"{rid} 裁决为不落地，但登记裁决表携带证据号")
        if (row["quote_sha256"] or "").strip() != (cand["quote_sha256"] or "").strip():
            raise LandingError(f"{rid} 登记裁决表 quote_sha256 与候选包不一致（引文被替换）")
        if (row["recorded_locator"] or "").strip() != (cand["recorded_locator"] or "").strip():
            raise LandingError(f"{rid} 登记裁决表 recorded_locator 与候选包不一致")
        want_type = spec["type_after"] or (cand["current_final_relation_type"] or "").strip()
        if (row["final_relation_type_after"] or "").strip() != want_type:
            raise LandingError(
                f"{rid} 裁决表落地后类型 {row['final_relation_type_after']!r} 与内置裁决 {want_type!r} 不一致"
            )
        if (row["publish_status_after"] or "").strip() != spec["expect_publish"]:
            raise LandingError(
                f"{rid} 裁决表落地后状态 {(row['publish_status_after'] or '').strip()!r} "
                f"与内置裁决 {spec['expect_publish']!r} 不一致"
            )

    # 反向防后门：未裁决行不得携带授权内容、批次标记或落地标记
    for rid in sorted(cands):
        if rid in ADJUDICATION:
            continue
        row = by_rid[rid]
        if (row["batch_marker"] or "").strip():
            raise LandingError(f"{rid} 未裁决行携带批次标记：{row['batch_marker']!r}")
        if (row["adjudication_2026_10_08"] or "").strip() != "unadjudicated":
            raise LandingError(f"{rid} 未裁决行裁决列异常：{(row['adjudication_2026_10_08'] or '').strip()!r}")
        if (row["evidence_landed"] or "").strip() != "no":
            raise LandingError(f"{rid} 未裁决行 evidence_landed 应为 no")
        for col in ("authorized_by", "authorized_at", "authorization_quote"):
            if (row[col] or "").strip():
                raise LandingError(f"{rid} 未裁决行不得携带授权列内容：{col}")

    return {rid: by_rid[rid] for rid in sorted(ADJUDICATION)}


def _resolve_text_file(raw: str) -> Path:
    p = Path(raw)
    if not p.exists():
        p = RUNTIME_SOURCES / Path(raw).name
    if not p.exists():
        raise LandingError(f"本地原文不存在，无法逐字复核：{raw}")
    return p


def verify_candidates(cands: dict[str, dict[str, str]]) -> dict[str, dict[str, str]]:
    """四条裁决行的字段自洽 + 逐字回定位 + sha256 自洽。返回 rid -> 复核结论。

    落地脚本只处理已裁决行，但四条全部复核（REL-01891 的复核结论用于留档，
    说明「引文可逐字取回」与「证据等级不足」是两回事）。
    """
    checks: dict[str, dict[str, str]] = {}
    cache: dict[str, tuple[str, str, list[int]]] = {}
    for rid in sorted(ADJUDICATION):
        row = cands[rid]
        quote = row["recaptured_quote"]
        if row["recapture_status"].strip() != "recaptured":
            raise LandingError(f"{rid} recapture_status 非 recaptured：{row['recapture_status']!r}")
        if row["quote_form"].strip() != "whitespace_normalized":
            raise LandingError(f"{rid} quote_form 非 whitespace_normalized")
        if row["locator_agrees"].strip() != "yes":
            raise LandingError(f"{rid} derived/recorded locator 不一致")
        if row["human_verdict"].strip() not in ("correct", "wrong_type"):
            raise LandingError(f"{rid} 候选包 human_verdict 非成立口径：{row['human_verdict']!r}")
        if str(len(quote)) != row["quote_char_len"].strip():
            raise LandingError(f"{rid} quote_char_len 与实际长度不符：{row['quote_char_len']} vs {len(quote)}")
        if _sha256_text(quote) != row["quote_sha256"].strip():
            raise LandingError(f"{rid} quote_sha256 不自洽，引文与候选包记录不一致")
        key = row["source_file"]
        if key not in cache:
            cache[key] = normalized_bundle(_resolve_text_file(key))
        _, flat, _ = cache[key]
        try:
            start, end = int(row["normalized_start"]), int(row["normalized_end"])
        except ValueError as exc:
            raise LandingError(f"{rid} normalized 偏移非法：{exc}") from exc
        if flat[start:end] != quote:
            raise LandingError(
                f"{rid} 按偏移 [{start},{end}) 在空白归一原文中未能逐字取回引文：{row['recorded_locator']}"
            )
        checks[rid] = {
            "recheck": "verbatim_offset_hit_whitespace_normalized",
            "quote": quote,
            "basis": "",
        }
    return checks


def gate_attestation(
    cands: dict[str, dict[str, str]],
    persons: dict[str, dict[str, str]],
    checks: dict[str, dict[str, str]],
) -> None:
    """三条落地引文过双方佐证门（REL-01891 同样复核留档，但不落地）。"""
    for rid in sorted(ADJUDICATION):
        row = cands[rid]
        pa = persons.get(row["person_a_id"].strip())
        pb = persons.get(row["person_b_id"].strip())
        if pa is None or pb is None:
            raise LandingError(f"{rid} 人物 ID 不在 persons.csv 中")
        quote = row["recaptured_quote"]
        diary = row["recorded_locator"].strip().startswith("鲁迅日记")
        ok, basis = quote_attests_pair(
            quote, name_candidates(pa), name_candidates(pb), diary_author_implicit=diary
        )
        checks[rid]["basis"] = basis
        if rid == "REL-01891":
            # 佐证门只回答「引文是否记载双方」；证据等级由人工裁决，不在此翻转。
            continue
        if not ok:
            raise LandingError(f"{rid} 落地引文未过双方佐证门：{basis}")


def resolve_sources(
    src_rows: list[dict[str, str]], rids: list[str], cands: dict[str, dict[str, str]]
) -> dict[str, dict[str, str]]:
    """按 citation 唯一命中既有 source_id；不得注册新来源。"""
    by_citation: dict[str, list[dict[str, str]]] = defaultdict(list)
    for row in src_rows:
        by_citation[(row.get("citation") or "").strip()].append(row)
    out: dict[str, dict[str, str]] = {}
    for rid in rids:
        locator = cands[rid]["recorded_locator"].strip()
        hits = by_citation.get(locator, [])
        if len(hits) != 1:
            raise LandingError(
                f"{rid} locator {locator!r} 在 sources.csv 中命中 {len(hits)} 条，"
                "不满足「唯一复用既有来源」，本批不允许注册新来源"
            )
        sid = hits[0]["source_id"].strip()
        if sid != EXPECTED_SOURCE_REUSE[rid]:
            raise LandingError(f"{rid} 复用来源 {sid} 与预期 {EXPECTED_SOURCE_REUSE[rid]} 不一致")
        strength = (hits[0].get("evidence_strength") or "").strip()
        if strength not in STRENGTH_TO_LEVEL:
            raise LandingError(f"{rid} 来源 {sid} evidence_strength 无法映射 source_level：{strength!r}")
        out[rid] = {"source_id": sid, "strength": strength, "level": STRENGTH_TO_LEVEL[strength]}
    return out


def build_evidence_row(
    cols: list[str],
    rele_id: str,
    rid: str,
    cand: dict[str, str],
    src: dict[str, str],
    recheck: dict[str, str],
) -> dict[str, str]:
    spec = ADJUDICATION[rid]
    tail = ""
    if rid == "REL-01161":
        tail = "本关系经人工裁决成立，但由发布门禁挡在 pending_review，不进入公开层"
    if rid == "REL-01368":
        tail = "本批同步类型更正 交游→签名联署（correction_reason 追加）"
    note = (
        f"{BATCH_MARKER}：Phase 5 重捕候选第二批落地（{LANDED_AT}）。"
        f"候选包行溯源=research/drafts/reports/phase5_quote_recapture_candidates.csv#relation_id={rid}"
        f"（旧证据行 {cand.get('old_recorded_quote', '')}）。"
        f"逐字复核={recheck['recheck']}：按 {Path(cand['source_file']).name} "
        f"偏移 [{cand['normalized_start']},{cand['normalized_end']})（空白归一）原样取回，"
        f"quote_sha256={cand['quote_sha256']} 自洽；双方佐证门通过（{recheck['basis']}）。"
        f"本行 source_level 按 sources.evidence_strength 既有映射取值（{src['strength']}→{src['level']}）。"
        f"裁决：{AUTHORIZED_BY} {AUTHORIZED_AT} 逐条独立裁决，逐字「{AUTHORIZATION_QUOTE}」。"
        + (tail + "。" if tail else "")
    )
    row = {col: "" for col in cols}
    row.update({
        "relation_evidence_id": rele_id,
        "relation_id": rid,
        "source_id": src["source_id"],
        "locator": cand["recorded_locator"].strip(),
        "quote": cand["recaptured_quote"],
        "context": (
            f"2026-10-08 逐条独立裁决（{spec['verdict']}）。夜间轮理由：{cand.get('overnight_reason', '')}"
        ),
        "quote_or_context": cand["recaptured_quote"],
        "evidence_support": "support",
        "source_level": src["level"],
        "review_status": "reviewed",
        "reviewer_note": note,
    })
    return row


def apply_landing(data_dir: Path, dry_run: bool = False) -> dict:
    data_dir = Path(data_dir)
    rel_cols, rel_rows = _read(data_dir / "person_relations.csv")
    ev_cols, ev_rows = _read(data_dir / "relation_evidences.csv")
    src_cols, src_rows = _read(data_dir / "sources.csv")
    pas_cols, pas_rows = _read(data_dir / "source_passages.csv")
    works_cols, works_rows = _read(data_dir / "source_works.csv")

    counts_before = {
        "person_relations.csv": len(rel_rows),
        "relation_evidences.csv": len(ev_rows),
        "sources.csv": len(src_rows),
        "source_passages.csv": len(pas_rows),
        "source_works.csv": len(works_rows),
    }

    marker_rows = [r for r in ev_rows if BATCH_MARKER in (r.get("reviewer_note") or "")]
    if marker_rows:
        # 幂等二跑：只校验本批标记行计数，不校验全局四表计数——对后续批次免疫。
        if len(marker_rows) != EXPECTED_LANDED_ROWS:
            raise LandingError(
                f"二跑：本批标记行应为 {EXPECTED_LANDED_ROWS} 条，实际 {len(marker_rows)} 条，疑似数据漂移"
            )
        got_rids = {r["relation_id"].strip() for r in marker_rows}
        if got_rids != set(LAND_IDS):
            raise LandingError(f"二跑：本批标记行关系集不符：{sorted(got_rids)}")
        return {
            "status": "no-op",
            "message": "无新增/已完成：本批落地痕迹已存在且标记行计数符合预期，跳过写入。",
        }

    # ---- 落地前基线 ----
    _check_counts(counts_before, EXPECTED_BASELINE_COUNTS, "落地前基线")
    status_dist = Counter((r.get("publish_status") or "").strip() for r in rel_rows)
    if dict(status_dist) != EXPECTED_BASELINE_STATUS:
        raise LandingError(f"落地前发布状态分布不符：期望 {EXPECTED_BASELINE_STATUS}，实际 {dict(status_dist)}")
    support_dist = Counter((r.get("evidence_support") or "").strip() for r in ev_rows)
    if dict(support_dist) != EXPECTED_BASELINE_SUPPORT:
        raise LandingError(f"落地前证据等级分布不符：{dict(support_dist)}")
    reviewed_before = sum(1 for r in ev_rows if (r.get("review_status") or "").strip() == "reviewed")
    if reviewed_before != EXPECTED_BASELINE_REVIEWED:
        raise LandingError(f"落地前 reviewed 证据应为 {EXPECTED_BASELINE_REVIEWED}，实际 {reviewed_before}")
    if not all(
        PREV_BATCH_MARKER in (r.get("reviewer_note") or "")
        for r in ev_rows if (r.get("review_status") or "").strip() == "reviewed"
    ):
        raise LandingError("落地前基线漂移：20 条 reviewed 行并非全部携带第一批批次标记")
    if {r.get("publish_status_origin", "").strip() for r in rel_rows} != {"derived"}:
        raise LandingError("落地前 origin 应全为 derived")
    critical_before = sum(1 for r in rel_rows if (r.get("relation_risk_level") or "").strip() == "critical")
    if critical_before != EXPECTED_BASELINE_CRITICAL:
        raise LandingError(f"落地前 critical 应为 {EXPECTED_BASELINE_CRITICAL}，实际 {critical_before}")

    # ---- 候选包复核 + 登记裁决表前置校验 + 佐证门 + 来源复用 ----
    cands = load_candidates(CANDIDATES)
    adj_file = validate_adjudicated(cands)
    persons = {r["person_id"].strip(): r for r in _read(PERSONS_CSV)[1]}
    checks = verify_candidates(cands)
    gate_attestation(cands, persons, checks)
    src_map = resolve_sources(src_rows, LAND_IDS + INSUFFICIENT_IDS, cands)

    # ---- 内存内落地 ----
    rel_by_id = {r["relation_id"].strip(): r for r in rel_rows}
    originals = {rid: dict(row) for rid, row in rel_by_id.items()}
    ev_snapshot = [dict(r) for r in ev_rows]

    for rid in LAND_IDS:
        row = rel_by_id.get(rid)
        if row is None:
            raise LandingError(f"{rid} 不在 person_relations.csv 中")
        if (row.get("review_status") or "").strip():
            raise LandingError(f"{rid} 关系行 review_status 非空，与基线不符")
        spec = ADJUDICATION[rid]
        if spec["type_after"]:
            suggested = cands[rid]["proposed_relation_type"].strip()
            if suggested != spec["type_after"]:
                raise LandingError(f"{rid} 候选包建议类型 {suggested!r} 与裁决目标 {spec['type_after']!r} 不一致")
            before = (row.get("final_relation_type") or "").strip()
            reason = (row.get("correction_reason") or "").strip()
            addition = (
                f"{BATCH_MARKER}：人工裁决 wrong_type 成立，类型由「{before}」改为「{spec['type_after']}」"
                f"（{AUTHORIZED_BY} {AUTHORIZED_AT} 逐条独立裁决，标准强于概括授权）。"
            )
            row["standard_relation_type"] = spec["type_after"]
            row["final_relation_type"] = spec["type_after"]
            row["correction_reason"] = f"{reason}｜{addition}" if reason else addition

    by_rel: dict[str, list[dict[str, str]]] = defaultdict(list)
    for r in ev_rows:
        by_rel[r["relation_id"].strip()].append(r)

    next_n = _next_rele_id([r["relation_evidence_id"] for r in ev_rows])
    new_evidence: list[dict[str, str]] = []
    ledger: list[dict[str, str]] = []
    for rid in LAND_IDS:
        cand = cands[rid]
        rele_id = f"RELE-{next_n:05d}"
        next_n += 1
        new_evidence.append(build_evidence_row(ev_cols, rele_id, rid, cand, src_map[rid], checks[rid]))
        by_rel[rid].append(new_evidence[-1])

    # 登记裁决表证据号必须与本次确定性计算一致（防「换一条证据」的静默篡改）。
    for e in new_evidence:
        want_eid = (adj_file[e["relation_id"]]["new_relation_evidence_id"] or "").strip()
        if want_eid != e["relation_evidence_id"]:
            raise LandingError(
                f"{e['relation_id']} 登记裁决表证据号 {want_eid!r} 与本次计算 "
                f"{e['relation_evidence_id']} 不一致，拒绝落地"
            )

    for rid in LAND_IDS:
        row = rel_by_id[rid]
        cols = assign_relation_status_columns(row, by_rel[rid], existing={
            "reviewer": row.get("reviewer", ""), "reviewed_at": row.get("reviewed_at", ""),
            "review_note": row.get("review_note", ""),
        })
        row["publish_status"] = cols["publish_status"]
        row["publish_status_origin"] = cols["publish_status_origin"]
        row["reviewer"] = cols["reviewer"]
        row["reviewed_at"] = cols["reviewed_at"]
        row["review_note"] = cols["review_note"]

    ev_rows_all = ev_rows + new_evidence

    # ---- 硬后置条件（任务书第 3 节）----
    dist_after = Counter((r.get("publish_status") or "").strip() for r in rel_rows)
    if dict(dist_after) != EXPECTED_POST_STATUS:
        raise LandingError(f"落地后发布状态分布不符：期望 {EXPECTED_POST_STATUS}，实际 {dict(dist_after)}")
    support_after = Counter((r.get("evidence_support") or "").strip() for r in ev_rows_all)
    if dict(support_after) != EXPECTED_POST_SUPPORT:
        raise LandingError(f"落地后证据等级分布不符：{dict(support_after)}")
    reviewed_after = sum(1 for r in ev_rows_all if (r.get("review_status") or "").strip() == "reviewed")
    if reviewed_after != EXPECTED_POST_REVIEWED:
        raise LandingError(f"落地后 reviewed 应为 {EXPECTED_POST_REVIEWED}，实际 {reviewed_after}")
    if len(ev_rows_all) != EXPECTED_POST_COUNTS["relation_evidences.csv"]:
        raise LandingError(f"落地后证据行数不符：{len(ev_rows_all)}")
    if {r.get("publish_status_origin", "").strip() for r in rel_rows} != {"derived"}:
        raise LandingError("本批只允许 derived origin，出现 human_adjudication")
    critical_after = sum(1 for r in rel_rows if (r.get("relation_risk_level") or "").strip() == "critical")
    if critical_after != critical_before:
        raise LandingError(f"relation_risk_level 被改写：critical {critical_before} -> {critical_after}（禁止反向降险）")

    public_ids = {
        r["relation_id"].strip() for r in rel_rows
        if (r.get("publish_status") or "").strip() in PUBLIC_RELATION_STATUSES
    }
    if public_ids != EXPECTED_PUBLIC_IDS:
        raise LandingError(f"公开层应为 {sorted(EXPECTED_PUBLIC_IDS)}，实际 {sorted(public_ids)}")
    for rid in LAND_IDS:
        expect = ADJUDICATION[rid]["expect_publish"]
        got = rel_by_id[rid]["publish_status"]
        if got != expect:
            raise LandingError(f"{rid} 落地后状态应为 {expect}，实际 {got}")

    # REL-01891：生产层零痕迹 + 状态保持 inferred
    rid = "REL-01891"
    if any(BATCH_MARKER in (e.get("reviewer_note") or "") for e in by_rel.get(rid, [])):
        raise LandingError("REL-01891 判证据不足却出现本批证据行")
    if rel_by_id[rid]["publish_status"] != "inferred":
        raise LandingError("REL-01891 应保持 inferred")
    if cands[rid]["secondary_description"].strip() != "yes":
        raise LandingError("REL-01891 候选行 secondary_description 非 yes，裁决理由与留档不符")

    # 类型更正恰 1 条，且 correction_reason 追加不覆盖
    corrections = []
    for rid, row in rel_by_id.items():
        if (originals[rid].get("final_relation_type") or "") != (row.get("final_relation_type") or ""):
            corrections.append(rid)
    if corrections != ["REL-01368"]:
        raise LandingError(f"类型更正应恰为 REL-01368 一条，实际 {corrections}")
    orig_reason = (originals["REL-01368"].get("correction_reason") or "").strip()
    if not rel_by_id["REL-01368"]["correction_reason"].startswith(orig_reason):
        raise LandingError("REL-01368 correction_reason 被覆盖而非追加")
    if EXPECTED_TYPE_CORRECTIONS != 1:  # 常量自检，防误改
        raise LandingError("EXPECTED_TYPE_CORRECTIONS 常量异常")

    # 除预期字段外零漂移：关系表逐行 diff，只允许 3 条落地行按预期变化
    allowed_changes = {
        "REL-00622": {"publish_status"},
        "REL-01161": set(),
        "REL-01368": {"publish_status", "standard_relation_type", "final_relation_type", "correction_reason"},
    }
    changed_rids: set[str] = set()
    for rid, row in rel_by_id.items():
        diffs = {c for c in rel_cols if (row.get(c) or "") != (originals[rid].get(c) or "")}
        if diffs:
            if rid not in allowed_changes:
                raise LandingError(f"{rid} 出现计划外字段改动：{sorted(diffs)}")
            if not diffs <= allowed_changes[rid]:
                raise LandingError(f"{rid} 改动字段超出预期：{sorted(diffs - allowed_changes[rid])}")
            changed_rids.add(rid)
    if changed_rids - {"REL-00622", "REL-01368"} != set():
        raise LandingError(f"出现计划外的关系行改动：{sorted(changed_rids)}")
    # 既有证据行原样保留
    if ev_rows_all[:len(ev_snapshot)] != ev_snapshot:
        raise LandingError("既有 relation_evidences 行被改写")
    if len({r["relation_evidence_id"] for r in ev_rows_all}) != len(ev_rows_all):
        raise LandingError("relation_evidence_id 出现重复")

    # 全表按门禁重算零漂移 + 公开层语义守门
    for row in rel_rows:
        st, origin = derive_relation_publish_status_with_origin(
            row, by_rel.get(row["relation_id"].strip(), [])
        )
        if (st, origin) != (row["publish_status"], row["publish_status_origin"]):
            raise LandingError(f"{row['relation_id']} 落地后门禁重算不一致：{origin}/{st}")
    for row in rel_rows:
        rid = row["relation_id"].strip()
        if (row.get("publish_status") or "").strip() not in PUBLIC_RELATION_STATUSES:
            continue
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
            and (BATCH_MARKER in (e.get("reviewer_note") or "")
                 or PREV_BATCH_MARKER in (e.get("reviewer_note") or ""))
        ]
        if not qualified:
            raise LandingError(f"{rid} 公开层关系缺少两批之一的合格 support 证据")
        pa = persons.get(row["source_person_id"].strip())
        pb = persons.get(row["target_person_id"].strip())
        ok_any = False
        for e in qualified:
            diary = (e.get("locator") or "").strip().startswith("鲁迅日记")
            ok, _ = quote_attests_pair(e["quote"], name_candidates(pa), name_candidates(pb),
                                       diary_author_implicit=diary)
            if ok:
                ok_any = True
                break
        if not ok_any:
            raise LandingError(f"{rid} 公开层合格 support 引文均未过双方佐证门")

    # ---- 台账与裁决记录（内存）----
    for rid in LAND_IDS + INSUFFICIENT_IDS:
        cand = cands[rid]
        spec = ADJUDICATION[rid]
        landed = spec["land"]
        ev_row = next((e for e in new_evidence if e["relation_id"] == rid), None)
        ledger.append({
            "relation_id": rid,
            "person_a_id": cand["person_a_id"], "person_a_name": cand["person_a_name"],
            "person_b_id": cand["person_b_id"], "person_b_name": cand["person_b_name"],
            "batch_marker": BATCH_MARKER,
            "adjudication": spec["verdict"],
            "landed_this_batch": "yes" if landed else "no",
            "evidence_landed": "yes" if ev_row else "no",
            "new_relation_evidence_id": ev_row["relation_evidence_id"] if ev_row else "",
            "candidate_human_verdict": cand["human_verdict"],
            "final_relation_type_before": originals[rid]["final_relation_type"],
            "final_relation_type_after": rel_by_id[rid]["final_relation_type"],
            "publish_status_before": originals[rid]["publish_status"],
            "publish_status_after": rel_by_id[rid]["publish_status"],
            "publish_status_origin": rel_by_id[rid]["publish_status_origin"],
            "reused_source_id": src_map[rid]["source_id"],
            "source_registered": "no",
            "locator": cand["recorded_locator"],
            "normalized_start": cand["normalized_start"],
            "normalized_end": cand["normalized_end"],
            "quote_sha256": cand["quote_sha256"],
            "quote_verbatim_recheck": checks[rid]["recheck"],
            "attestation_basis": checks[rid]["basis"],
            "authorized_by": AUTHORIZED_BY,
            "authorized_at": AUTHORIZED_AT,
            "authorization_quote": AUTHORIZATION_QUOTE,
            "candidate_row_source": f"phase5_quote_recapture_candidates.csv#relation_id={rid}",
            "note": spec["note"],
        })

    adjudicated_rows: list[dict[str, str]] = []
    ledger_by_rid = {e["relation_id"]: e for e in ledger}
    for cand in (cands[cand_id] for cand_id in sorted(cands)):
        rid = cand["relation_id"].strip()
        entry = ledger_by_rid.get(rid)
        item = {col: "" for col in ADJUDICATED_COLUMNS}
        item.update({
            "relation_id": rid,
            "person_a_id": cand["person_a_id"], "person_a_name": cand["person_a_name"],
            "person_b_id": cand["person_b_id"], "person_b_name": cand["person_b_name"],
            "current_final_relation_type": cand["current_final_relation_type"],
            "proposed_relation_type": cand["proposed_relation_type"],
            "recapture_status": cand["recapture_status"],
            "quote_sha256": cand["quote_sha256"],
            "recorded_locator": cand["recorded_locator"],
            "candidate_human_verdict": cand["human_verdict"],
        })
        if entry:
            item.update({
                "adjudication_2026_10_08": entry["adjudication"],
                "evidence_landed": entry["evidence_landed"],
                "new_relation_evidence_id": entry["new_relation_evidence_id"],
                "reused_source_id": entry["reused_source_id"],
                "final_relation_type_after": entry["final_relation_type_after"],
                "publish_status_after": entry["publish_status_after"],
                "quote_verbatim_recheck": entry["quote_verbatim_recheck"],
                "batch_marker": BATCH_MARKER,
                "authorized_by": AUTHORIZED_BY,
                "authorized_at": AUTHORIZED_AT,
                "authorization_quote": AUTHORIZATION_QUOTE,
                "note": entry["note"],
            })
        else:
            item.update({
                "adjudication_2026_10_08": "unadjudicated",
                "evidence_landed": "no",
                "final_relation_type_after": "",
                "publish_status_after": "",
                "note": "不在 2026-10-08 逐条裁决范围，维持 pending_human_review，生产层零改动。",
            })
        adjudicated_rows.append(item)

    summary = {
        "public_supported": len(public_ids),
        "new_evidence_rows": len(new_evidence),
        "type_corrections": len(corrections),
        "status_dist": dict(dist_after),
        "support_dist": dict(support_after),
        "reviewed_after": reviewed_after,
        "critical_after": critical_after,
        "evid_total": len(ev_rows_all),
        "public_ids": sorted(public_ids),
        "ledger": ledger,
        "adjudicated_rows": adjudicated_rows,
        "checks": checks,
    }
    if dry_run:
        return {"status": "dry-run", "message": "校验全部通过（未写入）。", **summary}

    _write(data_dir / "person_relations.csv", rel_cols, rel_rows)
    _write(data_dir / "relation_evidences.csv", ev_cols, ev_rows_all)
    return {"status": "applied", "message": "落地完成。", **summary}


def _import_kb_schema():
    """按文件路径加载根目录 kb_schema，避免把 ROOT 加进 sys.path（根 app.py 会遮蔽 app 包）。"""
    try:
        import kb_schema  # noqa: PLC0415

        return kb_schema
    except ModuleNotFoundError:
        import importlib.util

        spec = importlib.util.spec_from_file_location("kb_schema", ROOT / "kb_schema.py")
        module = importlib.util.module_from_spec(spec)
        # 必须先注册再执行：kb_schema 用 @dataclass(slots=True)，其注解解析要查 sys.modules。
        sys.modules["kb_schema"] = module
        spec.loader.exec_module(module)
        return module


def verify_post_state(data_dir: Path) -> dict[str, object]:
    validate_data_dir = _import_kb_schema().validate_data_dir

    counts = {}
    for name in EXPECTED_POST_COUNTS:
        counts[name] = len(_read(data_dir / name)[1])
    _check_counts(counts, EXPECTED_POST_COUNTS, "写后复核")
    result = validate_data_dir(data_dir)
    if result.errors:
        raise LandingError(f"写后 Schema 校验出现 {len(result.errors)} 个错误：{result.errors[:3]}")
    return {"counts": counts, "schema_errors": len(result.errors), "schema_warnings": len(result.warnings)}


def write_adjudicated(rows: list[dict[str, str]], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    _write(path, ADJUDICATED_COLUMNS, rows)


def write_ledger(rows: list[dict[str, str]], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    _write(path, LEDGER_COLUMNS, rows)


def write_adjudication_record(path: Path, result: dict) -> None:
    lines = [
        "# Phase 5 重捕候选第二批人工裁决记录（2026-10-08，逐条独立裁决）",
        "",
        "> 裁决载体：本文件 + `phase5_quote_recapture_adjudicated.csv`  ",
        "> 候选包：`phase5_quote_recapture_candidates.csv`（28 行，原样保留，不表示裁决状态）  ",
        "> 落地脚本：`research/analysis/apply_phase5_recapture_landing.py`（批次标记 `" + BATCH_MARKER + "`）",
        "",
        "## 0. 授权语（逐字，不得改写）",
        "",
        f"用户 2026-10-08 逐条裁决，原话：「{AUTHORIZATION_QUOTE}」。",
        "",
        "本批为**逐条独立裁决**，口径强于 2026-09-20 第一批的「授权按建议执行」"
        "（概括授权，签核路径 C，逐字「我都没问题，你自己看着办吧」）。"
        "两种口径在台账、站点文案与后续引用中必须分别说明，不得混同为一类「人工裁决」。",
        "",
        "## 1. 四条裁决与处置",
        "",
        "| relation_id | 人物对 | 现类型 | 裁决（语义） | 处置 | 落地后状态 |",
        "| --- | --- | --- | --- | --- | --- |",
    ]
    for entry in result["ledger"]:
        action = (
            f"新增 support 证据行 {entry['new_relation_evidence_id']}（复用 {entry['reused_source_id']}）"
            if entry["evidence_landed"] == "yes"
            else "生产层零改动"
        )
        if entry["final_relation_type_before"] != entry["final_relation_type_after"]:
            action += f"；类型 {entry['final_relation_type_before']}→{entry['final_relation_type_after']}"
        lines.append(
            f"| {entry['relation_id']} | {entry['person_a_name']}—{entry['person_b_name']} "
            f"| {entry['final_relation_type_before']} | {entry['adjudication']} | {action} "
            f"| {entry['publish_status_after']}（{entry['publish_status_origin']}） |"
        )
    lines += [
        "",
        "## 2. REL-01891「证据不足」而非 rejected 的理由",
        "",
        "重切段落为书目著录（"
        "「…传记小说。李克因作，载《东方纪事》1987年3、4期合刊。叙述叶紫…以及叶紫同陈企霞、聂绀弩、"
        "周颖夫妇、萧军、萧红夫妇等的交往」），属**二手著录**（候选包 secondary_description=yes），",
        "不足以定为 support 级直接记载。五态中 `rejected` 的语义是**人工否定/证伪**；"
        "把「未证实」写成 `rejected` 会夸大成「已证伪」。故本条**不新增任何证据行**，"
        "生产表零改动，关系保持 `inferred`，留在研究层待后续补证。",
        "",
        "## 3. REL-01161「成立但不进公开层」的门禁依据",
        "",
        "新增 support 证据行是对「裁决成立」的记录；但该关系 `relation_risk_level=critical`、"
        "`confidence=low`、类型 `同属组织` 属推断类型，按 `relation_publish_status` 判定顺序"
        "被三重拦截，自然停留在 `pending_review`。**禁止**（本批也未）使用 `human_adjudication` "
        "通道放行——公开层只接受 derived `supported`。风险列未做任何改写（critical 计数仍 1974）。",
        "",
        "## 4. 引文溯源",
        "",
        "四条重切引文、locator、quote_sha256、归一化偏移全部取自候选包，落地脚本运行时按 "
        "`source_file` + `normalized_start/end` 在空白归一后的本地原文中逐字回定位复核，未重新检索、未改写引文。",
        "逐条复核结论见 `phase5_recapture_landing_ledger.csv` 的 `quote_verbatim_recheck` 列"
        "（全部 `verbatim_offset_hit_whitespace_normalized`）。",
        "",
        f"授权：{AUTHORIZED_BY}，{AUTHORIZED_AT}。",
        "",
    ]
    _write_md(path, "\n".join(lines))


def write_report(result: dict, post: dict, path: Path) -> None:
    dist = result["status_dist"]
    entries = result["ledger"]
    lines = [
        "# Phase 5 重捕候选第二批落地报告（2026-10-08，逐条独立裁决）",
        "",
        "> 执行脚本：`research/analysis/apply_phase5_recapture_landing.py`（幂等）  ",
        f"> 批次标记：`{BATCH_MARKER}`  落地日期：{LANDED_AT}",
        "",
        "## 0. 授权与口径声明",
        "",
        f"- 授权人：{AUTHORIZED_BY}；授权时间：{AUTHORIZED_AT}。",
        f"- 授权语（逐字）：「{AUTHORIZATION_QUOTE}」。",
        "- 本批为**逐条独立人工裁决**，口径强于第一批 2026-09-20 的「授权按建议执行」"
        "（概括授权，签核路径 C）。两种口径必须分别引用：公开层现共 7 条，"
        "其中 5 条来自 2026-09-20 概括授权批次，2 条来自本批逐条裁决。",
        "- 公开层只接受 **derived `supported`**：未改门禁、未用 `human_adjudication`、"
        "未改 `relation_risk_level` / `needs_manual_review` / `confidence`。",
        "",
        "## 1. 落地明细",
        "",
        "| relation_id | 人物对 | 裁决 | 类型变化 | 落地后状态 | 新证据行 | 复用来源 |",
        "| --- | --- | --- | --- | --- | --- | --- |",
    ]
    for e in entries:
        tc = (
            f"{e['final_relation_type_before']}→{e['final_relation_type_after']}"
            if e["final_relation_type_before"] != e["final_relation_type_after"] else "不改"
        )
        lines.append(
            f"| {e['relation_id']} | {e['person_a_name']}—{e['person_b_name']} | {e['adjudication']} "
            f"| {tc} | {e['publish_status_after']} | {e['new_relation_evidence_id'] or '—（不落地）'} "
            f"| {e['reused_source_id']} |"
        )
    lines += [
        "",
        "- REL-01891 判「证据不足」：未新增任何证据行、未写 `rejected`（rejected 语义是人工否定），"
        "生产表零改动，保持 `inferred`。",
        "- REL-01161 落地 support 证据但被门禁（critical + low + 同属组织）自然挡在 `pending_review`，"
        "不进入公开层。",
        "- 引文逐字复核：4/4 按候选包偏移在空白归一原文中原样取回，`quote_sha256` 自洽；"
        "3 条落地引文另过双方佐证门。",
        "",
        "## 2. 实测终值 vs 预期终值",
        "",
        "| 项目 | 预期 | 实测 |",
        "| --- | ---: | ---: |",
        f"| person_relations 行数 | {EXPECTED_POST_COUNTS['person_relations.csv']} | "
        f"{post['counts']['person_relations.csv']} |",
        f"| relation_evidences 行数 | {EXPECTED_POST_COUNTS['relation_evidences.csv']} | "
        f"{post['counts']['relation_evidences.csv']} |",
        f"| sources / passages / works | 1178 / 1178 / 65 | "
        f"{post['counts']['sources.csv']} / {post['counts']['source_passages.csv']} / "
        f"{post['counts']['source_works.csv']}（未注册新来源、未重写这三张表） |",
        f"| publish_status 分布 supported/pending_review/inferred | 7 / 2451 / 1780 | "
        f"{dist.get('supported', 0)} / {dist.get('pending_review', 0)} / {dist.get('inferred', 0)} |",
        f"| support 证据行 | 23 | {result['support_dist'].get('support', 0)} |",
        f"| associated 证据行 | 10249 | {result['support_dist'].get('associated', 0)}（未改判既有行） |",
        f"| reviewed 证据行 | 23 | {result['reviewed_after']} |",
        f"| critical 计数 | 1974 | {result['critical_after']}（未反向降险） |",
        f"| 类型更正 | 恰 1 条（REL-01368） | {result['type_corrections']} 条 |",
        f"| 公开层关系 | 7 条 | {result['public_supported']} 条（{', '.join(result['public_ids'])}） |",
        "",
        "Schema：写后 "
        f"{post['schema_errors']} errors / {post['schema_warnings']} warnings。",
        "",
        "## 3. 明确未做的事",
        "",
        "- 未改 `relation_publish_status.py` 判定顺序、`INFERRED_RELATION_TYPES`、`PUBLIC_RELATION_STATUSES`。",
        "- 未把 10249 条 associated 或第一批 20 条 support 改判/改写。",
        "- 未注册新来源（三个落地 locator 全部复用既有 SRC-0779 / SRC-0054 / SRC-0739）。",
        "- 未把 REL-01891 写为 rejected；未对候选包与队列文件做任何改动。",
        "- 未使用 human_adjudication 通道；未动风险/置信/复核标记列。",
        "",
        "## 4. 复核方式",
        "",
        "```powershell",
        "python research/analysis/apply_phase5_recapture_landing.py --dry-run   # 只校验不写",
        "python research/analysis/apply_phase5_recapture_landing.py             # 落地（二跑输出「无新增/已完成」）",
        "python research/analysis/build_publish_data.py",
        "python build_static_site.py",
        "python research/analysis/build_trustworthy_network_analysis.py",
        "python -m pytest -q",
        "```",
        "",
        f"逐条痕迹见 `phase5_recapture_landing_ledger.csv`（{len(entries)} 行）；"
        "28 行候选的处置与授权溯源见 `phase5_quote_recapture_adjudicated.csv`；"
        "裁决原文与语义见 `phase5_quote_recapture_adjudication_record.md`。",
        "",
    ]
    path.parent.mkdir(parents=True, exist_ok=True)
    _write_md(path, "\n".join(lines))


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Phase 5 重捕候选第二批生产层落地（2026-10-08 逐条独立裁决，幂等）。"
    )
    parser.add_argument("--data-dir", type=Path, default=DATA, help="生产数据目录")
    parser.add_argument("--dry-run", action="store_true", help="只执行全部校验，不写任何文件")
    parser.add_argument("--skip-report", action="store_true", help="不写裁决/台账/报告（沙盒测试用）")
    args = parser.parse_args()

    try:
        result = apply_landing(args.data_dir, dry_run=args.dry_run)
    except LandingError as exc:
        print(f"落地失败，未写入任何文件：{exc}")
        return 1

    print(result["message"])
    if result["status"] == "no-op":
        return 0
    print(f"本批裁决 4 条：落地证据行 {result['new_evidence_rows']} 条 / "
          f"类型更正 {result['type_corrections']} 条 / 证据不足 1 条（REL-01891，零改动）")
    print(f"公开层 derived supported：{result['public_supported']} 条")

    if result["status"] == "applied" and not args.skip_report:
        write_adjudicated(result["adjudicated_rows"], ADJUDICATED)
        write_ledger(result["ledger"], LEDGER)
        write_adjudication_record(ADJUDICATION_RECORD, result)
        print(f"裁决记录：{ADJUDICATED.relative_to(ROOT)} / {ADJUDICATION_RECORD.relative_to(ROOT)}")
        print(f"台账：{LEDGER.relative_to(ROOT)}")

    if result["status"] == "applied":
        post = verify_post_state(args.data_dir)
        print(f"写后复核：计数 {post['counts']}；Schema {post['schema_errors']} errors / {post['schema_warnings']} warnings")
        if not args.skip_report:
            write_report(result, post, REPORT_MD)
            print(f"报告：{REPORT_MD.relative_to(ROOT)}")
    elif result["status"] == "dry-run":
        print(f"dry-run：校验通过；投影公开层 {result['public_supported']} 条、"
              f"新增证据 {result['new_evidence_rows']} 条、状态分布 {result['status_dist']}。未写入任何文件。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
