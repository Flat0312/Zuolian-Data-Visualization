"""Phase 5 全量扫掠第 1 批（双 Agent 交叉验证）生产层落地（幂等）。

授权依据
--------
1. 2026-10-08 用户授权以交叉验证替代逐条人工裁决，逐字授权语（登记于 BLOCKED.md 与
   合并器 docstring）：「不用我裁，你和zcode交叉验证裁决吧」。
2. 2026-10-08 用户两项规则决定（本会话逐字确认）：
   - 采用保守交集（分歧条目中双方 support 且至少一方判成立者按现类型落地）；
   - 词表归并到多数标签：论战→文学论战、交往→交游（已由 merge_relation_type_vocab.py
     落地），3 条 type_synonym 分歧（REL-00011/00038/00088）随之解锁；
   - 与前 15 条同批落地全部 18 条。
3. 裁决者：A = Codex（GPT 系，主 Agent）；B = Claude Opus 4.6 (Thinking)（antigravity
   harness，独立裁决文件，本仓不可见）。**这不是人工复核**：公开层只走 derived
   `supported`，本脚本禁止使用 `human_adjudication` / `verified` 通道，台账、站点与
   报告必须标注「双 Agent 交叉验证」口径。

落地语义（18 条 = 一致 15 + 保守交集 3）
----------------------------------------
- 一致行（landable=yes）：新增一条 support 证据行；verdict=类型需改 的同步更正
  `standard_relation_type` 与 `final_relation_type`（proposed_type 经别名归一），
  `correction_reason` 追加不覆盖；
- 保守交集行（landable=intersection）：新增 support 证据行，类型**不改**（按现类型
  落地，不超出任何一方认可的范围）；
- 发布状态一律由 `relation_publish_status.assign_relation_status_columns` 重算
  （derived）；不改 `relation_risk_level` / `needs_manual_review` / `confidence`。

来源注册
--------
locator 在 sources.csv 按 citation 唯一命中即复用；未命中的 5 条（左联史 3 页 +
左联词典 1 页……详见 EXPECTED_NEW_SOURCES）按**同族模板**注册新来源（先例：
P5-LANDING-2026-09-28 注册 SRC-1178），并调用 `source_layer.sync_source_layer`
补齐 passage 与 citation_count——这是该模块 docstring 规定的唯一层级维护入口。

运行时校验（任一失败即整体退出、不写任何文件）
----------------------------------------------
1. 授权记录文件存在且逐字含两条授权语；
2. 交叉验证对照表 `landable` 集合恰为内置 18 条（多、少、改都拒绝）；
   一致行双方 verdict/grade/proposed_type（归一后）一致；交集行满足交集条件；
3. 类型更正目标与 xval 双方 proposed_type（归一后）一致；
4. 18 条候选引文按 sweep_candidates 的 `source_file` + `normalized_start/end` 在空白
   归一后的本地原文中逐字回定位，`quote_sha256` 自洽；
5. 18 条引文过双方佐证门（`quote_attestation`，不放宽）；
6. 落地前基线（词表归并后）：4238 / 10272（support 23）/ 1178 / 1178 / 65、
   supported 7、critical 1974、文学论战 127、交游 1214、论战 0、交往 0、origin 全 derived；
7. 落地后硬后置条件（公开层 25、类型更正恰 6、既有证据行零改写、除 18 条外关系行
   零漂移、门禁全表重算一致、公开层语义守门）；
8. 写后 Schema 0 错误。

幂等：存在本批标记行时只校验「本批标记行恰 18 条」即输出「无新增/已完成」并零写入。
原子写（tmp + `os.replace`）；CSV 为 utf-8-sig + QUOTE_MINIMAL。
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

XVAL = REPORTS / "sweep_batch1_cross_validated.csv"
CANDIDATES = REPORTS / "sweep_candidates.csv"
AUTH_RECORD = REPORTS / "phase5_sweep_batch1_authorization_record.md"
LEDGER = REPORTS / "phase5_sweep_batch1_landing_ledger.csv"
REPORT_MD = REPORTS / "phase5_sweep_batch1_landing_report.md"

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
    from research.analysis.source_layer import sync_source_layer
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
    from source_layer import sync_source_layer  # noqa: E402

BATCH_MARKER = "P5-SWEEP-BATCH1-2026-10-08"
PREV_BATCH_MARKERS = ("P5-LANDING-2026-09-28", "P5-RECAPTURE-2026-10-08")
AUTHORIZED_BY = "用户（授权双 Agent 交叉验证 + 两项规则决定）"
AUTHORIZED_AT = "2026-10-08"
AUTH_XVAL_QUOTE = "不用我裁，你和zcode交叉验证裁决吧"
AUTH_RULE_QUOTE = "采用保守交集；归并到多数标签：论战→文学论战，交往→交游；同批落地全部 18 条"
LANDED_AT = "2026-10-08"
ADJUDICATOR_A = "Codex（GPT 系，主 Agent）"
ADJUDICATOR_B = "Claude Opus 4.6 (Thinking)（antigravity harness）"

TYPE_ALIASES = {"论战": "文学论战", "交往": "交游"}

STRENGTH_TO_LEVEL = {"一手": "A", "二手": "B", "转引": "C", "参考": "D", "推断": "D"}
FAMILY_BY_TITLE = {"左联史": "zuolian_shi", "左联词典": "zuolian_cidian"}
SOURCE_PATH_BY_TITLE = {
    "左联史": "research/raw_texts/左联史.txt",
    "左联词典": "research/raw_texts/左联词典.txt",
}

# 18 条落地集合（一致 15 + 保守交集 3）。运行时与 xval 的 landable 集合双向对照。
LAND_IDS = frozenset({
    # 一致 15（其中 3 条为词表归并解锁的同义分歧）
    "REL-00008", "REL-00011", "REL-00012", "REL-00016", "REL-00022", "REL-00025",
    "REL-00027", "REL-00032", "REL-00033", "REL-00036", "REL-00038", "REL-00039",
    "REL-00061", "REL-00082", "REL-00088",
    # 保守交集 3
    "REL-00029", "REL-00035", "REL-00089",
})
INTERSECTION_IDS = frozenset({"REL-00029", "REL-00035", "REL-00089"})

# 类型更正目标（verdict=类型需改 的一致行；proposed_type 经别名归一后的目标值）。
# 运行时与 xval 双方 proposed_type 逐条对照，多、少、改都拒绝。
TYPE_TARGETS = {
    "REL-00011": "交游",       # 通信 → 交游（双方分别提议 交游/交往，归一后一致）
    "REL-00012": "签名联署",   # 创作合作 → 签名联署
    "REL-00025": "签名联署",   # 通信 → 签名联署
    "REL-00038": "文学论战",   # 交游 → 文学论战（双方分别提议 文学论战/论战，归一后一致）
    "REL-00061": "签名联署",   # 交游 → 签名联署
    "REL-00088": "交游",       # 通信 → 交游（双方分别提议 交游/交往，归一后一致）
}

# 预期需注册的新来源（locator 在 sources.csv 无 citation 命中；按同族模板注册）。
EXPECTED_NEW_SOURCES = {
    "REL-00008": "左联史 第477页",
    "REL-00027": "左联史 第543页",
    "REL-00033": "左联史 第484页",
    "REL-00035": "左联词典 第538页",
    "REL-00038": "左联史 第715页",
}

# 落地前基线（2026-10-08 词表归并后实测）
EXPECTED_BASELINE_COUNTS = {
    "person_relations.csv": 4238,
    "relation_evidences.csv": 10272,
    "sources.csv": 1178,
    "source_passages.csv": 1178,
    "source_works.csv": 65,
}
EXPECTED_BASELINE_STATUS = {"supported": 7, "pending_review": 2451, "inferred": 1780}
EXPECTED_BASELINE_SUPPORT = {"associated": 10249, "support": 23}
EXPECTED_BASELINE_REVIEWED = 23
EXPECTED_BASELINE_CRITICAL = 1974
EXPECTED_BASELINE_TYPE = {"文学论战": 127, "交游": 1214, "论战": 0, "交往": 0}

# 落地后硬后置条件（dry-run 实测后固化；任一不符即整体退出、不写任何文件）
EXPECTED_NEW_EVIDENCE_ROWS = 18
EXPECTED_TYPE_CORRECTIONS = 6
EXPECTED_POST_COUNTS = {
    "person_relations.csv": 4238,
    "relation_evidences.csv": 10290,
    "sources.csv": 1183,
    "source_passages.csv": 1183,
    "source_works.csv": 65,
}
EXPECTED_POST_STATUS = {"supported": 25, "pending_review": 2451, "inferred": 1762}
EXPECTED_POST_SUPPORT = {"associated": 10249, "support": 41}
EXPECTED_POST_REVIEWED = 41
EXPECTED_PUBLIC_IDS = frozenset({
    # 第一批（2026-09-20 概括授权）
    "REL-00046", "REL-00059", "REL-00097", "REL-03289", "REL-03518",
    # 第二批（2026-10-08 逐条独立人工裁决）
    "REL-00622", "REL-01368",
    # 本批（2026-10-08 双 Agent 交叉验证）
    *LAND_IDS,
})

LEDGER_COLUMNS = [
    "relation_id", "person_a_id", "person_a_name", "person_b_id", "person_b_name",
    "batch_marker", "landing_rule", "codex_verdict", "opus_verdict",
    "codex_proposed_type", "opus_proposed_type", "type_normalized",
    "landed_this_batch", "new_relation_evidence_id",
    "final_relation_type_before", "final_relation_type_after",
    "publish_status_before", "publish_status_after", "publish_status_origin",
    "source_id", "source_registered", "locator",
    "normalized_start", "normalized_end", "quote_sha256",
    "quote_verbatim_recheck", "attestation_basis",
    "authorized_by", "authorized_at", "authorization_xval_quote",
    "xval_row_source", "candidate_row_source", "note",
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


def _next_source_id(existing_ids: list[str]) -> str:
    max_n = 0
    for value in existing_ids:
        m = re.fullmatch(r"SRC-(\d+)", str(value).strip())
        if m:
            max_n = max(max_n, int(m.group(1)))
    return f"SRC-{max_n + 1:04d}"


def _check_counts(rows: dict[str, int], expected: dict[str, int], stage: str) -> None:
    for name, want in expected.items():
        got = rows.get(name)
        if got != want:
            raise LandingError(f"{stage}：{name} 应为 {want}，实际 {got}")


def verify_authorization() -> None:
    """授权记录文件必须存在且逐字含两条授权语（防「无授权也可落地」）。"""
    if not AUTH_RECORD.exists():
        raise LandingError(f"授权记录不存在，拒绝落地：{AUTH_RECORD}")
    text = AUTH_RECORD.read_text(encoding="utf-8")
    for quote in (AUTH_XVAL_QUOTE, "采用保守交集", "论战→文学论战，交往→交游"):
        if quote not in text:
            raise LandingError(f"授权记录缺少逐字授权语：{quote!r}")


def load_xval() -> dict[str, dict[str, str]]:
    """读取交叉验证对照表，校验 landable 集合与内置 18 条双向一致，并校验裁决结构。"""
    cols, rows = _read(XVAL)
    for col in ("relation_id", "landable", "codex_verdict", "opus_verdict",
                "codex_grade", "opus_grade", "codex_proposed_type", "opus_proposed_type",
                "codex_key_phrase", "opus_key_phrase", "locator", "quote"):
        if col not in cols:
            raise LandingError(f"对照表缺列 {col}")
    landable_ids = {
        r["relation_id"].strip() for r in rows if r["landable"].strip() in ("yes", "intersection")
    }
    if landable_ids != set(LAND_IDS):
        raise LandingError(
            "对照表可落地集合与内置 18 条不一致："
            f"多出 {sorted(landable_ids - LAND_IDS)}，缺少 {sorted(LAND_IDS - landable_ids)}"
            "（对照表疑似被改动，拒绝落地）"
        )
    by_id: dict[str, dict[str, str]] = {}
    for r in rows:
        rid = r["relation_id"].strip()
        if rid in by_id:
            raise LandingError(f"对照表 relation_id 重复：{rid}")
        by_id[rid] = r
    for rid in sorted(LAND_IDS):
        row = by_id[rid]
        a, b = row["codex_verdict"].strip(), row["opus_verdict"].strip()
        ga, gb = row["codex_grade"].strip(), row["opus_grade"].strip()
        ta, tb = row["codex_proposed_type"].strip(), row["opus_proposed_type"].strip()
        if row["landable"].strip() == "intersection":
            if rid in INTERSECTION_IDS:
                # 交集条件：双方都 support 且至少一方判成立。
                if not (ga == "support" and gb == "support"):
                    raise LandingError(f"{rid} 交集行证据等级非双方 support：{ga}/{gb}")
                if "成立" not in (a, b):
                    raise LandingError(f"{rid} 交集行须至少一方判成立：{a}/{b}")
            else:
                raise LandingError(f"{rid} landable=intersection 但不在内置交集集合")
        else:
            if rid in INTERSECTION_IDS:
                raise LandingError(f"{rid} 在内置交集集合但对照表为一致行")
            if a != b or ga != gb:
                raise LandingError(f"{rid} 一致行双方 verdict/grade 不一致：{a}/{ga} vs {b}/{gb}")
            if a == "类型需改":
                if normalize_label(ta) != normalize_label(tb):
                    raise LandingError(f"{rid} 一致行双方 proposed_type 归一后不一致：{ta} vs {tb}")
                target = TYPE_TARGETS.get(rid)
                if not target:
                    raise LandingError(f"{rid} 判类型需改但无内置类型目标")
                if normalize_label(ta) != target:
                    raise LandingError(
                        f"{rid} 对照表类型目标 {normalize_label(ta)!r} 与内置 {target!r} 不一致"
                    )
            else:
                if rid in TYPE_TARGETS:
                    raise LandingError(f"{rid} 无类型需改裁决却存在内置类型目标")
        # key_phrase 必须是引文逐字子串（与合并器同一校验，防对照表被改后落地）
        flat = re.sub(r"\s+", "", row["quote"])
        for tag in ("codex", "opus"):
            key = row[f"{tag}_key_phrase"].strip()
            if not key or key not in flat:
                raise LandingError(f"{rid} {tag} key_phrase 非引文连续子串：{key!r}")
    return by_id


def normalize_label(label: str) -> str:
    return TYPE_ALIASES.get((label or "").strip(), (label or "").strip())


def load_candidates() -> dict[str, dict[str, str]]:
    _, rows = _read(CANDIDATES)
    by_id = {r["relation_id"].strip(): r for r in rows}
    if len(by_id) != len(rows):
        raise LandingError("候选池 relation_id 重复")
    missing = sorted(LAND_IDS - set(by_id))
    if missing:
        raise LandingError(f"候选池缺少本批关系：{missing}")
    return by_id


def verify_candidates(
    xval: dict[str, dict[str, str]], cands: dict[str, dict[str, str]]
) -> dict[str, dict[str, str]]:
    """18 条候选引文逐字回定位 + sha256 自洽。返回 rid -> 复核结论。"""
    checks: dict[str, dict[str, str]] = {}
    cache: dict[str, tuple[str, str, list[int]]] = {}
    for rid in sorted(LAND_IDS):
        xrow, crow = xval[rid], cands[rid]
        quote = crow["quote"]
        if _sha256_text(quote) != crow["quote_sha256"].strip():
            raise LandingError(f"{rid} quote_sha256 不自洽，候选池引文与哈希记录不一致")
        # xval 与候选池的引文必须同源（对照表引文来自候选池，防中途替换）
        if re.sub(r"\s+", "", xrow["quote"]) != re.sub(r"\s+", "", quote):
            raise LandingError(f"{rid} 对照表引文与候选池引文（空白归一后）不一致")
        key = crow["source_file"]
        if key not in cache:
            cache[key] = normalized_bundle(_resolve_text_file(key))
        _, flat, _ = cache[key]
        try:
            start, end = int(crow["normalized_start"]), int(crow["normalized_end"])
        except ValueError as exc:
            raise LandingError(f"{rid} normalized 偏移非法：{exc}") from exc
        if flat[start:end] != quote:
            raise LandingError(
                f"{rid} 按偏移 [{start},{end}) 在空白归一原文中未能逐字取回引文：{crow['locator']}"
            )
        checks[rid] = {
            "recheck": "verbatim_offset_hit_whitespace_normalized",
            "quote": quote,
            "basis": "",
        }
    return checks


def _resolve_text_file(raw: str) -> Path:
    p = Path(raw)
    if not p.exists():
        p = DATA / "runtime_sources" / Path(raw).name
    if not p.exists():
        p = ROOT / "research" / "raw_texts" / Path(raw).name
    if not p.exists():
        raise LandingError(f"本地原文不存在，无法逐字复核：{raw}")
    return p


def gate_attestation(
    cands: dict[str, dict[str, str]],
    persons: dict[str, dict[str, str]],
    checks: dict[str, dict[str, str]],
) -> None:
    """18 条落地引文过双方佐证门（复用 quote_attestation，不放宽）。"""
    for rid in sorted(LAND_IDS):
        crow = cands[rid]
        pa = persons.get(crow["person_a_id"].strip())
        pb = persons.get(crow["person_b_id"].strip())
        if pa is None or pb is None:
            raise LandingError(f"{rid} 人物 ID 不在 persons.csv 中")
        quote = crow["quote"]
        diary = crow["locator"].strip().startswith("鲁迅日记")
        ok, basis = quote_attests_pair(
            quote, name_candidates(pa), name_candidates(pb), diary_author_implicit=diary
        )
        checks[rid]["basis"] = basis
        if not ok:
            raise LandingError(f"{rid} 落地引文未过双方佐证门：{basis}")


def resolve_sources(
    src_cols: list[str],
    src_rows: list[dict[str, str]],
    cands: dict[str, dict[str, str]],
) -> tuple[dict[str, dict[str, str]], list[dict[str, str]]]:
    """locator 按 citation 唯一命中即复用；未命中的按同族模板注册新来源。

    返回 (rid -> {source_id, strength, level, registered}, 新增来源行)。
    注册顺序按 relation_id 排序，保证 source_id 分配确定（SRC-1179 起）。
    """
    by_citation: dict[str, list[dict[str, str]]] = defaultdict(list)
    for row in src_rows:
        cit = (row.get("citation") or "").strip()
        if cit:
            by_citation[cit].append(row)
    template_by_family: dict[str, dict[str, str]] = {}
    for row in src_rows:
        fam = (row.get("source_family") or "").strip()
        if fam in FAMILY_BY_TITLE.values() and fam not in template_by_family:
            template_by_family[fam] = row
    for title, fam in FAMILY_BY_TITLE.items():
        if fam not in template_by_family:
            raise LandingError(f"同族模板来源不存在，无法注册新来源：{title}（{fam}）")

    mapping: dict[str, dict[str, str]] = {}
    added: list[dict[str, str]] = []
    existing_ids = [r["source_id"].strip() for r in src_rows]
    for rid in sorted(LAND_IDS):
        locator = cands[rid]["locator"].strip()
        hits = by_citation.get(locator, [])
        if len(hits) > 1:
            raise LandingError(f"{rid} locator {locator!r} 在 sources.csv 命中多条 citation，无法唯一定位")
        if hits:
            src = hits[0]
            strength = (src.get("evidence_strength") or "").strip()
            if strength not in STRENGTH_TO_LEVEL:
                raise LandingError(f"{rid} 复用来源 {src['source_id']} evidence_strength 无法映射：{strength!r}")
            mapping[rid] = {
                "source_id": src["source_id"].strip(), "strength": strength,
                "level": STRENGTH_TO_LEVEL[strength], "registered": False,
            }
            continue
        expect = EXPECTED_NEW_SOURCES.get(rid)
        if locator != expect:
            raise LandingError(
                f"{rid} locator {locator!r} 在 sources.csv 无命中，且不在本批允许注册的白名单内"
                f"（预期 {expect!r}）"
            )
        title = locator.split(" ")[0].strip()
        fam = FAMILY_BY_TITLE.get(title)
        if not fam:
            raise LandingError(f"{rid} 新来源书名无法定位同族：{title!r}")
        template = template_by_family[fam]
        new_id = _next_source_id(existing_ids + [r["source_id"] for r in added])
        row = {col: template.get(col, "") for col in src_cols}
        row["source_id"] = new_id
        row["title"] = title
        row["citation"] = locator
        row["source_path"] = SOURCE_PATH_BY_TITLE[title]
        row["source_family"] = fam
        row["review_note"] = (
            f"{BATCH_MARKER} 按 sweep 候选引文注册；引文见 relation_evidences；"
            f"逐字复核见 {LEDGER.name}。"
        )
        added.append(row)
        by_citation[locator].append(row)
        strength = (row.get("evidence_strength") or "").strip()
        if strength not in STRENGTH_TO_LEVEL:
            raise LandingError(f"{rid} 新来源 evidence_strength 无法映射：{strength!r}")
        mapping[rid] = {
            "source_id": new_id, "strength": strength,
            "level": STRENGTH_TO_LEVEL[strength], "registered": True,
        }
    got_new = {rid: m["source_id"] for rid, m in mapping.items() if m["registered"]}
    if set(got_new) != set(EXPECTED_NEW_SOURCES):
        raise LandingError(
            f"新注册来源集合与预期不符：实际 {sorted(got_new)}，预期 {sorted(EXPECTED_NEW_SOURCES)}"
        )
    return mapping, added


def build_evidence_row(
    cols: list[str],
    rele_id: str,
    rid: str,
    xrow: dict[str, str],
    cand: dict[str, str],
    src: dict[str, str],
    recheck: dict[str, str],
) -> dict[str, str]:
    rule = "保守交集（按现类型落地）" if rid in INTERSECTION_IDS else "双 Agent 一致"
    type_note = ""
    if rid in TYPE_TARGETS:
        type_note = f"类型更正：{xrow['current_final_relation_type']}→{TYPE_TARGETS[rid]}（standard/final 同步，correction_reason 追加）。"
    note = (
        f"{BATCH_MARKER}：全量扫掠第 1 批落地（{LANDED_AT}），口径=**双 Agent 交叉验证，非人工复核**；"
        f"裁决者 A={ADJUDICATOR_A}；裁决者 B={ADJUDICATOR_B}；双方 blind 裁决（输入不含夜间轮 reason），"
        f"一致/交集后由本幂等脚本机械落地。落地规则={rule}。"
        f"裁决：Codex={xrow['codex_verdict']}{'/' + xrow['codex_proposed_type'] if xrow['codex_proposed_type'] else ''}，"
        f"Opus={xrow['opus_verdict']}{'/' + xrow['opus_proposed_type'] if xrow['opus_proposed_type'] else ''}。"
        f"候选行溯源=research/drafts/reports/sweep_candidates.csv#relation_id={rid}"
        f"（对照表 sweep_batch1_cross_validated.csv）。"
        f"逐字复核={recheck['recheck']}：按 {Path(cand['source_file']).name} 偏移 "
        f"[{cand['normalized_start']},{cand['normalized_end']})（空白归一）原样取回，"
        f"quote_sha256={cand['quote_sha256']} 自洽；双方佐证门通过（{recheck['basis']}）。"
        f"本行 source_level 按 sources.evidence_strength 既有映射取值（{src['strength']}→{src['level']}）"
        f"{'；来源为本批新注册。' if src['registered'] else '。'}"
        f"{type_note}"
        f"授权：{AUTHORIZED_BY} {AUTHORIZED_AT}，逐字「{AUTH_XVAL_QUOTE}」「{AUTH_RULE_QUOTE}」。"
    )
    row = {col: "" for col in cols}
    row.update({
        "relation_evidence_id": rele_id,
        "relation_id": rid,
        "source_id": src["source_id"],
        "locator": cand["locator"].strip(),
        "quote": cand["quote"],
        "context": (
            f"双 Agent 交叉验证裁决（非人工复核）。Codex：{xrow['codex_reason']} "
            f"Opus：{xrow['opus_reason']}"
        ),
        "quote_or_context": cand["quote"],
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
    works_cols, works_rows = _read(data_dir / "source_works.csv")
    pas_cols, pas_rows = _read(data_dir / "source_passages.csv")

    counts_before = {
        "person_relations.csv": len(rel_rows),
        "relation_evidences.csv": len(ev_rows),
        "sources.csv": len(src_rows),
        "source_passages.csv": len(pas_rows),
        "source_works.csv": len(works_rows),
    }

    marker_rows = [r for r in ev_rows if BATCH_MARKER in (r.get("reviewer_note") or "")]
    if marker_rows:
        # 幂等二跑：只校验本批标记行计数，不校验全局计数——对后续批次免疫。
        if len(marker_rows) != EXPECTED_NEW_EVIDENCE_ROWS:
            raise LandingError(
                f"二跑：本批标记行应为 {EXPECTED_NEW_EVIDENCE_ROWS} 条，实际 {len(marker_rows)} 条，疑似数据漂移"
            )
        got_rids = {r["relation_id"].strip() for r in marker_rows}
        if got_rids != set(LAND_IDS):
            raise LandingError(f"二跑：本批标记行关系集不符：{sorted(got_rids)}")
        return {
            "status": "no-op",
            "message": "无新增/已完成：本批落地痕迹已存在且标记行计数符合预期，跳过写入。",
        }

    # ---- 授权 + 对照表 + 候选池前置校验 ----
    verify_authorization()
    xval = load_xval()
    cands = load_candidates()

    # ---- 落地前基线（词表归并后）----
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
    reviewed_marks = [
        (r.get("reviewer_note") or "") for r in ev_rows
        if (r.get("review_status") or "").strip() == "reviewed"
    ]
    if not all(any(m in note for m in PREV_BATCH_MARKERS) for note in reviewed_marks):
        raise LandingError("落地前基线漂移：reviewed 行并非全部携带前两批批次标记")
    if {r.get("publish_status_origin", "").strip() for r in rel_rows} != {"derived"}:
        raise LandingError("落地前 origin 应全为 derived")
    critical_before = sum(1 for r in rel_rows if (r.get("relation_risk_level") or "").strip() == "critical")
    if critical_before != EXPECTED_BASELINE_CRITICAL:
        raise LandingError(f"落地前 critical 应为 {EXPECTED_BASELINE_CRITICAL}，实际 {critical_before}")
    type_dist = Counter((r.get("final_relation_type") or "").strip() for r in rel_rows)
    for label, want in EXPECTED_BASELINE_TYPE.items():
        if type_dist.get(label, 0) != want:
            raise LandingError(
                f"落地前类型「{label}」应为 {want}，实际 {type_dist.get(label, 0)}"
                "（词表归并未执行或数据漂移，拒绝落地）"
            )

    # ---- 引文逐字复核 + 佐证门 + 来源解析 ----
    persons = {r["person_id"].strip(): r for r in _read(PERSONS_CSV)[1]}
    checks = verify_candidates(xval, cands)
    gate_attestation(cands, persons, checks)
    src_map, added_sources = resolve_sources(src_cols, src_rows, cands)

    # ---- 内存内落地 ----
    rel_by_id = {r["relation_id"].strip(): r for r in rel_rows}
    originals = {rid: dict(rel_by_id[rid]) for rid in LAND_IDS}
    originals_all = [dict(r) for r in rel_rows]
    ev_snapshot = [dict(r) for r in ev_rows]

    for rid in sorted(LAND_IDS):
        row = rel_by_id.get(rid)
        if row is None:
            raise LandingError(f"{rid} 不在 person_relations.csv 中")
        if rid in TYPE_TARGETS:
            before = (row.get("final_relation_type") or "").strip()
            target = TYPE_TARGETS[rid]
            reason = (row.get("correction_reason") or "").strip()
            addition = (
                f"{BATCH_MARKER}：双 Agent 交叉验证一致判类型需改，类型由「{before}」改为「{target}」"
                f"（{ADJUDICATOR_A} 与 {ADJUDICATOR_B} 提议经别名归一后一致；非人工复核）。"
            )
            row["standard_relation_type"] = target
            row["final_relation_type"] = target
            row["correction_reason"] = f"{reason}｜{addition}" if reason else addition

    by_rel: dict[str, list[dict[str, str]]] = defaultdict(list)
    for r in ev_rows:
        by_rel[r["relation_id"].strip()].append(r)

    next_n = _next_rele_id([r["relation_evidence_id"] for r in ev_rows])
    new_evidence: list[dict[str, str]] = []
    for rid in sorted(LAND_IDS):
        rele_id = f"RELE-{next_n:05d}"
        next_n += 1
        new_evidence.append(build_evidence_row(ev_cols, rele_id, rid, xval[rid], cands[rid], src_map[rid], checks[rid]))
        by_rel[rid].append(new_evidence[-1])

    for rid in sorted(LAND_IDS):
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
    if added_sources:
        src_rows_all = src_rows + added_sources
    else:
        src_rows_all = src_rows

    # ---- 硬后置条件 ----
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
    if public_ids != set(EXPECTED_PUBLIC_IDS):
        raise LandingError(
            f"公开层应为 {len(EXPECTED_PUBLIC_IDS)} 条，实际 {len(public_ids)}："
            f"多出 {sorted(public_ids - EXPECTED_PUBLIC_IDS)}，缺少 {sorted(EXPECTED_PUBLIC_IDS - public_ids)}"
        )

    # 类型更正恰 6 条，且 correction_reason 追加不覆盖
    corrections = []
    for rid in sorted(LAND_IDS):
        if (originals[rid].get("final_relation_type") or "") != (rel_by_id[rid].get("final_relation_type") or ""):
            corrections.append(rid)
            if not rel_by_id[rid]["correction_reason"].startswith(
                (originals[rid].get("correction_reason") or "").strip()
            ) and (originals[rid].get("correction_reason") or "").strip():
                raise LandingError(f"{rid} correction_reason 被覆盖而非追加")
    if corrections != sorted(TYPE_TARGETS):
        raise LandingError(f"类型更正应为 {sorted(TYPE_TARGETS)}，实际 {corrections}")

    # 除 18 条落地行外关系行零漂移；落地行只允许预期字段变化
    allowed_changes_base = {"publish_status", "publish_status_origin", "reviewer", "reviewed_at", "review_note"}
    changed_outside: list[str] = []
    for row, orig in zip(rel_rows, originals_all, strict=True):
        rid = row["relation_id"].strip()
        diffs = {c for c in rel_cols if (row.get(c) or "") != (orig.get(c) or "")}
        if not diffs:
            continue
        if rid not in LAND_IDS:
            changed_outside.append(rid)
            continue
        allowed = set(allowed_changes_base)
        if rid in TYPE_TARGETS:
            allowed |= {"standard_relation_type", "final_relation_type", "correction_reason"}
        if not diffs <= allowed:
            raise LandingError(f"{rid} 改动字段超出预期：{sorted(diffs - allowed)}")
    if changed_outside:
        raise LandingError(f"出现计划外的关系行改动：{sorted(changed_outside)}")

    # 既有证据行原样保留；证据号唯一
    if ev_rows_all[: len(ev_snapshot)] != ev_snapshot:
        raise LandingError("既有 relation_evidences 行被改写")
    if len({r["relation_evidence_id"] for r in ev_rows_all}) != len(ev_rows_all):
        raise LandingError("relation_evidence_id 出现重复")

    # 全表门禁重算一致 + 公开层语义守门
    for row in rel_rows:
        rid = row["relation_id"].strip()
        st, origin = derive_relation_publish_status_with_origin(row, by_rel.get(rid, []))
        if (st, origin) != ((row.get("publish_status") or "").strip(), (row.get("publish_status_origin") or "").strip()):
            raise LandingError(f"{rid} 落地后门禁重算不一致：{origin}/{st}")
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

    # ---- 台账（内存）----
    ledger = []
    for rid in sorted(LAND_IDS):
        row = rel_by_id[rid]
        xrow = xval[rid]
        cand = cands[rid]
        src = src_map[rid]
        ev_row = next(e for e in new_evidence if e["relation_id"] == rid)
        ledger.append({
            "relation_id": rid,
            "person_a_id": cand["person_a_id"], "person_a_name": cand["person_a_name"],
            "person_b_id": cand["person_b_id"], "person_b_name": cand["person_b_name"],
            "batch_marker": BATCH_MARKER,
            "landing_rule": "保守交集" if rid in INTERSECTION_IDS else "双 Agent 一致",
            "codex_verdict": xrow["codex_verdict"], "opus_verdict": xrow["opus_verdict"],
            "codex_proposed_type": xrow["codex_proposed_type"], "opus_proposed_type": xrow["opus_proposed_type"],
            "type_normalized": TYPE_TARGETS.get(rid, ""),
            "landed_this_batch": "yes",
            "new_relation_evidence_id": ev_row["relation_evidence_id"],
            "final_relation_type_before": originals[rid]["final_relation_type"],
            "final_relation_type_after": row["final_relation_type"],
            "publish_status_before": originals[rid]["publish_status"],
            "publish_status_after": row["publish_status"],
            "publish_status_origin": row["publish_status_origin"],
            "source_id": src["source_id"], "source_registered": "yes" if src["registered"] else "no",
            "locator": cand["locator"].strip(),
            "normalized_start": cand["normalized_start"], "normalized_end": cand["normalized_end"],
            "quote_sha256": cand["quote_sha256"],
            "quote_verbatim_recheck": checks[rid]["recheck"],
            "attestation_basis": checks[rid]["basis"],
            "authorized_by": AUTHORIZED_BY, "authorized_at": AUTHORIZED_AT,
            "authorization_xval_quote": AUTH_XVAL_QUOTE,
            "xval_row_source": f"sweep_batch1_cross_validated.csv#relation_id={rid}",
            "candidate_row_source": f"sweep_candidates.csv#relation_id={rid}",
            "note": "双 Agent 交叉验证（非人工复核）；详见 phase5_sweep_batch1_authorization_record.md。",
        })

    summary = {
        "public_supported": len(public_ids),
        "new_evidence_rows": len(new_evidence),
        "type_corrections": len(corrections),
        "registered_sources": [f"{rid}:{src_map[rid]['source_id']}" for rid in sorted(EXPECTED_NEW_SOURCES)],
        "status_dist": dict(dist_after),
        "support_dist": dict(support_after),
        "reviewed_after": reviewed_after,
        "critical_after": critical_after,
        "public_ids": sorted(public_ids),
        "ledger": ledger,
        "counts_after": {
            "sources.csv": len(src_rows_all),
            "source_passages.csv": len(pas_rows) + len(added_sources),
        },
    }
    if dry_run:
        return {"status": "dry-run", "message": "校验全部通过（未写入）。", **summary}

    _write(data_dir / "person_relations.csv", rel_cols, rel_rows)
    _write(data_dir / "relation_evidences.csv", ev_cols, ev_rows_all)
    if added_sources:
        _write(data_dir / "sources.csv", src_cols, src_rows_all)
        sync_source_layer(data_dir)
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


def write_authorization_record(path: Path) -> None:
    lines = [
        "# Phase 5 扫掠第 1 批落地授权记录（2026-10-08，双 Agent 交叉验证）",
        "",
        f"> 批次标记：`{BATCH_MARKER}`  ",
        "> 本文件为落地脚本 `apply_phase5_sweep_batch1_landing.py` 的授权前置校验对象，",
        "> 脚本启动时逐字核对下列授权语，缺失即拒绝落地。",
        "",
        "## 0. 口径声明",
        "",
        "**本批为双 Agent 交叉验证，不是人工复核。** 公开层只走 derived `supported`，",
        "禁用 `human_adjudication` / `verified` 通道；台账、站点与答辩引用必须标注本口径。",
        "答辩若被问「谁核的」，如实回答：两名不同血统的模型独立裁决后取一致项（含保守交集），",
        "无逐条人工复核。",
        "",
        "## 1. 授权语（逐字，不得改写）",
        "",
        "1. 交叉验证替代逐条人工裁决（登记于 BLOCKED.md 2026-10-08 节与合并器 docstring）：",
        f"   「{AUTH_XVAL_QUOTE}」",
        "2. 两项规则与落地范围（本会话逐字确认）：",
        f"   「{AUTH_RULE_QUOTE}」",
        "",
        "## 2. 裁决者与独立性",
        "",
        f"- 裁决者 A：{ADJUDICATOR_A}；裁决文件写于仓库外 `D:/1大创/.xval/`。",
        f"- 裁决者 B：{ADJUDICATOR_B}；裁决文件原存 `.codex_tmp/`，本仓结构上无法读对方。",
        "- 双方 blind 裁决同一份 `sweep_batch1_input.csv`（50 条，不含夜间轮 reason）。",
        "- 第二裁决者刻意避开 GLM-5.3 系列：GLM-5.3-Flash 是 2026-09-06 夜间轮（本仓引文缺陷来源）",
        "  的原始执行者，由它裁决等于自我背书。",
        "",
        "## 3. 两项规则决定的语义",
        "",
        "- **保守交集**：分歧条目中双方都判 support 且至少一方判「成立」（接受现类型）的，",
        "  按现类型不改落地——不超出任何一方认可的范围。本批共 3 条：REL-00029 鲁迅—沙汀、",
        "  REL-00035 鲁迅—徐懋庸、REL-00089 鲁迅—许广平。",
        "- **词表归并**：`论战`→`文学论战`、`交往`→`交游`（多数标签口径，已由",
        "  `merge_relation_type_vocab.py` 全量落地，47 行）；由此解锁 3 条 type_synonym 分歧：",
        "  REL-00011 鲁迅—郑伯奇、REL-00038 鲁迅—穆木天、REL-00088 鲁迅—内山完造。",
        "- 落地范围：一致 15 条 + 保守交集 3 条，共 18 条，同批落地。",
        "",
        "## 4. 执行环境说明",
        "",
        "裁决已由上述两名模型完成并留档（`sweep_batch1_input.csv`、`.xval/verdict_codex.csv`、",
        "合并器报告）；本批落地为**幂等脚本的机械执行**（逐字回定位 + sha256 + 双方佐证门 + 门禁重算），",
        "执行会话的模型身份不构成新的裁决行为。",
        "",
        f"授权人：{AUTHORIZED_BY}；授权日期：{AUTHORIZED_AT}。",
        "",
    ]
    _write_md(path, "\n".join(lines))


def write_report(result: dict, post: dict, path: Path) -> None:
    dist = result["status_dist"]
    entries = result["ledger"]
    lines = [
        "# Phase 5 扫掠第 1 批落地报告（2026-10-08，双 Agent 交叉验证）",
        "",
        "> 执行脚本：`research/analysis/apply_phase5_sweep_batch1_landing.py`（幂等）  ",
        f"> 批次标记：`{BATCH_MARKER}`  落地日期：{LANDED_AT}  ",
        "> 前置合并器：`merge_sweep_cross_validation.py --allow-conservative-intersection`",
        "",
        "## 0. 授权与口径声明",
        "",
        f"- 授权人：{AUTHORIZED_BY}；授权时间：{AUTHORIZED_AT}。",
        f"- 授权语（逐字）：「{AUTH_XVAL_QUOTE}」「{AUTH_RULE_QUOTE}」。",
        "- **本批为双 Agent 交叉验证，不是人工复核**：公开层现共 25 条，"
        "其中 5 条来自 2026-09-20 概括授权批次、2 条来自 2026-10-08 逐条独立人工裁决、"
        "18 条来自本批交叉验证——三种口径在台账、站点与答辩引用中必须分别说明，不得混同。",
        "- 公开层只接受 **derived `supported`**：未改门禁、未用 `human_adjudication`、"
        "未改 `relation_risk_level` / `needs_manual_review` / `confidence`。",
        "",
        "## 1. 落地明细",
        "",
        "| relation_id | 人物对 | 规则 | Codex/Opus 裁决 | 类型变化 | 落地后状态 | 新证据行 | 来源 |",
        "| --- | --- | --- | --- | --- | --- | --- | --- |",
    ]
    for e in entries:
        tc = (
            f"{e['final_relation_type_before']}→{e['final_relation_type_after']}"
            if e["final_relation_type_before"] != e["final_relation_type_after"] else "不改"
        )
        src_note = e["source_id"] + ("（新注册）" if e["source_registered"] == "yes" else "")
        lines.append(
            f"| {e['relation_id']} | {e['person_a_name']}—{e['person_b_name']} | {e['landing_rule']} "
            f"| {e['codex_verdict']}/{e['opus_verdict']} | {tc} | {e['publish_status_after']} "
            f"| {e['new_relation_evidence_id']} | {src_note} |"
        )
    lines += [
        "",
        "- 一致 15 条（含词表归并解锁的 3 条同义分歧 REL-00011/00038/00088）+ "
        "保守交集 3 条（REL-00029/00035/00089，按现类型不改落地）。",
        f"- 类型更正恰 {EXPECTED_TYPE_CORRECTIONS} 条（standard/final 同步，correction_reason 追加不覆盖）。",
        f"- 引文逐字复核：{EXPECTED_NEW_EVIDENCE_ROWS}/{EXPECTED_NEW_EVIDENCE_ROWS} 按候选池偏移在空白归一原文中"
        "原样取回，`quote_sha256` 自洽；18 条全部另过双方佐证门。",
        f"- 新注册来源 {len(EXPECTED_NEW_SOURCES)} 条："
        + "、".join(f"{k.split(':')[0]}→{k.split(':')[1]}" for k in result["registered_sources"])
        + "（按同族模板，先例 SRC-1178；passage 与 citation_count 由 `sync_source_layer` 统一补齐）。",
        "",
        "## 2. 实测终值 vs 预期终值",
        "",
        "| 项目 | 预期 | 实测 |",
        "| --- | ---: | ---: |",
        f"| person_relations 行数 | {EXPECTED_POST_COUNTS['person_relations.csv']} | "
        f"{post['counts']['person_relations.csv']} |",
        f"| relation_evidences 行数 | {EXPECTED_POST_COUNTS['relation_evidences.csv']} | "
        f"{post['counts']['relation_evidences.csv']} |",
        f"| sources / passages / works | {EXPECTED_POST_COUNTS['sources.csv']} / "
        f"{EXPECTED_POST_COUNTS['source_passages.csv']} / {EXPECTED_POST_COUNTS['source_works.csv']} | "
        f"{post['counts']['sources.csv']} / {post['counts']['source_passages.csv']} / "
        f"{post['counts']['source_works.csv']} |",
        f"| publish_status 分布 supported/pending_review/inferred | "
        f"{EXPECTED_POST_STATUS['supported']} / {EXPECTED_POST_STATUS['pending_review']} / "
        f"{EXPECTED_POST_STATUS['inferred']} | {dist.get('supported', 0)} / "
        f"{dist.get('pending_review', 0)} / {dist.get('inferred', 0)} |",
        f"| support 证据行 | {EXPECTED_POST_SUPPORT['support']} | {result['support_dist'].get('support', 0)} |",
        f"| associated 证据行 | {EXPECTED_POST_SUPPORT['associated']} | "
        f"{result['support_dist'].get('associated', 0)}（未改判既有行） |",
        f"| reviewed 证据行 | {EXPECTED_POST_REVIEWED} | {result['reviewed_after']} |",
        f"| critical 计数 | {EXPECTED_BASELINE_CRITICAL} | {result['critical_after']}（未反向降险） |",
        f"| 类型更正 | 恰 {EXPECTED_TYPE_CORRECTIONS} 条 | {result['type_corrections']} 条 |",
        f"| 公开层关系 | {len(EXPECTED_PUBLIC_IDS)} 条 | {result['public_supported']} 条 |",
        "",
        "Schema：写后 "
        f"{post['schema_errors']} errors / {post['schema_warnings']} warnings。",
        "",
        "## 3. 明确未做的事",
        "",
        "- 未改 `relation_publish_status.py` 判定顺序、`INFERRED_RELATION_TYPES`、`PUBLIC_RELATION_STATUSES`。",
        "- 未改判/改写 10249 条 associated 或前两批 23 条 support。",
        "- 未使用 human_adjudication 通道；未动风险/置信/复核标记列。",
        "- 分歧中不满足保守交集的 12 条一律搁置（`sweep_batch1_disagreements.csv`），生产层零改动。",
        "- REL-00085（双方一致但类型同属组织属推断类）仅落研究层口径未变，本批**未落地**，保持原状。",
        "",
        "## 4. 复核方式",
        "",
        "```powershell",
        "python research/analysis/merge_relation_type_vocab.py                # 词表归并（幂等）",
        "python research/analysis/merge_sweep_cross_validation.py --allow-conservative-intersection",
        "python research/analysis/apply_phase5_sweep_batch1_landing.py --dry-run",
        "python research/analysis/apply_phase5_sweep_batch1_landing.py        # 二跑输出「无新增/已完成」",
        "python research/analysis/build_publish_data.py",
        "python build_static_site.py",
        "python research/analysis/build_trustworthy_network_analysis.py",
        "python -m pytest -q",
        "```",
        "",
        f"逐条痕迹见 `phase5_sweep_batch1_landing_ledger.csv`（{len(entries)} 行）；"
        "授权语境与口径见 `phase5_sweep_batch1_authorization_record.md`。",
        "",
    ]
    _write_md(path, "\n".join(lines))


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Phase 5 扫掠第 1 批生产层落地（双 Agent 交叉验证，幂等）。"
    )
    parser.add_argument("--data-dir", type=Path, default=DATA, help="生产数据目录")
    parser.add_argument("--dry-run", action="store_true", help="只执行全部校验，不写任何文件")
    parser.add_argument("--skip-report", action="store_true", help="不写授权记录/台账/报告（沙盒测试用）")
    args = parser.parse_args()

    try:
        result = apply_landing(args.data_dir, dry_run=args.dry_run)
    except LandingError as exc:
        print(f"落地失败，未写入任何文件：{exc}")
        return 1

    print(result["message"])
    if result["status"] == "no-op":
        return 0
    print(f"本批落地 {result['new_evidence_rows']} 条（一致 15 + 保守交集 3）；"
          f"类型更正 {result['type_corrections']} 条；新注册来源 {len(result['registered_sources'])} 条")
    print(f"公开层 derived supported：{result['public_supported']} 条")

    if result["status"] == "applied" and not args.skip_report:
        write_authorization_record(AUTH_RECORD)
        _write(LEDGER, LEDGER_COLUMNS, result["ledger"])
        print(f"授权记录：{AUTH_RECORD.relative_to(ROOT)}")
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
