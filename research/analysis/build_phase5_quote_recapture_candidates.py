"""Phase 5 引文重捕候选包：按句边界重切引文，全部 pending_human_review，生产层零改动。

背景
----
2026-09-28 落地时发现夜间轮记录的 ``quote`` 多为按页码截取的固定窗口，与同行 ``reason``
描述的史实不一致：48 条候选中只有 20 条引文真正同时记载双方当事人。28 条进入
``phase5_quote_recapture_queue.csv``，其中 9 条被标为 ``recapture_quote_then_regrade``
（同一本地原文中确实存在同时记载两人的叙述性段落）。

本脚本只做一件事：把那 9 条的**正确段落按句边界重切出来**，逐字校验后作为候选交人工裁决。
不写生产层、不改 ``relation_evidences``、不改 ``publish_status``、不降低 ``relation_risk_level``。

与直接采用夜间轮 ``reason`` 的区别
--------------------------------
``reason`` 是自然语言转述，不能当引文用（本批已证明它与 quote 经常不一致）。
本脚本一律从本地原文重新切取，切取结果必须：
1. 是归一化文本中的**连续片段**（可逐字回定位，记录起止偏移）；
2. 以句读边界收束，不是任意固定窗口；
3. 不跨页标记（``────``）、不超长度上限；
4. 不是纯顿号人名罗列（``looks_like_name_list``）；
5. 重切后仍通过双方佐证门（``quote_attests_pair``）。
任一不满足即标为 rejected_* 并给出原因，不静默降级、不猜测。

另外两类不做重捕，只登记建议：
- ``cooccurrence_only_keep_associated``：同窗只是人名罗列，属关联级，不得升 support；
- ``no_local_support_mark_insufficient``：本地全文找不到任何同窗共现，应判证据不足。
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import re
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
DATA = ROOT / "data" / "processed"
REPORTS = ROOT / "research" / "drafts" / "reports"
QUEUE = REPORTS / "phase5_quote_recapture_queue.csv"
OUT_CSV = REPORTS / "phase5_quote_recapture_candidates.csv"
OUT_MD = REPORTS / "phase5_quote_recapture_report.md"

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
        PUBLIC_RELATION_STATUSES,
        assign_relation_status_columns,
    )
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
        PUBLIC_RELATION_STATUSES,
        assign_relation_status_columns,
    )

SENTENCE_DELIMS = "。；！？"
PAGE_MARKER = "────"
MAX_QUOTE_CHARS = 400
PAGE_RE = re.compile(r"第(\d{1,4})页")
ENUM_TAIL_RE = re.compile(r"等共?\d{1,4}[余位人]?")
BOOK_BY_FILE = {"左联史.txt": "左联史", "左联词典.txt": "左联词典"}

COLUMNS = [
    "relation_id", "person_a_id", "person_a_name", "person_b_id", "person_b_name",
    "current_final_relation_type", "proposed_relation_type", "human_verdict",
    "queue_proposed_action", "recapture_status", "recaptured_quote", "quote_char_len",
    "quote_form", "quote_sha256", "derived_locator", "recorded_locator", "locator_agrees",
    "source_file", "normalized_start", "normalized_end", "attestation_basis", "list_like",
    "name_list_pattern",
    "projected_publish_status", "projected_public", "old_recorded_quote",
    "old_attestation_basis", "overnight_reason", "review_status",
]

SOURCE_DIRS = (DATA / "runtime_sources", ROOT / "research" / "raw_texts")


def resolve_source(raw: str) -> Path:
    """把队列里记录的来源路径解析成本机可读路径。

    队列的 ``local_file`` 是生成当日的本机绝对路径（含 Windows 盘符），换机器或换系统即失效。
    因此先按原样试，再按**文件名**在已知源目录里找——与
    ``audit_phase7_candidates_independent._resolve_source_path`` 同一做法。
    找不到就抛错，绝不静默跳过（跳过等于让候选包少一行而没人知道）。
    """
    text = str(raw or "").strip()
    if not text:
        raise SystemExit("队列行缺少 local_file，无法定位原文")
    direct = Path(text)
    if direct.exists():
        return direct
    name = direct.name
    for directory in SOURCE_DIRS:
        candidate = directory / name
        if candidate.exists():
            return candidate
    raise SystemExit(f"来源文件不可用（原样与按文件名均未命中）: {text}")


def _read(path: Path) -> list[dict[str, str]]:
    with open(path, encoding="utf-8-sig", newline="") as fh:
        return list(csv.DictReader(fh))


def _expand_sentence(flat: str, lo: int, hi: int) -> tuple[int, int]:
    """把 [lo, hi) 向两侧扩到句读边界，返回含末尾句读的区间。"""
    left = lo
    while left > 0 and flat[left - 1] not in SENTENCE_DELIMS:
        left -= 1
    right = hi
    while right < len(flat) and flat[right] not in SENTENCE_DELIMS:
        right += 1
    if right < len(flat):
        right += 1
    return left, right


def _derive_locator(flat: str, pos: int, book: str) -> str:
    """取该位置之前最后一个页码标记，拼成与既有登记一致的 locator。"""
    last = None
    for m in PAGE_RE.finditer(flat[:pos]):
        last = m
    if last is None:
        return f"{book} 页码未识别"
    return f"{book} 第{last.group(1)}页"


def name_list_pattern(text: str) -> bool:
    """识别「名单句式」：以「等N人／等共N人」收束，或含长串顿号并列的短词。

    比 ``looks_like_name_list`` 更严：后者只看顿号密度与动词，会被「…签名的有…等共189人」
    这类含动词的长名单放过。名单句式对「签名联署」可能仍是直接证据，但对「交游／同属组织」
    只是共现，因此这里只做**标注**交人工判断，不自动否决。
    """
    flat = re.sub(r"\s+", "", str(text or ""))
    if ENUM_TAIL_RE.search(flat):
        return True
    tokens = [t for t in re.split(r"[、,，]", flat) if 2 <= len(t) <= 4]
    return len(tokens) >= 8


def _project_status(
    rel_row: dict[str, str],
    existing_evidence: list[dict[str, str]],
    new_quote: str,
    new_locator: str,
    corrected_type: str,
) -> str:
    """只读投影：若把重切引文定为 support，该关系经现有门禁会得到什么 publish_status。"""
    row = dict(rel_row)
    if corrected_type:
        row["final_relation_type"] = corrected_type
    evidences = [dict(e) for e in existing_evidence]
    evidences.append({
        "relation_id": row.get("relation_id", ""),
        "evidence_support": "support",
        "review_status": "reviewed",
        "locator": new_locator,
        "quote": new_quote,
        "context": new_quote,
    })
    cols = assign_relation_status_columns(row, evidences, existing={
        "reviewer": row.get("reviewer", ""), "reviewed_at": row.get("reviewed_at", ""),
        "review_note": row.get("review_note", ""),
    })
    return cols["publish_status"]


def recapture_one(
    row: dict[str, str],
    persons: dict[str, dict[str, str]],
    bundles: dict[str, tuple[str, str, list[int]]],
    rel_row: dict[str, str],
    existing_evidence: list[dict[str, str]],
    corrected_type: str,
) -> dict[str, str]:
    out = {c: "" for c in COLUMNS}
    out.update({
        "relation_id": row["relation_id"],
        "person_a_id": row["person_a_id"], "person_a_name": row["person_a_name"],
        "person_b_id": row["person_b_id"], "person_b_name": row["person_b_name"],
        "current_final_relation_type": (rel_row.get("final_relation_type") or "").strip(),
        "proposed_relation_type": row.get("final_relation_type", ""),
        "human_verdict": row.get("human_verdict", ""),
        "queue_proposed_action": row.get("proposed_action", ""),
        "recapture_status": "",
        "quote_form": "whitespace_normalized",
        "derived_locator": "", "recorded_locator": row.get("recorded_locator", ""),
        "locator_agrees": "", "source_file": row.get("local_file", ""),
        "normalized_start": "", "normalized_end": "", "attestation_basis": "",
        "list_like": "", "name_list_pattern": "",
        "projected_publish_status": "", "projected_public": "",
        "old_attestation_basis": row.get("attestation_basis", ""),
        "old_recorded_quote": row.get("recorded_evidence_id", ""),
        "overnight_reason": row.get("overnight_reason", ""),
        "review_status": "pending_human_review",
    })
    action = row.get("proposed_action", "")
    if action != "recapture_quote_then_regrade":
        out["recapture_status"] = f"not_attempted_{action}"
        return out

    src = row.get("local_file", "")
    bundle = bundles.get(src)
    if bundle is None:
        out["recapture_status"] = "rejected_source_unavailable"
        return out
    text, flat, idx = bundle
    book = BOOK_BY_FILE.get(resolve_source(src).name, resolve_source(src).stem)
    pa = persons.get(row["person_a_id"], {})
    pb = persons.get(row["person_b_id"], {})
    na, nb = name_candidates(pa), name_candidates(pb)
    diary = book == "鲁迅日记"

    hits = find_cooccurrence((text, flat, idx), na, nb, limit=1)
    if not hits:
        out["recapture_status"] = "rejected_no_cooccurrence"
        return out
    hit = hits[0]
    pos_a = int(hit["normalized_pos"])
    span_lo, span_hi = pos_a, pos_a + len(str(hit["name_a"]))
    bpos = flat.find(str(hit["name_b"]), max(0, span_lo - 260))
    if bpos >= 0:
        span_lo = min(span_lo, bpos)
        span_hi = max(span_hi, bpos + len(str(hit["name_b"])))
    left, right = _expand_sentence(flat, span_lo, span_hi)
    quote = flat[left:right]

    if PAGE_MARKER in quote:
        out["recapture_status"] = "rejected_crosses_page_marker"
        out["recaptured_quote"] = quote[:MAX_QUOTE_CHARS]
        out["normalized_start"], out["normalized_end"] = str(left), str(right)
        return out
    if len(quote) > MAX_QUOTE_CHARS:
        out["recapture_status"] = "rejected_span_too_long"
        out["quote_char_len"] = str(len(quote))
        out["recaptured_quote"] = quote[:MAX_QUOTE_CHARS]
        out["normalized_start"], out["normalized_end"] = str(left), str(right)
        return out
    if looks_like_name_list(quote):
        out["recapture_status"] = "rejected_name_list_only"
        out["list_like"] = "yes"
        out["name_list_pattern"] = "yes" if name_list_pattern(quote) else "no"
        out["recaptured_quote"] = quote
        out["quote_char_len"] = str(len(quote))
        out["normalized_start"], out["normalized_end"] = str(left), str(right)
        return out
    locator = _derive_locator(flat, left, book)
    recorded = (row.get("recorded_locator") or "").strip()
    if recorded and locator != recorded:
        # 页码不一致意味着这不是「同一处引文重切」，而是在全书别处找到了另一段共现。
        # 实测这类段落全部不成立（阳翰笙—林淡秋 189 人签名名单、殷夫—李辉英 冯铿词条、
        # 叶紫—戴望舒 书目提要），故直接拒绝，交人工另行处理。
        out["recapture_status"] = "rejected_locator_page_mismatch"
        out["derived_locator"] = locator
        out["locator_agrees"] = "no"
        out["recaptured_quote"] = quote[:MAX_QUOTE_CHARS]
        out["quote_char_len"] = str(len(quote))
        out["normalized_start"], out["normalized_end"] = str(left), str(right)
        out["name_list_pattern"] = "yes" if name_list_pattern(quote) else "no"
        return out
    ok, basis = quote_attests_pair(quote, na, nb, diary_author_implicit=diary)
    if not ok:
        out["recapture_status"] = f"rejected_attestation_{basis}"
        out["attestation_basis"] = basis
        out["recaptured_quote"] = quote
        out["quote_char_len"] = str(len(quote))
        return out
    # 逐字回定位：切出的片段必须原样存在于归一化文本中
    if flat.count(quote) < 1 or flat[left:right] != quote:
        out["recapture_status"] = "rejected_relocate_failed"
        return out

    status = _project_status(rel_row, existing_evidence, quote, locator, corrected_type)
    out.update({
        "recapture_status": "recaptured",
        "recaptured_quote": quote,
        "quote_char_len": str(len(quote)),
        "quote_sha256": hashlib.sha256(quote.encode("utf-8")).hexdigest(),
        "derived_locator": locator,
        "locator_agrees": "yes" if locator == (row.get("recorded_locator") or "").strip() else "no",
        "normalized_start": str(left), "normalized_end": str(right),
        "attestation_basis": basis,
        "list_like": "no",
        "name_list_pattern": "yes" if name_list_pattern(quote) else "no",
        "projected_publish_status": status,
        "projected_public": "yes" if status in PUBLIC_RELATION_STATUSES else "no",
    })
    return out

def build(queue_path: Path = QUEUE) -> list[dict[str, str]]:
    queue = _read(queue_path)
    if not queue:
        raise SystemExit(f"重捕队列为空或不存在：{queue_path}")
    persons = {r["person_id"].strip(): r for r in _read(DATA / "persons.csv")}
    relations = {r["relation_id"].strip(): r for r in _read(DATA / "person_relations.csv")}
    evidence_by_rel: dict[str, list[dict[str, str]]] = {}
    for r in _read(DATA / "relation_evidences.csv"):
        evidence_by_rel.setdefault(r["relation_id"].strip(), []).append(r)
    adj = {r["relation_id"]: r for r in _read(REPORTS / "phase5_relation_review_adjudicated.csv")}

    bundles: dict[str, tuple[str, str, list[int]]] = {}
    for row in queue:
        src = (row.get("local_file") or "").strip()
        if src and src not in bundles:
            bundles[src] = normalized_bundle(resolve_source(src))

    out: list[dict[str, str]] = []
    for row in sorted(queue, key=lambda r: r["relation_id"]):
        rid = row["relation_id"]
        rel_row = relations.get(rid)
        if rel_row is None:
            raise SystemExit(f"{rid} 不在 person_relations.csv 中")
        verdict = (row.get("human_verdict") or adj.get(rid, {}).get("human_verdict") or "").strip()
        corrected = ""
        if verdict == "wrong_type":
            corrected = (adj.get(rid, {}).get("ai_suggested_type") or "").strip()
        item = recapture_one(
            row, persons, bundles, rel_row,
            evidence_by_rel.get(rid, []), corrected,
        )
        if not item["human_verdict"]:
            item["human_verdict"] = verdict
        out.append(item)
    return out


def write_report(rows: list[dict[str, str]], path: Path) -> None:
    status = Counter(r["recapture_status"] for r in rows)
    recaptured = [r for r in rows if r["recapture_status"] == "recaptured"]
    would_public = [r for r in recaptured if r["projected_public"] == "yes"]
    loc_agree = Counter(r["locator_agrees"] for r in recaptured)
    lines = [
        "# Phase 5 引文重捕候选包（全部 pending_human_review，生产层零改动）",
        "",
        "> 生成脚本：`research/analysis/build_phase5_quote_recapture_candidates.py`（只读）  ",
        "> 输入：`phase5_quote_recapture_queue.csv`（28 条，2026-09-28 双方佐证门未通过者）  ",
        "> 输出：`phase5_quote_recapture_candidates.csv`",
        "",
        "## 0. 这份文件不是什么",
        "",
        "- **不是已证史实**，也不是已转正的证据。每一行 `review_status=pending_human_review`。",
        "- 本脚本不写 `relation_evidences.csv`、不改 `publish_status`、不改 `relation_risk_level`。",
        "- 重切出的引文只是「同一本地原文中确实存在、且同时提到双方当事人」的候选段落；",
        "  它是否真的证明该条关系（而不只是同页并列、同名异人、职务关联被读成交游），",
        "  必须由人读了原句再判。夜间轮的 `reason` 字段已被证明不可直接采信。",
        "",
        "## 1. 处理结果",
        "",
        "| recapture_status | 条数 | 含义 |",
        "| --- | ---: | --- |",
    ]
    meaning = {
        "recaptured": "按句边界重切成功，逐字可回定位，通过双方佐证门",
        "rejected_name_list_only": "同窗只是顿号人名罗列，属关联级，不得升 support",
        "rejected_crosses_page_marker": "句子跨页标记，切取不连续，需人工另选段落",
        "rejected_span_too_long": f"句 span 超过 {MAX_QUOTE_CHARS} 字上限，需人工收窄",
        "rejected_no_cooccurrence": "本地全文找不到同窗共现",
        "rejected_locator_page_mismatch": "重切段落不在夜间轮登记的同一页——属全书别处的另一段共现，不是同一处引文的重切",
        "rejected_source_unavailable": "本地原文不可用",
        "not_attempted_cooccurrence_only_keep_associated": "队列已判为仅共现，不重捕",
        "not_attempted_no_local_support_mark_insufficient": "队列已判本地无佐证，不重捕",
    }
    for key, count in status.most_common():
        lines.append(f"| `{key}` | {count} | {meaning.get(key, '见脚本')} |")
    lines += [
        "",
        f"- 重切成功：**{len(recaptured)}** 条；其中若把引文定为 support，"
        f"经现有发布门禁投影可进入公开层的为 **{len(would_public)}** 条"
        f"（`projected_public=yes`）。其余被 critical/high 风险或推断类型挡住，投影仅供参考。",
        f"- 重切后 locator 与夜间轮登记的页码一致：{loc_agree.get('yes', 0)} 条；"
        f"不一致：{loc_agree.get('no', 0)} 条（不一致本身是夜间轮定位漂移的旁证，须人工确认以哪个为准）。",
        "- **投影不等于落地**：`projected_publish_status` 是用 `assign_relation_status_columns` "
        "对「假设该引文被人工定为 support」做的只读推演，用于估计工作量与收益，不构成任何状态变更。",
        "",
        "## 2. 重切成功的候选（逐条待人工裁决）",
        "",
    ]
    for r in sorted(recaptured, key=lambda x: x["relation_id"]):
        lines += [
            f"### {r['relation_id']}　{r['person_a_name']} — {r['person_b_name']}"
            f"（现类型「{r['current_final_relation_type']}」"
            + (f"，裁决建议改为「{r['proposed_relation_type']}」" if r["human_verdict"] == "wrong_type" else "")
            + "）",
            "",
            f"- 重切 locator：`{r['derived_locator']}`（夜间轮登记：`{r['recorded_locator']}`，"
            f"一致={r['locator_agrees']}）",
            f"- 归一化偏移：{r['normalized_start']}–{r['normalized_end']}；"
            f"长度 {r['quote_char_len']} 字；quote_sha256 `{r['quote_sha256'][:16]}…`",
            f"- 佐证依据：`{r['attestation_basis']}`；人名罗列={r['list_like']}；"
            f"名单句式={r['name_list_pattern']}",
            f"- 投影 publish_status：`{r['projected_publish_status']}`（公开={r['projected_public']}）",
            f"- 夜间轮理由（**不可直接采信**）：{r['overnight_reason'][:200]}",
            "",
            "> 重切引文（空白归一后的原文连续片段）：",
            ">",
            f"> {r['recaptured_quote']}",
            "",
        ]
        if r["name_list_pattern"] == "yes":
            lines += [
                "> ⚠ **名单句式**：本段以「等N人」类并列名单收束。若该关系类型是 `签名联署`，"
                "同列一份名单可以是直接证据；若是 `交游`/`同属组织`，同列名单只算共现，"
                "不足以支持。**还须核对重切段落与夜间轮 `reason` 所指是否为同一份文献**——"
                "本批已出现 reason 指甲文献、重切段落实为乙文献的情况，两者都真但不是同一件事。",
                "",
            ]
    lines += [
        "## 3. 人工裁决需要回答的问题",
        "",
        "1. 重切的这句原文，是否**直接**记载了这两人之间的该种关系？"
        "（同页并列、共同出现在名单里、职务关联，都不等于交游或联署。）",
        "2. 若成立，`evidence_support` 应定 `support` 还是 `associated`？",
        "3. `locator_agrees=no` 的条目，以重切页码还是夜间轮登记页码为准？",
        "4. 关系类型是否按 `proposed_relation_type` 更正？",
        "5. 被 critical/high 风险挡住但证据成立的条目，是否愿意逐条独立复核后走 "
        "`human_adjudication` → `verified` 通道（需写 reviewer/reviewed_at/review_note）？",
        "",
        "## 4. 裁决之后的落地方式",
        "",
        "参照 `apply_phase5_relation_landing.py` 另起一个幂等脚本：以本候选包为输入、"
        "只处理人工填了裁决的行、运行时重新逐字回定位并校验 quote_sha256、"
        "写前置基线与后置计数断言、二跑输出「无新增/已完成」、产出台账。"
        "不得直接改本文件或队列文件来「表示已裁决」。",
        "",
    ]
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines), encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description="生成 Phase 5 引文重捕候选包（只读，不写生产层）。")
    parser.add_argument("--queue", type=Path, default=QUEUE, help="重捕队列 CSV")
    parser.add_argument("--out", type=Path, default=OUT_CSV, help="候选包输出路径")
    parser.add_argument("--report", type=Path, default=OUT_MD, help="报告输出路径")
    args = parser.parse_args()

    rows = build(args.queue)
    args.out.parent.mkdir(parents=True, exist_ok=True)
    with open(args.out, "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=COLUMNS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)
    write_report(rows, args.report)

    status = Counter(r["recapture_status"] for r in rows)
    recaptured = [r for r in rows if r["recapture_status"] == "recaptured"]
    public = [r for r in recaptured if r["projected_public"] == "yes"]
    print(f"候选 {len(rows)} 条（生产层零改动，全部 pending_human_review）")
    for key, count in status.most_common():
        print(f"  {key}: {count}")
    print(f"重切成功 {len(recaptured)} 条；投影可进公开层 {len(public)} 条：")
    for r in sorted(public, key=lambda x: x["relation_id"]):
        print(f"  {r['relation_id']} {r['person_a_name']}—{r['person_b_name']}"
              f" | {r['derived_locator']} | {r['quote_char_len']}字")
    print(f"候选包：{args.out.relative_to(ROOT)}")
    print(f"报告：{args.report.relative_to(ROOT)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
