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
        looks_like_name_list,
        name_candidates,
        normalized_bundle,
        quote_attests_pair,
    )
    from relation_publish_status import (  # noqa: E402
        PUBLIC_RELATION_STATUSES,
        assign_relation_status_columns,
    )

# 这两本书的 OCR 大量用 ASCII 句点 "." 当句号（如"…联名签署.同年12月,…"），
# 只认 "。；！？" 会让句子向左扩张成一整页、随后因超长被拒，反而丢掉正确的那一处。
SENTENCE_DELIMS = "。；！？.．"
PAGE_MARKER = "────"
MAX_QUOTE_CHARS = 400
PAGE_RE = re.compile(r"第(\d{1,4})页")
ENUM_TAIL_RE = re.compile(r"等共?\d{1,4}[余位人]?")
TITLE_RE = re.compile(r"《([^》]{2,40})》")
# 同一页里可能有多处同窗共现（REL-01368 即命中两份不同的联名名单）。
# 只按"姓名距离最近"挑会挑错，因此多取几个候选再按下面三条排序。
COOCCURRENCE_CANDIDATES = 12
# 对这些关系类型而言，"同列一份具体文件的签署名单"本身就是直接证据，
# 不能按"纯人名罗列＝只是共现"一律否决（REL-01368 即被误杀：
# 《为横死之小林遗族募捐启》9 人签署名单正是 签名联署 的直接记载）。
# 其余类型（交游／交往／通信／创作合作…）里名单只算共现，仍然否决。
LIST_ACCEPTABLE_TYPES = ("签名联署", "同属组织")
BOOK_BY_FILE = {"左联史.txt": "左联史", "左联词典.txt": "左联词典"}

COLUMNS = [
    "relation_id", "person_a_id", "person_a_name", "person_b_id", "person_b_name",
    "current_final_relation_type", "proposed_relation_type", "human_verdict",
    "queue_proposed_action", "recapture_status", "recaptured_quote", "quote_char_len",
    "quote_form", "quote_sha256", "derived_locator", "recorded_locator", "locator_agrees",
    "source_file", "normalized_start", "normalized_end", "attestation_basis", "list_like",
    "name_list_pattern", "secondary_description",
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


SECONDARY_MARKERS = (
    "传记小说", "评传", "年谱", "论文", "书评", "研究资料", "回忆录", "摘编", "转载",
    "一书", "文中", "记述", "著录", "李克因作", "作,载", "载《",
)
SECONDARY_YEAR_RE = re.compile(r"载\d{4}")


def secondary_description(text: str) -> bool:
    """粗判该段是否为**书目著录／二手评述**而非原始记载。

    实测教训：REL-01891（叶紫—萧军）切到的是《左联词典》第587页的书目条——
    "叶紫——一颗富有而又饥饿的星 传记小说。李克因作，载《东方纪事》1987年3、4期合刊。
    叙述……叶紫同陈企霞、聂绀弩、周颖夫妇、萧军、萧红夫妇等的交往"。
    它确实同时提到两人且不是纯名单，但语义是"某本传记小说描写了他们的交往"，
    属二手著录，证据强度低于原始记载，不应径直定为 support。
    本函数只做**标注**，是否可用交人工判断。
    """
    flat = re.sub(r"\s+", "", str(text or ""))
    if SECONDARY_YEAR_RE.search(flat):
        return True
    return any(marker in flat for marker in SECONDARY_MARKERS)


def enumerate_page_candidates(
    flat: str,
    names_a: list[str],
    names_b: list[str],
    book: str,
    recorded_locator: str,
    window: int = 200,
    cap: int = 40,
) -> list[dict]:
    """枚举登记页上所有 A/B 同窗共现，各自扩成句子并算出页码。

    必须自己枚举而不能用 quote_attestation.find_cooccurrence：后者按"姓名距离最近"
    排序后截断到 limit 条，而同一页里往往有多处共现（REL-01368 的第132页既有
    《为横死之小林遗族募捐启》9 人名单，也有营救丁潘的 38 人联名致电），
    正确的那处可能因距离略远被截断挤掉，选择器就再也看不到它。
    """
    positions: dict[str, list[int]] = {}
    for name in dict.fromkeys([*names_a, *names_b]):
        found: list[int] = []
        start = 0
        while True:
            pos = flat.find(name, start)
            if pos < 0:
                break
            found.append(pos)
            start = pos + 1
        if found:
            positions[name] = found
    seen: set[tuple[int, int]] = set()
    out: list[dict] = []
    for a in names_a:
        for pa in positions.get(a, []):
            for b in names_b:
                for pb in positions.get(b, []):
                    if abs(pa - pb) > window:
                        continue
                    lo, hi = min(pa, pb), max(pa + len(a), pb + len(b))
                    key = (lo, hi)
                    if key in seen:
                        continue
                    seen.add(key)
                    left, right = _expand_sentence(flat, lo, hi)
                    quote = flat[left:right]
                    out.append({
                        "left": left, "right": right, "quote": quote,
                        "locator": _derive_locator(flat, left, book),
                        "distance": abs(pa - pb),
                    })
                    if len(out) >= cap:
                        return out
    return out


def reason_titles(reason: str) -> list[str]:
    """从夜间轮 reason 里抽出《篇名》，用作同页多候选的消歧锚。

    reason 是转述、不能当引文用，但它提到的**篇名**是可核对的客观锚点：
    同页若有多处共现，优先取含该篇名的那处，否则可能切到另一份文献
    （REL-01368 的实测教训：同页既有《为横死之小林遗族募捐启》9 人名单，
    也有营救丁玲潘梓年的 38 人联名致电，两者都真但不是同一件事）。
    """
    return [t.strip() for t in TITLE_RE.findall(str(reason or "")) if len(t.strip()) >= 2]


def pick_cooccurrence(cands: list[dict], reason: str, recorded_locator: str) -> dict | None:
    """在同页多个同窗候选里选最可能是 reason 所指那一处。

    排序键（依次）：①页码与夜间轮登记一致 → ②命中 reason 里的《篇名》 →
    ③非名单句式 → ④非书目著录／二手评述 → ⑤片段更短（更聚焦） → ⑥双方姓名距离更近。

    ①必须排在最前：同一份文献在全书会出现多次（不同人物词条各引一次），
    只按篇名挑会挑到别的词条页上，随后被同页硬门拒绝，反而丢掉本来正确的那一处。
    返回 None 表示无候选。
    """
    if not cands:
        return None
    titles = reason_titles(reason)
    recorded = (recorded_locator or "").strip()

    def key(c: dict):
        quote = c["quote"]
        page_ok = 0 if (not recorded or c["locator"] == recorded) else 1
        title_ok = 0 if any(t in quote for t in titles) else 1
        listed = 1 if looks_like_name_list(quote) else 0
        secondary = 1 if secondary_description(quote) else 0
        return (page_ok, title_ok, listed, secondary, len(quote), int(c["distance"]))

    return min(cands, key=key)


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
        "list_like": "", "name_list_pattern": "", "secondary_description": "",
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

    recorded = (row.get("recorded_locator") or "").strip()
    cands = enumerate_page_candidates(flat, na, nb, book, recorded)
    if not cands:
        out["recapture_status"] = "rejected_no_cooccurrence"
        return out
    chosen = pick_cooccurrence(cands, row.get("overnight_reason", ""), recorded)
    if chosen is None:
        out["recapture_status"] = "rejected_no_cooccurrence"
        return out
    left, right = chosen["left"], chosen["right"]
    quote = chosen["quote"]

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
    effective_type = corrected_type or (rel_row.get("final_relation_type") or "").strip()
    listed = looks_like_name_list(quote)
    if listed and effective_type not in LIST_ACCEPTABLE_TYPES:
        out["recapture_status"] = "rejected_name_list_only"
        out["list_like"] = "yes"
        out["name_list_pattern"] = "yes" if name_list_pattern(quote) else "no"
        out["secondary_description"] = "yes" if secondary_description(quote) else "no"
        out["recaptured_quote"] = quote
        out["quote_char_len"] = str(len(quote))
        out["normalized_start"], out["normalized_end"] = str(left), str(right)
        return out
    locator = chosen["locator"]
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
        out["secondary_description"] = "yes" if secondary_description(quote) else "no"
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
        "secondary_description": "yes" if secondary_description(quote) else "no",
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
                "> ⚠ **名单句式**：本段以「等N人」类并列名单收束。若该关系类型是 `签名联署` 或 `同属组织`，"
                "同列一份**具体文件**的签署名单／任职名单可以是直接证据；若是 `交游`／`交往`／`通信`，"
                "同列名单只算共现，不足以支持（这类已被自动否决，不会出现在这里）。"
                "选择器已优先取命中 reason 里《篇名》的同页段落，但仍须人工确认这份名单"
                "就是该关系记录所指的那一份——同一页可能并列多份名单（REL-01368 的第132页"
                "同时有《为横死之小林遗族募捐启》9 人名单与营救丁潘的 38 人联名致电）。",
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
