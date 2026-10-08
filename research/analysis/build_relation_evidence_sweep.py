"""全量关系证据扫掠：从本地三本全文为每条关系找可定位的句级候选证据（只读，生产零改动）。

为什么需要它
------------
2026-09-06 夜间轮只核查了 411 条关系，且用「页码 + 固定窗口」截取引文，导致约 56% 的
引文并未真正记载该关系（见 `evidence_verbatim_audit_2026-09-28.md`）。本脚本改为对**全部
4238 条关系**做句级检索：先把句子切对，再要求引文本身同时记载双方当事人。

与夜间轮做法的三点关键差别
--------------------------
1. **切句而非切窗口**：以句读边界（。；！？.．）收束；这两本书的 OCR 大量用 ASCII 句点当句号，
   只认「。；！？」会让句子扩张成一整页。
2. **佐证门作用在最终引文上**：不是「两人在附近出现过」就算，而是切出的那句里必须同时含双方
   姓名或别名（鲁迅日记为日记体，作者即甲方，只需引文出现乙方）。早期 sizing 脚本漏了这一步，
   把 755 条高估成可公开，实际会低很多。
3. **同页优先 + 不跨页标记**：句段不得含页码分隔线（`────`），页码由该句之前最后一个
   `第N页` 标记推出，与既有 `sources.csv` 的 citation 口径一致。

本脚本只做检索与标注，**不做任何史实判断**。是否成立由两名独立裁决者判定
（见 `sweep_batch*_input.csv`），且裁决输入**不含**夜间轮的 `reason` 字段——那条 reason 由
GLM-5.3-Flash 生成、已被证明与引文经常不一致，带上它只会锚定裁决者。

标注列语义（全部只作提示，不自动否决，除注明者外）
--------------------------------------------------
- `relation_verb`：句中命中的交往类动词，用于区分「叙述性互动」与「单纯并列」。
- `list_pattern`：名单句式（「等N人」收束或长串顿号并列）。对 `签名联署`/`同属组织`
  可能是直接证据，对 `交游`/`交往`/`通信` 只算共现。
- `secondary_description`：书目著录／二手评述（「载《」「载1987年」「传记小说」「评传」等）。
  证据强度低于原始记载。**硬门**：本批不把二手著录放进可公开候选。
- `projected_publish_status`：若该引文被定为 support，经现有发布门禁会得到什么状态。
  只做只读投影，不写任何状态。
"""

from __future__ import annotations

import argparse
import bisect
import csv
import hashlib
import re
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
DATA = ROOT / "data" / "processed"
REPORTS = ROOT / "research" / "drafts" / "reports"
RAW_TEXTS = ROOT / "research" / "raw_texts"

OUT_POOL = REPORTS / "sweep_candidates.csv"
OUT_REPORT = REPORTS / "sweep_report.md"

_ANALYSIS_DIR = Path(__file__).resolve().parent
try:
    from research.analysis.quote_attestation import name_candidates, quote_attests_pair
    from research.analysis.relation_publish_status import assign_relation_status_columns
except ImportError:
    if str(_ANALYSIS_DIR) not in sys.path:
        sys.path.append(str(_ANALYSIS_DIR))
    from quote_attestation import (  # noqa: E402
        name_candidates,
        quote_attests_pair,
    )
    from relation_publish_status import (  # noqa: E402
        assign_relation_status_columns,
    )

# 三本可逐字复核的本地全文；鲁迅日记按 .gitignore 政策只在本地，不在版本库内。
SOURCE_BOOKS = {
    "左联史": ("zuolian_shi", DATA / "runtime_sources" / "左联史.txt"),
    "左联词典": ("zuolian_cidian", DATA / "runtime_sources" / "左联词典.txt"),
    "鲁迅日记": ("luxun_diary", RAW_TEXTS / "日记全编：全2册 (鲁迅 著) (Z-Library).txt"),
}
DIARY_AUTHOR = "鲁迅"

SENTENCE_DELIMS = "。；！？.．"
PAGE_MARKER = "────"
PAGE_RE = re.compile(r"第(\d{1,4})页")
ENUM_TAIL_RE = re.compile(r"等共?\d{1,4}[余位人]?")
SECONDARY_YEAR_RE = re.compile(r"载\d{4}")
SECONDARY_MARKERS = (
    "传记小说", "评传", "年谱", "论文", "书评", "研究资料", "回忆录", "摘编",
    "转载", "一书", "文中", "记述", "著录", "作,载", "载《",
)
RELATION_VERBS = (
    "拜访", "拜会", "会见", "会面", "见面", "探望", "探视", "看望", "通信", "致函",
    "去信", "来信", "复信", "回信", "寄信", "写信", "联名", "联署", "签署", "签名",
    "合编", "合著", "合作", "共同", "同往", "同行", "一同", "一起", "介绍", "推荐",
    "发起", "组织", "参加", "编辑", "主编", "撰稿", "投稿", "约", "邀", "宴请",
    "聚会", "座谈", "讨论", "论战", "批判", "支持", "援助", "营救", "访", "晤", "赠",
)
# 名单句式对这些类型可能是直接证据；其余类型只算共现（硬门：其余类型直接排除名单句）
LIST_ACCEPTABLE_TYPES = ("签名联署", "同属组织")
WINDOW = 160
MAX_QUOTE = 400
MAX_SPANS_PER_RELATION = 240
POOL_COLUMNS = [
    "relation_id", "person_a_id", "person_a_name", "person_b_id", "person_b_name",
    "current_final_relation_type", "relation_risk_level", "confidence", "needs_manual_review",
    "current_publish_status", "book", "source_family", "page", "locator",
    "quote", "quote_char_len", "quote_sha256", "normalized_start", "normalized_end",
    "source_file", "attestation_basis", "relation_verb", "list_pattern",
    "secondary_description", "projected_publish_status", "projected_public",
    "review_status",
]


def list_pattern(text: str) -> bool:
    """名单句式：以「等N人／等共N人」收束，或含长串顿号并列短词。"""
    flat = re.sub(r"\s+", "", str(text or ""))
    if ENUM_TAIL_RE.search(flat):
        return True
    return len([t for t in re.split(r"[、,，]", flat) if 2 <= len(t) <= 4]) >= 8


def secondary_description(text: str) -> bool:
    """书目著录／二手评述（证据强度低于原始记载）。"""
    flat = re.sub(r"\s+", "", str(text or ""))
    if SECONDARY_YEAR_RE.search(flat):
        return True
    return any(m in flat for m in SECONDARY_MARKERS)


def sentence_span(flat: str, pos: int) -> tuple[int, int]:
    left = pos
    while left > 0 and flat[left - 1] not in SENTENCE_DELIMS:
        left -= 1
    right = pos
    while right < len(flat) and flat[right] not in SENTENCE_DELIMS:
        right += 1
    if right < len(flat):
        right += 1
    return left, right


def page_of(flat: str, pos: int) -> str:
    last = None
    for m in PAGE_RE.finditer(flat[:pos]):
        last = m
    return last.group(1) if last else ""


class BookIndex:
    """一本全文的归一化文本 + 人名倒排索引。"""

    def __init__(self, book: str, family: str, path: Path) -> None:
        self.book = book
        self.family = family
        self.path = path
        self.flat, self.idx = self._normalize(path)
        self.occ: dict[str, list[int]] = {}

    @staticmethod
    def _normalize(path: Path) -> tuple[str, list[int]]:
        text = path.read_text(encoding="utf-8", errors="replace")
        buf: list[str] = []
        idx: list[int] = []
        for i, ch in enumerate(text):
            if not ch.isspace():
                buf.append(ch)
                idx.append(i)
        return "".join(buf), idx

    def index_names(self, names: set[str]) -> None:
        pattern = re.compile("|".join(re.escape(n) for n in sorted(names, key=len, reverse=True)))
        occ: dict[str, list[int]] = {}
        for m in pattern.finditer(self.flat):
            occ.setdefault(m.group(0), []).append(m.start())
        self.occ = occ

    def positions(self, name: str) -> list[int]:
        return self.occ.get(name, [])

    def locator(self, pos: int) -> str:
        page = page_of(self.flat, pos)
        return f"{self.book} 第{page}页" if page else f"{self.book} 页码未识别"


def find_best_candidate(
    bi: BookIndex,
    names_a: list[str],
    names_b: list[str],
    relation_type: str,
    diary_author_implicit: bool,
) -> dict | None:
    """在该书内为一条关系找最佳句级候选；找不到返回 None。

    排序键（依次）：非名单 → 非二手著录 → 有关系动词 → 句更短 → 双方距离更近。
    硬门：跨页标记、超长、引文未同时记载双方、名单句式且类型不接受名单。
    """
    flat = bi.flat
    seen: set[tuple[int, int]] = set()
    scored: list[tuple[tuple, dict]] = []
    budget = MAX_SPANS_PER_RELATION
    for a in names_a:
        for pa in bi.positions(a):
            for b in names_b:
                plist = bi.positions(b)
                if not plist:
                    continue
                i = bisect.bisect_left(plist, pa - WINDOW)
                while i < len(plist) and plist[i] <= pa + WINDOW:
                    pb = plist[i]
                    i += 1
                    lo, hi = min(pa, pb), max(pa + len(a), pb + len(b))
                    if (lo, hi) in seen:
                        continue
                    seen.add((lo, hi))
                    budget -= 1
                    if budget < 0:
                        break
                    left, right = sentence_span(flat, lo)
                    quote = flat[left:right]
                    if len(quote) > MAX_QUOTE or PAGE_MARKER in quote:
                        continue
                    attested, basis = quote_attests_pair(
                        quote, names_a, names_b, diary_author_implicit=diary_author_implicit
                    )
                    if not attested:
                        continue
                    listed = list_pattern(quote)
                    if listed and relation_type not in LIST_ACCEPTABLE_TYPES:
                        continue
                    sec = secondary_description(quote)
                    verb = next((v for v in RELATION_VERBS if v in quote), "")
                    score = (
                        1 if listed else 0,
                        1 if sec else 0,
                        0 if verb else 1,
                        len(quote),
                        abs(pa - pb),
                    )
                    scored.append((score, {
                        "left": left, "right": right, "quote": quote, "basis": basis,
                        "verb": verb, "list": listed, "secondary": sec,
                        "distance": abs(pa - pb),
                    }))
                if budget < 0:
                    break
            if budget < 0:
                break
        if budget < 0:
            break
    if not scored:
        return None
    scored.sort(key=lambda item: item[0])
    return scored[0][1]


def project_status(rel_row: dict[str, str], candidate: dict, evidence_rows: list[dict]) -> str:
    """只读投影：若该引文被定为 support，现有门禁会给出什么 publish_status。"""
    row = dict(rel_row)
    rows = [dict(e) for e in evidence_rows]
    rows.append({
        "relation_id": row.get("relation_id", ""),
        "evidence_support": "support",
        "review_status": "reviewed",
        "locator": candidate["locator"],
        "quote": candidate["quote"],
        "context": candidate["quote"],
    })
    cols = assign_relation_status_columns(row, rows, existing={
        "reviewer": row.get("reviewer", ""), "reviewed_at": row.get("reviewed_at", ""),
        "review_note": row.get("review_note", ""),
    })
    return cols["publish_status"]
BATCH_COLUMNS = [
    "relation_id", "person_a_name", "person_b_name", "current_final_relation_type",
    "book", "page", "locator", "quote", "relation_verb", "list_pattern",
    "secondary_description", "projected_publish_status",
    "verdict", "evidence_grade", "proposed_type", "adjudicator_reason",
    "adjudicator", "adjudicated_at",
]
VALID_VERDICTS = ("成立", "类型需改", "证据不足")
VALID_GRADES = ("support", "associated")


def _read(path: Path) -> list[dict[str, str]]:
    with open(path, encoding="utf-8-sig", newline="") as fh:
        return list(csv.DictReader(fh))


def sweep(batch_size: int, batch_index: int) -> dict:
    persons = {r["person_id"].strip(): r for r in _read(DATA / "persons.csv")}
    relations = _read(DATA / "person_relations.csv")
    evidence_by_rel: dict[str, list[dict[str, str]]] = {}
    for r in _read(DATA / "relation_evidences.csv"):
        evidence_by_rel.setdefault(r["relation_id"].strip(), []).append(r)
    src_by_cit: dict[tuple[str, str], str] = {}
    for r in _read(DATA / "sources.csv"):
        src_by_cit.setdefault(
            ((r.get("source_family") or "").strip(), (r.get("citation") or "").strip()),
            (r.get("source_id") or "").strip(),
        )

    names: dict[str, list[str]] = {pid: name_candidates(p) for pid, p in persons.items()}
    all_names = {n for v in names.values() for n in v}
    books: list[BookIndex] = []
    for book, (family, path) in SOURCE_BOOKS.items():
        if not path.exists():
            raise SystemExit(f"本地全文缺失，无法扫掠：{path}")
        bi = BookIndex(book, family, path)
        bi.index_names(all_names)
        books.append(bi)

    pool: list[dict[str, str]] = []
    stats: Counter[str] = Counter()
    for rel in relations:
        rid = rel["relation_id"].strip()
        na = names.get(rel["source_person_id"].strip(), [])
        nb = names.get(rel["target_person_id"].strip(), [])
        if not na or not nb:
            stats["skip_no_names"] += 1
            continue
        rtype = (rel.get("final_relation_type") or "").strip()
        best: tuple[tuple, dict, BookIndex] | None = None
        for bi in books:
            diary = bi.book == "鲁迅日记" and DIARY_AUTHOR in (*na, *nb)
            cand = find_best_candidate(bi, na, nb, rtype, diary)
            if cand is None:
                continue
            key = (1 if cand["list"] else 0, 1 if cand["secondary"] else 0,
                   0 if cand["verb"] else 1, len(cand["quote"]), cand["distance"])
            if best is None or key < best[0]:
                best = (key, cand, bi)
        if best is None:
            stats["no_qualified_sentence"] += 1
            continue
        _, cand, bi = best
        stats["has_candidate"] += 1
        if cand["secondary"]:
            stats["flag_secondary"] += 1
        if cand["list"]:
            stats["flag_list"] += 1
        status = project_status(rel, {**cand, "locator": bi.locator(cand["left"])},
                                evidence_by_rel.get(rid, []))
        locator = bi.locator(cand["left"])
        public = status in ("supported", "verified")
        stats["projected_public" if public else "projected_not_public"] += 1
        pool.append({
            "relation_id": rid,
            "person_a_id": rel["source_person_id"].strip(),
            "person_a_name": (persons.get(rel["source_person_id"].strip(), {}).get("standard_name") or "").strip(),
            "person_b_id": rel["target_person_id"].strip(),
            "person_b_name": (persons.get(rel["target_person_id"].strip(), {}).get("standard_name") or "").strip(),
            "current_final_relation_type": rtype,
            "relation_risk_level": (rel.get("relation_risk_level") or "").strip(),
            "confidence": (rel.get("confidence") or "").strip(),
            "needs_manual_review": (rel.get("needs_manual_review") or "").strip(),
            "current_publish_status": (rel.get("publish_status") or "").strip(),
            "book": bi.book, "source_family": bi.family,
            "page": page_of(bi.flat, cand["left"]), "locator": locator,
            "quote": cand["quote"], "quote_char_len": str(len(cand["quote"])),
            "quote_sha256": hashlib.sha256(cand["quote"].encode("utf-8")).hexdigest(),
            "normalized_start": str(cand["left"]), "normalized_end": str(cand["right"]),
            "source_file": str(bi.path),
            "attestation_basis": cand["basis"], "relation_verb": cand["verb"],
            "list_pattern": "yes" if cand["list"] else "no",
            "secondary_description": "yes" if cand["secondary"] else "no",
            "projected_publish_status": status,
            "projected_public": "yes" if public else "no",
            "review_status": "pending_cross_validation",
            "resolved_source_id": src_by_cit.get((bi.family, locator), ""),
        })

    pool.sort(key=lambda r: r["relation_id"])
    eligible = [r for r in pool
                if r["projected_public"] == "yes" and r["secondary_description"] == "no"]
    start = (batch_index - 1) * batch_size
    batch = eligible[start:start + batch_size]
    return {"pool": pool, "batch": batch, "eligible": eligible, "stats": stats,
            "batch_size": batch_size, "batch_index": batch_index}


def write_pool(pool: list[dict[str, str]], path: Path) -> None:
    cols = [c for c in POOL_COLUMNS if c != "review_status"] + ["review_status"]
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=cols, extrasaction="ignore")
        w.writeheader()
        w.writerows(pool)


def write_batch(batch: list[dict[str, str]], path: Path) -> None:
    """裁决输入：**不含**夜间轮 reason，也不含另一名裁决者的任何结论。"""
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=BATCH_COLUMNS, extrasaction="ignore")
        w.writeheader()
        for row in batch:
            item = {c: row.get(c, "") for c in BATCH_COLUMNS}
            for c in ("verdict", "evidence_grade", "proposed_type",
                      "adjudicator_reason", "adjudicator", "adjudicated_at"):
                item[c] = ""
            w.writerow(item)


def write_report(result: dict, path: Path, batch_path: Path) -> None:
    stats = result["stats"]
    pool = result["pool"]
    lines = [
        "# 全量关系证据扫掠报告（只读，生产层零改动）",
        "",
        "> 生成脚本：`research/analysis/build_relation_evidence_sweep.py`  ",
        f"> 候选池：`{OUT_POOL.name}`（{len(pool)} 行）  ",
        f"> 本批裁决输入：`{batch_path.name}`（{len(result['batch'])} 行）",
        "",
        "## 1. 扫掠口径",
        "",
        "- 输入关系：4238 条（全量，不抽样）。",
        "- 本地全文：左联史 / 左联词典 / 鲁迅日记（空白归一后检索；OCR 字间带空格）。",
        f"- 句级候选硬门：句读边界切句、不跨页标记（`────`）、长度 ≤{MAX_QUOTE} 字、"
        f"**引文本身须同时记载双方**（鲁迅日记按日记体免甲方自名）、"
        f"名单句式仅在 `签名联署`/`同属组织` 类型下保留。",
        "- 排序：非名单 → 非二手著录 → 有关系动词 → 句更短 → 双方距离更近。",
        "",
        "## 2. 结果",
        "",
        "| 项 | 数量 |",
        "| --- | ---: |",
        f"| 有句级候选 | {stats['has_candidate']} |",
        f"| 无合格句级候选 | {stats['no_qualified_sentence']} |",
        f"| 缺人名（无法检索） | {stats['skip_no_names']} |",
        f"| 标记为名单句式 | {stats['flag_list']} |",
        f"| 标记为二手著录 | {stats['flag_secondary']} |",
        f"| 投影可进公开层 | {stats['projected_public']} |",
        f"| 投影仍被保守门禁挡住 | {stats['projected_not_public']} |",
        f"| 本批可裁决（可公开 ∧ 非二手） | {len(result['eligible'])} |",
        "",
        "投影不等于结论：`projected_publish_status` 是「假设该引文被定为 support」时现有门禁的只读推演。",
        "",
        "## 3. 裁决方式：双 Agent 独立交叉验证",
        "",
        "2026-10-08 用户授权以交叉验证替代逐条人工裁决（原话：「不用我裁，你和zcode交叉验证裁决吧」，"
        "后因不用 GLM-5.3 系列，第二裁决者改为 antigravity harness 的 Claude Opus 4.6 (Thinking)）。",
        "",
        "- 两名裁决者**各自独立**填写同一份 `sweep_batch*_input.csv` 的副本，互不可见对方结论；",
        "- 裁决输入**不含**夜间轮 `reason`（该字段由 GLM-5.3-Flash 生成、已证明与引文经常不一致，带上会锚定）；",
        "- **只有两方一致**的条目才可落地；分歧条目进 `sweep_disagreements.csv`，不落地、不改状态；",
        "- 裁决口径为「双 Agent 交叉验证」，**不是人工复核**：公开层只走 derived `supported`，"
        "禁止使用 `human_adjudication` / `verified` 通道，台账与站点必须标明此口径。",
        "",
        "## 4. 裁决词表",
        "",
        "| 列 | 取值 | 含义 |",
        "| --- | --- | --- |",
        "| `verdict` | 成立 | 引文直接记载了这两人之间的该种关系 |",
        "| | 类型需改 | 关系成立但类型不当，`proposed_type` 填建议类型 |",
        "| | 证据不足 | 引文只是同页并列、共现、职务关联或二手转述，不足以支持该关系 |",
        "| `evidence_grade` | support | 直接记载，可作定级依据 |",
        "| | associated | 仅来源关联／共现，不得进入公开层 |",
        "",
        "判断要点：同列一份**具体文件**的签署名单对 `签名联署` 是直接证据，对 `交游` 只是共现；"
        "同一句里分别提到两人但各自与第三人发生关系，不算佐证；"
        "书目著录（「某传记小说叙述了……的交往」）属二手，不足以定为 support。",
        "",
    ]
    path.write_text("\n".join(lines), encoding="utf-8")


def main() -> int:
    ap = argparse.ArgumentParser(description="全量关系证据扫掠（只读，生产层零改动）。")
    ap.add_argument("--batch-size", type=int, default=50)
    ap.add_argument("--batch-index", type=int, default=1)
    ap.add_argument("--pool-out", type=Path, default=OUT_POOL)
    ap.add_argument("--report-out", type=Path, default=OUT_REPORT)
    args = ap.parse_args()

    result = sweep(args.batch_size, args.batch_index)
    batch_path = REPORTS / f"sweep_batch{args.batch_index}_input.csv"
    write_pool(result["pool"], args.pool_out)
    write_batch(result["batch"], batch_path)
    write_report(result, args.report_out, batch_path)

    print(f"候选池 {len(result['pool'])} 条 -> {args.pool_out.relative_to(ROOT)}")
    print(f"本批可裁决 {len(result['eligible'])} 条，取第 {args.batch_index} 批 {len(result['batch'])} 条")
    print(f"裁决输入 -> {batch_path.relative_to(ROOT)}（不含夜间轮 reason，防锚定）")
    for k, v in result["stats"].most_common():
        print(f"  {k}: {v}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
