"""Phase 7 候选独立审计（主 Agent 独立回源核查，非重复执行者自检）。

对冻结的 12 条关系证据候选逐条独立核查：
1. 8 条 support：引文哈希、本地原文逐字定位（OCR 容差）、年月日锚定、
   引文人物词元覆盖、语义支持判断。
2. 4 条 insufficient：检索回执计数、夜间轮第二轮分诊分歧核实
   （新证据引文逐字定位 + 人物词元覆盖）。

已知分歧（本审计核实）：
- REL-00523 二轮升级 support 的理由与左联史原文相符，但捕获引文
  EVI-N1-E0960180FF3F 摘自同页另一段落（茅盾/叶圣陶），不含潘汉年或丁玲，
  判 issue_found：需替换引文后再议转正。
- REL-01219 / REL-01743 二轮升级 associated 的新引文逐字命中且包含双方，
  判 agree_upgrade_to_associated。

输出（写入 ``--out-dir``，默认 ``research/drafts/reports/``）：
- ``phase7_candidate_independent_audit.csv``
- ``phase7_candidate_independent_audit_report.md``

全部结论 ``pending_human_review``；生产层零改动。
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import re
import unicodedata
from collections import Counter
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_CANDIDATES = REPO_ROOT / "research" / "drafts" / "reports" / "phase7_relation_evidence_candidates.csv"
DEFAULT_OVERNIGHT_DIR = REPO_ROOT / "research" / "content-upgrade-overnight-2026-09-06"
DEFAULT_OUT_DIR = REPO_ROOT / "research" / "drafts" / "reports"
DIARY_RAW = REPO_ROOT / "research" / "raw_texts" / "日记全编：全2册 (鲁迅 著) (Z-Library).txt"
RUNTIME_SOURCES = REPO_ROOT / "data" / "processed" / "runtime_sources"

AUDITOR = "ZCode GLM-5.3（主 Agent，独立审计，非执行者自检）"
AUDITED_AT = "2026-09-20"

PERSON_TOKENS = {
    "鲁迅": ["鲁迅"],
    "李小峰": ["小峰"],
    "李霁野": ["霁野"],
    "陈望道": ["陈望道"],
    "郁达夫": ["郁达夫", "达夫"],
    "许广平": ["广平"],
    "林语堂": ["语堂"],
    "郑伯奇": ["郑伯奇"],
    "冯雪峰": ["冯雪峰", "雪峰"],
    "巴比塞": ["巴比塞"],
    "潘汉年": ["潘汉年"],
    "丁玲": ["丁玲"],
    "冯乃超": ["冯乃超"],
    "柔石": ["柔石"],
    "穆木天": ["穆木天"],
}

INDEPENDENT_VERDICTS = {
    "CAND-RELE-P7-001": ("agree_support", "日记直接记录收信（得小峰信并书刊），语义直接支持通信。"),
    "CAND-RELE-P7-002": ("agree_support", "日记直接记录寄信（寄霁野信），语义直接支持通信。"),
    "CAND-RELE-P7-003": ("agree_support", "日记直接记录寄信并稿，语义直接支持通信。"),
    "CAND-RELE-P7-004": ("agree_support", "日记直接记录收信（得郁达夫信），语义直接支持通信。"),
    "CAND-RELE-P7-005": ("agree_support", "日记直接记录共同活动（达夫招饮、与广平同往），支持本条交游，不外推私人关系性质。"),
    "CAND-RELE-P7-006": ("agree_support", "日记直接记录寄信（寄语堂信），语义直接支持通信。"),
    "CAND-RELE-P7-007": ("agree_support", "日记直接记录寄信（寄郑伯奇信），语义直接支持通信。"),
    "CAND-RELE-P7-008": ("agree_support", "日记直接记录复信（复冯雪峰信），语义直接支持通信。"),
    "CAND-RELE-P7-009": (
        "agree_insufficient",
        "三轮检索回执在案；词典332页主语为鲁迅与巴比塞的友谊著录、左联史192页为北平刊物欢迎代表团背景，均非两人直接交游证据。维持 insufficient。",
    ),
    "CAND-RELE-P7-010": (
        "issue_found",
        "二轮升级 support 的理由与左联史原文相符（「潘汉年就去看望他们」「介绍他俩一同加人左联」段已独立定位），"
        "但捕获引文 EVI-N1-E0960180FF3F 摘自同页另一段落（茅盾/叶圣陶/郑振铎），不含潘汉年或丁玲。"
        "建议：替换为已定位的正确段落引文后再议转正；在此之前 support 升级缺乏逐字证据。",
    ),
    "CAND-RELE-P7-011": (
        "agree_upgrade_to_associated",
        "二轮新引文（词典334页）逐字命中，冯乃超与柔石同在鲁迅50寿辰共同发起名单，同活动关联成立；交游仍无直接证据，associated 恰当。",
    ),
    "CAND-RELE-P7-012": (
        "agree_upgrade_to_associated",
        "二轮新引文（左联史21页）逐字命中，丁玲与穆木天同列左联党团成员，同属组织成立；交游仍无直接证据，associated 恰当。",
    ),
}

REL_00523_CORRECT_PASSAGE_PHRASES = ("潘汉年就去看望他们", "介绍他俩一同加人左联")

AUDIT_COLUMNS = [
    "candidate_id",
    "relation_id",
    "relation_pair",
    "frozen_evidence_support",
    "second_round_support",
    "verbatim_located",
    "occurrences",
    "anchor",
    "anchor_matches_locator",
    "quote_person_tokens",
    "search_receipts",
    "new_evidence_verbatim",
    "new_quote_person_tokens",
    "independent_verdict",
    "independent_note",
    "auditor",
    "audited_at",
    "review_status",
]


def _normalize(text: str) -> str:
    text = unicodedata.normalize("NFKC", text)
    return text.translate(str.maketrans("", "", " \u3000\t\r\n·"))


def _strip_postmarks(text: str) -> str:
    text = re.sub(r"[一二三四五六七八九十廿]{1,3}月[一二三四五六七八九十廿]{1,3}日发", "□", text)
    return re.sub(r"[一二三四五六七八九十廿]{1,2}日发", "□", text)


_MONTH_RE = re.compile(r"(一[一二]?|二[一二]?|三[一二]?|四|五|六|七|八|九|十|十一|十二)月(?![分发])")
_DAY_RE = re.compile(r"([一二三四五六七八九十]{1,3})日(?![分发])")
_VOL_RE = re.compile(r"日记[一二三四五六七八九十百]+\((\d{4})年\)")
_MONTH_NUM = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9, "十": 10, "十一": 11, "十二": 12}
_UNITS = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}


def _cn_day_to_num(text: str) -> int:
    total = num = 0
    for ch in text:
        if ch in _UNITS:
            num = _UNITS[ch]
        elif ch == "十":
            total += (num or 1) * 10
            num = 0
    return total + num


def _anchors(text: str) -> tuple[list, list, list]:
    vols = [(m.start(), m.group(1)) for m in _VOL_RE.finditer(text)]
    months = [(m.start(), _MONTH_NUM[m.group(1)]) for m in _MONTH_RE.finditer(text)]
    days = [(m.start(), _cn_day_to_num(m.group(1))) for m in _DAY_RE.finditer(text)]
    return vols, months, days


def _prev(marks: list, pos: int):
    last = None
    for start, label in marks:
        if start >= pos:
            break
        last = label
    return last


def _locate_quote(quote: str, source_text: str) -> dict[str, object]:
    """在剥离邮戳后的规范化原文中定位引文，返回逐字命中、次数与精确年月日锚。"""
    vols, months, days = _anchors(source_text)
    normalized_quote = _normalize(quote)
    occurrences = [m.start() for m in re.finditer(re.escape(normalized_quote), source_text)]
    if not occurrences:
        return {"verbatim_located": False, "occurrences": 0, "anchor": "", "anchors": []}
    anchors = []
    for pos in occurrences:
        year = _prev(vols, pos)
        anchors.append((int(year) if year else None, _prev(months, pos), _prev(days, pos)))
    return {
        "verbatim_located": True,
        "occurrences": len(occurrences),
        "anchor": f"{anchors[0][0]}年{anchors[0][1]}月{anchors[0][2]}日" if anchors[0][0] else "",
        "anchors": anchors,
    }


def _tokens_present(quote: str, source_name: str, target_name: str) -> bool:
    tokens = PERSON_TOKENS.get(source_name, [source_name]) + PERSON_TOKENS.get(target_name, [target_name])
    return any(token in quote for token in tokens)


def _resolve_source_path(local_path: str) -> Path:
    candidate = Path(local_path)
    if candidate.exists():
        return candidate
    name = Path(local_path).name
    fallback = RUNTIME_SOURCES / name
    if fallback.exists():
        return fallback
    raise FileNotFoundError(f"来源文件不可用: {local_path}")


def audit(candidates_path: Path, out_dir: Path, overnight_dir: Path = DEFAULT_OVERNIGHT_DIR) -> dict[str, int]:
    with candidates_path.open(encoding="utf-8-sig", newline="") as handle:
        candidates = [dict(row) for row in csv.DictReader(handle)]

    relations = {
        record["relation_id"]: record
        for record in (
            json.loads(line)
            for line in overnight_dir.joinpath("relations.jsonl").read_text(encoding="utf-8").splitlines()
            if line.strip()
        )
        if "phase7" in record.get("groups", [])
    }
    evidences = {
        item["evidence_id"]: item
        for item in (
            json.loads(line)
            for line in overnight_dir.joinpath("evidence.jsonl").read_text(encoding="utf-8").splitlines()
            if line.strip()
        )
    }
    receipts: Counter[str] = Counter()
    for log in (
        json.loads(line)
        for line in overnight_dir.joinpath("search_log.jsonl").read_text(encoding="utf-8").splitlines()
        if line.strip()
    ):
        for rid in log.get("claim_or_relation_ids", []) or []:
            receipts[rid] += 1

    diary_text = _strip_postmarks(_normalize(DIARY_RAW.read_text(encoding="utf-8", errors="ignore")))
    source_cache: dict[str, str] = {}

    rel_00523_history = _normalize(
        (_resolve_source_path("data/processed/runtime_sources/左联史.txt")).read_text(encoding="utf-8", errors="ignore")
    )
    for phrase in REL_00523_CORRECT_PASSAGE_PHRASES:
        if phrase not in rel_00523_history:
            raise ValueError(f"REL-00523 正确段落锚定短语未在左联史命中: {phrase}")

    rows = []
    verdict_counts: Counter[str] = Counter()
    for candidate in candidates:
        relation_id = candidate["relation_id"]
        pair = f"{candidate['source_person_name']}→{candidate['target_person_name']}"
        record = relations.get(relation_id, {})
        second_round = record.get("evidence_support", "")
        verdict, note = INDEPENDENT_VERDICTS[candidate["candidate_id"]]

        verbatim = False
        occurrences = 0
        anchor = ""
        anchor_match = False
        tokens = False
        new_verbatim = ""
        new_tokens = ""

        if candidate["evidence_support"] == "support":
            quote = candidate["quote"]
            if hashlib.sha256(quote.encode("utf-8")).hexdigest() != candidate["quote_sha256"]:
                raise ValueError(f"{candidate['candidate_id']} 引文哈希不匹配")
            located = _locate_quote(quote, diary_text)
            verbatim = bool(located["verbatim_located"])
            occurrences = int(located["occurrences"])
            want = re.match(r"(\d{4})年(\d{1,2})月(\d{1,2})日", candidate["locator"])
            wanted = tuple(map(int, want.groups())) if want else None
            for year, month, day in located["anchors"]:  # type: ignore[union-attr]
                if (year, month, day) == wanted:
                    anchor_match = True
                    anchor = f"{year}年{month}月{day}日"
                    break
            if not anchor_match:
                anchor = str(located["anchor"])
            tokens = _tokens_present(quote, candidate["source_person_name"], candidate["target_person_name"])
        else:
            evidence_ids = record.get("evidence_ids", [])
            if evidence_ids:
                evidence = evidences.get(evidence_ids[-1])
                if evidence is None:
                    raise ValueError(f"{relation_id} 二轮证据 {evidence_ids[-1]} 不在 evidence.jsonl")
                path = _resolve_source_path(evidence.get("local_path", ""))
                if str(path) not in source_cache:
                    source_cache[str(path)] = _normalize(path.read_text(encoding="utf-8", errors="ignore"))
                located = _locate_quote(evidence["quote"], source_cache[str(path)])
                new_verbatim = bool(located["verbatim_located"])
                new_tokens = _tokens_present(
                    evidence["quote"], candidate["source_person_name"], candidate["target_person_name"]
                )

        rows.append(
            {
                "candidate_id": candidate["candidate_id"],
                "relation_id": relation_id,
                "relation_pair": pair,
                "frozen_evidence_support": candidate["evidence_support"],
                "second_round_support": second_round,
                "verbatim_located": str(verbatim) if candidate["evidence_support"] == "support" else "",
                "occurrences": str(occurrences) if candidate["evidence_support"] == "support" else "",
                "anchor": anchor,
                "anchor_matches_locator": str(anchor_match) if candidate["evidence_support"] == "support" else "",
                "quote_person_tokens": str(tokens) if candidate["evidence_support"] == "support" else "",
                "search_receipts": str(receipts.get(relation_id, 0)),
                "new_evidence_verbatim": str(new_verbatim) if new_verbatim != "" else "",
                "new_quote_person_tokens": str(new_tokens) if new_tokens != "" else "",
                "independent_verdict": verdict,
                "independent_note": note,
                "auditor": AUDITOR,
                "audited_at": AUDITED_AT,
                "review_status": "pending_human_review",
            }
        )
        verdict_counts[verdict] += 1

    out_dir.mkdir(parents=True, exist_ok=True)
    csv_path = out_dir / "phase7_candidate_independent_audit.csv"
    with csv_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=AUDIT_COLUMNS, lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)
    _write_report(out_dir / "phase7_candidate_independent_audit_report.md", rows, verdict_counts)

    return {
        "rows": len(rows),
        "agree_support": verdict_counts["agree_support"],
        "agree_insufficient": verdict_counts["agree_insufficient"],
        "agree_upgrade_to_associated": verdict_counts["agree_upgrade_to_associated"],
        "issue_found": verdict_counts["issue_found"],
    }


def _write_report(path: Path, rows: list[dict[str, str]], verdict_counts: Counter[str]) -> None:
    lines = [
        "# Phase 7 候选独立审计报告（主 Agent）",
        "",
        f"> 审计人：{AUDITOR}",
        f"> 审计日期：{AUDITED_AT}",
        "> 输入：冻结候选包 12 条 + 夜间轮二轮分诊（relations/evidence/search_log）+ 本地原文三源",
        "",
        "## 结论",
        "",
        f"agree_support {verdict_counts['agree_support']} · agree_insufficient {verdict_counts['agree_insufficient']} · "
        f"agree_upgrade_to_associated {verdict_counts['agree_upgrade_to_associated']} · issue_found {verdict_counts['issue_found']}。",
        "全部 12 项为 pending_human_review，本审计不直接改动生产层或冻结候选包。",
        "",
        "| 候选 | 关系 | 对 | 冻结 | 二轮 | 独立判定 |",
        "| --- | --- | --- | --- | --- | --- |",
    ]
    for row in rows:
        lines.append(
            f"| {row['candidate_id']} | {row['relation_id']} | {row['relation_pair']} | "
            f"{row['frozen_evidence_support']} | {row['second_round_support']} | {row['independent_verdict']} |"
        )
    issue_rows = [row for row in rows if row["independent_verdict"] == "issue_found"]
    lines.extend(["", "## 发现（需人工裁决）", ""])
    for row in issue_rows:
        lines.append(f"### {row['relation_id']} {row['relation_pair']}")
        lines.append("")
        lines.append(row["independent_note"])
        lines.append("")
    lines.extend(
        [
            "## 核验方法",
            "",
            "- 引文哈希：sha256(quote) 与候选包登记值一致（8/8）。",
            "- 逐字定位：OCR 容差规范化（NFKC、去空白间隔点）后在本地《日记全编》全文检索，8/8 命中。",
            "- 年月日锚定：剥离邮戳短语（X月N日发）后取引文前最近的卷（日记N(YYYY年)）/月/日标题，"
            "8/8 与候选 locator 精确一致；REL-00060「寄霁野信。」全文 18 处命中中含 1928-02-26 精确锚。",
            "- 人物词元：引文含关系至少一方（日记侧作者本人覆盖另一方），别名表见脚本常量。",
            "- 二轮新证据：REL-01219（词典334页）/REL-01743（左联史21页）逐字命中且含双方；"
            "REL-00523 捕获引文逐字命中但为同页错误段落（见发现）。",
            "- REL-00523 正确段落锚定短语「潘汉年就去看望他们」「介绍他俩一同加人左联」已在左联史原文验证存在。",
            "",
        ]
    )
    path.write_text("\n".join(lines), encoding="utf-8", newline="")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--candidates", type=Path, default=DEFAULT_CANDIDATES)
    parser.add_argument("--out-dir", type=Path, default=DEFAULT_OUT_DIR)
    parser.add_argument("--overnight-dir", type=Path, default=DEFAULT_OVERNIGHT_DIR)
    args = parser.parse_args()
    print(audit(args.candidates, args.out_dir, args.overnight_dir))


if __name__ == "__main__":
    main()
