"""按用户授权裁决 Phase 7 关系证据候选（研究层裁决包，不动生产）。

输入：冻结候选包 ``phase7_relation_evidence_candidates.csv``、夜间轮
``evidence.jsonl``（二轮新证据引文）、本地《左联史》《左联词典》文本。

输出（写入 ``--out-dir``，默认 ``research/drafts/reports/``）：
- ``phase7_relation_evidence_candidates_adjudicated.csv``：12 条全列保留＋裁决列。
- ``phase7_adjudication_record.md``：授权登记与逐条裁决说明。

裁决内容（依据 2026-09-20 主 Agent 独立审计 + 用户授权）：
- 8 条日记 support：agree_support，保持。
- REL-00523（潘汉年—丁玲）：**换引文后升 support**——原二轮引文
  EVI-N1-E0960180FF3F 摘自同页茅盾段落（审计判 issue_found）；替换为左联史
  第30页「潘汉年就去看望他们……介绍他俩一同加人左联」段（运行时重新切取、
  逐字定位与哈希自洽校验，OCR「加人」保留原样）。
- REL-01219（冯乃超—柔石）/ REL-01743（丁玲—穆木天）：升 associated，
  引文取夜间轮 EVI-N1-FBD6D22370A6 / EVI-N1-5B8C3B7E6622（逐字与哈希校验）。
- REL-00113（鲁迅—巴比塞）：agree_insufficient，维持。

边界：``review_status`` 置 ``adjudicated_authorized``；**不改生产层**——
relation_evidences 落地与 publish_status 转换属下一批次，须另行幂等脚本与门禁。
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import re
import unicodedata
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_CANDIDATES = REPO_ROOT / "research" / "drafts" / "reports" / "phase7_relation_evidence_candidates.csv"
DEFAULT_OVERNIGHT_DIR = REPO_ROOT / "research" / "content-upgrade-overnight-2026-09-06"
DEFAULT_OUT_DIR = REPO_ROOT / "research" / "drafts" / "reports"
HISTORY_PATH = REPO_ROOT / "data" / "processed" / "runtime_sources" / "左联史.txt"
DICTIONARY_PATH = REPO_ROOT / "data" / "processed" / "runtime_sources" / "左联词典.txt"

AUTHORIZED_BY = "用户（会话授权）"
AUTHORIZED_AT = "2026-09-20"

# REL-00523 替换引文的两端锚（规范化后检索，OCR 原样保留）
QUOTE_00523_START = "党的文艺工作领导人潘汉年就去看望他们"
QUOTE_00523_END = "介绍他俩一同加人左联"

UPGRADES = {
    "REL-01219": ("associated", "EVI-N1-FBD6D22370A6", "agree_upgrade_to_associated"),
    "REL-01743": ("associated", "EVI-N1-5B8C3B7E6622", "agree_upgrade_to_associated"),
}
KEEP_INSUFFICIENT = {"REL-00113"}

NEW_SOURCE_IDS = {
    "左联史": ("CAND-SRC-P7-003", "左联史（runtime 本地文本）", str(HISTORY_PATH)),
    "左联词典": ("CAND-SRC-P7-004", "左联词典（runtime 本地文本）", str(DICTIONARY_PATH)),
}

ADJUDICATION_COLUMNS = [
    "adjudicated_verdict",
    "adjudication_note",
    "authorized_by",
    "authorized_at",
    "authorization_quote",
]


def _normalize(text: str) -> str:
    text = unicodedata.normalize("NFKC", text)
    return text.translate(str.maketrans("", "", " \u3000\t\r\n·"))


def _extract_quote_00523(history_raw: str) -> str:
    """按锚短语从原文连续切取 REL-00523 替换引文（紧凑存储，逐字对应原文）。"""
    positions: list[int] = []
    chars: list[str] = []
    for i, ch in enumerate(history_raw):
        if ch in " \u3000\t\r\n·":
            continue
        for c in unicodedata.normalize("NFKC", ch):
            positions.append(i)
            chars.append(c)
    norm = "".join(chars)
    start_n = norm.find(QUOTE_00523_START)
    if start_n < 0:
        raise ValueError("REL-00523 替换引文起点锚未在左联史命中")
    end_n = norm.find("。", norm.find(QUOTE_00523_END, start_n)) + 1
    if end_n <= start_n:
        raise ValueError("REL-00523 替换引文终点锚未在左联史命中")
    raw_slice = history_raw[positions[start_n] : positions[end_n - 1] + 1]
    return re.sub(r"[ \u3000\t\r\n·]", "", raw_slice)


def adjudicate(candidates_path: Path, out_dir: Path, authorization_quote: str, overnight_dir: Path = DEFAULT_OVERNIGHT_DIR) -> dict[str, int]:
    quote = authorization_quote.strip()
    if not quote:
        raise ValueError("授权语为空：候选裁决必须有用户逐字授权语")

    with candidates_path.open(encoding="utf-8-sig", newline="") as handle:
        rows = [dict(row) for row in csv.DictReader(handle)]

    evidences = {
        item["evidence_id"]: item
        for item in (
            json.loads(line)
            for line in overnight_dir.joinpath("evidence.jsonl").read_text(encoding="utf-8").splitlines()
            if line.strip()
        )
    }

    history_raw = HISTORY_PATH.read_text(encoding="utf-8", errors="ignore")
    dictionary_raw = DICTIONARY_PATH.read_text(encoding="utf-8", errors="ignore")
    history_norm = _normalize(history_raw)
    dictionary_norm = _normalize(dictionary_raw)

    quote_00523 = _extract_quote_00523(history_raw)
    if _normalize(quote_00523) not in history_norm:
        raise ValueError("REL-00523 替换引文未能逐字回定位左联史")

    counts = {"support": 0, "associated": 0, "insufficient": 0}
    quote_corrected = 0
    for row in rows:
        relation_id = row["relation_id"]
        note = "维持冻结 support：日记直接记载，独立审计逐字与年月日锚定通过。"
        verdict = "agree_support"
        if relation_id == "REL-00523":
            source_id, source_title, source_path = NEW_SOURCE_IDS["左联史"]
            row["evidence_support"] = "support"
            row["quote"] = quote_00523
            row["quote_sha256"] = hashlib.sha256(quote_00523.encode("utf-8")).hexdigest()
            row["locator"] = "左联史 第30页（潘汉年看望丁玲、胡也频段）"
            row["candidate_source_id"] = source_id
            row["source_title"] = source_title
            row["source_path_or_url"] = source_path
            row["source_family"] = "左联史"
            row["source_level"] = "B（公开转录或OCR史著；本地逐字核对）"
            row["access_date"] = AUTHORIZED_AT
            verdict = "quote_corrected_to_support"
            note = (
                "二轮捕获引文 EVI-N1-E0960180FF3F 摘自同页茅盾段落（独立审计 issue_found）；"
                "按授权替换为同页正确段落（含「潘汉年就去看望他们」「介绍他俩一同加人左联」，"
                "运行时重新切取并逐字回定位校验，OCR「加人」保留原样），升 support。"
            )
            quote_corrected = 1
        elif relation_id in UPGRADES:
            support, evidence_id, verdict = UPGRADES[relation_id]
            evidence = evidences.get(evidence_id)
            if evidence is None:
                raise ValueError(f"{relation_id} 二轮证据 {evidence_id} 不在 evidence.jsonl")
            if hashlib.sha256(evidence["quote"].encode("utf-8")).hexdigest() != evidence["quote_sha256"]:
                raise ValueError(f"{relation_id} 二轮证据 {evidence_id} 引文哈希不匹配")
            family = "左联词典" if "词典" in evidence["source_title"] else "左联史"
            target_norm = dictionary_norm if family == "左联词典" else history_norm
            if _normalize(evidence["quote"]) not in target_norm:
                raise ValueError(f"{relation_id} 二轮证据引文未能逐字回定位 {family}")
            source_id, source_title, source_path = NEW_SOURCE_IDS[family]
            row["evidence_support"] = support
            row["quote"] = evidence["quote"]
            row["quote_sha256"] = evidence["quote_sha256"]
            row["locator"] = evidence["locator"]
            row["candidate_source_id"] = source_id
            row["source_title"] = source_title
            row["source_path_or_url"] = source_path
            row["source_family"] = family
            row["source_level"] = evidence["source_level"]
            row["access_date"] = AUTHORIZED_AT
            note = f"按授权升 associated：二轮证据 {evidence_id}（{evidence['locator']}）逐字与哈希校验通过，双方同列记载；交游无直接证据。"
        elif relation_id in KEEP_INSUFFICIENT:
            verdict = "agree_insufficient"
            note = "维持 insufficient：三轮检索回执在案，本地两源均无直接交游证据；不得以背景共现确认关系。"
        row["adjudicated_verdict"] = verdict
        row["adjudication_note"] = note
        row["authorized_by"] = AUTHORIZED_BY
        row["authorized_at"] = AUTHORIZED_AT
        row["authorization_quote"] = quote
        row["review_status"] = "adjudicated_authorized"
        counts[row["evidence_support"]] = counts.get(row["evidence_support"], 0) + 1

    out_dir.mkdir(parents=True, exist_ok=True)
    fieldnames = list(rows[0].keys())
    for column in ADJUDICATION_COLUMNS:
        if column not in fieldnames:
            fieldnames.append(column)
    out_path = out_dir / "phase7_relation_evidence_candidates_adjudicated.csv"
    with out_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames, lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)

    _write_record(out_dir / "phase7_adjudication_record.md", quote, counts, quote_corrected, rows)
    return {"rows": len(rows), **counts, "quote_corrected": quote_corrected}


def _write_record(path: Path, quote: str, counts: dict[str, int], quote_corrected: int, rows: list[dict[str, str]]) -> None:
    lines = [
        "# Phase 7 候选裁决登记（研究层）",
        "",
        f"- 授权人：{AUTHORIZED_BY}",
        f"- 授权时间：{AUTHORIZED_AT}",
        f"- 授权语（逐字）：「{quote}」",
        "- 依据：2026-09-20 主 Agent 独立审计 `phase7_candidate_independent_audit.csv` + 用户对三项待裁决清单的整体授权。",
        "",
        "## 裁决结果",
        "",
        f"- support {counts['support']}（含 REL-00523 引文替换后升级 1 条）",
        f"- associated {counts['associated']}（REL-01219 / REL-01743 升级）",
        f"- insufficient {counts['insufficient']}（REL-00113 维持）",
        f"- 引文替换 {quote_corrected} 条",
        "",
        "## 逐条裁决",
        "",
        "| 候选 | 关系 | 裁决 | 说明 |",
        "| --- | --- | --- | --- |",
    ]
    for row in rows:
        pair = f"{row.get('source_person_name', '')}→{row.get('target_person_name', '')}"
        lines.append(f"| {row['candidate_id']} | {row['relation_id']} {pair} | {row['adjudicated_verdict']} | {row['adjudication_note'][:60]}… |")
    lines.extend(
        [
            "",
            "## 边界",
            "",
            "本裁决为研究层产物（`review_status=adjudicated_authorized`），未改动生产层：",
            "`relation_evidences.csv` 落地、`person_relations.publish_status` 转换与发布层重建属下一批次，",
            "须另行幂等脚本、守门测试与验收（参照第三/四批A模式）。",
            "",
        ]
    )
    path.write_text("\n".join(lines), encoding="utf-8", newline="")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--candidates", type=Path, default=DEFAULT_CANDIDATES)
    parser.add_argument("--out-dir", type=Path, default=DEFAULT_OUT_DIR)
    parser.add_argument("--overnight-dir", type=Path, default=DEFAULT_OVERNIGHT_DIR)
    parser.add_argument("--authorization-quote", required=True)
    args = parser.parse_args()
    print(adjudicate(args.candidates, args.out_dir, args.authorization_quote, args.overnight_dir))


if __name__ == "__main__":
    main()
