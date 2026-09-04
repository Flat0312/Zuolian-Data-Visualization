"""Phase7 首批核心人物关系证据候选包（只读生产数据，可重复运行）。

任务：冻结12条审计对象（鲁迅/丁玲/柔石/冯雪峰相关），逐条核验真实证据，
只建候选包，不写生产数据。

选择规则（任务书任务1，确定性）：
1. 至少一端属于 ZLH-001、ZLH-021、ZLH-016、ZLH-005；
2. final_relation_type 属于通信、交往、交游、创作合作；
3. risk 为 low/medium、needs_manual_review=no；
4. 按 source_ids 去重数量降序、relation_quality_score 降序（空值按 -1 计，最低）、
   relation_id 升序取前12条。

开始查证后不得换样本；二跑相同12条和相同顺序。

证据核验结论（VERIFIED_EVIDENCE）为人工逐条核验后登记的常量：
- 只允许 support / associated / conflict / insufficient；
- support 要求原文直接证明该具体关系，且 locator、逐字短引文均非空；
- 共现/同组织/背景一律记 associated，不得标 support；
- 全部行 review_status=pending_human_review；新来源 CAND-SRC-P7-*、新证据 CAND-RELE-P7-*；
- 引文不超过200字，登记访问日期与引文SHA-256；转载同一原文只算一个来源族；
- 每条关系最多核验2个独立来源族，全任务不超过30个来源。
"""
from __future__ import annotations

import argparse
import hashlib
from pathlib import Path

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_DATA_DIR = PROJECT_ROOT / "data" / "processed"
DEFAULT_OUT_DIR = PROJECT_ROOT / "research" / "drafts" / "reports"

SNAPSHOT_NOTE = "2026-09-04"
N_SELECT = 12
CORE_IDS = ("ZLH-001", "ZLH-021", "ZLH-016", "ZLH-005")
ALLOWED_TYPES = ("通信", "交往", "交游", "创作合作")
ALLOWED_RISK = ("low", "medium")

SELECTION_COLUMNS = (
    "selection_rank",
    "relation_id",
    "source_person_id",
    "source_person_name",
    "target_person_id",
    "target_person_name",
    "current_relation_type",
    "relation_risk_level",
    "needs_manual_review",
    "source_ids",
    "source_id_count",
    "relation_quality_score",
    "context",
    "evidence_ref",
    "publish_status",
)

CANDIDATE_COLUMNS = (
    "candidate_id",
    "relation_id",
    "source_person_id",
    "source_person_name",
    "target_person_id",
    "target_person_name",
    "current_relation_type",
    "proposed_relation_type",
    "evidence_support",
    "candidate_source_id",
    "source_title",
    "source_path_or_url",
    "source_family",
    "source_level",
    "locator",
    "quote",
    "quote_sha256",
    "access_date",
    "review_status",
    "receipt_id",
    "researcher_note",
)

RECEIPT_COLUMNS = (
    "receipt_id",
    "candidate_source_id",
    "source_title",
    "source_path_or_url",
    "retrieval_status",
    "access_date",
    "content_hash",
    "note",
)

PENDING = "pending_human_review"
MAX_SOURCES_PER_RELATION = 2
MAX_SOURCES_TOTAL = 30

# 人工逐条核验登记：relation_id -> 证据清单（按核验顺序，最多2个独立来源族）。
# 引文均在本地保存的原文中逐字复核；P7-001 的公开转录页面本轮 HTTP 200，
# 其余公开链接保留为人工复开入口。这里不把同组织共现升级为交游。
VERIFIED_EVIDENCE: dict[str, list[dict[str, str]]] = {
    "REL-00092": [{
        "proposed_relation_type": "通信", "evidence_support": "support",
        "candidate_source_id": "CAND-SRC-P7-001", "source_title": "鲁迅日记·日记十七（1928年）",
        "source_path_or_url": "https://zh.wikisource.org/zh-hans/鲁迅日记/日记十七",
        "source_family": "鲁迅日记", "source_level": "B（公开转录的一手日记；本地逐字核对）",
        "locator": "1928年11月26日条", "quote": "夜得小峰信及《而已集》、《语丝》。",
        "access_date": "2026-09-04", "receipt_id": "P7-RCP-001",
        "retrieval_status": "verified_http_200; local_text_checked",
        "content_hash": "c9d41730702add6034416c22ad363b3b256e1aaffec6ddab02583c94903a34c3",
        "receipt_note": "公开页响应字节 SHA-256；引文另在本地《鲁迅日记》全文逐字核对。",
        "researcher_note": "直接记录收信，支持通信；不外推通信频率。",
    }],
    "REL-00060": [{
        "proposed_relation_type": "通信", "evidence_support": "support",
        "candidate_source_id": "CAND-SRC-P7-001", "source_title": "鲁迅日记·日记十七（1928年）",
        "source_path_or_url": "https://zh.wikisource.org/zh-hans/鲁迅日记/日记十七",
        "source_family": "鲁迅日记", "source_level": "B（公开转录的一手日记；本地逐字核对）",
        "locator": "1928年2月26日条", "quote": "晚复。寄霁野信。",
        "access_date": "2026-09-04", "receipt_id": "P7-RCP-001",
        "researcher_note": "直接记录寄信，支持通信。",
    }],
    "REL-00109": [{
        "proposed_relation_type": "通信", "evidence_support": "support",
        "candidate_source_id": "CAND-SRC-P7-001", "source_title": "鲁迅日记·日记十七（1928年）",
        "source_path_or_url": "https://zh.wikisource.org/zh-hans/鲁迅日记/日记十七",
        "source_family": "鲁迅日记", "source_level": "B（公开转录的一手日记；本地逐字核对）",
        "locator": "1928年9月26日条", "quote": "午后寄陈望道信并稿。",
        "access_date": "2026-09-04", "receipt_id": "P7-RCP-001",
        "researcher_note": "直接记录寄信，支持通信。",
    }],
    "REL-00019": [{
        "proposed_relation_type": "通信", "evidence_support": "support",
        "candidate_source_id": "CAND-SRC-P7-001", "source_title": "鲁迅日记·日记十七（1928年）",
        "source_path_or_url": "https://zh.wikisource.org/zh-hans/鲁迅日记/日记十七",
        "source_family": "鲁迅日记", "source_level": "B（公开转录的一手日记；本地逐字核对）",
        "locator": "1928年4月1日条", "quote": "得郁达夫信。",
        "access_date": "2026-09-04", "receipt_id": "P7-RCP-001",
        "researcher_note": "直接记录收信，支持通信。",
    }],
    "REL-00089": [{
        "proposed_relation_type": "交游", "evidence_support": "support",
        "candidate_source_id": "CAND-SRC-P7-001", "source_title": "鲁迅日记·日记十七（1928年）",
        "source_path_or_url": "https://zh.wikisource.org/zh-hans/鲁迅日记/日记十七",
        "source_family": "鲁迅日记", "source_level": "B（公开转录的一手日记；本地逐字核对）",
        "locator": "1928年4月2日条", "quote": "达夫招饮于陶乐春，与广平同往。",
        "access_date": "2026-09-04", "receipt_id": "P7-RCP-001",
        "researcher_note": "直接共同活动记录，支持此条交游；不据此断言私人关系性质。",
    }],
    "REL-00097": [{
        "proposed_relation_type": "通信", "evidence_support": "support",
        "candidate_source_id": "CAND-SRC-P7-001", "source_title": "鲁迅日记·日记十七（1928年）",
        "source_path_or_url": "https://zh.wikisource.org/zh-hans/鲁迅日记/日记十七",
        "source_family": "鲁迅日记", "source_level": "B（公开转录的一手日记；本地逐字核对）",
        "locator": "1928年11月24日条", "quote": "午后寄语堂信。",
        "access_date": "2026-09-04", "receipt_id": "P7-RCP-001",
        "researcher_note": "直接记录寄信，支持通信。",
    }],
    "REL-00011": [{
        "proposed_relation_type": "通信", "evidence_support": "support",
        "candidate_source_id": "CAND-SRC-P7-002", "source_title": "鲁迅日记·日记二十四（1935年）",
        "source_path_or_url": "https://zh.wikisource.org/zh-hans/鲁迅日记/日记二十四",
        "source_family": "鲁迅日记", "source_level": "B（公开转录的一手日记；本地逐字核对）",
        "locator": "1935年2月21日条", "quote": "午后寄郑伯奇信。",
        "access_date": "2026-09-04", "receipt_id": "P7-RCP-002",
        "retrieval_status": "local_text_checked; public_locator_recorded",
        "content_hash": "22819f64addecb1529bc4c92e655f0589c26dd44203cc9dcba155123ab6be000",
        "receipt_note": "本地《鲁迅日记》全文 SHA-256；公开页链接已记录供人工复开。",
        "researcher_note": "直接记录寄信，支持通信。",
    }],
    "REL-00006": [{
        "proposed_relation_type": "通信", "evidence_support": "support",
        "candidate_source_id": "CAND-SRC-P7-001", "source_title": "鲁迅日记·日记十七（1928年）",
        "source_path_or_url": "https://zh.wikisource.org/zh-hans/鲁迅日记/日记十七",
        "source_family": "鲁迅日记", "source_level": "B（公开转录的一手日记；本地逐字核对）",
        "locator": "1928年7月20日条", "quote": "复冯雪峰信。",
        "access_date": "2026-09-04", "receipt_id": "P7-RCP-001",
        "researcher_note": "直接记录复信，支持通信。",
    }],
    # 现有 OCR 片段不能逐字复原，也没有取得可定位的一手或授权版本；故不填伪引文。
    "REL-00113": [],
    "REL-00523": [],
    "REL-01219": [],
    "REL-01743": [],
}

INSUFFICIENT_NOTES = {
    "REL-00113": "现有词典 OCR 仅见欢迎词题名，未取到可逐字复核的原文或可定位授权版本；不得据此确认交游。",
    "REL-00523": "现有词典 OCR 有救援丁玲的线索，但未取到可逐字复核的原文上下文；不得把线索或救援行动等同交游。",
    "REL-01219": "现有词典 OCR 有共同发起活动的线索，但未取到可逐字复核的原文；不得把共同活动等同交游。",
    "REL-01743": "现有材料仅为左联成员或工作背景共现，未找到两人直接关系的可定位原文。",
}


def _read(data_dir: Path, name: str) -> pd.DataFrame:
    return pd.read_csv(data_dir / name, encoding="utf-8-sig", dtype=str).fillna("")


def _split_ids(value: object) -> list[str]:
    return [i.strip() for i in str(value).replace("；", ";").replace("、", ";").split(";") if i.strip()]


def _quality_number(value: object) -> float:
    try:
        return float(str(value).strip())
    except (ValueError, TypeError):
        return -1.0


def select_relations(data_dir: Path) -> pd.DataFrame:
    """按任务书规则选出12条关系（确定性排序）。"""
    rels = _read(data_dir, "person_relations.csv")
    persons = _read(data_dir, "persons.csv")
    name_by_id = dict(zip(persons["person_id"], persons["standard_name"]))
    pool = rels[
        (
            rels["source_person_id"].isin(CORE_IDS)
            | rels["target_person_id"].isin(CORE_IDS)
        )
        & (rels["final_relation_type"].isin(ALLOWED_TYPES))
        & (rels["relation_risk_level"].isin(ALLOWED_RISK))
        & (rels["needs_manual_review"] == "no")
    ].copy()
    pool["source_id_count"] = pool["source_ids"].map(lambda v: len(set(_split_ids(v))))
    pool["_qnum"] = pool["relation_quality_score"].map(_quality_number)
    pool = pool.sort_values(
        ["source_id_count", "_qnum", "relation_id"], ascending=[False, False, True], kind="mergesort"
    ).head(N_SELECT)
    pool = pool.reset_index(drop=True)
    pool["selection_rank"] = range(1, len(pool) + 1)
    pool["source_person_name"] = pool["source_person_id"].map(lambda v: name_by_id.get(v, ""))
    pool["target_person_name"] = pool["target_person_id"].map(lambda v: name_by_id.get(v, ""))
    pool["current_relation_type"] = pool["final_relation_type"]
    return _finalize_selection(pool)


def _finalize_selection(pool: pd.DataFrame) -> pd.DataFrame:
    out = pd.DataFrame()
    out["selection_rank"] = pool["selection_rank"]
    out["relation_id"] = pool["relation_id"]
    out["source_person_id"] = pool["source_person_id"]
    out["source_person_name"] = pool["source_person_name"]
    out["target_person_id"] = pool["target_person_id"]
    out["target_person_name"] = pool["target_person_name"]
    out["current_relation_type"] = pool["current_relation_type"]
    out["relation_risk_level"] = pool["relation_risk_level"]
    out["needs_manual_review"] = pool["needs_manual_review"]
    out["source_ids"] = pool["source_ids"]
    out["source_id_count"] = pool["source_id_count"]
    out["relation_quality_score"] = pool["relation_quality_score"]
    out["context"] = pool["context"]
    out["evidence_ref"] = pool["evidence_ref"]
    out["publish_status"] = pool["publish_status"] if "publish_status" in pool.columns else ""
    return out[list(SELECTION_COLUMNS)]


def _quote_hash(quote: str) -> str:
    return hashlib.sha256(quote.strip().encode("utf-8")).hexdigest() if quote.strip() else ""


def build_evidence_rows(selection: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """由 VERIFIED_EVIDENCE 生成候选表与回执表（确定性，无时间戳）。"""
    cand_rows: list[dict[str, str]] = []
    receipt_rows: list[dict[str, str]] = []
    seen_receipts: set[str] = set()
    counter = 0
    for _, sel in selection.iterrows():
        rid = str(sel["relation_id"])
        evidences = VERIFIED_EVIDENCE.get(rid, [])
        if not evidences:
            counter += 1
            cand_rows.append(
                {
                    "candidate_id": f"CAND-RELE-P7-{counter:03d}",
                    "relation_id": rid,
                    "source_person_id": str(sel["source_person_id"]),
                    "source_person_name": str(sel["source_person_name"]),
                    "target_person_id": str(sel["target_person_id"]),
                    "target_person_name": str(sel["target_person_name"]),
                    "current_relation_type": str(sel["current_relation_type"]),
                    "proposed_relation_type": "",
                    "evidence_support": "insufficient",
                    "candidate_source_id": "",
                    "source_title": "",
                    "source_path_or_url": "",
                    "source_family": "",
                    "source_level": "",
                    "locator": "",
                    "quote": "",
                    "quote_sha256": "",
                    "access_date": "",
                    "review_status": PENDING,
                    "receipt_id": "",
                    "researcher_note": INSUFFICIENT_NOTES.get(rid, "尚未完成人工核验；缺证如实记录，待查。"),
                }
            )
            continue
        for ev in evidences[:MAX_SOURCES_PER_RELATION]:
            counter += 1
            quote = ev.get("quote", "")
            receipt_id = ev.get("receipt_id", "")
            cand_rows.append(
                {
                    "candidate_id": f"CAND-RELE-P7-{counter:03d}",
                    "relation_id": rid,
                    "source_person_id": str(sel["source_person_id"]),
                    "source_person_name": str(sel["source_person_name"]),
                    "target_person_id": str(sel["target_person_id"]),
                    "target_person_name": str(sel["target_person_name"]),
                    "current_relation_type": str(sel["current_relation_type"]),
                    "proposed_relation_type": ev.get("proposed_relation_type", ""),
                    "evidence_support": ev.get("evidence_support", ""),
                    "candidate_source_id": ev.get("candidate_source_id", ""),
                    "source_title": ev.get("source_title", ""),
                    "source_path_or_url": ev.get("source_path_or_url", ""),
                    "source_family": ev.get("source_family", ""),
                    "source_level": ev.get("source_level", ""),
                    "locator": ev.get("locator", ""),
                    "quote": quote,
                    "quote_sha256": _quote_hash(quote),
                    "access_date": ev.get("access_date", ""),
                    "review_status": PENDING,
                    "receipt_id": receipt_id,
                    "researcher_note": ev.get("researcher_note", ""),
                }
            )
            if receipt_id and receipt_id not in seen_receipts:
                seen_receipts.add(receipt_id)
                receipt_rows.append(
                    {
                        "receipt_id": receipt_id,
                        "candidate_source_id": ev.get("candidate_source_id", ""),
                        "source_title": ev.get("source_title", ""),
                        "source_path_or_url": ev.get("source_path_or_url", ""),
                        "retrieval_status": ev.get("retrieval_status", ""),
                        "access_date": ev.get("access_date", ""),
                        "content_hash": ev.get("content_hash", ""),
                        "note": ev.get("receipt_note", ""),
                    }
                )
    cand_df = pd.DataFrame(cand_rows, columns=list(CANDIDATE_COLUMNS))
    receipt_df = pd.DataFrame(receipt_rows, columns=list(RECEIPT_COLUMNS))
    return cand_df, receipt_df


def build(data_dir: Path, out_dir: Path) -> dict[str, object]:
    selection = select_relations(data_dir)
    cand_df, receipt_df = build_evidence_rows(selection)
    out_dir.mkdir(parents=True, exist_ok=True)
    selection.to_csv(out_dir / "phase7_relation_selection.csv", index=False, encoding="utf-8-sig")
    cand_df.to_csv(out_dir / "phase7_relation_evidence_candidates.csv", index=False, encoding="utf-8-sig")
    receipt_df.to_csv(out_dir / "phase7_source_receipts.csv", index=False, encoding="utf-8-sig")
    _write_report(out_dir, selection, cand_df, receipt_df)
    support = int((cand_df["evidence_support"] == "support").sum()) if len(cand_df) else 0
    associated = int((cand_df["evidence_support"] == "associated").sum()) if len(cand_df) else 0
    conflict = int((cand_df["evidence_support"] == "conflict").sum()) if len(cand_df) else 0
    insufficient = int((cand_df["evidence_support"] == "insufficient").sum()) if len(cand_df) else 0
    return {
        "selected": len(selection),
        "candidates": len(cand_df),
        "receipts": len(receipt_df),
        "support": support,
        "associated": associated,
        "conflict": conflict,
        "insufficient": insufficient,
    }


def _write_report(
    out_dir: Path, selection: pd.DataFrame, cand_df: pd.DataFrame, receipt_df: pd.DataFrame
) -> None:
    lines = [
        "# Phase7 首批核心人物关系证据候选报告（候选包，未经人工裁决）",
        "",
        f"- 数据快照说明：只读生产数据（{SNAPSHOT_NOTE}实测基线）；本轮不写生产数据，公开关系仍为0。",
        f"- 冻结样本：{len(selection)} 条关系（见 phase7_relation_selection.csv，开始查证后未替换）。",
        f"- 候选行：{len(cand_df)} 行；独立来源（回执）：{len(receipt_df)} 个（上限30，每条关系最多2个来源族）。",
        "- 结论口径（仅计数，非准确率、非总体可信度）：",
    ]
    for label in ("support", "associated", "conflict", "insufficient"):
        n = int((cand_df["evidence_support"] == label).sum()) if len(cand_df) else 0
        lines.append(f"  - {label}：{n}")
    lines += ["", "## 逐条结论", ""]
    by_rel: dict[str, list[dict[str, str]]] = {}
    for row in cand_df.to_dict("records"):
        by_rel.setdefault(str(row["relation_id"]), []).append({k: str(v) for k, v in row.items()})
    for _, sel in selection.iterrows():
        rid = str(sel["relation_id"])
        lines.append(
            f"### {rid}（{sel['source_person_name']}—{sel['target_person_name']}，当前类型：{sel['current_relation_type']}）"
        )
        lines.append("")
        lines.append(
            f"- 既有生产来源引用：{sel['source_id_count']} 个；本轮不以既有聚合字段替代逐字引文。"
        )
        rows = by_rel.get(rid, [])
        if not rows:
            lines.append("- 结论：insufficient（尚未核验）。")
        for r in rows:
            lines.append(
                f"- 结论：{r['evidence_support']}；来源：{r['source_title'] or '（无）'}"
                f"（{r['candidate_source_id'] or '无来源'}，{r['source_level'] or '未分级'}）；"
                f"定位：{r['locator'] or '（无）'}；回执：{r['receipt_id'] or '（无）'}"
            )
            if r["researcher_note"]:
                lines.append(f"  - 说明：{r['researcher_note']}")
        lines.append("")
    lines += [
        "## 缺口与局限",
        "",
        "- 本轮只建候选包，不写生产数据；全部行 pending_human_review，转正须人工裁决。",
        "- 共现/同组织/背景材料一律记 associated，不标 support；找不到即记 insufficient。",
        "- 转载同一原文只计一个来源族；百科与普通媒体只记D级线索。",
        "",
    ]
    (out_dir / "phase7_relation_evidence_report.md").write_text("\n".join(lines), encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="构建Phase7首批关系证据候选包（只读）")
    parser.add_argument("--data-dir", type=Path, default=DEFAULT_DATA_DIR)
    parser.add_argument("--out-dir", type=Path, default=DEFAULT_OUT_DIR)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    summary = build(args.data_dir, args.out_dir)
    print(f"phase7 relation candidates: {summary}")
    print(f"out_dir: {args.out_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
