from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
AUDIT_DIR = BASE_DIR / "audit"
RAW_EDGES_PATH = DATA_DIR / "edges.csv"
AUDITED_EDGES_PATH = DATA_DIR / "edges_audited.csv"
REPORT_PATH = AUDIT_DIR / "source_data_audit.md"

INTERNAL_ROLES = {"核心领导", "骨干成员", "普通成员"}
EXTERNAL_ROLES = {"外围联络人", "相关人士"}
COMMUNICATION_KEYWORDS = ("来信", "复信", "寄", "函", "通信", "电", "回信")
DEBATE_KEYWORDS = ("论战", "辩论", "笔战", "驳斥", "批判", "反驳", "回击")
ORGANIZATION_KEYWORDS = ("左联", "联盟", "大会", "委员", "机关刊", "宣言", "签名", "联名", "欢迎词")

# Verified conservatively. Anything not in this set is treated as pending review
# rather than being forced into a possibly wrong family label.
VERIFIED_KINSHIP_PAIRS: dict[frozenset[str], str] = {
    frozenset({"ZLH-001", "ZLH-122"}): "人工复核：鲁迅与许广平为伴侣。",
    frozenset({"ZLH-027", "ZLH-028"}): "人工复核：萧军与萧红为共同生活的伴侣。",
    frozenset({"ZLH-134", "ZLH-135"}): "上下文直接写有“鹿地亘和池田幸子夫妇”。",
}


@dataclass(slots=True)
class AuditSummary:
    duplicate_node_ids: int
    duplicate_node_labels: int
    bad_birth_death_rows: int
    missing_edge_nodes: int
    self_loops: int
    duplicate_edge_rows: int
    flagged_kinship_rows: int
    verified_kinship_rows: int
    flagged_kinship_pairs: int
    suspicious_list_like_rows: int
    invalid_event_entity_refs: int
    invalid_event_timestamps: int
    invalid_event_coordinates: int
    out_of_lifespan_events: int


def canonical_pair_key(source: str, target: str) -> str:
    return "__".join(sorted((str(source), str(target))))


def clean_text(value: object, limit: int = 180) -> str:
    text = " ".join(str(value or "").split())
    return text if len(text) <= limit else f"{text[:limit].rstrip()}..."


def split_ids(value: object) -> list[str]:
    return [item.strip() for item in str(value or "").split(";") if item.strip()]


def append_note(current: str, note: str) -> str:
    if not current:
        return note
    if note in current:
        return current
    return f"{current}；{note}"


def append_flag(current: str, flag: str) -> str:
    items = [item for item in str(current or "").split(",") if item]
    if flag not in items:
        items.append(flag)
    return ",".join(items)


def suggestion_for_row(row: pd.Series) -> str:
    text = f"{row.get('Evidence_Ref', '')} {row.get('Context', '')}"
    source_role = str(row.get("Source_Role", ""))
    target_role = str(row.get("Target_Role", ""))

    if any(keyword in text for keyword in COMMUNICATION_KEYWORDS):
        return "弱关联-通信"
    if any(keyword in text for keyword in DEBATE_KEYWORDS):
        return "弱关联-论战"
    if source_role in INTERNAL_ROLES and target_role in INTERNAL_ROLES:
        return "强关联-组织隶属"
    if any(keyword in text for keyword in ORGANIZATION_KEYWORDS):
        return "强关联-组织隶属"
    if source_role in EXTERNAL_ROLES or target_role in EXTERNAL_ROLES:
        return "弱关联-交游"
    return "弱关联-交游"


def is_list_like(text: str, evidence_ref: str) -> bool:
    if len(text) < 80:
        return False
    separators = sum(text.count(token) for token in ("、", "，", ",", ";", "；"))
    ref_count = len([item for item in str(evidence_ref).split(";") if item.strip()])
    return separators >= 8 and ref_count >= 3


def load_frames() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    nodes = pd.read_csv(DATA_DIR / "nodes.csv", encoding="utf-8-sig").fillna("")
    edges = pd.read_csv(RAW_EDGES_PATH, encoding="utf-8-sig").fillna("")
    events = pd.read_csv(DATA_DIR / "events.csv", encoding="utf-8-sig").fillna("")
    return nodes, edges, events


def audit_nodes(nodes: pd.DataFrame) -> tuple[pd.DataFrame, int, int, int]:
    audited = nodes.copy()
    duplicate_ids = int(audited["Id"].duplicated().sum())
    duplicate_labels = int(audited["Label"].duplicated().sum())
    birth_death_pattern = re.compile(r"^\d{4}\s*-\s*(?:\d{4}|\?)$")
    bad_birth_death = int(
        audited["Birth_Death"].astype(str).map(lambda value: bool(value) and not bool(birth_death_pattern.match(value))).sum()
    )
    return audited, duplicate_ids, duplicate_labels, bad_birth_death


def audit_edges(nodes: pd.DataFrame, edges: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, int], list[dict[str, object]], list[dict[str, object]]]:
    audited = edges.copy()
    audited["csv_line"] = audited.index + 2
    audited["pair_key"] = audited.apply(lambda row: canonical_pair_key(row["Source"], row["Target"]), axis=1)
    audited["Original_Relation_Type"] = audited["Relation_Type"]
    audited["Suggested_Relation_Type"] = ""
    audited["Audit_Flags"] = ""
    audited["Audit_Note"] = ""

    name_map = nodes.set_index("Id")["Label"].to_dict()
    role_map = nodes.set_index("Id")["Role"].to_dict()
    valid_ids = set(nodes["Id"])

    audited["Source_Name"] = audited["Source"].map(name_map).fillna(audited["Source"])
    audited["Target_Name"] = audited["Target"].map(name_map).fillna(audited["Target"])
    audited["Source_Role"] = audited["Source"].map(role_map).fillna("")
    audited["Target_Role"] = audited["Target"].map(role_map).fillna("")

    duplicate_edge_rows = int(
        audited.duplicated(subset=["Source", "Target", "Relation_Type", "Context", "Evidence_Ref", "Weight"]).sum()
    )
    missing_edge_nodes = int((~audited["Source"].isin(valid_ids) | ~audited["Target"].isin(valid_ids)).sum())
    self_loops = int((audited["Source"] == audited["Target"]).sum())

    flagged_kinship_samples: list[dict[str, object]] = []
    list_like_samples: list[dict[str, object]] = []

    for idx, row in audited.iterrows():
        pair = frozenset((str(row["Source"]), str(row["Target"])))
        relation_type = str(row["Relation_Type"])
        text = f"{row.get('Evidence_Ref', '')} {row.get('Context', '')}"

        if "亲属" in relation_type:
            if pair in VERIFIED_KINSHIP_PAIRS:
                audited.at[idx, "Audit_Flags"] = append_flag(audited.at[idx, "Audit_Flags"], "verified_kinship")
                audited.at[idx, "Audit_Note"] = append_note(audited.at[idx, "Audit_Note"], VERIFIED_KINSHIP_PAIRS[pair])
            else:
                suggestion = suggestion_for_row(row)
                audited.at[idx, "Relation_Type"] = "待核-疑似误标"
                audited.at[idx, "Suggested_Relation_Type"] = suggestion
                audited.at[idx, "Audit_Flags"] = append_flag(audited.at[idx, "Audit_Flags"], "kinship_mismatch")
                audited.at[idx, "Audit_Note"] = append_note(
                    audited.at[idx, "Audit_Note"],
                    "原始关系标为亲属，但证据更像并列共现/组织材料，需人工复核。",
                )
                if len(flagged_kinship_samples) < 30:
                    flagged_kinship_samples.append(
                        {
                            "line": int(row["csv_line"]),
                            "source": row["Source_Name"],
                            "target": row["Target_Name"],
                            "source_role": row["Source_Role"],
                            "target_role": row["Target_Role"],
                            "original_relation": relation_type,
                            "suggested_relation": suggestion,
                            "evidence": clean_text(row["Evidence_Ref"], 120),
                            "context": clean_text(row["Context"], 160),
                        }
                    )

        if is_list_like(str(row.get("Context", "")), str(row.get("Evidence_Ref", ""))):
            audited.at[idx, "Audit_Flags"] = append_flag(audited.at[idx, "Audit_Flags"], "list_like_evidence")
            audited.at[idx, "Audit_Note"] = append_note(
                audited.at[idx, "Audit_Note"],
                "上下文主要是长名单并列，可能不足以支持直接关系。",
            )
            if len(list_like_samples) < 30:
                list_like_samples.append(
                    {
                        "line": int(row["csv_line"]),
                        "source": row["Source_Name"],
                        "target": row["Target_Name"],
                        "relation": audited.at[idx, "Relation_Type"],
                        "original_relation": row["Original_Relation_Type"],
                        "evidence": clean_text(row["Evidence_Ref"], 120),
                        "context": clean_text(row["Context"], 160),
                    }
                )

    kinship_rows = audited["Original_Relation_Type"].astype(str).str.contains("亲属", regex=False)
    verified_kinship_rows = int((kinship_rows & audited["Audit_Flags"].str.contains("verified_kinship", regex=False)).sum())
    flagged_kinship_rows = int((kinship_rows & audited["Audit_Flags"].str.contains("kinship_mismatch", regex=False)).sum())
    suspicious_list_like_rows = int(audited["Audit_Flags"].str.contains("list_like_evidence", regex=False).sum())

    flagged_pair_count = int(
        audited.loc[audited["Audit_Flags"].str.contains("kinship_mismatch", regex=False), "pair_key"].nunique()
    )

    stats = {
        "missing_edge_nodes": missing_edge_nodes,
        "self_loops": self_loops,
        "duplicate_edge_rows": duplicate_edge_rows,
        "verified_kinship_rows": verified_kinship_rows,
        "flagged_kinship_rows": flagged_kinship_rows,
        "flagged_kinship_pairs": flagged_pair_count,
        "suspicious_list_like_rows": suspicious_list_like_rows,
    }
    return audited, stats, flagged_kinship_samples, list_like_samples


def audit_events(nodes: pd.DataFrame, events: pd.DataFrame) -> tuple[pd.DataFrame, int, int, int, int]:
    audited = events.copy()
    valid_ids = set(nodes["Id"])
    audited["csv_line"] = audited.index + 2
    audited["Datetime"] = pd.to_datetime(audited["Timestamp"], errors="coerce", format="mixed")
    audited["Longitude_Num"] = pd.to_numeric(audited["Longitude"], errors="coerce")
    audited["Latitude_Num"] = pd.to_numeric(audited["Latitude"], errors="coerce")
    audited["Entity_List"] = audited["Entity_ID"].apply(split_ids)

    invalid_refs = int(audited["Entity_List"].map(lambda ids: any(item not in valid_ids for item in ids)).sum())
    invalid_timestamps = int(audited["Datetime"].isna().sum())
    invalid_coordinates = int(
        (
            audited["Longitude_Num"].isna()
            | audited["Latitude_Num"].isna()
            | ~audited["Longitude_Num"].between(70, 140)
            | ~audited["Latitude_Num"].between(0, 55)
        ).sum()
    )

    lifespan_map: dict[str, tuple[int | None, int | None]] = {}
    for _, row in nodes.iterrows():
        birth_death = str(row["Birth_Death"])
        match = re.match(r"^(\d{4})\s*-\s*(\d{4}|\?)$", birth_death)
        if not match:
            lifespan_map[str(row["Id"])] = (None, None)
            continue
        birth = int(match.group(1))
        death = None if match.group(2) == "?" else int(match.group(2))
        lifespan_map[str(row["Id"])] = (birth, death)

    def out_of_lifespan(row: pd.Series) -> bool:
        if pd.isna(row["Datetime"]):
            return False
        year = int(row["Datetime"].year)
        for entity_id in row["Entity_List"]:
            birth, death = lifespan_map.get(entity_id, (None, None))
            if birth is not None and year < birth:
                return True
            if death is not None and year > death:
                return True
        return False

    audited["Out_Of_Lifespan"] = audited.apply(out_of_lifespan, axis=1)
    out_of_lifespan_events = int(audited["Out_Of_Lifespan"].sum())
    return audited, invalid_refs, invalid_timestamps, invalid_coordinates, out_of_lifespan_events


def build_report(
    summary: AuditSummary,
    flagged_kinship_samples: list[dict[str, object]],
    list_like_samples: list[dict[str, object]],
) -> str:
    lines: list[str] = []
    lines.append("# 左联知识库源数据审计报告")
    lines.append("")
    lines.append("## 结论")
    lines.append("")
    lines.append(f"- 当前 `edges.csv` 中共有 **{summary.flagged_kinship_rows + summary.verified_kinship_rows}** 条原始“亲属”边。")
    lines.append(f"- 其中仅 **{summary.verified_kinship_rows}** 条被保留为已核亲属；其余 **{summary.flagged_kinship_rows}** 条已在 `edges_audited.csv` 中改标为 `待核-疑似误标`。")
    lines.append(f"- 被标记为疑似误标的原始亲属边覆盖 **{summary.flagged_kinship_pairs}** 组人物对。")
    lines.append(f"- 另有 **{summary.suspicious_list_like_rows}** 条关系边存在“长名单并列”风险，可能并不构成直接关系。")
    lines.append("")
    lines.append("## 结构校验")
    lines.append("")
    lines.append(f"- 重复人物 ID：{summary.duplicate_node_ids}")
    lines.append(f"- 重复人物名称：{summary.duplicate_node_labels}")
    lines.append(f"- 生卒年格式异常：{summary.bad_birth_death_rows}")
    lines.append(f"- 关系边缺失人物引用：{summary.missing_edge_nodes}")
    lines.append(f"- 关系自环：{summary.self_loops}")
    lines.append(f"- 完全重复关系行：{summary.duplicate_edge_rows}")
    lines.append(f"- 事件缺失人物引用：{summary.invalid_event_entity_refs}")
    lines.append(f"- 事件时间无法解析：{summary.invalid_event_timestamps}")
    lines.append(f"- 事件坐标异常：{summary.invalid_event_coordinates}")
    lines.append(f"- 事件时间落在人物生卒范围之外：{summary.out_of_lifespan_events}")
    lines.append("")
    lines.append("## 高风险样本")
    lines.append("")
    lines.append("### 原始亲属误标样本")
    lines.append("")
    for sample in flagged_kinship_samples[:20]:
        lines.append(
            f"- line {sample['line']}: {sample['source']}（{sample['source_role']}） -> "
            f"{sample['target']}（{sample['target_role']}），原标 `{sample['original_relation']}`，"
            f"建议 `{sample['suggested_relation']}`。证据：{sample['evidence']}；上下文：{sample['context']}"
        )
    lines.append("")
    lines.append("### 长名单并列风险样本")
    lines.append("")
    for sample in list_like_samples[:20]:
        lines.append(
            f"- line {sample['line']}: {sample['source']} -> {sample['target']}，现标 `{sample['relation']}`"
            f"（原始 `{sample['original_relation']}`）。证据：{sample['evidence']}；上下文：{sample['context']}"
        )
    lines.append("")
    lines.append("## 输出文件")
    lines.append("")
    lines.append(f"- 审计后关系表：`{AUDITED_EDGES_PATH}`")
    lines.append(f"- 本报告：`{REPORT_PATH}`")
    lines.append("")
    return "\n".join(lines)


def main() -> None:
    AUDIT_DIR.mkdir(parents=True, exist_ok=True)

    nodes, edges, events = load_frames()
    audited_nodes, duplicate_node_ids, duplicate_node_labels, bad_birth_death_rows = audit_nodes(nodes)
    audited_edges, edge_stats, flagged_kinship_samples, list_like_samples = audit_edges(audited_nodes, edges)
    _, invalid_event_refs, invalid_event_timestamps, invalid_event_coordinates, out_of_lifespan_events = audit_events(
        audited_nodes, events
    )

    summary = AuditSummary(
        duplicate_node_ids=duplicate_node_ids,
        duplicate_node_labels=duplicate_node_labels,
        bad_birth_death_rows=bad_birth_death_rows,
        missing_edge_nodes=edge_stats["missing_edge_nodes"],
        self_loops=edge_stats["self_loops"],
        duplicate_edge_rows=edge_stats["duplicate_edge_rows"],
        flagged_kinship_rows=edge_stats["flagged_kinship_rows"],
        verified_kinship_rows=edge_stats["verified_kinship_rows"],
        flagged_kinship_pairs=edge_stats["flagged_kinship_pairs"],
        suspicious_list_like_rows=edge_stats["suspicious_list_like_rows"],
        invalid_event_entity_refs=invalid_event_refs,
        invalid_event_timestamps=invalid_event_timestamps,
        invalid_event_coordinates=invalid_event_coordinates,
        out_of_lifespan_events=out_of_lifespan_events,
    )

    audited_edges.to_csv(AUDITED_EDGES_PATH, index=False, encoding="utf-8-sig")
    REPORT_PATH.write_text(build_report(summary, flagged_kinship_samples, list_like_samples), encoding="utf-8")

    print(f"审计完成：{AUDITED_EDGES_PATH}")
    print(f"报告位置：{REPORT_PATH}")
    print(
        "关键统计："
        f" flagged_kinship_rows={summary.flagged_kinship_rows},"
        f" verified_kinship_rows={summary.verified_kinship_rows},"
        f" suspicious_list_like_rows={summary.suspicious_list_like_rows}"
    )


if __name__ == "__main__":
    main()
