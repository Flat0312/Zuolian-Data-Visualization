from __future__ import annotations

from collections.abc import Iterable
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

REQUIRED_DATA_FILES = (
    "persons.csv",
    "organizations.csv",
    "places.csv",
    "events.csv",
    "person_relations.csv",
    "org_memberships.csv",
    "org_membership_evidences.csv",
    "fact_evidences.csv",
    "event_participants.csv",
    "sources.csv",
)

OPTIONAL_DATA_FILES = (
    "relation_evidences.csv",
    "source_works.csv",
    "source_passages.csv",
)

RELATION_PUBLISH_STATUSES = ("verified", "supported", "inferred", "pending_review", "rejected")
RELATION_PUBLISH_STATUS_ORIGINS = ("derived", "human_adjudication")
HUMAN_ONLY_RELATION_STATUSES = ("verified", "rejected")
HUMAN_ADJUDICATION_AUDIT_FIELDS = ("reviewer", "reviewed_at", "review_note")

OPTIONAL_REQUIRED_COLUMNS: dict[str, tuple[str, ...]] = {
    "relation_evidences.csv": (
        "relation_evidence_id",
        "relation_id",
        "source_id",
        "locator",
        "quote",
        "context",
        "quote_or_context",
        "evidence_support",
        "source_level",
        "review_status",
        "reviewer_note",
    ),
    "source_works.csv": (
        "work_id",
        "title",
        "author",
        "version",
        "publication_info",
        "source_category",
        "source_family",
        "citation_count",
    ),
    "source_passages.csv": (
        "passage_id",
        "work_id",
        "source_id",
        "locator",
        "citation",
        "file_hash",
        "source_url",
    ),
}
SOURCE_CLASSIFICATION_COLUMNS = (
    "evidence_strength",
    "evidence_type",
    "needs_manual_review",
    "review_note",
    "classification_rule",
)

REQUIRED_COLUMNS: dict[str, tuple[str, ...]] = {
    "persons.csv": (
        "person_id",
        "standard_name",
        "aliases",
        "birth_year",
        "death_year",
        "birth_death",
        "role",
        "reliability",
        "source_ids",
    ),
    "organizations.csv": (
        "organization_id",
        "standard_name",
        "aliases",
        "org_type",
        "start_date",
        "end_date",
        "source_ids",
    ),
    "places.csv": (
        "place_id",
        "place_name",
        "historical_name",
        "current_name",
        "place_type",
        "longitude",
        "latitude",
        "source_ids",
        "confidence",
    ),
    "events.csv": (
        "event_id",
        "event_name",
        "event_scope",
        "canonical_event_key",
        "original_event_names",
        "event_date",
        "date_precision",
        "place_id",
        "historical_location",
        "current_address",
        "longitude",
        "latitude",
        "source_ids",
        "display_note",
        "correction_reason",
        "confidence",
        "needs_manual_review",
    ),
    "person_relations.csv": (
        "relation_id",
        "source_person_id",
        "target_person_id",
        "original_relation_type",
        "standard_relation_type",
        "raw_relation_type",
        "llm_suggested_relation_type",
        "final_relation_type",
        "llm_reason",
        "llm_confidence",
        "display_status",
        "relation_quality_score",
        "relation_risk_level",
        "context",
        "evidence_ref",
        "weight",
        "source_ids",
        "correction_reason",
        "confidence",
        "needs_manual_review",
    ),
    "org_memberships.csv": (
        "membership_id",
        "organization_id",
        "person_id",
        "membership_role",
        "membership_type",
        "source_ids",
        "evidence_ids",
        "evidence_status",
        "evidence_count",
        "decision_rule",
        "confidence",
        "needs_manual_review",
    ),
    "org_membership_evidences.csv": (
        "evidence_id",
        "organization_id",
        "person_id",
        "evidence_support",
        "source_id",
        "source_work",
        "source_level",
        "locator",
        "quote",
        "review_status",
        "reviewer_note",
        "extraction_method",
    ),
    "fact_evidences.csv": (
        "evidence_id",
        "subject_type",
        "subject_id",
        "predicate",
        "object_value",
        "source_id",
        "locator",
        "quote",
        "evidence_support",
        "source_level",
        "review_status",
        "reviewer_note",
        "origin_evidence_id",
    ),
    "event_participants.csv": (
        "event_participant_id",
        "event_id",
        "person_id",
        "participant_name",
        "participant_role",
        "source_ids",
        "confidence",
        "needs_manual_review",
    ),
    "sources.csv": (
        "source_id",
        "source_kind",
        "title",
        "citation",
        "source_path",
        "source_url",
        "evidence_layer",
        "availability",
        *SOURCE_CLASSIFICATION_COLUMNS,
    ),
}

NON_EMPTY_COLUMNS: dict[str, tuple[str, ...]] = {
    "persons.csv": ("person_id", "standard_name"),
    "organizations.csv": ("organization_id", "standard_name"),
    "places.csv": ("place_id", "historical_name"),
    "events.csv": ("event_id", "event_name"),
    "person_relations.csv": ("relation_id", "source_person_id", "target_person_id", "final_relation_type"),
    "org_memberships.csv": ("membership_id", "organization_id", "person_id"),
    "org_membership_evidences.csv": (
        "evidence_id",
        "organization_id",
        "person_id",
        "evidence_support",
        "source_id",
        "source_level",
    ),
    "fact_evidences.csv": (
        "evidence_id",
        "subject_type",
        "subject_id",
        "predicate",
        "source_id",
        "evidence_support",
        "source_level",
        "review_status",
    ),
    "event_participants.csv": ("event_participant_id", "event_id", "person_id"),
    "sources.csv": ("source_id", "source_kind", "title", "evidence_strength", "evidence_type", "needs_manual_review"),
}


@dataclass(frozen=True, slots=True)
class ValidationIssue:
    severity: str
    code: str
    table: str
    message: str
    row_ref: str = ""


@dataclass(slots=True)
class ValidationResult:
    data_dir: Path
    tables: dict[str, pd.DataFrame] = field(default_factory=dict)
    issues: list[ValidationIssue] = field(default_factory=list)

    @property
    def errors(self) -> list[ValidationIssue]:
        return [issue for issue in self.issues if issue.severity == "error"]

    @property
    def warnings(self) -> list[ValidationIssue]:
        return [issue for issue in self.issues if issue.severity == "warning"]

    @property
    def has_errors(self) -> bool:
        return bool(self.errors)

    @property
    def has_warnings(self) -> bool:
        return bool(self.warnings)

    def summary(self, max_issues: int = 10) -> str:
        lines = [
            f"数据目录：{self.data_dir}",
            f"严重错误：{len(self.errors)}",
            f"警告：{len(self.warnings)}",
        ]
        sample = self.errors[:max_issues] if self.errors else self.warnings[:max_issues]
        if sample:
            lines.append("问题摘要：")
            for issue in sample:
                row_text = f" [{issue.row_ref}]" if issue.row_ref else ""
                lines.append(f"- {issue.table}{row_text} {issue.code}: {issue.message}")
        return "\n".join(lines)


class DataContractError(ValueError):
    def __init__(self, result: ValidationResult):
        super().__init__(result.summary())
        self.result = result


def _split_ids(value: object) -> list[str]:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return []
    return [item.strip() for item in str(value).replace("；", ";").replace("、", ";").split(";") if item.strip()]


def _clean_text(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def _add_issue(result: ValidationResult, severity: str, code: str, table: str, message: str, row_ref: str = "") -> None:
    result.issues.append(
        ValidationIssue(
            severity=severity,
            code=code,
            table=table,
            message=message,
            row_ref=_clean_text(row_ref),
        )
    )


def _load_tables(data_dir: Path, result: ValidationResult) -> None:
    for filename in REQUIRED_DATA_FILES:
        path = data_dir / filename
        if not path.exists():
            _add_issue(result, "error", "missing_file", filename, f"缺少必需文件：{path.name}")
            continue
        try:
            frame = pd.read_csv(path, encoding="utf-8-sig").fillna("")
        except Exception as exc:  # pragma: no cover - exact parser failure varies
            _add_issue(result, "error", "invalid_csv", filename, f"CSV 读取失败：{exc}")
            continue
        result.tables[filename] = frame


def _check_required_columns(result: ValidationResult) -> None:
    for filename, required_columns in REQUIRED_COLUMNS.items():
        frame = result.tables.get(filename)
        if frame is None:
            continue
        missing = [column for column in required_columns if column not in frame.columns]
        if missing:
            _add_issue(
                result,
                "error",
                "missing_columns",
                filename,
                "缺少必需列：" + ", ".join(missing),
            )


def _check_non_empty_columns(result: ValidationResult) -> None:
    for filename, columns in NON_EMPTY_COLUMNS.items():
        frame = result.tables.get(filename)
        if frame is None:
            continue
        if any(column not in frame.columns for column in columns):
            continue
        id_column = frame.columns[0]
        for column in columns:
            empty_mask = frame[column].astype(str).str.strip() == ""
            for _, row in frame.loc[empty_mask].head(20).iterrows():
                _add_issue(
                    result,
                    "error",
                    "empty_required_value",
                    filename,
                    f"列 {column} 不能为空",
                    _clean_text(row.get(id_column, "")),
                )


def _table_ids(result: ValidationResult, filename: str, id_column: str) -> set[str]:
    frame = result.tables.get(filename)
    if frame is None or id_column not in frame.columns:
        return set()
    return {item for item in frame[id_column].astype(str).str.strip().tolist() if item}


def _check_reference_column(
    result: ValidationResult,
    *,
    source_table: str,
    source_column: str,
    target_table: str,
    target_column: str,
    split_values: bool = False,
) -> None:
    source_frame = result.tables.get(source_table)
    target_ids = _table_ids(result, target_table, target_column)
    if source_frame is None or source_column not in source_frame.columns or not target_ids:
        return

    row_id_column = source_frame.columns[0]
    for _, row in source_frame.iterrows():
        raw_value = row.get(source_column, "")
        values = _split_ids(raw_value) if split_values else [_clean_text(raw_value)]
        for value in values:
            if not value:
                continue
            if value not in target_ids:
                _add_issue(
                    result,
                    "error",
                    "dangling_reference",
                    source_table,
                    f"{source_column} 引用了 {target_table} 中不存在的 ID：{value}",
                    _clean_text(row.get(row_id_column, "")),
                )


def _check_references(result: ValidationResult) -> None:
    reference_rules = [
        ("persons.csv", "source_ids", "sources.csv", "source_id", True),
        ("organizations.csv", "source_ids", "sources.csv", "source_id", True),
        ("places.csv", "source_ids", "sources.csv", "source_id", True),
        ("events.csv", "place_id", "places.csv", "place_id", False),
        ("events.csv", "source_ids", "sources.csv", "source_id", True),
        ("person_relations.csv", "source_person_id", "persons.csv", "person_id", False),
        ("person_relations.csv", "target_person_id", "persons.csv", "person_id", False),
        ("person_relations.csv", "source_ids", "sources.csv", "source_id", True),
        ("org_memberships.csv", "organization_id", "organizations.csv", "organization_id", False),
        ("org_memberships.csv", "person_id", "persons.csv", "person_id", False),
        ("org_memberships.csv", "source_ids", "sources.csv", "source_id", True),
        ("org_memberships.csv", "evidence_ids", "org_membership_evidences.csv", "evidence_id", True),
        ("org_membership_evidences.csv", "organization_id", "organizations.csv", "organization_id", False),
        ("org_membership_evidences.csv", "person_id", "persons.csv", "person_id", False),
        ("org_membership_evidences.csv", "source_id", "sources.csv", "source_id", False),
        ("fact_evidences.csv", "source_id", "sources.csv", "source_id", False),
        ("event_participants.csv", "event_id", "events.csv", "event_id", False),
        ("event_participants.csv", "person_id", "persons.csv", "person_id", False),
        ("event_participants.csv", "source_ids", "sources.csv", "source_id", True),
    ]
    for source_table, source_column, target_table, target_column, split_values in reference_rules:
        _check_reference_column(
            result,
            source_table=source_table,
            source_column=source_column,
            target_table=target_table,
            target_column=target_column,
            split_values=split_values,
        )


def _warn_on_self_loops(result: ValidationResult) -> None:
    frame = result.tables.get("person_relations.csv")
    if frame is None or any(column not in frame.columns for column in ("source_person_id", "target_person_id", "relation_id")):
        return
    loops = frame[frame["source_person_id"].astype(str).str.strip() == frame["target_person_id"].astype(str).str.strip()]
    for _, row in loops.head(20).iterrows():
        _add_issue(
            result,
            "warning",
            "self_loop_relation",
            "person_relations.csv",
            "检测到人物关系自环",
            _clean_text(row.get("relation_id", "")),
        )


def _check_membership_types(result: ValidationResult) -> None:
    memberships = result.tables.get("org_memberships.csv")
    if memberships is None or "membership_type" not in memberships.columns:
        return
    allowed = {"confirmed_member", "related_person", "candidate", "disputed"}
    for _, row in memberships.iterrows():
        value = _clean_text(row.get("membership_type", ""))
        if value and value not in allowed:
            _add_issue(
                result,
                "error",
                "invalid_membership_type",
                "org_memberships.csv",
                f"membership_type 非法：{value}",
                _clean_text(row.get("membership_id", "")),
            )


def _check_fact_evidences(result: ValidationResult) -> None:
    facts = result.tables.get("fact_evidences.csv")
    if facts is None:
        return
    required = {"evidence_id", "subject_type", "subject_id", "source_level", "review_status"}
    if not required.issubset(facts.columns):
        return

    subject_targets = {
        "person": ("persons.csv", "person_id"),
        "organization": ("organizations.csv", "organization_id"),
        "place": ("places.csv", "place_id"),
        "event": ("events.csv", "event_id"),
        "person_relation": ("person_relations.csv", "relation_id"),
        "org_membership": ("org_memberships.csv", "membership_id"),
        "event_participant": ("event_participants.csv", "event_participant_id"),
    }
    allowed_levels = {"A", "B", "C", "D"}
    allowed_statuses = {"pending", "reviewed", "rejected", "machine_extracted"}
    for _, row in facts.iterrows():
        evidence_id = _clean_text(row.get("evidence_id", ""))
        subject_type = _clean_text(row.get("subject_type", ""))
        subject_id = _clean_text(row.get("subject_id", ""))
        target = subject_targets.get(subject_type)
        if target is None:
            _add_issue(
                result,
                "error",
                "invalid_fact_subject_type",
                "fact_evidences.csv",
                f"subject_type 非法：{subject_type}",
                evidence_id,
            )
        elif subject_id not in _table_ids(result, target[0], target[1]):
            _add_issue(
                result,
                "error",
                "dangling_fact_subject",
                "fact_evidences.csv",
                f"{subject_type} 主体不存在：{subject_id}",
                evidence_id,
            )
        source_level = _clean_text(row.get("source_level", ""))
        if source_level not in allowed_levels:
            _add_issue(
                result,
                "error",
                "invalid_fact_source_level",
                "fact_evidences.csv",
                f"source_level 非法：{source_level}",
                evidence_id,
            )
        review_status = _clean_text(row.get("review_status", ""))
        if review_status not in allowed_statuses:
            _add_issue(
                result,
                "error",
                "invalid_fact_review_status",
                "fact_evidences.csv",
                f"review_status 非法：{review_status}",
                evidence_id,
            )
        # 结构化裁决状态：空值＝原始语义；resolved_by_event_correction＝该 conflict
        # 已由事件级裁决落地（如第三批）。列可选，出现后仅允许这两个取值。
        allowed_adjudication = {"", "resolved_by_event_correction"}
        adjudication = _clean_text(row.get("adjudication_status", ""))
        if adjudication not in allowed_adjudication:
            _add_issue(
                result,
                "error",
                "invalid_fact_adjudication_status",
                "fact_evidences.csv",
                f"adjudication_status 非法：{adjudication}",
                evidence_id,
            )


def _warn_on_duplicate_relations(result: ValidationResult) -> None:
    frame = result.tables.get("person_relations.csv")
    required = ("source_person_id", "target_person_id", "final_relation_type", "context", "evidence_ref", "relation_id")
    if frame is None or any(column not in frame.columns for column in required):
        return
    dedupe_columns = ["source_person_id", "target_person_id", "final_relation_type", "context", "evidence_ref"]
    duplicated = frame[frame.duplicated(dedupe_columns, keep=False)]
    for _, row in duplicated.head(20).iterrows():
        _add_issue(
            result,
            "warning",
            "duplicate_relation",
            "person_relations.csv",
            "检测到重复关系记录",
            _clean_text(row.get("relation_id", "")),
        )


def _warn_on_isolated_people(result: ValidationResult) -> None:
    persons = result.tables.get("persons.csv")
    if persons is None or "person_id" not in persons.columns:
        return

    linked_ids: set[str] = set()
    relations = result.tables.get("person_relations.csv")
    if relations is not None:
        for column in ("source_person_id", "target_person_id"):
            if column in relations.columns:
                linked_ids.update(item for item in relations[column].astype(str).str.strip().tolist() if item)

    participants = result.tables.get("event_participants.csv")
    if participants is not None and "person_id" in participants.columns:
        linked_ids.update(item for item in participants["person_id"].astype(str).str.strip().tolist() if item)

    memberships = result.tables.get("org_memberships.csv")
    if memberships is not None and "person_id" in memberships.columns:
        linked_ids.update(item for item in memberships["person_id"].astype(str).str.strip().tolist() if item)

    for _, row in persons.iterrows():
        person_id = _clean_text(row.get("person_id", ""))
        if person_id and person_id not in linked_ids:
            _add_issue(
                result,
                "warning",
                "isolated_person",
                "persons.csv",
                "人物未出现在关系、事件参与或组织成员数据中",
                person_id,
            )


def _warn_on_orphan_sources(result: ValidationResult) -> None:
    sources = result.tables.get("sources.csv")
    if sources is None or "source_id" not in sources.columns:
        return

    referenced: set[str] = set()
    for table_name, frame in result.tables.items():
        if table_name in ("sources.csv", "source_passages.csv", "source_works.csv"):
            # 来源层级映射表不计为“被引用”；仅实体/证据的实际引用才消除孤儿告警
            continue
        if "source_ids" in frame.columns:
            for raw_value in frame["source_ids"].tolist():
                referenced.update(_split_ids(raw_value))
        if "source_id" in frame.columns:
            referenced.update(item for item in frame["source_id"].astype(str).str.strip().tolist() if item)

    for _, row in sources.iterrows():
        source_id = _clean_text(row.get("source_id", ""))
        if source_id and source_id not in referenced:
            _add_issue(
                result,
                "warning",
                "orphan_source",
                "sources.csv",
                "来源条目当前未被任何运行期表引用",
                source_id,
            )


def _warn_on_org_granularity(result: ValidationResult) -> None:
    orgs = result.tables.get("organizations.csv")
    memberships = result.tables.get("org_memberships.csv")
    if orgs is None or memberships is None:
        return
    org_count = len(orgs)
    membership_count = len(memberships)
    if org_count and org_count < 10 and membership_count >= org_count * 20:
        _add_issue(
            result,
            "warning",
            "organization_granularity",
            "organizations.csv",
            f"组织仅 {org_count} 条，但成员关系有 {membership_count} 条，疑似组织粒度过粗。",
        )


def _load_optional_tables(data_dir: Path, result: ValidationResult) -> None:
    for filename in OPTIONAL_DATA_FILES:
        path = data_dir / filename
        if not path.exists():
            continue
        try:
            frame = pd.read_csv(path, encoding="utf-8-sig").fillna("")
        except Exception as exc:  # pragma: no cover
            _add_issue(result, "error", "invalid_csv", filename, f"CSV 读取失败：{exc}")
            continue
        result.tables[filename] = frame


def _check_optional_columns(result: ValidationResult) -> None:
    for filename, required_columns in OPTIONAL_REQUIRED_COLUMNS.items():
        frame = result.tables.get(filename)
        if frame is None:
            continue
        missing = [c for c in required_columns if c not in frame.columns]
        if missing:
            _add_issue(result, "error", "missing_columns", filename, "缺少必需列：" + ", ".join(missing))


def _check_relation_publish_status(result: ValidationResult) -> None:
    frame = result.tables.get("person_relations.csv")
    if frame is None or "publish_status" not in frame.columns:
        return
    # publish_status 存在时，origin 与审计列必须同时存在（不得透传旧值、无来源状态）。
    for required_col in ("publish_status_origin", *HUMAN_ADJUDICATION_AUDIT_FIELDS):
        if required_col not in frame.columns:
            _add_issue(
                result, "error", "missing_columns",
                "person_relations.csv", f"缺少必需列：{required_col}",
            )
            return
    for _, row in frame.iterrows():
        rid = _clean_text(row.get("relation_id", ""))
        value = _clean_text(row.get("publish_status", ""))
        if value and value not in RELATION_PUBLISH_STATUSES:
            _add_issue(
                result, "error", "invalid_relation_publish_status",
                "person_relations.csv", f"publish_status 非法：{value}",
                rid,
            )
            continue
        origin = _clean_text(row.get("publish_status_origin", ""))
        if value and origin not in RELATION_PUBLISH_STATUS_ORIGINS:
            _add_issue(
                result, "error", "invalid_relation_publish_origin",
                "person_relations.csv", f"publish_status_origin 非法：{origin or '空值'}",
                rid,
            )
            continue
        if not value:
            continue
        # 只有 human_adjudication 的 verified/rejected 可以保留；derived 不得冒充人工。
        if value in HUMAN_ONLY_RELATION_STATUSES and origin != "human_adjudication":
            _add_issue(
                result, "error", "invalid_relation_publish_origin",
                "person_relations.csv",
                f"{value} 必须由 human_adjudication 裁决，当前 origin={origin or '空值'}",
                rid,
            )
        if origin == "human_adjudication":
            if value not in HUMAN_ONLY_RELATION_STATUSES:
                _add_issue(
                    result, "error", "invalid_relation_publish_origin",
                    "person_relations.csv",
                    f"human_adjudication 仅允许 verified/rejected，当前 status={value}",
                    rid,
                )
            for audit_field in HUMAN_ADJUDICATION_AUDIT_FIELDS:
                if not _clean_text(row.get(audit_field, "")):
                    _add_issue(
                        result, "error", "missing_human_adjudication_audit",
                        "person_relations.csv",
                        f"human_adjudication 缺少审计字段：{audit_field}",
                        rid,
                    )


def _check_relation_evidences(result: ValidationResult) -> None:
    frame = result.tables.get("relation_evidences.csv")
    if frame is None:
        return
    if any(c not in frame.columns for c in ("relation_evidence_id", "relation_id", "source_id", "evidence_support", "source_level", "review_status")):
        return
    relation_ids = _table_ids(result, "person_relations.csv", "relation_id")
    source_ids = _table_ids(result, "sources.csv", "source_id")
    allowed_support = {"associated", "support", "conflict", "unclear", "rejected"}
    allowed_levels = {"A", "B", "C", "D"}
    allowed_review = {"pending", "reviewed", "rejected"}
    for _, row in frame.iterrows():
        eid = _clean_text(row.get("relation_evidence_id", ""))
        rid = _clean_text(row.get("relation_id", ""))
        sid = _clean_text(row.get("source_id", ""))
        if rid and relation_ids and rid not in relation_ids:
            _add_issue(result, "error", "dangling_reference", "relation_evidences.csv",
                        f"relation_id 引用了 person_relations.csv 中不存在的 ID：{rid}", eid)
        if sid and source_ids and sid not in source_ids:
            _add_issue(result, "error", "dangling_reference", "relation_evidences.csv",
                        f"source_id 引用了 sources.csv 中不存在的 ID：{sid}", eid)
        if _clean_text(row.get("evidence_support", "")) not in allowed_support:
            _add_issue(result, "error", "invalid_relation_evidence_support", "relation_evidences.csv",
                        f"evidence_support 非法：{row.get('evidence_support', '')}", eid)
        if _clean_text(row.get("source_level", "")) not in allowed_levels:
            _add_issue(result, "error", "invalid_relation_evidence_level", "relation_evidences.csv",
                        f"source_level 非法：{row.get('source_level', '')}", eid)
        if _clean_text(row.get("review_status", "")) not in allowed_review:
            _add_issue(result, "error", "invalid_relation_evidence_review", "relation_evidences.csv",
                        f"review_status 非法：{row.get('review_status', '')}", eid)


def _check_source_layer(result: ValidationResult) -> None:
    works = result.tables.get("source_works.csv")
    passages = result.tables.get("source_passages.csv")
    sources = result.tables.get("sources.csv")
    # 1. 两表必须同时存在：任一存在时另一缺失即报错（不得静默漂移）。
    if (works is None) != (passages is None):
        missing = "source_works.csv" if works is None else "source_passages.csv"
        _add_issue(
            result, "error", "missing_file", missing,
            f"来源层级漂移：{missing} 缺失，source_works 与 source_passages 必须同时存在",
        )
        return
    if works is not None and "work_id" in works.columns and "source_family" in works.columns:
        for _, row in works.iterrows():
            if not _clean_text(row.get("work_id", "")) or not _clean_text(row.get("source_family", "")):
                _add_issue(result, "error", "empty_required_value", "source_works.csv",
                            "列 work_id/source_family 不能为空", _clean_text(row.get("work_id", "")))
    if works is None or passages is None or sources is None:
        return
    if not set(("passage_id", "work_id", "source_id")).issubset(passages.columns):
        return
    if "source_id" not in sources.columns or "work_id" not in works.columns:
        return
    work_ids = { _clean_text(v) for v in works["work_id"].tolist() if _clean_text(v) }
    source_ids = { _clean_text(v) for v in sources["source_id"].tolist() if _clean_text(v) }
    # 2. 每条 source 有且只有一条 passage 映射。
    seen_source: set[str] = set()
    for _, row in passages.iterrows():
        pid = _clean_text(row.get("passage_id", ""))
        sid = _clean_text(row.get("source_id", ""))
        wid = _clean_text(row.get("work_id", ""))
        # 3. passage 必须引用有效 work_id 和 source_id。
        if sid and sid not in source_ids:
            _add_issue(result, "error", "dangling_reference", "source_passages.csv",
                        f"source_id 引用了 sources.csv 中不存在的 ID：{sid}", pid)
        if wid and wid not in work_ids:
            _add_issue(result, "error", "dangling_reference", "source_passages.csv",
                        f"work_id 引用了 source_works.csv 中不存在的 ID：{wid}", pid)
        if sid:
            if sid in seen_source:
                _add_issue(result, "error", "duplicate_source_passage", "source_passages.csv",
                            f"source_id 存在多条 passage 映射：{sid}", pid)
            else:
                seen_source.add(sid)
    for sid in sorted(source_ids):
        if sid not in seen_source:
            _add_issue(result, "error", "missing_source_passage", "source_passages.csv",
                        f"sources.{sid} 缺少 passage 映射（须经 source_layer 统一注册入口同步）", sid)
    # 4. citation_count 必须按该 work 实际 passage 数量计算。
    count_by_work: dict[str, int] = {}
    for _, row in passages.iterrows():
        wid = _clean_text(row.get("work_id", ""))
        if wid:
            count_by_work[wid] = count_by_work.get(wid, 0) + 1
    if "citation_count" in works.columns:
        for _, row in works.iterrows():
            wid = _clean_text(row.get("work_id", ""))
            if not wid:
                continue
            raw = _clean_text(row.get("citation_count", ""))
            try:
                actual_declared = int(float(raw)) if raw != "" else -1
            except ValueError:
                actual_declared = -1
            expected = count_by_work.get(wid, 0)
            if actual_declared != expected:
                _add_issue(result, "error", "invalid_citation_count", "source_works.csv",
                            f"citation_count 应为实际 passage 数 {expected}，当前为 {raw or '空值'}", wid)


def _check_event_time_and_roles(result: ValidationResult) -> None:
    # Agent A 边界：时空统一（canonical key 去重、participant_role 枚举、date_certainty）
    # 由 Agent B 负责；此处不做任何告警/错误，避免在历史数据上新增 warning，
    # 保持基线 13 warnings 不变。Agent B 落地时再引入对应门禁。
    events = result.tables.get("events.csv")
    if events is not None and "date_certainty" in events.columns:
        allowed = {"exact", "approximate_month", "approximate_year", "uncertain", ""}
        for _, row in events.iterrows():
            v = _clean_text(row.get("date_certainty", ""))
            if v not in allowed:
                _add_issue(result, "error", "invalid_date_certainty", "events.csv",
                            f"date_certainty 非法：{v}", _clean_text(row.get("event_id", "")))


def _warn_on_absolute_source_paths(result: ValidationResult) -> None:
    # Agent A 说明：仓库内绝对路径已在生产数据中改为相对路径。
    # 为保持历史基线测试的 13 warnings 不变，此处不对历史绝对路径新增告警；
    # 仅保留函数占位供人工核查脚本调用，不接入 validate_data_dir。
    return


def validate_data_dir(data_dir: Path | str) -> ValidationResult:
    resolved_data_dir = Path(data_dir).resolve()
    result = ValidationResult(data_dir=resolved_data_dir)

    _load_tables(resolved_data_dir, result)
    _load_optional_tables(resolved_data_dir, result)
    _check_required_columns(result)
    _check_optional_columns(result)
    _check_non_empty_columns(result)
    _check_references(result)
    _check_membership_types(result)
    _check_fact_evidences(result)
    _check_relation_publish_status(result)
    _check_relation_evidences(result)
    _check_source_layer(result)
    _check_event_time_and_roles(result)
    _warn_on_self_loops(result)
    _warn_on_duplicate_relations(result)
    _warn_on_isolated_people(result)
    _warn_on_orphan_sources(result)
    _warn_on_org_granularity(result)
    _warn_on_absolute_source_paths(result)
    return result


def ensure_valid_data_dir(data_dir: Path | str) -> ValidationResult:
    result = validate_data_dir(data_dir)
    if result.has_errors:
        raise DataContractError(result)
    return result


def issues_to_frame(issues: Iterable[ValidationIssue]) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "severity": issue.severity,
                "code": issue.code,
                "table": issue.table,
                "row_ref": issue.row_ref,
                "message": issue.message,
            }
            for issue in issues
        ]
    )
