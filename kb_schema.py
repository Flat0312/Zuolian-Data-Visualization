from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable

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
        if table_name == "sources.csv":
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


def validate_data_dir(data_dir: Path | str) -> ValidationResult:
    resolved_data_dir = Path(data_dir).resolve()
    result = ValidationResult(data_dir=resolved_data_dir)

    _load_tables(resolved_data_dir, result)
    _check_required_columns(result)
    _check_non_empty_columns(result)
    _check_references(result)
    _check_membership_types(result)
    _check_fact_evidences(result)
    _warn_on_self_loops(result)
    _warn_on_duplicate_relations(result)
    _warn_on_isolated_people(result)
    _warn_on_orphan_sources(result)
    _warn_on_org_granularity(result)
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
