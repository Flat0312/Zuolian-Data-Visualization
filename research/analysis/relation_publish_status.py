"""关系发布状态：互斥五态 + 派生/人工双来源（返修版）。

五态定义（互斥，语义以 relation_evidences 证据为准）：
- verified: 人工确认，且至少一条可定位的 support 证据（origin=human_adjudication）
- supported: 至少一条 evidence_support=support、未 rejected、带 locator 且 quote/context 非空的证据
- inferred: 只有 associated、共现、同组织或规则推断（不公开）
- pending_review: 待核验、critical/high 风险、needs_manual_review=yes 或证据冲突（不公开）
- rejected: 人工否定（origin=human_adjudication）

来源（publish_status_origin）：
- derived: 规则自动派生。**每次全量重算，绝不透传旧值**——旧 supported 一旦风险
  变为 critical 或证据被撤，重算后必须离开公开层。
- human_adjudication: 人工裁决。仅 verified/rejected 可保留，且必须带
  reviewer / reviewed_at / review_note 审计字段（Schema 强制）。

判定顺序（derived）：待核验 → 需人工复核 → critical/high → 证据冲突 →
低置信(low) → 推断类型（同属组织/空间共现/时空共现） → 合格 support 证据 → inferred。
low 与推断类型即使持有 support 证据也不得进入公开层（保守原则：史料语义正确优先）。

另提供 is_low_risk_heuristic_relation：历史 1760 条启发式口径（只看类型/风险/
置信/复核标记，不看证据、不看 publish_status）。该口径**不得称为可信关系**，
仅用于"较低风险规则筛选网络"等研究对照。

禁止反向降险：本模块绝不改写 relation_risk_level。
"""
from __future__ import annotations

RELATION_PUBLISH_STATUSES = ("verified", "supported", "inferred", "pending_review", "rejected")
PUBLIC_RELATION_STATUSES = ("verified", "supported")
RELATION_PUBLISH_STATUS_ORIGINS = ("derived", "human_adjudication")
HUMAN_ONLY_STATUSES = ("verified", "rejected")
HUMAN_ADJUDICATION_AUDIT_FIELDS = ("reviewer", "reviewed_at", "review_note")

INFERRED_RELATION_TYPES = {"同属组织", "空间共现", "时空共现"}


def _s(value: object) -> str:
    if value is None:
        return ""
    try:
        import pandas as pd  # type: ignore

        if isinstance(value, float) and pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def _get(row: object, key: str) -> str:
    if isinstance(row, dict):
        return _s(row.get(key, ""))
    try:
        return _s(row.get(key, ""))  # type: ignore[union-attr]
    except Exception:
        return ""


def has_qualifying_support_evidence(evidence_rows: object) -> bool:
    """是否存在合格的直接支持证据：support + 未 rejected + locator + (quote|context)。"""
    if not evidence_rows:
        return False
    for ev in evidence_rows:  # type: ignore[union-attr]
        if _get(ev, "evidence_support") != "support":
            continue
        if _get(ev, "review_status") == "rejected":
            continue
        if _get(ev, "locator") and (_get(ev, "quote") or _get(ev, "context")):
            return True
    return False


def has_conflict_evidence(evidence_rows: object) -> bool:
    if not evidence_rows:
        return False
    return any(_get(ev, "evidence_support") == "conflict" for ev in evidence_rows)  # type: ignore[union-attr]


def has_human_adjudication(row: object) -> bool:
    """人工裁决仅当 origin 标记为 human_adjudication 且状态属于 verified/rejected。"""
    return _get(row, "publish_status_origin") == "human_adjudication" and _get(row, "publish_status") in HUMAN_ONLY_STATUSES


def derive_relation_publish_status_with_origin(row: object, evidence_rows: object = None) -> tuple[str, str]:
    """返回 (publish_status, publish_status_origin)。

    - human_adjudication 的 verified/rejected 予以保留（人工裁决优先）；
    - 其余一切状态（包括旧 supported）按当前行 + 当前证据全量重算，不透传。
    """
    if has_human_adjudication(row):
        return _get(row, "publish_status"), "human_adjudication"

    if _get(row, "final_relation_type") == "待核验":
        return "pending_review", "derived"
    if _get(row, "needs_manual_review").lower() == "yes":
        return "pending_review", "derived"
    if _get(row, "relation_risk_level").lower() in ("critical", "high"):
        return "pending_review", "derived"
    if has_conflict_evidence(evidence_rows):
        return "pending_review", "derived"
    if _get(row, "confidence").lower() == "low":
        return "inferred", "derived"
    if _get(row, "final_relation_type") in INFERRED_RELATION_TYPES:
        return "inferred", "derived"
    if has_qualifying_support_evidence(evidence_rows):
        return "supported", "derived"
    return "inferred", "derived"


def derive_relation_publish_status(row: object, evidence_rows: object = None) -> str:
    status, _ = derive_relation_publish_status_with_origin(row, evidence_rows)
    return status


def is_public_relation_status(status: object) -> bool:
    return _s(status) in PUBLIC_RELATION_STATUSES


def is_low_risk_heuristic_relation(row: object) -> bool:
    """历史启发式筛选口径（原 1760 条）：只看类型/风险/置信/复核标记。

    不使用 relation_evidences，也不读 publish_status；命名上不得称为"可信关系"，
    仅用于较低风险规则筛选网络的对照研究。
    """
    if _get(row, "final_relation_type") == "待核验":
        return False
    if _get(row, "needs_manual_review").lower() == "yes":
        return False
    if _get(row, "relation_risk_level").lower() in ("critical", "high"):
        return False
    if _get(row, "confidence").lower() == "low":
        return False
    if _get(row, "final_relation_type") in INFERRED_RELATION_TYPES:
        return False
    return True


def assign_relation_status_columns(
    relations_row: object, evidence_rows: object, existing: dict[str, str] | None = None
) -> dict[str, str]:
    """为单条关系计算 publish_status 四列（含 origin 与审计字段）。

    existing 为该行现有列值 dict；human_adjudication 的审计字段被保留，
    派生行的审计字段一律清空（防止 derived 残留人工痕迹）。
    """
    existing = existing or {}
    status, origin = derive_relation_publish_status_with_origin(relations_row, evidence_rows)
    if origin == "human_adjudication":
        return {
            "publish_status": status,
            "publish_status_origin": origin,
            "reviewer": existing.get("reviewer", ""),
            "reviewed_at": existing.get("reviewed_at", ""),
            "review_note": existing.get("review_note", ""),
        }
    return {
        "publish_status": status,
        "publish_status_origin": "derived",
        "reviewer": "",
        "reviewed_at": "",
        "review_note": "",
    }
