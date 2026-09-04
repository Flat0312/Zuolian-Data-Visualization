"""关系发布状态：互斥五态 + 可复现派生规则。

五态定义（互斥）：
- verified: 人工确认且有可定位证据（当前生产为 0，须经人工裁决后方可标记）
- supported: 材料直接支持，但尚未完成人工确认（可进入发布层）
- inferred: 同组织、共现或规则推断（不进入发布层，仅研究探索）
- pending_review: 需要人工审核（不进入发布层）
- rejected: 证据不能支持（不进入发布层，预留）

派生规则（只迁移现有证据和状态，不新增史实判断）：
1. final_relation_type == 待核验 -> pending_review
2. needs_manual_review == yes -> pending_review
3. relation_risk_level in (critical, high) -> pending_review
4. confidence == low -> inferred（缺乏佐证应降低展示等级，不降低风险警示）
5. final_relation_type in (同属组织, 空间共现, 时空共现) -> inferred（共现/同组织不写成已确认史实）
6. 其余 -> supported（材料直接支持但尚未人工确认；verified 须人工明确裁决）

禁止反向降险：本模块绝不改写 relation_risk_level，仅派生展示等级。
"""
from __future__ import annotations

RELATION_PUBLISH_STATUSES = ("verified", "supported", "inferred", "pending_review", "rejected")
PUBLIC_RELATION_STATUSES = ("verified", "supported")

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


def derive_relation_publish_status(row: object) -> str:
    """接受 dict / Series，返回五态之一。缺字段时按保守原则返回 pending_review。"""
    if isinstance(row, dict):
        get = row.get
    else:
        def get(k: str, d: str = "") -> str:  # type: ignore
            try:
                return row.get(k, d)  # type: ignore
            except Exception:
                return d

    # 已有明确人工裁决列则尊重：publish_status 本身若为合法值则透传（幂等）
    existing = _s(get("publish_status", ""))
    if existing in RELATION_PUBLISH_STATUSES:
        return existing
    # 预留 rejected 透传：若调用方已标记 rejected，不降级为其他状态
    if _s(get("review_status", "")) == "rejected":
        return "rejected"

    final_type = _s(get("final_relation_type", ""))
    needs_review = _s(get("needs_manual_review", "")).lower()
    risk = _s(get("relation_risk_level", "")).lower()
    confidence = _s(get("confidence", "")).lower()

    if final_type == "待核验":
        return "pending_review"
    if needs_review == "yes":
        return "pending_review"
    if risk in ("critical", "high"):
        return "pending_review"
    if confidence == "low":
        return "inferred"
    if final_type in INFERRED_RELATION_TYPES:
        return "inferred"
    return "supported"


def is_public_relation_status(status: object) -> bool:
    return _s(status) in PUBLIC_RELATION_STATUSES
