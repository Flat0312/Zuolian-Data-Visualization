from __future__ import annotations

import re

import pandas as pd


def clean_text(value: object, limit: int = 0, fallback: str = "") -> str:
    """Normalize whitespace, strip NaN. Truncate if limit > 0."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return fallback
    cleaned = " ".join(str(value).replace("\r", " ").split())
    if not cleaned:
        return fallback
    if limit and len(cleaned) > limit:
        return f"{cleaned[:limit].rstrip()}..."
    return cleaned


def split_ids(value: object) -> list[str]:
    """Split a delimited ID string on ; ； 、"""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return []
    return [item.strip() for item in re.split(r"[;；、]", str(value)) if item.strip()]
