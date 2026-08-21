from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest
from conftest import create_standard_dataset


def test_load_data_raises_contract_error_for_invalid_standard_data(sandbox_tmp_path: Path) -> None:
    create_standard_dataset(sandbox_tmp_path)
    app_dir = sandbox_tmp_path / "app" / "frontend"
    app_dir.mkdir(parents=True, exist_ok=True)

    sources_path = sandbox_tmp_path / "data" / "processed" / "sources.csv"
    sources = pd.read_csv(sources_path)
    sources.loc[0, "evidence_strength"] = ""
    sources.to_csv(sources_path, index=False, encoding="utf-8-sig")

    from data_loader import load_data
    from kb_schema import DataContractError

    with pytest.raises(DataContractError) as exc_info:
        load_data(app_dir)

    assert "sources.csv" in str(exc_info.value)
    assert "evidence_strength" in str(exc_info.value)


def test_load_data_includes_membership_conclusions_and_evidence(sandbox_tmp_path: Path) -> None:
    create_standard_dataset(sandbox_tmp_path)
    app_dir = sandbox_tmp_path / "app" / "frontend"
    app_dir.mkdir(parents=True, exist_ok=True)

    from data_loader import load_data

    data = load_data(app_dir)

    assert data.memberships.loc[0, "membership_type"] == "confirmed_member"
    assert data.membership_evidences.loc[0, "evidence_id"] == "OME1"
    assert data.fact_evidences.loc[0, "evidence_id"] == "FE1"
