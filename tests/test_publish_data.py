from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
from conftest import create_standard_dataset


def test_build_publish_data_filters_candidate_memberships_and_keeps_references_closed(sandbox_tmp_path: Path) -> None:
    processed_dir = create_standard_dataset(sandbox_tmp_path)
    publish_dir = sandbox_tmp_path / "data" / "publish"
    report_path = sandbox_tmp_path / "publish_report.md"

    memberships_path = processed_dir / "org_memberships.csv"
    memberships = pd.read_csv(memberships_path)
    candidate = memberships.iloc[0].copy()
    candidate["membership_id"] = "M2"
    candidate["person_id"] = "P2"
    candidate["membership_type"] = "candidate"
    candidate["evidence_ids"] = "OME2"
    memberships = pd.concat([memberships, candidate.to_frame().T], ignore_index=True)
    memberships.to_csv(memberships_path, index=False, encoding="utf-8-sig")

    org_evidences_path = processed_dir / "org_membership_evidences.csv"
    org_evidences = pd.read_csv(org_evidences_path)
    candidate_evidence = org_evidences.iloc[0].copy()
    candidate_evidence["evidence_id"] = "OME2"
    candidate_evidence["person_id"] = "P2"
    org_evidences = pd.concat([org_evidences, candidate_evidence.to_frame().T], ignore_index=True)
    org_evidences.to_csv(org_evidences_path, index=False, encoding="utf-8-sig")

    facts_path = processed_dir / "fact_evidences.csv"
    facts = pd.read_csv(facts_path)
    candidate_fact = facts.iloc[0].copy()
    candidate_fact["evidence_id"] = "FE2"
    candidate_fact["subject_id"] = "P2"
    candidate_fact["origin_evidence_id"] = "OME2"
    facts = pd.concat([facts, candidate_fact.to_frame().T], ignore_index=True)
    facts.to_csv(facts_path, index=False, encoding="utf-8-sig")

    from research.analysis.build_publish_data import build_publish_data

    manifest = build_publish_data(processed_dir, publish_dir, report_path)

    published_memberships = pd.read_csv(publish_dir / "org_memberships.csv")
    published_org_evidences = pd.read_csv(publish_dir / "org_membership_evidences.csv")
    published_facts = pd.read_csv(publish_dir / "fact_evidences.csv")
    assert set(published_memberships["membership_type"]) == {"confirmed_member"}
    assert set(published_org_evidences["evidence_id"]) == {"OME1"}
    assert set(published_facts["evidence_id"]) == {"FE1"}
    assert manifest["tables"]["org_memberships.csv"]["filtered"] == 1
    assert json.loads((publish_dir / "publish_manifest.json").read_text(encoding="utf-8"))["schema_errors"] == 0
    assert report_path.exists()


def test_resolve_data_dir_can_select_public_or_research_mode(sandbox_tmp_path: Path) -> None:
    create_standard_dataset(sandbox_tmp_path)
    publish_dir = sandbox_tmp_path / "data" / "publish"
    publish_dir.mkdir(parents=True)

    from data_paths import resolve_data_dir

    app_dir = sandbox_tmp_path / "app" / "frontend"
    assert resolve_data_dir(app_dir, required_files=(), mode="public") == publish_dir
    assert resolve_data_dir(app_dir, required_files=(), mode="research") == sandbox_tmp_path / "data" / "processed"
