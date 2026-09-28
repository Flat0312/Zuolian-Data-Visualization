# Phase 1 ORG-001 Evidence-Driven Membership Rebuild Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace role-based ORG-001 membership inference with an evidence-ledger-driven, reproducible classification pipeline.

**Architecture:** A focused research script extracts explicit membership evidence from local paginated sources, combines it with the existing raw-workbook lead, applies deterministic source-level rules, and writes both the evidence ledger and the derived membership conclusions. The schema validator treats the ledger as a required standard table and enforces references and allowed values.

**Tech Stack:** Python 3.10+, standard library CSV/regex/pathlib, pandas-based existing schema validator, pytest.

---

## File Structure

- Create `research/analysis/rebuild_org_memberships.py`: evidence extraction, classification and CSV generation.
- Create `tests/test_org_membership_rebuild.py`: TDD coverage for source levels, explicit evidence extraction and decision rules.
- Create `data/processed/org_membership_evidences.csv`: generated ORG-001 evidence ledger.
- Modify `data/processed/org_memberships.csv`: generated ORG-001 conclusions.
- Modify `research/analysis/build_standard_kb_pipeline.py`: remove role-based confirmed-member inference.
- Modify `kb_schema.py`: register and validate the evidence ledger and membership state values.
- Modify `app/frontend/data_paths.py`: load the evidence ledger as standard data.
- Modify `tests/conftest.py`: include a valid evidence-ledger fixture.
- Modify `README.md`: document current data snapshot and evidence-ledger workflow.

### Task 1: Define the Evidence Decision API

- [x] Write failing tests for A-level single-source confirmation, two independent B-level confirmation, insufficient evidence, related-person fallback, and disputed evidence.
- [x] Run `python -m pytest tests/test_org_membership_rebuild.py -v` and confirm failures are caused by the missing module.
- [x] Implement source-level mapping and deterministic decision functions in `rebuild_org_memberships.py`.
- [x] Re-run the focused tests and confirm they pass.

### Task 2: Extract Explicit Evidence from Local Paginated Sources

- [x] Write failing tests using a small OCR-spaced paginated text fixture.
- [x] Verify the tests fail because extraction is missing.
- [x] Implement whitespace-normalized matching, page detection, explicit membership phrase filtering, and source-page lookup.
- [x] Re-run the focused tests and confirm they pass.

### Task 3: Generate the ORG-001 Ledger and Conclusions

- [x] Write failing integration tests using a temporary mini dataset.
- [x] Verify the tests fail because generation is missing.
- [x] Implement ledger generation for all current ORG-001 leads and derived conclusion generation.
- [x] Run the focused tests.
- [x] Run the script against `data/processed/`.
- [x] Verify all 150 existing ORG-001 leads remain represented and every conclusion points to ledger evidence.

### Task 4: Enforce the Contract in Schema Validation

- [x] Add failing schema tests for missing evidence references and invalid membership states.
- [x] Verify the tests fail.
- [x] Register `org_membership_evidences.csv` and add allowed-value and evidence-link checks.
- [x] Run schema tests and the full test suite.

### Task 5: Stop the Legacy Pipeline from Recreating False Confirmations

- [x] Add a failing regression test proving direct-looking `persons.role` values cannot create confirmed membership without evidence.
- [x] Verify the regression test fails against the legacy behavior.
- [x] Modify `build_standard_kb_pipeline.py` so initial role-derived rows are candidates or related persons only.
- [x] Run the regression test and full suite.

### Task 6: Document and Verify Phase 1

- [x] Update README data snapshot and workflow notes.
- [x] Run `python kb_schema.py`.
- [x] Run `python -m pytest -v`.
- [x] Run `python build_static_site.py`.
- [x] Run `python -m ruff check app.py build_static_site.py kb_schema.py app research/analysis`（已执行；发现 463 项阶段外历史 lint 债务，Phase 1 新增脚本与测试的定向 lint 通过）。
- [x] Record Phase 1 metrics: confirmed, related, candidate, disputed, evidence count, source-level distribution.

## Phase 1 Acceptance Criteria

- `persons.role` is not sufficient to generate `confirmed_member`.
- Every ORG-001 conclusion has at least one ledger record.
- Every `confirmed_member` satisfies one-A or two-independent-B evidence rules.
- Candidate and related records remain available for research and display.
- The full test suite, schema validation and static build pass.
