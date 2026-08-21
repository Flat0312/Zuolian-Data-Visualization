from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pandas as pd

from conftest import PROJECT_ROOT, create_standard_dataset


def test_audit_kb_returns_non_zero_and_writes_reports_for_broken_data(sandbox_tmp_path: Path) -> None:
    data_dir = create_standard_dataset(sandbox_tmp_path)
    report_dir = sandbox_tmp_path / "reports"

    participants_path = data_dir / "event_participants.csv"
    participants = pd.read_csv(participants_path)
    participants.loc[0, "person_id"] = "P999"
    participants.to_csv(participants_path, index=False, encoding="utf-8-sig")

    completed = subprocess.run(
        [
            sys.executable,
            str(PROJECT_ROOT / "research" / "analysis" / "audit_kb.py"),
            "--data-dir",
            str(data_dir),
            "--report-dir",
            str(report_dir),
        ],
        cwd=PROJECT_ROOT,
        capture_output=True,
        text=True,
        check=False,
    )

    assert completed.returncode != 0
    assert (report_dir / "audit_kb_report.md").exists()
    assert (report_dir / "audit_kb_issues.csv").exists()
