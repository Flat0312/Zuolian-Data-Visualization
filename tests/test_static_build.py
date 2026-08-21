from __future__ import annotations

import subprocess
import sys
from pathlib import Path

from conftest import PROJECT_ROOT, create_standard_dataset


def test_build_static_site_cli_writes_to_custom_output_dir(sandbox_tmp_path: Path) -> None:
    data_dir = create_standard_dataset(sandbox_tmp_path)
    output_dir = sandbox_tmp_path / "site-out"

    completed = subprocess.run(
        [
            sys.executable,
            str(PROJECT_ROOT / "build_static_site.py"),
            "--data-dir",
            str(data_dir),
            "--output-dir",
            str(output_dir),
        ],
        cwd=PROJECT_ROOT,
        capture_output=True,
        text=True,
        check=False,
    )

    assert completed.returncode == 0, completed.stderr
    assert (output_dir / "index.html").exists()
    assert (output_dir / "people" / "index.html").exists()
    assert (output_dir / "assets" / "search-index.json").exists()
