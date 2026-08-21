"""第三批合并脚本在"真实旧提交基线"上的回归测试。

与 test_merge_batch3_review.py 的合成剥离基线不同，本测试用 `git archive`
从固定的历史提交 0f7d445（第三批数据合并前最后一个提交）原样取出
data/processed 与第三批候选草稿，在其上执行合并，验证：

- 增量恰为 sources +12、fact_evidences +20、events −2、event_participants −4；
- schema 保持 0 错误、≤13 警告；
- 二次执行零写入（幂等）。

若该提交在历史中被重写或丢失，本测试应失败——这是有意的防回归锚点。
"""

import csv
import io
import subprocess
import tarfile
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[1]
PINNED_BASELINE_COMMIT = "0f7d445"

EXPECTED_COUNTS = {
    "sources.csv": (1165, 1177),
    "fact_evidences.csv": (608, 628),
    "events.csv": (150, 148),
    "event_participants.csv": (228, 224),
}

DRAFT_FILES = [
    "research/drafts/reports/phase2_batch3_event_sources.csv",
    "research/drafts/reports/phase2_batch3_event_evidences.csv",
]


def _materialize_commit(ref: str, dest: Path) -> None:
    """从指定提交原样取出 data/processed 与第三批草稿到临时目录。"""
    result = subprocess.run(
        ["git", "archive", "--format=tar", ref,
         "data/processed",
         "research/drafts/reports/phase2_batch3_event_sources.csv",
         "research/drafts/reports/phase2_batch3_event_evidences.csv"],
        cwd=REPO_ROOT,
        check=True,
        capture_output=True,
    )
    with tarfile.open(fileobj=io.BytesIO(result.stdout)) as tar:
        tar.extractall(dest, filter="data")


def _row_count(path: Path) -> int:
    with open(path, encoding="utf-8-sig", newline="") as fh:
        return sum(1 for _ in csv.DictReader(fh))


@pytest.fixture()
def baseline_dir(tmp_path: Path) -> Path:
    dest = tmp_path / "baseline"
    _materialize_commit(PINNED_BASELINE_COMMIT, dest)
    return dest


def test_merge_on_pinned_premerge_commit(baseline_dir: Path, tmp_path: Path, monkeypatch, capsys):
    import importlib.util

    spec = importlib.util.spec_from_file_location(
        "merge_batch3_event_review_real",
        REPO_ROOT / "research" / "analysis" / "merge_batch3_event_review.py",
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)

    monkeypatch.setattr(module, "DATA", baseline_dir / "data" / "processed")
    monkeypatch.setattr(
        module, "DRAFTS", baseline_dir / "research" / "drafts" / "reports"
    )

    module.main()
    captured = capsys.readouterr().out
    assert "无新增" not in captured

    data_dir = baseline_dir / "data" / "processed"
    for name, (_before, after) in EXPECTED_COUNTS.items():
        assert _row_count(data_dir / name) == after, f"{name}: 期望 {after}"

    with open(data_dir / "events.csv", encoding="utf-8-sig", newline="") as fh:
        events = {r["event_id"] for r in csv.DictReader(fh)}
    assert "EVT-00007" not in events and "EVT-00119" not in events

    from kb_schema import validate_data_dir

    result = validate_data_dir(data_dir)
    assert len(result.errors) == 0, result.errors[:3]
    assert len(result.warnings) <= 13, [str(w) for w in result.warnings]

    # 幂等：同一基线上二跑必须早退且零写入。
    snapshots = {
        name: (data_dir / name).read_bytes() for name in EXPECTED_COUNTS
    }
    module.main()
    assert "无新增" in capsys.readouterr().out
    for name, blob in snapshots.items():
        assert (data_dir / name).read_bytes() == blob, f"{name} 二跑发生变化"
