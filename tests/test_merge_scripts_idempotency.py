"""合并脚本幂等性隔离测试。

在临时目录中复制生产 CSV 与候选草稿，替换脚本数据路径后连续执行两次：
- 已完整合并的生产副本上两连跑：第二次必须零写入（SHA256 与行数不变）；
- 剥离本批痕迹后的副本上首跑应正常增量，二跑同样零变化；
- 生产数据中 12 条 FE-EVP-* 与 2 条 FE-EVP2-* 的 origin 编号各自唯一。

本测试只读取生产文件用于复制，绝不修改生产 CSV。
"""

import csv
import hashlib
import importlib.util
import shutil
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
PROD_DATA = REPO_ROOT / "data" / "processed"
PROD_DRAFTS = REPO_ROOT / "research" / "drafts" / "reports"
ANALYSIS_DIR = REPO_ROOT / "research" / "analysis"

BATCH1_DATA_FILES = ["sources.csv", "fact_evidences.csv", "events.csv"]
BATCH1_DRAFT_FILES = [
    "phase2_event_evidence_pilot_sources.csv",
    "phase2_event_evidence_pilot.csv",
    "phase1_p1_evidence_supplement.csv",
]
BATCH2_DATA_FILES = ["sources.csv", "fact_evidences.csv", "events.csv"]
BATCH2_DRAFT_FILES = [
    "phase2_batch2_longhua_roster_sources.csv",
    "phase2_batch2_longhua_roster.csv",
]


def _load_module(name: str):
    spec = importlib.util.spec_from_file_location(name, ANALYSIS_DIR / f"{name}.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _row_count(path: Path) -> int:
    with open(path, encoding="utf-8-sig", newline="") as fh:
        return sum(1 for _ in csv.DictReader(fh))


def _read_rows(path: Path) -> list[dict[str, str]]:
    with open(path, encoding="utf-8-sig", newline="") as fh:
        return list(csv.DictReader(fh))


def _write_rows(path: Path, rows: list[dict[str, str]]) -> None:
    if not rows:
        raise AssertionError(f"{path} 意外为空")
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


def _stage(tmp_path: Path, data_files: list[str], draft_files: list[str]) -> tuple[Path, Path]:
    data_dir = tmp_path / "data" / "processed"
    drafts_dir = tmp_path / "research" / "drafts" / "reports"
    data_dir.mkdir(parents=True)
    drafts_dir.mkdir(parents=True)
    for name in data_files:
        shutil.copyfile(PROD_DATA / name, data_dir / name)
    for name in draft_files:
        shutil.copyfile(PROD_DRAFTS / name, drafts_dir / name)
    return data_dir, drafts_dir


def _snapshot(paths: list[Path]) -> dict[str, tuple[int, str]]:
    return {str(p): (_row_count(p), _sha256(p)) for p in paths}


def _assert_unchanged(before: dict[str, tuple[int, str]]) -> None:
    for path_str, (count, digest) in before.items():
        p = Path(path_str)
        assert _row_count(p) == count, f"{p.name} 行数发生变化"
        assert _sha256(p) == digest, f"{p.name} 字节内容发生变化"


def _configure_batch1(monkeypatch, module, data_dir: Path, drafts_dir: Path) -> None:
    monkeypatch.setattr(module, "DATA", data_dir)
    monkeypatch.setattr(module, "DRAFTS", drafts_dir)
    monkeypatch.setattr(module, "PILOT_SOURCES", drafts_dir / "phase2_event_evidence_pilot_sources.csv")
    monkeypatch.setattr(module, "PILOT_EVIDENCES", drafts_dir / "phase2_event_evidence_pilot.csv")
    monkeypatch.setattr(module, "P1_SUPPLEMENT", drafts_dir / "phase1_p1_evidence_supplement.csv")


def _configure_batch2(monkeypatch, module, data_dir: Path, drafts_dir: Path) -> None:
    monkeypatch.setattr(module, "DATA", data_dir)
    monkeypatch.setattr(module, "DRAFTS", drafts_dir)
    monkeypatch.setattr(
        module, "PILOT_SOURCES", drafts_dir / "phase2_batch2_longhua_roster_sources.csv"
    )
    monkeypatch.setattr(module, "PILOT_EVIDENCES", drafts_dir / "phase2_batch2_longhua_roster.csv")


def _strip_batch_traces(data_dir: Path, drafts_dir: Path, origin_prefix: str, pilot_sources_csv: str) -> None:
    """从副本中剥离某一批已合并的来源与证据，模拟该批合并前的状态。"""
    pilot_urls = {
        r["source_url"] for r in _read_rows(drafts_dir / pilot_sources_csv) if r["source_url"]
    }
    evidences = [r for r in _read_rows(data_dir / "fact_evidences.csv")
                 if not r["origin_evidence_id"].startswith(origin_prefix)]
    sources = [r for r in _read_rows(data_dir / "sources.csv") if r["source_url"] not in pilot_urls]
    _write_rows(data_dir / "fact_evidences.csv", evidences)
    _write_rows(data_dir / "sources.csv", sources)


def test_batch1_second_run_on_merged_copy_is_noop(tmp_path, monkeypatch, capsys):
    module = _load_module("merge_event_evidence_pilot")
    data_dir, drafts_dir = _stage(tmp_path, BATCH1_DATA_FILES, BATCH1_DRAFT_FILES)
    _configure_batch1(monkeypatch, module, data_dir, drafts_dir)

    module.main()
    tracked = [
        data_dir / "sources.csv",
        data_dir / "fact_evidences.csv",
        data_dir / "events.csv",
        drafts_dir / "phase1_p1_evidence_supplement.csv",
    ]
    before = _snapshot(tracked)

    module.main()
    captured = capsys.readouterr().out
    assert "无新增" in captured
    _assert_unchanged(before)


def test_batch2_second_run_on_merged_copy_is_noop(tmp_path, monkeypatch, capsys):
    module = _load_module("merge_longhua_roster")
    data_dir, drafts_dir = _stage(tmp_path, BATCH2_DATA_FILES, BATCH2_DRAFT_FILES)
    _configure_batch2(monkeypatch, module, data_dir, drafts_dir)

    module.main()
    tracked = [
        data_dir / "sources.csv",
        data_dir / "fact_evidences.csv",
        data_dir / "events.csv",
    ]
    before = _snapshot(tracked)

    module.main()
    captured = capsys.readouterr().out
    assert "无新增" in captured
    _assert_unchanged(before)


def test_batch1_fresh_add_then_second_run_is_noop(tmp_path, monkeypatch, capsys):
    module = _load_module("merge_event_evidence_pilot")
    data_dir, drafts_dir = _stage(tmp_path, BATCH1_DATA_FILES, BATCH1_DRAFT_FILES)
    _strip_batch_traces(
        data_dir, drafts_dir, "FE-EVP-", "phase2_event_evidence_pilot_sources.csv"
    )
    _configure_batch1(monkeypatch, module, data_dir, drafts_dir)

    src_before = _row_count(data_dir / "sources.csv")
    ev_before = _row_count(data_dir / "fact_evidences.csv")
    module.main()
    assert _row_count(data_dir / "sources.csv") == src_before + 10
    assert _row_count(data_dir / "fact_evidences.csv") == ev_before + 12

    p1_rows = _read_rows(drafts_dir / "phase1_p1_evidence_supplement.csv")
    rejected = [r for r in p1_rows if r["evidence_id"] == "FE-SUP-0137"]
    assert len(rejected) == 1 and rejected[0]["review_status"] == "rejected"
    assert rejected[0]["reviewer_note"].count("废弃不合并") == 1

    tracked = [
        data_dir / "sources.csv",
        data_dir / "fact_evidences.csv",
        data_dir / "events.csv",
        drafts_dir / "phase1_p1_evidence_supplement.csv",
    ]
    before = _snapshot(tracked)
    module.main()
    assert "无新增" in capsys.readouterr().out
    _assert_unchanged(before)


def test_batch2_fresh_add_then_second_run_is_noop(tmp_path, monkeypatch, capsys):
    module = _load_module("merge_longhua_roster")
    data_dir, drafts_dir = _stage(tmp_path, BATCH2_DATA_FILES, BATCH2_DRAFT_FILES)
    _strip_batch_traces(
        data_dir, drafts_dir, "FE-EVP2-", "phase2_batch2_longhua_roster_sources.csv"
    )
    _configure_batch2(monkeypatch, module, data_dir, drafts_dir)

    src_before = _row_count(data_dir / "sources.csv")
    ev_before = _row_count(data_dir / "fact_evidences.csv")
    module.main()
    assert _row_count(data_dir / "sources.csv") == src_before + 2
    assert _row_count(data_dir / "fact_evidences.csv") == ev_before + 2

    tracked = [
        data_dir / "sources.csv",
        data_dir / "fact_evidences.csv",
        data_dir / "events.csv",
    ]
    before = _snapshot(tracked)
    module.main()
    assert "无新增" in capsys.readouterr().out
    _assert_unchanged(before)


def test_promoted_origin_ids_unique_in_production():
    rows = _read_rows(PROD_DATA / "fact_evidences.csv")
    batch1 = [r["origin_evidence_id"] for r in rows if r["origin_evidence_id"].startswith("FE-EVP-")]
    batch2 = [r["origin_evidence_id"] for r in rows if r["origin_evidence_id"].startswith("FE-EVP2-")]
    assert len(batch1) == 12 and len(set(batch1)) == 12
    assert len(batch2) == 2 and len(set(batch2)) == 2


def test_batch2_merge_from_old_baseline_schema_clean_and_src1165_referenced(tmp_path, monkeypatch):
    """从旧基线（剥离第二批全部痕迹）合并后：schema 0错误、≤13警告、SRC-1165已被事件引用。"""
    data_dir = tmp_path / "data" / "processed"
    drafts_dir = tmp_path / "research" / "drafts" / "reports"
    shutil.copytree(PROD_DATA, data_dir)
    drafts_dir.mkdir(parents=True)
    for name in BATCH2_DRAFT_FILES:
        shutil.copyfile(PROD_DRAFTS / name, drafts_dir / name)

    pilot_rows = _read_rows(drafts_dir / "phase2_batch2_longhua_roster_sources.csv")
    pilot_urls = {r["source_url"] for r in pilot_rows if r["source_url"]}

    sources = _read_rows(data_dir / "sources.csv")
    dropped_ids = {r["source_id"] for r in sources if r["source_url"] in pilot_urls}
    _write_rows(data_dir / "sources.csv", [r for r in sources if r["source_url"] not in pilot_urls])

    evidences = [
        r for r in _read_rows(data_dir / "fact_evidences.csv")
        if not r["origin_evidence_id"].startswith("FE-EVP2-")
    ]
    _write_rows(data_dir / "fact_evidences.csv", evidences)

    events = _read_rows(data_dir / "events.csv")
    for row in events:
        if row["event_id"] == "EVT-00148":
            row["source_ids"] = ";".join(
                s for s in row["source_ids"].split(";") if s and s not in dropped_ids
            )
    _write_rows(data_dir / "events.csv", events)

    module = _load_module("merge_longhua_roster")
    _configure_batch2(monkeypatch, module, data_dir, drafts_dir)
    module.main()

    from kb_schema import validate_data_dir

    result = validate_data_dir(data_dir)
    assert len(result.errors) == 0, result.errors[:3]
    assert len(result.warnings) <= 13, [str(w) for w in result.warnings]

    evt148 = next(r for r in _read_rows(data_dir / "events.csv") if r["event_id"] == "EVT-00148")
    attached = set(evt148["source_ids"].split(";"))
    assert "SRC-1165" in attached, evt148["source_ids"]
    roster_formal_id = next(
        r["source_id"]
        for r in _read_rows(data_dir / "sources.csv")
        if r["source_url"] in pilot_urls and "名录" in r["title"]
    )
    assert roster_formal_id in attached, evt148["source_ids"]
