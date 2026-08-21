"""第三批事件裁决合并脚本（merge_batch3_event_review）的隔离与幂等测试。

场景与既有两批一致：
- 已完整合并的生产副本上二连跑：第二次必须零写入（行数与 SHA256 不变）；
- 剥离本批证据/来源痕迹后的副本首跑应恢复增量，二跑零变化；
- 物理合并 EVT-00007→EVT-00006、EVT-00119→EVT-00120 后无悬空引用；
- 从剥离基线合并后 schema 保持 0 错误、≤13 警告。

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

BATCH3_DATA_FILES = [
    "sources.csv",
    "fact_evidences.csv",
    "events.csv",
    "event_participants.csv",
]
BATCH3_DRAFT_FILES = [
    "phase2_batch3_event_sources.csv",
    "phase2_batch3_event_evidences.csv",
]

# SRC-EVP3-010 复用生产 SRC-1163（同 URL），剥离基线时不得误删该生产行。
REUSED_SOURCE_URL_SUFFIX = "gdxk.southcn.com/st/ztp/content/post_371048.html"


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


def _stage(tmp_path: Path) -> tuple[Path, Path]:
    data_dir = tmp_path / "data" / "processed"
    drafts_dir = tmp_path / "research" / "drafts" / "reports"
    data_dir.mkdir(parents=True)
    drafts_dir.mkdir(parents=True)
    for name in BATCH3_DATA_FILES:
        shutil.copyfile(PROD_DATA / name, data_dir / name)
    for name in BATCH3_DRAFT_FILES:
        shutil.copyfile(PROD_DRAFTS / name, drafts_dir / name)
    return data_dir, drafts_dir


def _configure(monkeypatch, module, data_dir: Path, drafts_dir: Path) -> None:
    monkeypatch.setattr(module, "DATA", data_dir)
    monkeypatch.setattr(module, "DRAFTS", drafts_dir)


def _strip_batch3_traces(data_dir: Path, drafts_dir: Path) -> tuple[int, int]:
    """剥离本批已合并的证据与来源（保留被复用的 SRC-1163），返回剥离后基数。"""
    draft_sources = _read_rows(drafts_dir / "phase2_batch3_event_sources.csv")
    pilot_urls = {
        r["source_url"]
        for r in draft_sources
        if r["source_url"] and REUSED_SOURCE_URL_SUFFIX not in r["source_url"]
    }
    sources = [
        r
        for r in _read_rows(data_dir / "sources.csv")
        if r["source_url"] not in pilot_urls
    ]
    _write_rows(data_dir / "sources.csv", sources)

    evidences = [
        r
        for r in _read_rows(data_dir / "fact_evidences.csv")
        if not r["origin_evidence_id"].startswith("FE-EVP3-")
    ]
    _write_rows(data_dir / "fact_evidences.csv", evidences)
    return len(sources), len(evidences)


def test_batch3_second_run_on_merged_copy_is_noop(tmp_path, monkeypatch, capsys):
    module = _load_module("merge_batch3_event_review")
    data_dir, drafts_dir = _stage(tmp_path)
    _configure(monkeypatch, module, data_dir, drafts_dir)

    module.main()
    tracked = [
        data_dir / "sources.csv",
        data_dir / "fact_evidences.csv",
        data_dir / "events.csv",
        data_dir / "event_participants.csv",
    ]
    before = _snapshot(tracked)

    module.main()
    captured = capsys.readouterr().out
    assert "无新增" in captured
    _assert_unchanged(before)


def _snapshot(paths: list[Path]) -> dict[str, tuple[int, str]]:
    return {str(p): (_row_count(p), _sha256(p)) for p in paths}


def _assert_unchanged(before: dict[str, tuple[int, str]]) -> None:
    for path_str, (count, digest) in before.items():
        p = Path(path_str)
        assert _row_count(p) == count, f"{p.name} 行数发生变化"
        assert _sha256(p) == digest, f"{p.name} 字节内容发生变化"


def test_batch3_fresh_add_counts_merge_and_remap(tmp_path, monkeypatch, capsys):
    module = _load_module("merge_batch3_event_review")
    data_dir, drafts_dir = _stage(tmp_path)
    src_base, ev_base = _strip_batch3_traces(data_dir, drafts_dir)
    _configure(monkeypatch, module, data_dir, drafts_dir)

    module.main()

    assert _row_count(data_dir / "sources.csv") == src_base + 12
    assert _row_count(data_dir / "fact_evidences.csv") == ev_base + 20

    # 物理删除重复事件及其参与者；保留条参与者完整。
    events = _read_rows(data_dir / "events.csv")
    event_ids = {r["event_id"] for r in events}
    assert "EVT-00007" not in event_ids and "EVT-00119" not in event_ids
    assert len(events) == 148

    parts = _read_rows(data_dir / "event_participants.csv")
    part_events = {r["event_id"] for r in parts}
    assert "EVT-00007" not in part_events and "EVT-00119" not in part_events
    kept6 = [r["person_id"] for r in parts if r["event_id"] == "EVT-00006"]
    kept20 = [r["person_id"] for r in parts if r["event_id"] == "EVT-00120"]
    assert sorted(kept6) == ["ZLH-001", "ZLH-016"]
    assert sorted(kept20) == ["ZLH-014", "ZLH-034"]

    # 被删事件的证据主体改指保留条。
    evidences = _read_rows(data_dir / "fact_evidences.csv")
    subjects = {
        r["origin_evidence_id"]: r["subject_id"]
        for r in evidences
        if r["origin_evidence_id"].startswith("FE-EVP3-")
    }
    assert subjects["FE-EVP3-0006"] == "EVT-00006"
    assert subjects["FE-EVP3-0015"] == "EVT-00120"
    origins = [r["origin_evidence_id"] for r in evidences if r["origin_evidence_id"].startswith("FE-EVP3-")]
    assert len(origins) == 20 and len(set(origins)) == 20

    by_id = {r["event_id"]: r for r in events}
    assert by_id["EVT-00258"]["event_date"] == "1931-09-20"
    assert by_id["EVT-00236"]["event_date"] == "1925-10-11"
    assert by_id["EVT-00236"]["date_precision"] == "日"
    assert by_id["EVT-00020"]["event_date"] == "1929-07-06"
    assert by_id["EVT-00257"]["event_name"] == "东方旅社秘密会议"
    assert by_id["EVT-00120"]["event_name"] == "革命文学论争"
    assert by_id["EVT-00120"]["event_date"] == "1928"
    assert by_id["EVT-00187"]["event_date"] == "1931-01-16"
    assert by_id["EVT-00187"]["historical_location"].startswith("静安寺路洛阳书店")
    assert by_id["EVT-00187"]["longitude"] == "" and by_id["EVT-00187"]["latitude"] == ""
    assert by_id["EVT-00188"]["event_date"] == "1931-05"
    assert by_id["EVT-00188"]["date_precision"] == "月"
    assert by_id["EVT-00261"]["confidence"] == "high"

    tracked = [
        data_dir / "sources.csv",
        data_dir / "fact_evidences.csv",
        data_dir / "events.csv",
        data_dir / "event_participants.csv",
    ]
    before = _snapshot(tracked)
    module.main()
    assert "无新增" in capsys.readouterr().out
    _assert_unchanged(before)


def test_batch3_merge_schema_clean_no_dangling(tmp_path, monkeypatch):
    """从剥离基线合并后：schema 0 错误、≤13 警告、无悬空 fact 主体。"""
    data_dir = tmp_path / "data" / "processed"
    drafts_dir = tmp_path / "research" / "drafts" / "reports"
    shutil.copytree(PROD_DATA, data_dir)
    drafts_dir.mkdir(parents=True)
    for name in BATCH3_DRAFT_FILES:
        shutil.copyfile(PROD_DRAFTS / name, drafts_dir / name)
    _strip_batch3_traces(data_dir, drafts_dir)
    module = _load_module("merge_batch3_event_review")
    _configure(monkeypatch, module, data_dir, drafts_dir)
    module.main()

    from kb_schema import validate_data_dir

    result = validate_data_dir(data_dir)
    assert len(result.errors) == 0, result.errors[:3]
    assert len(result.warnings) <= 13, [str(w) for w in result.warnings]

    event_ids = {r["event_id"] for r in _read_rows(data_dir / "events.csv")}
    facts = _read_rows(data_dir / "fact_evidences.csv")
    dangling = [
        r["evidence_id"]
        for r in facts
        if r["subject_type"] == "event" and r["subject_id"] not in event_ids
    ]
    assert dangling == []


def test_promoted_batch3_origin_ids_unique_in_production():
    rows = _read_rows(PROD_DATA / "fact_evidences.csv")
    batch3 = [r["origin_evidence_id"] for r in rows if r["origin_evidence_id"].startswith("FE-EVP3-")]
    assert len(batch3) == 20 and len(set(batch3)) == 20
