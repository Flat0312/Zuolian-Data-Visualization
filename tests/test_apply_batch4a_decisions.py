"""第四批A五项人工批准裁决的生产落地回归测试。

以固定提交 30d5f77（执行前最后基线）为真实旧基线，`git archive` 原样取出
data/processed，在隔离沙盒中运行 apply_batch4a_decisions 后验证：

- 五项决定的字段终值与证据状态；
- EVT-00005 及其 2 条参与者零残留、无悬空引用；
- 四表增量 sources 0 / fact_evidences −2 / events −1 / event_participants −2；
- 二跑字节零变化（幂等）；
- rejected 不进入覆盖率三口径与发布层；
- Schema 0 错误、警告 ≤13。

提交被历史重写或丢失时测试失败——有意防回归锚点。
"""

import csv
import io
import subprocess
import tarfile
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[1]
PINNED_PRE_APPLY_COMMIT = "30d5f77"

EXPECTED_COUNTS = {
    "sources.csv": (1177, 1177),
    "fact_evidences.csv": (628, 626),
    "events.csv": (148, 147),
    "event_participants.csv": (224, 222),
}

POST_EVT_00004 = {
    "event_name": "鲁迅与柔石赴北四川路看屋未成",
    "canonical_event_key": "EVT-00004|鲁迅与柔石赴北四川路看屋未成|1930-03-28",
    "needs_manual_review": "no",
    "source_ids": "SRC-1122;SRC-1143;SRC-1144;SRC-1167",
}
POST_EVT_00029 = {
    "event_name": "丁玲就读平民女校",
    "canonical_event_key": "EVT-00029|丁玲就读平民女校|1922",
    "source_ids": "SRC-1126;SRC-1154",
}


def _materialize_commit(ref: str, dest: Path) -> None:
    result = subprocess.run(
        ["git", "archive", "--format=tar", ref, "data/processed"],
        cwd=REPO_ROOT,
        check=True,
        capture_output=True,
    )
    with tarfile.open(fileobj=io.BytesIO(result.stdout)) as tar:
        tar.extractall(dest, filter="data")


def _read_csv(path: Path) -> list[dict[str, str]]:
    with open(path, encoding="utf-8-sig", newline="") as fh:
        return list(csv.DictReader(fh))


def _run_apply(data_dir: Path):
    import importlib.util

    spec = importlib.util.spec_from_file_location(
        "apply_batch4a_decisions_module",
        REPO_ROOT / "research" / "analysis" / "apply_batch4a_decisions.py",
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    module.DATA = data_dir
    out = module.main()
    return module, (out or "")


@pytest.fixture()
def applied_dir(tmp_path: Path) -> Path:
    dest = tmp_path / "baseline"
    _materialize_commit(PINNED_PRE_APPLY_COMMIT, dest)
    data_dir = dest / "data" / "processed"
    _, message = _run_apply(data_dir)
    assert "无新增" not in message, "首次执行不得早退"
    return data_dir


def test_five_decisions_final_field_values(applied_dir: Path) -> None:
    facts = {r["evidence_id"]: r for r in _read_csv(applied_dir / "fact_evidences.csv")}
    events = {r["event_id"]: r for r in _read_csv(applied_dir / "events.csv")}

    # 决定1：EVT-00001 五条转正、两条拒绝、一条降线索。
    assert facts["FE-EVI-3B84F7AC63"]["review_status"] == "reviewed"
    assert facts["FE-EVI-489692805D"]["review_status"] == "reviewed"
    assert facts["FE-EVI-58338B3D55"]["review_status"] == "reviewed"
    assert facts["FE-EVI-8DB5BD3637"]["review_status"] == "reviewed"
    assert facts["FE-EVI-FD8EBDD93E"]["review_status"] == "reviewed"
    assert facts["FE-EVI-21007B9057"]["review_status"] == "rejected"
    assert facts["FE-EVI-6883382B0E"]["review_status"] == "rejected"
    aud001 = facts["FE-EVI-0528D7CD44"]
    assert aud001["evidence_support"] == "lead" and aud001["review_status"] == "pending"

    # 决定2：EVT-00004 改写 + FE-EVP3-0005 正式证据改指并转正。
    row = events["EVT-00004"]
    for key, value in POST_EVT_00004.items():
        assert row[key] == value, f"EVT-00004.{key}={row[key]!r}"
    e0357 = facts["FE-EVI-0357C10A69"]
    assert e0357["review_status"] == "rejected"
    fe0005 = facts["FE-EVI-7792AD0E80"]
    assert fe0005["subject_id"] == "EVT-00004"
    assert fe0005["evidence_support"] == "support" and fe0005["review_status"] == "reviewed"
    assert fe0005["adjudication_status"] == ""
    # 引文三要素禁改
    baseline_notes = subprocess.run(
        ["git", "show", f"{PINNED_PRE_APPLY_COMMIT}:data/processed/fact_evidences.csv"],
        cwd=REPO_ROOT, check=True, capture_output=True,
    ).stdout.decode("utf-8-sig")
    base_facts = {r["evidence_id"]: r for r in csv.DictReader(io.StringIO(baseline_notes))}
    for field in ("quote", "locator", "source_id"):
        assert fe0005[field] == base_facts[fe0005["evidence_id"]][field], f"{field} 禁改被违反"

    # 决定3：EVT-00005 删除；010/011 行删除；012 改指 EVT-00008。
    ids_in_table = set(facts)
    assert "FE-EVI-04DD852F1C" not in ids_in_table
    assert "FE-EVI-43ECB964FE" not in ids_in_table
    d012 = facts["FE-EVI-D016A6A994"]
    assert d012["subject_id"] == "EVT-00008"
    assert d012["evidence_support"] == "support" and d012["review_status"] == "reviewed"
    assert d012["object_value"] == (
        "柔石等左联五烈士于1931年2月7日夜或2月8日凌晨在上海龙华警备司令部遇害，鲁迅约于2月10日获悉"
    )

    # 决定4：EVT-00017 唯一错挂证据拒绝；事件保留且维持待核标记。
    ev17 = events["EVT-00017"]
    assert ev17["confidence"] == "low" and ev17["needs_manual_review"] == "yes"
    assert facts["FE-EVI-2EC5596E2B"]["review_status"] == "rejected"

    # 决定5：EVT-00029 收窄 + 证据转支持。
    row29 = events["EVT-00029"]
    for key, value in POST_EVT_00029.items():
        assert row29[key] == value, f"EVT-00029.{key}={row29[key]!r}"
    assert "抵沪" not in row29["event_name"] and "抵沪" not in row29["display_note"]
    cda = facts["FE-EVI-CDAFBAFA48"]
    assert cda["evidence_support"] == "support" and cda["review_status"] == "reviewed"

    participants = {r["event_participant_id"]: r for r in _read_csv(applied_dir / "event_participants.csv")}
    assert participants["EVP-00006"]["source_ids"] == "SRC-1122;SRC-1167"
    assert participants["EVP-00007"]["source_ids"] == "SRC-1122;SRC-1167"
    assert participants["EVP-00043"]["source_ids"] == "SRC-1126;SRC-1154"


def test_evt00005_zero_residue_and_no_dangling_refs(applied_dir: Path) -> None:
    for name in ("fact_evidences.csv", "events.csv", "event_participants.csv"):
        text = (applied_dir / name).read_text(encoding="utf-8-sig")
        assert "EVT-00005" not in text, f"{name} 仍残留 EVT-00005"

    facts = _read_csv(applied_dir / "fact_evidences.csv")
    event_ids = {r["event_id"] for r in _read_csv(applied_dir / "events.csv")}
    person_ids = {r["person_id"] for r in _read_csv(applied_dir / "persons.csv")}
    for row in facts:
        if row["subject_type"] == "event":
            assert row["subject_id"] in event_ids, f"悬空事件主体 {row['subject_id']}"
    parts = _read_csv(applied_dir / "event_participants.csv")
    for row in parts:
        assert row["event_id"] in event_ids
        assert row["person_id"] in person_ids


def test_table_deltas_and_second_run_byte_stable(tmp_path: Path, capsys) -> None:
    dest = tmp_path / "idem"
    _materialize_commit(PINNED_PRE_APPLY_COMMIT, dest)
    data_dir = dest / "data" / "processed"

    # 执行前基线行数
    before = {n: len(_read_csv(data_dir / n)) for n in EXPECTED_COUNTS}
    _run_apply(data_dir)
    after = {n: len(_read_csv(data_dir / n)) for n in EXPECTED_COUNTS}

    deltas = {
        "sources.csv": (1177, 1177),
        "fact_evidences.csv": (628, 626),
        "events.csv": (148, 147),
        "event_participants.csv": (224, 222),
    }
    for name in EXPECTED_COUNTS:
        expected_before, expected_after = deltas[name]
        assert before[name] == expected_before and after[name] == expected_after, (
            f"{name}: {before[name]}→{after[name]} 与预期 {expected_before}→{expected_after} 不符"
        )

    snapshots_first = {n: (data_dir / n).read_bytes() for n in EXPECTED_COUNTS}
    _, second_message = _run_apply(data_dir)

    assert "无新增/已完成" in second_message, "二跑必须早退输出『无新增/已完成』"
    for name, blob in snapshots_first.items():
        assert (data_dir / name).read_bytes() == blob, f"{name} 二跑发生变化"


def test_baseline_drift_fails_without_writing(tmp_path: Path) -> None:
    dest = tmp_path / "drift"
    _materialize_commit(PINNED_PRE_APPLY_COMMIT, dest)
    data_dir = dest / "data" / "processed"
    events_path = data_dir / "events.csv"
    events = _read_csv(events_path)
    events[0]["event_date"] = "1930-03-29"
    with open(events_path, "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=list(events[0]))
        writer.writeheader()
        writer.writerows(events)
    before = {name: (data_dir / name).read_bytes() for name in EXPECTED_COUNTS}

    import importlib.util

    spec = importlib.util.spec_from_file_location(
        "apply_batch4a_decisions_drift_module",
        REPO_ROOT / "research" / "analysis" / "apply_batch4a_decisions.py",
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    module.DATA = data_dir
    with pytest.raises(RuntimeError, match="前置校验失败"):
        module.main()
    after = {name: (data_dir / name).read_bytes() for name in EXPECTED_COUNTS}
    assert after == before, "基线异常时不得写入任何目标表"


@pytest.fixture()
def baseline_sandbox(tmp_path: Path) -> Path:
    dest = tmp_path / "idem"
    _materialize_commit(PINNED_PRE_APPLY_COMMIT, dest)
    data_dir = dest / "data" / "processed"
    _run_apply(data_dir)
    return data_dir


def test_rejected_excluded_from_coverage_and_publish_layer(baseline_sandbox: Path, tmp_path: Path) -> None:
    from research.analysis.report_evidence_coverage import report_evidence_coverage

    summary = report_evidence_coverage(
        baseline_sandbox,
        tmp_path / "coverage.md",
        tmp_path / "queue.csv",
    )
    assert summary["event_attached_any"]["covered"] == 26
    assert summary["event_direct_support"]["covered"] == 21
    assert summary["event_confirmed"]["covered"] == 26
    for key in ("event_attached_any", "event_direct_support", "event_confirmed"):
        assert summary[key]["total"] == 147
    # rate 为函数内四舍五入值，与手工四位小数一致即视为同源。
    import math

    assert math.isclose(summary["event_attached_any"]["rate"], 26 / 147, abs_tol=5e-5)
    assert math.isclose(summary["event_direct_support"]["rate"], 21 / 147, abs_tol=5e-5)
    assert math.isclose(summary["event_confirmed"]["rate"], 26 / 147, abs_tol=5e-5)

    # EVT-00017 仅剩 rejected 证据 → 必须重新进入核心缺证队列。
    queue_text = (tmp_path / "queue.csv").read_text(encoding="utf-8-sig")
    assert "EVT-00017" in queue_text

    from research.analysis.build_publish_data import build_publish_data

    publish_dir = tmp_path / "publish_out"
    build_publish_data(baseline_sandbox, publish_dir, tmp_path / "gate.md")
    published_facts = _read_csv(publish_dir / "fact_evidences.csv")
    assert all(r["review_status"] != "rejected" for r in published_facts), "发布层混入 rejected"
    ids = {r["evidence_id"] for r in published_facts}
    for removed in ("FE-EVI-04DD852F1C", "FE-EVI-43ECB964FE"):
        assert removed not in ids
    published_events = {r["event_id"] for r in _read_csv(publish_dir / "events.csv")}
    assert "EVT-00005" not in published_events and "EVT-00004" in published_events


def test_schema_gates_on_executed_baseline(applied_dir: Path) -> None:
    from kb_schema import validate_data_dir

    result = validate_data_dir(applied_dir)
    assert len(result.errors) == 0, result.errors[:5]
    assert len(result.warnings) <= 13, [str(w) for w in result.warnings]


def test_third_batch_production_conflicts_drop_to_eight() -> None:
    """生产层 resolved conflict 应由 9 降为 8（FE-EVP3-0005 转 support）。

    pinned 旧提交独立重放仍应保持 9 的断言位于 test_merge_batch3_real_baseline.py，
    本处只锚定生产层终态，不削弱重放锚点。
    """
    facts_path = REPO_ROOT / "data" / "processed" / "fact_evidences.csv"
    rows = _read_csv(facts_path)
    resolved = [
        r
        for r in rows
        if r.get("adjudication_status") == "resolved_by_event_correction"
    ]
    assert len(resolved) == 8, f"生产层已裁决冲突应为 8，实际 {len(resolved)}"
    b3_conflict_left = [
        r
        for r in rows
        if r.get("origin_evidence_id", "").startswith("FE-EVP3-") and r["evidence_support"] == "conflict"
    ]
    assert all(r["adjudication_status"] == "resolved_by_event_correction" for r in b3_conflict_left)
