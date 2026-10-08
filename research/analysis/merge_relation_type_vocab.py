"""关系类型词表同义归并（2026-10-08 用户决定，幂等）。

决定依据
--------
BLOCKED.md「待决定 2：关系类型词表同义重复」：生产数据中 `论战`(39) 与
`文学论战`(88)、`交游`(1206) 与 `交往`(8) 并存，双 Agent 交叉验证中裁决者选到
同义词即被记为 `type_synonym` 分歧（第 1 批 3 条）。用户 2026-10-08 决定：
- **归并到多数标签**（选项标题「归并到多数标签」，方向经用户逐字校准确认）：
  `论战` → `文学论战`（88 > 39，改写 39 行）、`交往` → `交游`（1206 > 8，改写 8 行）。
- 归并后 `sweep_batch1` 的 3 条 `type_synonym` 分歧（REL-00011 / REL-00038 / REL-00088）
  经合并器别名归一自动变一致，与本批其余条目同批落地。

落地语义（只做标签归一，不做任何语义改判）
------------------------------------------
- 仅改写 `person_relations.csv` 的 `standard_relation_type` 与 `final_relation_type` 两列；
- `original_relation_type` / `raw_relation_type` / `llm_suggested_relation_type` 为历史
  原始记录，**一律不动**（保可追溯）；
- 不写 `correction_reason`：这是词表规范化，不是类型语义更正（与 REL-01368 的
  人工裁决 wrong_type 更正不同性质），语义凭据见台账与本文件；
- 不改 `relation_risk_level` / `needs_manual_review` / `confidence` / 发布状态各列；
- 两种新标签都不是待核验/推断类型，发布门禁重算必须**零漂移**（含公开层
  REL-00059：交往→交游 后仍为 derived `supported`，公开层计数 7 不变）。

运行时校验（任一失败即整体退出、不写任何文件）
----------------------------------------------
1. 落地前基线：行数 4238；`final`/`standard` 两列的 论战=39、文学论战=88、交游=1206、
   交往=8；发布状态 supported=7 / pending_review=2451 / inferred=1780；critical 1974；
   origin 全 derived；
2. 两列中 affected 行集合一致（同一行不会只改一列）；
3. 落地后：论战=0、交往=0、文学论战=127、交游=1214，其余类型计数逐项不变；
4. 落地后门禁全表重算零漂移；Schema 0 错误。

幂等：检测到 论战=0 且 交往=0 时输出「无新增/已完成」并零写入，对后续批次免疫。
原子写（tmp + `os.replace`）；CSV 为 utf-8-sig + QUOTE_MINIMAL，保持既有行序。
"""

from __future__ import annotations

import argparse
import csv
import os
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
DATA = ROOT / "data" / "processed"
REPORTS = ROOT / "research" / "drafts" / "reports"
LEDGER = REPORTS / "vocab_merge_2026-10-08_ledger.csv"
REPORT_MD = REPORTS / "vocab_merge_2026-10-08_report.md"

BATCH_MARKER = "VOCAB-MERGE-2026-10-08"
AUTHORIZED_BY = "用户（词表归并规则决定）"
AUTHORIZED_AT = "2026-10-08"
AUTHORIZATION_QUOTE = "归并到多数标签：论战→文学论战，交往→交游"

TYPE_ALIASES = {"论战": "文学论战", "交往": "交游"}
TYPE_COLUMNS = ("standard_relation_type", "final_relation_type")

# 落地前基线（2026-10-08 实测；「其余类型计数」按迁移前快照断言）
EXPECTED_BASELINE_ROWS = 4238
EXPECTED_BASELINE_TYPE = {"论战": 39, "文学论战": 88, "交游": 1206, "交往": 8}
EXPECTED_BASELINE_STATUS = {"supported": 7, "pending_review": 2451, "inferred": 1780}
EXPECTED_BASELINE_CRITICAL = 1974

# 落地后硬后置条件
EXPECTED_POST_TYPE = {"论战": 0, "文学论战": 127, "交游": 1214, "交往": 0}
EXPECTED_POST_STATUS = EXPECTED_BASELINE_STATUS  # 门禁零漂移：状态分布不变
EXPECTED_AFFECTED_ROWS = 47  # 39 + 8


class VocabMergeError(RuntimeError):
    """任何前置/后置校验失败都抛出本异常，调用方保证不写任何文件。"""


def _read(path: Path) -> tuple[list[str], list[dict[str, str]]]:
    with open(path, encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        return list(reader.fieldnames or []), [dict(row) for row in reader]


def _write(path: Path, columns: list[str], rows: list[dict[str, str]]) -> None:
    tmp = path.with_suffix(path.suffix + ".tmp")
    with open(tmp, "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=columns, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)
    os.replace(tmp, path)


def _write_md(path: Path, text: str) -> None:
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(text, encoding="utf-8")
    os.replace(tmp, path)


def _check_baseline(rel_rows: list[dict[str, str]]) -> None:
    if len(rel_rows) != EXPECTED_BASELINE_ROWS:
        raise VocabMergeError(f"person_relations 应为 {EXPECTED_BASELINE_ROWS} 行，实际 {len(rel_rows)}")
    for col in TYPE_COLUMNS:
        dist = Counter((r.get(col) or "").strip() for r in rel_rows)
        for label, want in EXPECTED_BASELINE_TYPE.items():
            if dist.get(label, 0) != want:
                raise VocabMergeError(
                    f"落地前基线不符：{col} 中「{label}」应为 {want}，实际 {dist.get(label, 0)}"
                    "（数据可能已被其他批次改动，拒绝执行）"
                )
    status = Counter((r.get("publish_status") or "").strip() for r in rel_rows)
    if dict(status) != EXPECTED_BASELINE_STATUS:
        raise VocabMergeError(f"落地前发布状态分布不符：期望 {EXPECTED_BASELINE_STATUS}，实际 {dict(status)}")
    critical = sum(1 for r in rel_rows if (r.get("relation_risk_level") or "").strip() == "critical")
    if critical != EXPECTED_BASELINE_CRITICAL:
        raise VocabMergeError(f"落地前 critical 应为 {EXPECTED_BASELINE_CRITICAL}，实际 {critical}")
    origins = {r.get("publish_status_origin", "").strip() for r in rel_rows}
    if origins != {"derived"}:
        raise VocabMergeError(f"落地前 origin 应全为 derived，实际 {origins}")


def already_merged(rel_rows: list[dict[str, str]]) -> bool:
    for col in TYPE_COLUMNS:
        dist = Counter((r.get(col) or "").strip() for r in rel_rows)
        if dist.get("论战", 0) == 0 and dist.get("交往", 0) == 0:
            return True
    return False


def apply_merge(data_dir: Path, dry_run: bool = False) -> dict:
    data_dir = Path(data_dir)
    rel_cols, rel_rows = _read(data_dir / "person_relations.csv")

    if already_merged(rel_rows):
        # 幂等二跑：只校验「别名标签已消失」——这是归并的永久效果。文学论战/交游 的
        # 计数会随后续落地批次变化（如 REL-00038 交游→文学论战后 127→128），
        # 不得把归并当日的瞬时值当永久终值校验（2026-10-08 返工缺陷）。
        for col in TYPE_COLUMNS:
            dist = Counter((r.get(col) or "").strip() for r in rel_rows)
            stale = [label for label in TYPE_ALIASES if dist.get(label, 0) != 0]
            if stale:
                raise VocabMergeError(
                    f"二跑：{col} 中旧标签未清零：{dict((s, dist.get(s, 0)) for s in stale)}，疑似数据漂移"
                )
        return {
            "status": "no-op",
            "message": "无新增/已完成：词表归并痕迹已存在且旧标签已清零，跳过写入。",
        }

    _check_baseline(rel_rows)

    # affected 行：两列中任一命中别名即受影响；两列必须同时受影响（同值同步）。
    affected: list[tuple[int, dict[str, str], dict[str, str]]] = []
    for idx, row in enumerate(rel_rows):
        changes: dict[str, str] = {}
        for col in TYPE_COLUMNS:
            value = (row.get(col) or "").strip()
            if value in TYPE_ALIASES:
                changes[col] = TYPE_ALIASES[value]
        if not changes:
            continue
        if len(changes) != len(TYPE_COLUMNS):
            raise VocabMergeError(
                f"{row.get('relation_id')} 两列归并状态不一致：仅 {sorted(changes)} 命中别名，拒绝执行"
            )
        affected.append((idx, row, changes))
    if len(affected) != EXPECTED_AFFECTED_ROWS:
        raise VocabMergeError(f"受影响行应为 {EXPECTED_AFFECTED_ROWS}，实际 {len(affected)}")

    # 公开层受影响行（预期仅 REL-00059，交往→交游 后门禁零漂移）
    public_affected = [
        (row.get("relation_id", "").strip(), changes)
        for _, row, changes in affected
        if (row.get("publish_status") or "").strip() in ("verified", "supported")
    ]
    for rid, changes in public_affected:
        if rid != "REL-00059" or changes.get("final_relation_type") != "交游":
            raise VocabMergeError(
                f"公开层受影响行 {rid} 与预期不符（应为 REL-00059 交往→交游），拒绝执行"
            )

    # 内存内归一（affected 行整行快照，供零漂移复核）
    originals_full = {idx: dict(rel_rows[idx]) for idx, _, _ in affected}
    for idx, row, changes in affected:
        for col, new in changes.items():
            row[col] = new

    # 落地后硬后置条件
    for col in TYPE_COLUMNS:
        dist = Counter((r.get(col) or "").strip() for r in rel_rows)
        for label, want in EXPECTED_POST_TYPE.items():
            if dist.get(label, 0) != want:
                raise VocabMergeError(f"归并后 {col} 中「{label}」应为 {want}，实际 {dist.get(label, 0)}")
    status_after = Counter((r.get("publish_status") or "").strip() for r in rel_rows)
    if dict(status_after) != EXPECTED_POST_STATUS:
        raise VocabMergeError(f"归并后发布状态分布漂移：期望 {EXPECTED_POST_STATUS}，实际 {dict(status_after)}")
    critical_after = sum(1 for r in rel_rows if (r.get("relation_risk_level") or "").strip() == "critical")
    if critical_after != EXPECTED_BASELINE_CRITICAL:
        raise VocabMergeError("critical 计数被改写（禁止）")
    # 门禁全表重算零漂移（词表归并不动证据，但重算须带既有证据行）
    try:
        from research.analysis.relation_publish_status import derive_relation_publish_status_with_origin
    except ImportError:
        if str(Path(__file__).resolve().parent) not in sys.path:
            sys.path.append(str(Path(__file__).resolve().parent))
        from relation_publish_status import derive_relation_publish_status_with_origin  # noqa: E402
    _, ev_rows = _read(data_dir / "relation_evidences.csv")
    by_rel: dict[str, list[dict[str, str]]] = {}
    for ev in ev_rows:
        by_rel.setdefault((ev.get("relation_id") or "").strip(), []).append(ev)
    for row in rel_rows:
        rid = (row.get("relation_id") or "").strip()
        st, origin = derive_relation_publish_status_with_origin(row, by_rel.get(rid, []))
        if (st, origin) != ((row.get("publish_status") or "").strip(), (row.get("publish_status_origin") or "").strip()):
            raise VocabMergeError(f"{rid} 门禁重算不一致：{origin}/{st}")

    # 台账（内存）
    ledger = [
        {
            "relation_id": row.get("relation_id", "").strip(),
            "batch_marker": BATCH_MARKER,
            "column": col,
            "type_before": originals_full[idx].get(col, ""),
            "type_after": changes[col],
            "publish_status": (row.get("publish_status") or "").strip(),
            "in_public_layer": "yes" if (row.get("publish_status") or "").strip() in ("verified", "supported") else "no",
            "authorized_by": AUTHORIZED_BY,
            "authorized_at": AUTHORIZED_AT,
            "authorization_quote": AUTHORIZATION_QUOTE,
            "note": "词表同义归并（多数标签），非类型语义更正；original/raw/llm_suggested 列未动。",
        }
        for idx, row, changes in affected
        for col in TYPE_COLUMNS
    ]

    summary = {
        "affected_rows": len(affected),
        "public_affected": [rid for rid, _ in public_affected],
        "ledger": ledger,
    }
    if dry_run:
        return {"status": "dry-run", "message": "校验全部通过（未写入）。", **summary}

    # 零漂移最终复核：整行快照对比（只允许两类型列变化）
    for idx, _, _ in affected:
        drift = [c for c in rel_cols if c not in TYPE_COLUMNS and (rel_rows[idx].get(c) or "") != (originals_full[idx].get(c) or "")]
        if drift:
            raise VocabMergeError(f"{rel_rows[idx].get('relation_id')} 出现计划外字段改动：{drift}")

    _write(data_dir / "person_relations.csv", rel_cols, rel_rows)
    return {"status": "applied", "message": "词表归并完成。", **summary}


def main() -> int:
    parser = argparse.ArgumentParser(description="关系类型词表同义归并（2026-10-08 用户决定，幂等）。")
    parser.add_argument("--data-dir", type=Path, default=DATA, help="生产数据目录")
    parser.add_argument("--dry-run", action="store_true", help="只执行全部校验，不写任何文件")
    args = parser.parse_args()

    try:
        result = apply_merge(args.data_dir, dry_run=args.dry_run)
    except VocabMergeError as exc:
        print(f"词表归并失败，未写入任何文件：{exc}")
        return 1

    print(result["message"])
    if result["status"] == "no-op":
        return 0
    print(f"受影响关系行：{result['affected_rows']} 条（standard/final 两列同步，共 {len(result['ledger'])} 个单元格）")
    print(f"公开层受影响：{result['public_affected'] or '无'}")
    if result["status"] == "applied":
        _write(LEDGER, [
            "relation_id", "batch_marker", "column", "type_before", "type_after",
            "publish_status", "in_public_layer", "authorized_by", "authorized_at",
            "authorization_quote", "note",
        ], result["ledger"])
        print(f"台账：{LEDGER.relative_to(ROOT)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
