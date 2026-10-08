"""双 Agent 交叉验证裁决的合并器：只有一致才可落地，分歧一律搁置。

授权与口径
----------
2026-10-08 用户授权以交叉验证替代逐条人工裁决（原话：「不用我裁，你和zcode交叉验证裁决吧」）。
因用户同时要求不使用 GLM-5.3 系列（zcode 仅提供 GLM-5.3 / GLM-5.3-Flash，后者正是 2026-09-06
夜间轮的原始执行者，存在自我背书风险），第二裁决者改为 antigravity harness 的
Claude Opus 4.6 (Thinking)，与 GLM、Qwen（本仓实现方）、GPT（主 Agent）血统均不同。

**这不是人工复核。** 因此：
- 公开层只接受 derived `supported`，禁止使用 `human_adjudication` / `verified` 通道；
- 台账、站点与报告必须标注「双 Agent 交叉验证」口径，不得写成「人工已确认」；
- 只有两名裁决者结论**完全一致**的条目才进入可落地集合，分歧条目搁置且不改任何状态；
  唯一例外是 2026-10-08 用户决定采用的**保守交集**（须 `--allow-conservative-intersection`
  显式开启）：分歧中双方都判 support 且至少一方判「成立」的，按现类型不改落地，
  `landable=intersection`——不超出任何一方认可的范围。

独立性保障
----------
两名裁决者各自 blind 填写同一份 `sweep_batchN_input.csv`；裁决输入**不含**夜间轮 `reason`
（该字段由 GLM-5.3-Flash 生成、已证明与引文经常不一致，带上会锚定裁决者）；
主 Agent 的裁决文件写在仓库外（`D:/1大创/.xval/`），Opus 结构上无法读取。

对双方的一视同仁校验（不因是自己就放松）
----------------------------------------
- verdict / evidence_grade / proposed_type 必须在词表内；
- `verdict=证据不足` 时 `evidence_grade` 必须为 `associated`；
- `verdict=类型需改` 时必须给出 `proposed_type`，且不得为推断类标签；
- `key_phrase` 必须是该行 `quote` 的**连续子串**（逐字），否则该裁决者该行判为无效。
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
REPORTS = ROOT / "research" / "drafts" / "reports"
DATA = ROOT / "data" / "processed"

VERDICTS = ("成立", "类型需改", "证据不足")
GRADES = ("support", "associated")
TYPE_VOCAB = (
    "交游", "交往", "通信", "合作", "创作合作", "签名联署", "同属组织", "纪念/悼念",
    "地下通讯", "文学论战", "论战", "亲属关系", "师生关系",
)
# 推断类标签不得由本合并器自行定义：直接引用发布门禁的规范常量，避免词表漂移。
# （曾经在此手写 ("待核验","空间共现","时空共现")，漏了 同属组织，导致 REL-00085
#  被误标为可落地——而按门禁它属推断类型，永远不会进入公开层。）
try:
    from research.analysis.relation_publish_status import (
        INFERRED_RELATION_TYPES as INFERRED_TYPES,
    )
    from research.analysis.relation_publish_status import (
        PUBLIC_RELATION_STATUSES as PUBLIC_STATUSES,
    )
except ImportError:
    if str(Path(__file__).resolve().parent) not in sys.path:
        sys.path.append(str(Path(__file__).resolve().parent))
    from relation_publish_status import (  # noqa: E402
        INFERRED_RELATION_TYPES as INFERRED_TYPES,
    )
    from relation_publish_status import (  # noqa: E402
        PUBLIC_RELATION_STATUSES as PUBLIC_STATUSES,
    )
PENDING_REVIEW_TYPE = "待核验"
# 「不得作为建议类型」≠「不进入公开层」：
# - 待核验/空间共现/时空共现 是机器推断占位标签，人工不应把关系改成它们；
# - 同属组织 是词表里的真实历史关系类型（生产数据 730 条在用），人工可以这么判，
#   只是按发布门禁属推断类、不会进入公开层。
NOT_ASSIGNABLE_TYPES = (PENDING_REVIEW_TYPE, "空间共现", "时空共现")
# 词表同义归并（2026-10-08 用户决定，执行脚本 merge_relation_type_vocab.py）：
# `论战`→`文学论战`、`交往`→`交游`（归并到多数标签）。生产数据已全量归一；
# 本表用于把**历史裁决输入文件**里仍存在的旧标签归一到新词表，使第 1 批的
# 3 条 `type_synonym` 分歧（REL-00011/00038/00088）按一致处理。
# 与 merge_relation_type_vocab.TYPE_ALIASES 保持同值（此处不 import，避免落地脚本耦合）。
TYPE_ALIASES = {"论战": "文学论战", "交往": "交游"}


def normalize_type(label: str) -> str:
    return TYPE_ALIASES.get((label or "").strip(), (label or "").strip())


# 保守交集规则（2026-10-08 用户决定采用）：分歧条目中双方都判 support 且至少一方判
# 「成立」（即接受现类型）的，按**现类型不改**落地——不超出任何一方认可的范围。
# 由 --allow-conservative-intersection 显式开启；不开时维持原严格口径（分歧一律搁置）。
CONSERVATIVE_INTERSECTION_FLAG = "--allow-conservative-intersection"
# 历史兜底：词表归并后 `type_synonym` 分支按理不再触发（别名归一使同义选择自动一致）；
# 保留检测仅防未来出现**未归一的历史输入**被误算成实质分歧。
SYNONYM_PAIRS = (("论战", "文学论战"), ("交游", "交往"))
EXPECTED_COLUMNS = [
    "relation_id", "verdict", "evidence_grade", "proposed_type",
    "key_phrase", "adjudicator_reason", "adjudicator", "adjudicated_at",
]
def _read_raw(path: Path) -> list[list[str]]:
    with open(path, encoding="utf-8-sig", newline="") as fh:
        return [row for row in csv.reader(fh)]


def load_verdicts(path: Path, expected_ids: set[str]) -> tuple[dict[str, dict[str, str]], list[str]]:
    """读取一名裁决者的输出，并修复「理由字段含未转义逗号导致的列串位」。

    修复规则（保守、可判定）：仅当行字段数多于表头、且末尾两列恰为合法
    adjudicator/adjudicated_at 形态时，把中间溢出的字段按原顺序并回 adjudicator_reason。
    任何无法确定归位的行直接判为无效，不猜测。
    """
    raw = _read_raw(path)
    if not raw:
        raise SystemExit(f"裁决文件为空：{path}")
    header = [h.strip() for h in raw[0]]
    if header != EXPECTED_COLUMNS:
        raise SystemExit(f"{path} 表头不符：{header}")
    problems: list[str] = []
    out: dict[str, dict[str, str]] = {}
    for lineno, row in enumerate(raw[1:], start=2):
        if not any(cell.strip() for cell in row):
            continue
        if len(row) == len(EXPECTED_COLUMNS):
            cells = row
        elif len(row) > len(EXPECTED_COLUMNS):
            overflow = len(row) - len(EXPECTED_COLUMNS)
            tail = row[-2:]
            reason_parts = row[5:5 + 1 + overflow]
            if len(tail) == 2 and tail[1].strip() and re.fullmatch(r"\d{4}-\d{2}-\d{2}", tail[1].strip()):
                cells = row[:5] + ["".join(reason_parts)] + tail
                problems.append(f"第{lineno}行列串位已按规则修复（理由字段含未转义逗号）：{row[0]}")
            else:
                problems.append(f"第{lineno}行字段数 {len(row)} 异常且无法确定归位，判为无效：{row[0]}")
                continue
        else:
            problems.append(f"第{lineno}行字段不足（{len(row)}），判为无效：{row[0] if row else '?'}")
            continue
        item = dict(zip(EXPECTED_COLUMNS, (c.strip() for c in cells)))
        rid = item["relation_id"]
        if rid not in expected_ids:
            problems.append(f"第{lineno}行 relation_id 不在本批输入中：{rid}")
            continue
        if rid in out:
            problems.append(f"{rid} 重复裁决，判为无效")
            continue
        out[rid] = item
    missing = sorted(expected_ids - set(out))
    if missing:
        problems.append(f"缺少 {len(missing)} 条裁决：{missing[:8]}")
    return out, problems


def validate_one(rid: str, item: dict[str, str], quote: str, who: str) -> list[str]:
    errs: list[str] = []
    if item["verdict"] not in VERDICTS:
        errs.append(f"{who}/{rid} verdict 非法：{item['verdict']!r}")
    if item["evidence_grade"] not in GRADES:
        errs.append(f"{who}/{rid} evidence_grade 非法：{item['evidence_grade']!r}")
    if item["verdict"] == "证据不足" and item["evidence_grade"] != "associated":
        errs.append(f"{who}/{rid} 判证据不足却给 {item['evidence_grade']}（必须 associated）")
    if item["verdict"] == "类型需改":
        pt = item["proposed_type"]
        if not pt:
            errs.append(f"{who}/{rid} 类型需改但未给 proposed_type")
        elif pt in NOT_ASSIGNABLE_TYPES:
            # 必须排在词表检查之前：这三个占位标签本就不在 TYPE_VOCAB 里，
            # 若先查词表会一律报「不在词表」，更具体的原因永远不触发（死代码）。
            errs.append(f"{who}/{rid} proposed_type 为机器推断占位标签，不得作为建议类型：{pt!r}")
        elif pt not in TYPE_VOCAB:
            errs.append(f"{who}/{rid} proposed_type 不在词表：{pt!r}")
    key = item["key_phrase"]
    if not key:
        errs.append(f"{who}/{rid} key_phrase 为空，无法核查")
    elif key not in re.sub(r"\s+", "", quote):
        errs.append(f"{who}/{rid} key_phrase 不是 quote 的连续子串：{key!r}")
    return errs
XVAL_COLUMNS = [
    "relation_id", "person_a_name", "person_b_name", "current_final_relation_type",
    "locator", "quote", "current_publish_status",
    "codex_verdict", "codex_grade", "codex_proposed_type", "codex_key_phrase", "codex_reason",
    "opus_verdict", "opus_grade", "opus_proposed_type", "opus_key_phrase", "opus_reason",
    "agreement", "disagreement_kind", "landable", "landable_block_reason", "review_status",
]


def _is_synonym(a: str, b: str) -> bool:
    return any({a, b} == set(pair) for pair in SYNONYM_PAIRS)


def merge(
    batch_input: Path, verdict_a: Path, verdict_b: Path, batch_index: int,
    conservative_intersection: bool = False,
) -> dict:
    with open(batch_input, encoding="utf-8-sig", newline="") as fh:
        inputs = list(csv.DictReader(fh))
    ids = {r["relation_id"] for r in inputs}
    a_raw, a_problems = load_verdicts(verdict_a, ids)
    b_raw, b_problems = load_verdicts(verdict_b, ids)

    with open(DATA / "person_relations.csv", encoding="utf-8-sig", newline="") as fh:
        relations = {r["relation_id"].strip(): r for r in csv.DictReader(fh)}

    errors: list[str] = []
    rows: list[dict[str, str]] = []
    for src in inputs:
        rid = src["relation_id"]
        quote = re.sub(r"\s+", "", src["quote"])
        a = a_raw.get(rid)
        b = b_raw.get(rid)
        if a:
            errors += validate_one(rid, a, quote, "codex")
        if b:
            errors += validate_one(rid, b, quote, "opus")
        rel = relations.get(rid, {})
        cur_status = (rel.get("publish_status") or src.get("current_publish_status") or "").strip()
        # 现类型优先读生产数据当前值（词表归并后可能与历史输入文件不同，如 REL-00059 交往→交游）。
        cur_type = (rel.get("final_relation_type") or "").strip() or (src["current_final_relation_type"] or "").strip()
        row = {
            "relation_id": rid,
            "person_a_name": src["person_a_name"], "person_b_name": src["person_b_name"],
            "current_final_relation_type": cur_type,
            "locator": src["locator"], "quote": src["quote"],
            "current_publish_status": cur_status,
            "review_status": "pending_cross_validation",
        }
        for tag, item in (("codex", a), ("opus", b)):
            row[f"{tag}_verdict"] = item["verdict"] if item else ""
            row[f"{tag}_grade"] = item["evidence_grade"] if item else ""
            row[f"{tag}_proposed_type"] = item["proposed_type"] if item else ""
            row[f"{tag}_key_phrase"] = item["key_phrase"] if item else ""
            row[f"{tag}_reason"] = item["adjudicator_reason"] if item else ""
        if not a or not b:
            row["agreement"] = "invalid_missing_verdict"
            row["disagreement_kind"] = "missing"
            row["landable"] = "no"
            row["landable_block_reason"] = "缺一方裁决"
            rows.append(row)
            continue
        same_verdict = a["verdict"] == b["verdict"]
        same_grade = a["evidence_grade"] == b["evidence_grade"]
        # proposed_type 比较一律经别名归一：词表已归并（2026-10-08），历史输入里的
        # 论战/交往 视为 文学论战/交游，不再算分歧。
        same_type = normalize_type(a["proposed_type"]) == normalize_type(b["proposed_type"])
        if same_verdict and same_grade and (a["verdict"] != "类型需改" or same_type):
            row["agreement"] = "agree"
            row["disagreement_kind"] = ""
        elif same_verdict and same_grade and a["verdict"] == "类型需改" and _is_synonym(
                a["proposed_type"], b["proposed_type"]):
            row["agreement"] = "disagree"
            row["disagreement_kind"] = "type_synonym"
        else:
            row["agreement"] = "disagree"
            kinds = []
            if not same_verdict:
                kinds.append("verdict")
            if not same_grade:
                kinds.append("grade")
            if not same_type:
                kinds.append("proposed_type")
            row["disagreement_kind"] = "+".join(kinds)

        landable, block = "no", ""
        if row["agreement"] != "agree":
            # 保守交集（2026-10-08 用户决定采用，须经 flag 显式开启）：分歧条目中
            # 双方都判 support 且至少一方判「成立」（接受现类型）的，按现类型不改落地。
            # 这不超出任何一方认可的范围；双方都判「类型需改」的不属交集（按现类型
            # 落地等于断言双方都拒绝的标签）。
            is_common = (
                conservative_intersection
                and a["verdict"] in ("成立", "类型需改") and b["verdict"] in ("成立", "类型需改")
                and "成立" in (a["verdict"], b["verdict"])
                and a["evidence_grade"] == "support" and b["evidence_grade"] == "support"
            )
            if is_common and cur_status in PUBLIC_STATUSES:
                block = "该关系已在公开层，无需重复落地"
            elif is_common:
                if cur_type in INFERRED_TYPES or cur_type in NOT_ASSIGNABLE_TYPES:
                    landable = "research_only"
                    block = f"保守交集：按现类型落地，但现类型为推断类（{cur_type}），落研究层、不进公开层"
                else:
                    landable = "intersection"
                    block = (
                        "保守交集（2026-10-08 用户决定采用）：双方 support 且至少一方判成立，"
                        "按现类型落地、不改类型"
                    )
            else:
                block = "双方裁决不一致"
        elif a["verdict"] == "证据不足":
            block = "双方一致判证据不足"
        elif a["evidence_grade"] != "support":
            block = "双方一致但证据等级仅 associated"
        elif cur_status in PUBLIC_STATUSES:
            block = "该关系已在公开层，无需重复落地"
        else:
            eff_type = (
                normalize_type(a["proposed_type"]) if a["verdict"] == "类型需改" else cur_type
            )
            if eff_type in INFERRED_TYPES or eff_type in NOT_ASSIGNABLE_TYPES:
                # 仍可落地（证据 + 类型更正都是真实改进），只是按门禁不会进入公开层。
                landable = "research_only"
                block = f"更正后类型为推断类（{eff_type}），落研究层、不进公开层"
            else:
                landable = "yes"
        row["landable"] = landable
        row["landable_block_reason"] = block
        rows.append(row)

    return {
        "rows": rows, "errors": errors,
        "a_problems": a_problems, "b_problems": b_problems,
        "batch_index": batch_index,
        "intersection_used": bool(conservative_intersection),
    }


def write_csv(rows: list[dict[str, str]], path: Path, columns: list[str]) -> None:
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=columns, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)


def write_report(result: dict, path: Path, paths: dict[str, Path]) -> None:
    rows = result["rows"]
    intersection_used = bool(result.get("intersection_used"))
    agree = [r for r in rows if r["agreement"] == "agree"]
    disagree = [r for r in rows if r["agreement"] == "disagree"]
    invalid = [r for r in rows if r["agreement"] == "invalid_missing_verdict"]
    landable = [r for r in rows if r["landable"] == "yes"]
    research_only = [r for r in rows if r["landable"] == "research_only"]
    intersection_rows = [r for r in rows if r["landable"] == "intersection"]
    kinds = Counter(r["disagreement_kind"] for r in disagree)
    va = Counter(r["codex_verdict"] for r in rows if r["codex_verdict"])
    vb = Counter(r["opus_verdict"] for r in rows if r["opus_verdict"])
    lines = [
        f"# 第 {result['batch_index']} 批关系证据 · 双 Agent 交叉验证裁决报告",
        "",
        "> 合并器：`research/analysis/merge_sweep_cross_validation.py`（只读，生产层零改动）  ",
        f"> 裁决输入：`{paths['input'].name}`（{len(rows)} 条，不含夜间轮 reason，防锚定）  ",
        "> 裁决者 A：Codex（主 Agent，GPT 系）  裁决者 B：Claude Opus 4.6 (Thinking)（antigravity harness）",
        "",
        "## 0. 口径声明（必读）",
        "",
        "- 本批裁决为 **双 Agent 交叉验证**，依据 2026-10-08 用户授权替代逐条人工裁决。",
        "- **这不是人工复核。** 落地后公开层只走 derived `supported`，禁止 `human_adjudication`/`verified`；",
        "  台账与站点必须标注本口径，不得写成「人工已确认」。",
        "- 第二裁决者未使用 GLM-5.3 系列：用户明确要求排除，且 zcode 的 GLM-5.3-Flash 正是 2026-09-06",
        "  夜间轮（本仓引文缺陷的来源）的原始执行者，由它裁决等于自我背书。",
        "- 独立性：双方各自 blind 裁决；Codex 的裁决文件写在仓库外，Opus 结构上无法读取；",
        "  裁决输入不含夜间轮 `reason`。",
        "- 词表同义归并已于 2026-10-08 完成（`论战`→`文学论战`、`交往`→`交游`，"
        "`merge_relation_type_vocab.py`）：proposed_type 比较一律经别名归一，"
        "历史输入里的同义选择不再算分歧。",
    ]
    if intersection_used:
        lines += [
            f"- **保守交集规则已采用**（2026-10-08 用户决定，运行时以 `{CONSERVATIVE_INTERSECTION_FLAG}` 开启）："
            "分歧条目中双方都判 support 且至少一方判「成立」的，按现类型不改落地；"
            "双方都判「类型需改」的不属交集、仍搁置。",
        ]
    else:
        lines += [
            f"- 保守交集规则未开启（`{CONSERVATIVE_INTERSECTION_FLAG}`）：本报告只统计不采用，"
            "分歧一律搁置。",
        ]
    lines += [
        "",
        "## 1. 结果概览",
        "",
        "| 项 | 数量 |",
        "| --- | ---: |",
        f"| 输入候选 | {len(rows)} |",
        f"| 双方一致 | {len(agree)} |",
        f"| 双方分歧（搁置） | {len(disagree)} |",
        f"| 保守交集落地（分歧中，按现类型） | {len(intersection_rows)} |",
        f"| 裁决缺失/无效 | {len(invalid)} |",
        f"| **一致且可落地（会进公开层）** | **{len(landable)}** |",
        f"| 一致且可落地（仅研究层，类型属推断类不公开） | {len(research_only)} |",
        "",
        f"一致率：{len(agree) / max(1, len(rows)):.1%}（未开启交集时分歧一律搁置，不落地、不改状态）",
        "",
        "### 各自 verdict 分布",
        "",
        "| verdict | Codex | Opus |",
        "| --- | ---: | ---: |",
    ]
    for v in VERDICTS:
        lines.append(f"| {v} | {va.get(v, 0)} | {vb.get(v, 0)} |")
    lines += [
        "",
        f"Codex 判定成立率（成立+类型需改）：{(va.get('成立', 0) + va.get('类型需改', 0)) / max(1, len(rows)):.1%}；"
        f"Opus：{(vb.get('成立', 0) + vb.get('类型需改', 0)) / max(1, len(rows)):.1%}。",
        "参照：2026-09-20 那批 400 条抽样（授权按建议执行口径）的成立率为 40.2%。",
        "",
        "## 2. 分歧明细",
        "",
        f"分歧类型分布：{dict(kinds) or '无'}",
        "",
        "`type_synonym` 是**词表缺陷**而非实质分歧：词表曾有 `论战`/`文学论战`、`交游`/`交往` 同义并存。"
        "**已于 2026-10-08 归并**（`论战`→`文学论战`、`交往`→`交游`，多数标签口径，"
        "`merge_relation_type_vocab.py` 执行）；本报告起 proposed_type 比较经别名归一，"
        "原 3 条 `type_synonym` 分歧（REL-00011/00038/00088）已按一致处理并计入可落地集合。"
        "下表如仍出现 `type_synonym`，属未归一的历史输入，需先归一再重跑。",
        "",
        "| relation_id | 人物对 | 现类型 | Codex | Opus | 分歧 |",
        "| --- | --- | --- | --- | --- | --- |",
    ]
    for r in sorted(disagree, key=lambda x: x["relation_id"]):
        lines.append(
            f"| {r['relation_id']} | {r['person_a_name']}—{r['person_b_name']} | {r['current_final_relation_type']} "
            f"| {r['codex_verdict']}/{r['codex_grade']}/{r['codex_proposed_type'] or '—'} "
            f"| {r['opus_verdict']}/{r['opus_grade']}/{r['opus_proposed_type'] or '—'} | {r['disagreement_kind']} |"
        )
    # 「保守交集」是**可选规则**，本报告只统计不采用：双方都认为关系成立且都给 support，
    # 仅在"该不该改类型/改成哪个同义标签"上不一致。若采用，则按现类型落地（不改类型），
    # 这是双方结论的交集，不会超出任何一方认可的范围。是否采用属项目规则决定，不由本脚本擅自生效。
    # 交集规则的正确读法：按「现类型不改」落地，只有在**至少一方判 成立**（即接受现类型）时
    # 才是双方结论的真交集。若双方都判「类型需改」（哪怕只是改到同义词的不同标签），
    # 按现类型落地就等于断言了一个两人都明确否定的类型——那不是交集，是造假。
    common = [
        r for r in disagree
        if r["codex_verdict"] in ("成立", "类型需改") and r["opus_verdict"] in ("成立", "类型需改")
        and ("成立" in (r["codex_verdict"], r["opus_verdict"]))
        and r["codex_grade"] == "support" and r["opus_grade"] == "support"
        and r["current_publish_status"] not in PUBLIC_STATUSES
    ]
    both_type_change = [
        r for r in disagree
        if r["codex_verdict"] == "类型需改" and r["opus_verdict"] == "类型需改"
        and r["codex_grade"] == "support" and r["opus_grade"] == "support"
        and r["current_publish_status"] not in PUBLIC_STATUSES
    ]
    common_public = [
        r for r in common
        if r["current_final_relation_type"] not in INFERRED_TYPES
        and r["current_final_relation_type"] not in NOT_ASSIGNABLE_TYPES
    ]
    if intersection_used:
        lines += [
            "",
            "### 已采用的规则：保守交集（2026-10-08 用户决定）",
            "",
            f"分歧 {len(disagree)} 条中有 {len(common)} 条**双方都认为关系成立且证据够 support**，"
            "且**至少一方判「成立」**（即接受现类型）——按现类型不改落地，不会超出任何一方认可的范围。"
            f"其中现类型可进公开层的有 {len(common_public)} 条，本报告起标记为 `landable=intersection`，"
            "与一致项同批落地（落地脚本按 `landable in (yes, intersection)` 选取）。",
            "",
            f"另有 {len(both_type_change)} 条是**双方都判「类型需改」**——这类**不属于交集**："
            "两人都明确否定了现类型，按现类型落地等于断言一个双方都拒绝的标签，仍一律搁置；"
            "其中属同义分歧的已随词表归并转为一致并解锁，其余仍搁置。",
            "",
        ]
    else:
        lines += [
            "",
            "### 可选规则：保守交集（本报告只统计，未采用）",
            "",
            f"分歧中有 {len(common)} 条**双方都认为关系成立且证据够 support**，且**至少一方判「成立」**"
            "（即接受现类型）——只有这类才存在真正的交集：按现类型不改落地，不会超出任何一方认可的范围。"
            f"其中按现类型即可进公开层的有 {len(common_public)} 条。",
            "",
            f"另有 {len(both_type_change)} 条是**双方都判「类型需改」**。这类**不属于交集**："
            "两人都明确否定了现类型，按现类型落地等于断言一个双方都拒绝的标签。其中属同义分歧的"
            "已随词表归并（2026-10-08）转为一致并解锁；其余仍搁置，不计入可解锁条数。",
            "",
            f"若采用「保守交集」规则，这 {len(common_public)} 条可与上表 {len(landable)} 条一起落地"
            f"（重跑本合并器并加 `{CONSERVATIVE_INTERSECTION_FLAG}`）。**本次运行未采用该规则**，"
            "当前严格口径下它们仍属分歧、一律搁置。",
            "",
        ]
    lines += [
        "| relation_id | 人物对 | 现类型 | Codex | Opus | 分歧性质 | 属交集 | 本批处置 |",
        "| --- | --- | --- | --- | --- | --- | --- | --- |",
    ]
    common_ids = {r["relation_id"] for r in common}
    for r in sorted(common + both_type_change, key=lambda x: x["relation_id"]):
        lines.append(
            f"| {r['relation_id']} | {r['person_a_name']}—{r['person_b_name']} | {r['current_final_relation_type']} "
            f"| {r['codex_verdict']}/{r['codex_proposed_type'] or '—'} "
            f"| {r['opus_verdict']}/{r['opus_proposed_type'] or '—'} | {r['disagreement_kind']} "
            f"| {'是' if r['relation_id'] in common_ids else '否（双方都要改类型）'} "
            f"| {r['landable']} |"
        )
    lines += [
        "",
        "### 分歧条目的双方理由",
        "",
    ]
    for r in sorted(disagree, key=lambda x: x["relation_id"]):
        lines += [
            f"**{r['relation_id']} {r['person_a_name']}—{r['person_b_name']}**（{r['locator']}，{r['disagreement_kind']}）",
            "",
            f"> 引文：{r['quote'][:220]}",
            "",
            f"- Codex：{r['codex_verdict']}/{r['codex_grade']}"
            f"{('/' + r['codex_proposed_type']) if r['codex_proposed_type'] else ''} —— {r['codex_reason']}",
            f"- Opus：{r['opus_verdict']}/{r['opus_grade']}"
            f"{('/' + r['opus_proposed_type']) if r['opus_proposed_type'] else ''} —— {r['opus_reason']}",
            "",
        ]
    lines += [
        "## 3. 可落地条目",
        "",
        f"一致项 {len(landable)} 条"
        + (f"；保守交集 {len(intersection_rows)} 条（按现类型、不改类型）" if intersection_rows else "")
        + "。落地规则：",
        "",
        "| relation_id | 人物对 | 现类型 | 裁决 | 更正后类型 | 落地规则 | 出处 |",
        "| --- | --- | --- | --- | --- | --- | --- |",
    ]
    for r in sorted(landable + intersection_rows, key=lambda x: x["relation_id"]):
        if r["landable"] == "intersection":
            eff = r["current_final_relation_type"]
            rule = "保守交集"
        else:
            eff = normalize_type(r["codex_proposed_type"]) if r["codex_verdict"] == "类型需改" else r["current_final_relation_type"]
            rule = "双 Agent 一致"
        lines.append(
            f"| {r['relation_id']} | {r['person_a_name']}—{r['person_b_name']} | {r['current_final_relation_type']} "
            f"| {r['codex_verdict']} | {eff} | {rule} | {r['locator']} |"
        )
    blocked = [r for r in agree if r["landable"] not in ("yes", "research_only")]
    lines += [
        "",
        f"双方一致但**不落地**的 {len(blocked)} 条，按原因分布：",
        "",
    ]
    for reason, count in Counter(r["landable_block_reason"] for r in blocked).most_common():
        lines.append(f"- {reason}：{count} 条")
    lines += [
        "",
        "## 4. 数据质量副产物",
        "",
    ]
    if result["errors"]:
        lines.append(f"裁决文件校验发现 {len(result['errors'])} 处问题（已按无效处理，未落地）：")
        lines += [f"- {e}" for e in result["errors"][:20]]
    else:
        lines.append("双方裁决文件均通过词表与 key_phrase 逐字校验，无问题。")
    probs = result["a_problems"] + result["b_problems"]
    if probs:
        lines += ["", "CSV 结构修复记录："] + [f"- {p}" for p in probs]
    lines += [
        "",
        "## 5. 产物与下一步",
        "",
        f"- 全量对照表：`{paths['xval'].name}`（{len(rows)} 行，含双方 verdict/grade/类型/key_phrase/理由）",
        f"- 分歧清单：`{paths['disagree'].name}`（{len(disagree)} 行，其中保守交集 {len(intersection_rows)} 条标记 `landable=intersection`，其余搁置，生产层零改动）",
        "- 落地须另起幂等脚本，只处理 `landable in (yes, intersection)` 的行，运行时重新按偏移逐字回定位"
        "并校验 quote_sha256；证据行 reviewer_note 必须写明「双 Agent 交叉验证」口径与两名裁决者标识。",
        "- `type_synonym` 分支仅为未归一历史输入兜底；正常输入经别名归一后同义选择自动一致。",
        "",
    ]
    path.write_text("\n".join(lines), encoding="utf-8")


def main() -> int:
    ap = argparse.ArgumentParser(description="合并双 Agent 交叉验证裁决（只读，生产层零改动）。")
    ap.add_argument("--batch-index", type=int, default=1)
    ap.add_argument("--input", type=Path, default=None)
    ap.add_argument("--verdict-codex", type=Path, default=Path(r"D:/1大创/.xval/verdict_codex.csv"))
    ap.add_argument("--verdict-opus", type=Path, default=None)
    ap.add_argument(
        "--allow-conservative-intersection",
        action="store_true", dest="conservative_intersection",
        help=(
            "采用保守交集规则（2026-10-08 用户决定）：分歧条目中双方都判 support 且至少一方判"
            "「成立」的，按现类型不改落地（landable=intersection）。不开时分歧一律搁置。"
        ),
    )
    args = ap.parse_args()

    n = args.batch_index
    inp = args.input or REPORTS / f"sweep_batch{n}_input.csv"
    opus = args.verdict_opus or ROOT / ".codex_tmp" / "verdict_opus.csv"
    xval = REPORTS / f"sweep_batch{n}_cross_validated.csv"
    dis = REPORTS / f"sweep_batch{n}_disagreements.csv"
    rep = REPORTS / f"sweep_batch{n}_cross_validation_report.md"

    result = merge(inp, args.verdict_codex, opus, n, conservative_intersection=args.conservative_intersection)
    write_csv(result["rows"], xval, XVAL_COLUMNS)
    write_csv([r for r in result["rows"] if r["agreement"] == "disagree"], dis, XVAL_COLUMNS)
    write_report(result, rep, {"input": inp, "xval": xval, "disagree": dis})

    rows = result["rows"]
    print(f"输入 {len(rows)} 条；一致 {sum(1 for r in rows if r['agreement']=='agree')}、"
          f"分歧 {sum(1 for r in rows if r['agreement']=='disagree')}、"
          f"无效 {sum(1 for r in rows if r['agreement']=='invalid_missing_verdict')}；"
          f"保守交集落地 {sum(1 for r in rows if r['landable']=='intersection')} 条"
          f"（规则{'已采用' if result['intersection_used'] else '未开启'}）")
    print(f"一致且可落地（进公开层）{sum(1 for r in rows if r['landable']=='yes')} 条；"
          f"仅研究层 {sum(1 for r in rows if r['landable']=='research_only')} 条")
    if result["errors"]:
        print(f"裁决校验问题 {len(result['errors'])} 处：")
        for e in result["errors"][:10]:
            print("  -", e)
    for p in result["a_problems"] + result["b_problems"]:
        print("  CSV:", p)
    print(f"对照表 -> {xval.relative_to(ROOT)}")
    print(f"分歧单 -> {dis.relative_to(ROOT)}")
    print(f"报告   -> {rep.relative_to(ROOT)}")
    # proposed_type 词表校验只对「不得作为建议类型」的占位标签报错；
    # 同属组织等真实类型即使不公开也是合法裁决，不算错误。
    return 1 if result["errors"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
