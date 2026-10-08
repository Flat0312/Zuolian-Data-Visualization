# Phase 5 重捕候选第二批落地报告（2026-10-08，逐条独立裁决）

> 执行脚本：`research/analysis/apply_phase5_recapture_landing.py`（幂等）  
> 批次标记：`P5-RECAPTURE-2026-10-08`  落地日期：2026-10-08

## 0. 授权与口径声明

- 授权人：用户（逐条独立裁决）；授权时间：2026-10-08。
- 授权语（逐字）：「第一优先REL-00622 和 REL-01368直接过，REL-01891判"证据不足"，REL-01161 成立但不影响公开层」。
- 本批为**逐条独立人工裁决**，口径强于第一批 2026-09-20 的「授权按建议执行」（概括授权，签核路径 C）。两种口径必须分别引用：公开层现共 7 条，其中 5 条来自 2026-09-20 概括授权批次，2 条来自本批逐条裁决。
- 公开层只接受 **derived `supported`**：未改门禁、未用 `human_adjudication`、未改 `relation_risk_level` / `needs_manual_review` / `confidence`。

## 1. 落地明细

| relation_id | 人物对 | 裁决 | 类型变化 | 落地后状态 | 新证据行 | 复用来源 |
| --- | --- | --- | --- | --- | --- | --- |
| REL-00622 | 周扬—邵荃麟 | 过（直接进公开层） | 不改 | supported | RELE-10270 | SRC-0779 |
| REL-01161 | 阳翰笙—林淡秋 | 成立但不影响公开层 | 不改 | pending_review | RELE-10271 | SRC-0739 |
| REL-01368 | 郁达夫—陈望道 | 过（直接进公开层）；类型更正 交游→签名联署 | 交游→签名联署 | supported | RELE-10272 | SRC-0054 |
| REL-01891 | 叶紫—萧军 | 证据不足（生产层零改动） | 不改 | inferred | —（不落地） | SRC-0946 |

- REL-01891 判「证据不足」：未新增任何证据行、未写 `rejected`（rejected 语义是人工否定），生产表零改动，保持 `inferred`。
- REL-01161 落地 support 证据但被门禁（critical + low + 同属组织）自然挡在 `pending_review`，不进入公开层。
- 引文逐字复核：4/4 按候选包偏移在空白归一原文中原样取回，`quote_sha256` 自洽；3 条落地引文另过双方佐证门。

## 2. 实测终值 vs 预期终值

| 项目 | 预期 | 实测 |
| --- | ---: | ---: |
| person_relations 行数 | 4238 | 4238 |
| relation_evidences 行数 | 10272 | 10272 |
| sources / passages / works | 1178 / 1178 / 65 | 1178 / 1178 / 65（未注册新来源、未重写这三张表） |
| publish_status 分布 supported/pending_review/inferred | 7 / 2451 / 1780 | 7 / 2451 / 1780 |
| support 证据行 | 23 | 23 |
| associated 证据行 | 10249 | 10249（未改判既有行） |
| reviewed 证据行 | 23 | 23 |
| critical 计数 | 1974 | 1974（未反向降险） |
| 类型更正 | 恰 1 条（REL-01368） | 1 条 |
| 公开层关系 | 7 条 | 7 条（REL-00046, REL-00059, REL-00097, REL-00622, REL-01368, REL-03289, REL-03518） |

Schema：写后 0 errors / 13 warnings。

## 3. 明确未做的事

- 未改 `relation_publish_status.py` 判定顺序、`INFERRED_RELATION_TYPES`、`PUBLIC_RELATION_STATUSES`。
- 未把 10249 条 associated 或第一批 20 条 support 改判/改写。
- 未注册新来源（三个落地 locator 全部复用既有 SRC-0779 / SRC-0054 / SRC-0739）。
- 未把 REL-01891 写为 rejected；未对候选包与队列文件做任何改动。
- 未使用 human_adjudication 通道；未动风险/置信/复核标记列。

## 4. 复核方式

```powershell
python research/analysis/apply_phase5_recapture_landing.py --dry-run   # 只校验不写
python research/analysis/apply_phase5_recapture_landing.py             # 落地（二跑输出「无新增/已完成」）
python research/analysis/build_publish_data.py
python build_static_site.py
python research/analysis/build_trustworthy_network_analysis.py
python -m pytest -q
```

逐条痕迹见 `phase5_recapture_landing_ledger.csv`（4 行）；28 行候选的处置与授权溯源见 `phase5_quote_recapture_adjudicated.csv`；裁决原文与语义见 `phase5_quote_recapture_adjudication_record.md`。
