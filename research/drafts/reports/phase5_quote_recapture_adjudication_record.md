# Phase 5 重捕候选第二批人工裁决记录（2026-10-08，逐条独立裁决）

> 裁决载体：本文件 + `phase5_quote_recapture_adjudicated.csv`  
> 候选包：`phase5_quote_recapture_candidates.csv`（28 行，原样保留，不表示裁决状态）  
> 落地脚本：`research/analysis/apply_phase5_recapture_landing.py`（批次标记 `P5-RECAPTURE-2026-10-08`）

## 0. 授权语（逐字，不得改写）

用户 2026-10-08 逐条裁决，原话：「第一优先REL-00622 和 REL-01368直接过，REL-01891判"证据不足"，REL-01161 成立但不影响公开层」。

本批为**逐条独立裁决**，口径强于 2026-09-20 第一批的「授权按建议执行」（概括授权，签核路径 C，逐字「我都没问题，你自己看着办吧」）。两种口径在台账、站点文案与后续引用中必须分别说明，不得混同为一类「人工裁决」。

## 1. 四条裁决与处置

| relation_id | 人物对 | 现类型 | 裁决（语义） | 处置 | 落地后状态 |
| --- | --- | --- | --- | --- | --- |
| REL-00622 | 周扬—邵荃麟 | 交游 | 过（直接进公开层） | 新增 support 证据行 RELE-10270（复用 SRC-0779） | supported（derived） |
| REL-01161 | 阳翰笙—林淡秋 | 同属组织 | 成立但不影响公开层 | 新增 support 证据行 RELE-10271（复用 SRC-0739） | pending_review（derived） |
| REL-01368 | 郁达夫—陈望道 | 交游 | 过（直接进公开层）；类型更正 交游→签名联署 | 新增 support 证据行 RELE-10272（复用 SRC-0054）；类型 交游→签名联署 | supported（derived） |
| REL-01891 | 叶紫—萧军 | 交游 | 证据不足（生产层零改动） | 生产层零改动 | inferred（derived） |

## 2. REL-01891「证据不足」而非 rejected 的理由

重切段落为书目著录（「…传记小说。李克因作，载《东方纪事》1987年3、4期合刊。叙述叶紫…以及叶紫同陈企霞、聂绀弩、周颖夫妇、萧军、萧红夫妇等的交往」），属**二手著录**（候选包 secondary_description=yes），
不足以定为 support 级直接记载。五态中 `rejected` 的语义是**人工否定/证伪**；把「未证实」写成 `rejected` 会夸大成「已证伪」。故本条**不新增任何证据行**，生产表零改动，关系保持 `inferred`，留在研究层待后续补证。

## 3. REL-01161「成立但不进公开层」的门禁依据

新增 support 证据行是对「裁决成立」的记录；但该关系 `relation_risk_level=critical`、`confidence=low`、类型 `同属组织` 属推断类型，按 `relation_publish_status` 判定顺序被三重拦截，自然停留在 `pending_review`。**禁止**（本批也未）使用 `human_adjudication` 通道放行——公开层只接受 derived `supported`。风险列未做任何改写（critical 计数仍 1974）。

## 4. 引文溯源

四条重切引文、locator、quote_sha256、归一化偏移全部取自候选包，落地脚本运行时按 `source_file` + `normalized_start/end` 在空白归一后的本地原文中逐字回定位复核，未重新检索、未改写引文。
逐条复核结论见 `phase5_recapture_landing_ledger.csv` 的 `quote_verbatim_recheck` 列（全部 `verbatim_offset_hit_whitespace_normalized`）。

授权：用户（逐条独立裁决），2026-10-08。
