# Phase 5 · 400 条关系审核签核单

> 生成脚本：``research/analysis/build_phase5_review_package.py``（幂等，重跑字节不变）
> 数据来源：``phase5_relation_review_template.csv`` × 夜间轮 ``relations.jsonl`` sample400 组（2026-09-06 回源核查，全 pending_human_review）

## 0. 边界声明

``ai_suggested_verdict`` 是夜间轮回源核查的机器建议，不是人工判定。
本签核单及审核包整体状态为 pending_human_review；``human_verdict``/``human_note``
两列必须由人工评审员填写，执行者不得代填。人工裁决落地前，本包不产生任何准确率结论。

## 1. 裁决词表（human_verdict 合法值）

| 值 | 含义 |
| --- | --- |
| `correct` | 关系成立且 standard_relation_type 恰当，可作为已证关系保留 |
| `wrong_type` | 关系成立但类型应改，human_note 填建议类型 |
| `not_supported` | 现有证据不足以支持，不可作为已证关系保留 |
| `contradicted` | 证据显示关系不成立或错挂，应删除或降级 |

## 2. 分层统计（AI 建议口径）

共 400 条。夜间轮证据分诊：support 48 / associated 113 / insufficient 239。
AI 建议裁决：correct 135 / wrong_type 26 / not_supported 239 / contradicted 0。

| 证据分诊 × 风险 | critical | high | medium | low |
| --- | ---: | ---: | ---: | ---: |
| support | 33 | 2 | 0 | 13 |
| associated | 55 | 7 | 1 | 50 |
| insufficient | 96 | 37 | 9 | 97 |

## 3. 签核路径

1. 路径 A（逐条）：按 `phase5_human_review_queue.csv` 顺序或直接在审核包 CSV 中逐条填写 `human_verdict`。
2. 路径 B（分层批量授权）：对某一分层（如 insufficient × critical）整层授权按 `ai_suggested_verdict` 执行，需在授权语中点名分层。
3. 路径 C（全量按建议执行）：对 400 条全部按 `ai_suggested_verdict` 落地，需明确授权语（参照第三批追认模式，落地脚本另行任务书）。

任何路径下，`wrong_type` 的目标类型与 `not_supported` 的处置（降级/删除）以人工裁决为准；
与 AI 建议不一致的行请在 `human_note` 写明理由。

## 4. 关联文件

- 审核包：`research/drafts/reports/phase5_relation_review_package.csv`
- 夜间轮证据与回执：`research/content-upgrade-overnight-2026-09-06/`（relations.jsonl、evidence.jsonl、search_log.jsonl）
- 旧 AI 启发式预审（已被本包取代，保留历史）：`phase5_relation_review_ai_filled.csv`
