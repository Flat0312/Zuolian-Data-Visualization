# Phase 7 候选裁决登记（研究层）

- 授权人：用户（会话授权）
- 授权时间：2026-09-20
- 授权语（逐字）：「我都没问题，你自己看着办吧」
- 依据：2026-09-20 主 Agent 独立审计 `phase7_candidate_independent_audit.csv` + 用户对三项待裁决清单的整体授权。

## 裁决结果

- support 9（含 REL-00523 引文替换后升级 1 条）
- associated 2（REL-01219 / REL-01743 升级）
- insufficient 1（REL-00113 维持）
- 引文替换 1 条

## 逐条裁决

| 候选 | 关系 | 裁决 | 说明 |
| --- | --- | --- | --- |
| CAND-RELE-P7-001 | REL-00092 鲁迅→李小峰 | agree_support | 维持冻结 support：日记直接记载，独立审计逐字与年月日锚定通过。… |
| CAND-RELE-P7-002 | REL-00060 鲁迅→李霁野 | agree_support | 维持冻结 support：日记直接记载，独立审计逐字与年月日锚定通过。… |
| CAND-RELE-P7-003 | REL-00109 鲁迅→陈望道 | agree_support | 维持冻结 support：日记直接记载，独立审计逐字与年月日锚定通过。… |
| CAND-RELE-P7-004 | REL-00019 鲁迅→郁达夫 | agree_support | 维持冻结 support：日记直接记载，独立审计逐字与年月日锚定通过。… |
| CAND-RELE-P7-005 | REL-00089 鲁迅→许广平 | agree_support | 维持冻结 support：日记直接记载，独立审计逐字与年月日锚定通过。… |
| CAND-RELE-P7-006 | REL-00097 鲁迅→林语堂 | agree_support | 维持冻结 support：日记直接记载，独立审计逐字与年月日锚定通过。… |
| CAND-RELE-P7-007 | REL-00011 鲁迅→郑伯奇 | agree_support | 维持冻结 support：日记直接记载，独立审计逐字与年月日锚定通过。… |
| CAND-RELE-P7-008 | REL-00006 鲁迅→冯雪峰 | agree_support | 维持冻结 support：日记直接记载，独立审计逐字与年月日锚定通过。… |
| CAND-RELE-P7-009 | REL-00113 鲁迅→巴比塞 | agree_insufficient | 维持 insufficient：三轮检索回执在案，本地两源均无直接交游证据；不得以背景共现确认关系。… |
| CAND-RELE-P7-010 | REL-00523 潘汉年→丁玲 | quote_corrected_to_support | 二轮捕获引文 EVI-N1-E0960180FF3F 摘自同页茅盾段落（独立审计 issue_found）；按授权替换为… |
| CAND-RELE-P7-011 | REL-01219 冯乃超→柔石 | agree_upgrade_to_associated | 按授权升 associated：二轮证据 EVI-N1-FBD6D22370A6（左联词典 第334页）逐字与哈希校验通… |
| CAND-RELE-P7-012 | REL-01743 丁玲→穆木天 | agree_upgrade_to_associated | 按授权升 associated：二轮证据 EVI-N1-5B8C3B7E6622（左联史 第21页）逐字与哈希校验通过，… |

## 边界

本裁决为研究层产物（`review_status=adjudicated_authorized`），未改动生产层：
`relation_evidences.csv` 落地、`person_relations.publish_status` 转换与发布层重建属下一批次，
须另行幂等脚本、守门测试与验收（参照第三/四批A模式）。
