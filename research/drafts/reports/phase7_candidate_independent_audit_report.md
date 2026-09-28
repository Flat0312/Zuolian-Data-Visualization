# Phase 7 候选独立审计报告（主 Agent）

> 审计人：ZCode GLM-5.3（主 Agent，独立审计，非执行者自检）
> 审计日期：2026-09-20
> 输入：冻结候选包 12 条 + 夜间轮二轮分诊（relations/evidence/search_log）+ 本地原文三源

## 结论

agree_support 8 · agree_insufficient 1 · agree_upgrade_to_associated 2 · issue_found 1。
全部 12 项为 pending_human_review，本审计不直接改动生产层或冻结候选包。

| 候选 | 关系 | 对 | 冻结 | 二轮 | 独立判定 |
| --- | --- | --- | --- | --- | --- |
| CAND-RELE-P7-001 | REL-00092 | 鲁迅→李小峰 | support | support | agree_support |
| CAND-RELE-P7-002 | REL-00060 | 鲁迅→李霁野 | support | support | agree_support |
| CAND-RELE-P7-003 | REL-00109 | 鲁迅→陈望道 | support | support | agree_support |
| CAND-RELE-P7-004 | REL-00019 | 鲁迅→郁达夫 | support | support | agree_support |
| CAND-RELE-P7-005 | REL-00089 | 鲁迅→许广平 | support | support | agree_support |
| CAND-RELE-P7-006 | REL-00097 | 鲁迅→林语堂 | support | support | agree_support |
| CAND-RELE-P7-007 | REL-00011 | 鲁迅→郑伯奇 | support | support | agree_support |
| CAND-RELE-P7-008 | REL-00006 | 鲁迅→冯雪峰 | support | support | agree_support |
| CAND-RELE-P7-009 | REL-00113 | 鲁迅→巴比塞 | insufficient | insufficient | agree_insufficient |
| CAND-RELE-P7-010 | REL-00523 | 潘汉年→丁玲 | insufficient | support | issue_found |
| CAND-RELE-P7-011 | REL-01219 | 冯乃超→柔石 | insufficient | associated | agree_upgrade_to_associated |
| CAND-RELE-P7-012 | REL-01743 | 丁玲→穆木天 | insufficient | associated | agree_upgrade_to_associated |

## 发现（需人工裁决）

### REL-00523 潘汉年→丁玲

二轮升级 support 的理由与左联史原文相符（「潘汉年就去看望他们」「介绍他俩一同加人左联」段已独立定位），但捕获引文 EVI-N1-E0960180FF3F 摘自同页另一段落（茅盾/叶圣陶/郑振铎），不含潘汉年或丁玲。建议：替换为已定位的正确段落引文后再议转正；在此之前 support 升级缺乏逐字证据。

## 核验方法

- 引文哈希：sha256(quote) 与候选包登记值一致（8/8）。
- 逐字定位：OCR 容差规范化（NFKC、去空白间隔点）后在本地《日记全编》全文检索，8/8 命中。
- 年月日锚定：剥离邮戳短语（X月N日发）后取引文前最近的卷（日记N(YYYY年)）/月/日标题，8/8 与候选 locator 精确一致；REL-00060「寄霁野信。」全文 18 处命中中含 1928-02-26 精确锚。
- 人物词元：引文含关系至少一方（日记侧作者本人覆盖另一方），别名表见脚本常量。
- 二轮新证据：REL-01219（词典334页）/REL-01743（左联史21页）逐字命中且含双方；REL-00523 捕获引文逐字命中但为同页错误段落（见发现）。
- REL-00523 正确段落锚定短语「潘汉年就去看望他们」「介绍他俩一同加人左联」已在左联史原文验证存在。
