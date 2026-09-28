# repair-report.md — 独立验收阻断项返工（2026-09-06 12:59–14:45）

依据：verification/independent_acceptance_2026-09-06.md（只读，未改动）。终态：**not_complete → repair_completed_pending_independent_acceptance**，由主Agent复验改判。

## 改动前后计数

| 项 | 返工前 | 返工后 |
|---|---|---|
| scope_ledger 60个实体状态 | 60 queued、result_path 全空 | 476/476 checked，实体行均带档案路径与缺口说明 |
| dossiers.jsonl 无claim_ids | 59/60 | 0/60（49份有正向claim，11份登记negative-finding事实，claim_ids全闭合） |
| facts.jsonl | 6条（其中6条证据引用为占位EVI-N1） | 213条（新增207条：24人物claim+事件/地点claim+11条负面发现；6条旧事实证据全部替换为页内逐字核实的真实EVI） |
| 240条insufficient正向search_log_ids | 全空 | 全部回填（每条≥1个真实存在SRCH） |
| 悬空search_log引用 | 13 | 0（13条SRCH-R2-*补建真实回执日志，receipt=work/round2_review.md） |
| not_attempted日志 | 15 | 0（15条逐项实际补查：18个SRCH-R2E结果中local_checked 15/not_found 3——准确分布见search_log.jsonl；未取得者如实not_found，不称完成） |
| topics execution_status | 5缺 | 5补（全checked） |
| T2–T5 content_status | review_ready_draft（独立审核未完成） | 降级 limited_report（独立非重复审核仍待，理由写入limits） |
| relations support | 58（含2条语义误判） | 56（REL-02693、REL-04139降级associated并建议改标同属组织；REL-00063证据换为词典287页柳倩条正确原文——原288页窗口误挂前一人物） |
| 鲁迅档案词典页 | 第307页（他人条目） | 第326页「鲁迅(1881一1936)」条目头，严格“名+（生卒）”模式重建全部30个映射 |

## 全量语义复核

- support：verification/support_semantic_review.jsonl 覆盖原58条全集（56维持+2降级），逐条含双方、关系动作、证据、支持方式、结论、修改。
- associated：verification/associated_review.jsonl 115条逐条核查未越界；26条proposed类型超出名单类证据强度的，标注flagged_for_human（不自行改判，待人工）。
- 语义原则执行：仅共现/同组织/同名单/同场/职务关联一律不得support；本轮support仅保留①一手日记直接动作②同文件联名③组织成员/职务的直接词条记载④纪念/合编/传记交往的直接记载。

## 档案降级与保留（60份终态）

- person：24 review_ready_draft（词典条目头经严格模式+OCR错字变体检索确认：含「柔石→标石(294页)」「艾芜→艾药(155页)」两个OCR错字命中），6 limited_report（茅盾/潘汉年/林语堂/邹韬奋/斯诺/高尔基两轮均无词典个人条目，如实降级）。
- event：4 review_ready_draft（≥3段页内逐字命中摘录）、16 limited_report；place：6 review_ready_draft、4 limited_report。所有档案正文含CLAIM链段与缺口节。

## 未完成项与续跑ID

1. T2–T5独立"非重复灌水"审核：未完成（执行者无权替代）；续跑=将topics/T2–T5送独立审核，通过后回升review_ready_draft。
2. 240条insufficient：两轮检索预算（上一轮）已用满，本轮未再扩检；续跑=人工裁决或授权新增来源族（影印原刊/档案馆藏），自REL-00036（work/verdicts_b001.txt第2条）起。
3. 26条associated的proposed类型（flagged_for_human）：待人工裁决，见verification/associated_review.jsonl。
4. 15条事件not_attempted补查中3条not_found（生产日期缺失/日记该年缺佚）：续跑=授权查原刊影印本，自EVT-00030起（search_log.jsonl SRCH-R2E-*）。

## 验证回执

- 五种注入先红：verification/injected-rework/inj1..inj5.log（悬空claim/悬空日志/queued却称完成/not_attempted计完成/缺根报告，各自命中目标错误）
- 还原后真实数据：python tools/self_check.py → errors=0（见下终验节）
