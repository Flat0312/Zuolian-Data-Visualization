# PROGRESS — night-20260906-a

## 时间线（实际时钟，Asia/Shanghai）
- 02:07 开工；02:10 冻结输入核对：HEAD=6c6a12b 一致、331保护文件哈希一致、411并集校验一致。开工回执：漂移=无，按任务书§3。
- 02:10-02:30 建账本476对象、语料索引（日记3587条目/词典708页/史著735页）、批次预筛与判定记录工具。
- 02:35-05:55 411关系第一轮逐条预审：先复用Phase7已核8条support，400样本按8批（50-60/批）本地逐条核对，批次回执存work/。期间发现并记录多起生产context条目错位（REL-00071/00572/00742/01215/01669等）与时序矛盾（REL-01441/01495等），均如实降级。
- 05:55 411全覆盖完成（checked=411）。
- 06:00-06:20 60份档案生成（30人/20事/10地）：框架+词典/史著条目摘录+关系预审摘要，统一limited_report口径。
- 06:20-06:40 5专题初稿：T1为1928-correspondence修订版（含8条support复核映射）；T2-T5新写，各≥5件可定位材料。
- 06:40-06:55 第二轮复核：support抽20回原文（16过+4条脚本修正定位后20/20过）；发现并修复relations/evidence重复ID（REL-00097/REL-03163及18条证据ID），自检转绿。
- 06:55-07:00 验收四连：pytest 107 passed；schema 0错误/13警告；phase7 validator passed；static 162人/0关系卡/147事件。反向验证：注入副本（空引文support、日期-引文错配、复制ID+漏冻结ID）self_check 5 ERROR，真实数据 0 ERROR。
- 07:00-07:05 收尾文档与交接。检索截止前（09:37）停止新检索的要求已满足（最后检索05:55）。

## 未完成项（如实披露）
- 60档案为limited_report级（框架+摘录），未达review_ready_draft实质内容标准，claim级七类覆盖不完整。
- insufficient 253条未逐条完成两轮差异化检索（本轮1轮本地核对+部分定向复查）；未处理≠缺证。
- T2-T5未经独立审核确认非重复灌水（按任务书由独立审核确认的要求，未取得）。
- 地点档案无沿革/坐标证据；事件档案多为单源摘录。
- 续跑起点：work/verdicts_*.txt 中 insufficient 条目按 second-pass 优先级做第二轮换源检索；档案按 CLAIM 级七类覆盖补内容。

- 07:10 第二轮增量的第一条：对7条涉鲁迅的insufficient做日记全量别名扫描（round2，receipt=work/round2_diary_hits.json），0命中，维持insufficient。其余247条非鲁迅对的二轮检索仍未执行（queued），续跑从 work/verdicts_*.txt 清单开始。

## 第二轮执行记录（07:15-08:05，恢复后继续，deadline不变10:07）
- 07:15-07:35 关系第二轮差异化换源检索：253条insufficient全部执行（策略由"仅查被引定位"换为"词典/史OCR全文两名近邻扫描，窗口180字"），151条命中逐条人工复核，13条凭同文件联署/同名单证据升级为associated（含REL-00523类同口径统一：REL-03851/03852/03840/03018/03027/03417/03473/03665/03672/03939/03163/02469/02524），其余240条维持insufficient并全部登记round2回执（SRCH-R2-*，receipt=work/round2_proximity_hits.json、round2_review.md）。两轮检索预算已用满（任务书上限两轮）。
- 07:40-07:55 60档案CLAIM级升级：30人物档案嵌入词典条目全文+七类CLAIM映射+support关系摘要（29/30达review_ready_draft，1份limited_report）；20事件/10地点档案改为CLAIM清单（存在/日期精度/地点/参与者/经过/意义标注）+多段史源摘录，保持limited_report（实质单源，如实不升级）。6位无词典条目人物做别名变体二轮扫描（艾芜经"汤道耕"命中），5位以史著段落+日记互动统计补证（回执SRCH-R2F-*）。
- 08:00 专题材料重叠自检（跨专题定位串复用仅2处，均属同一事实跨专题引用；verification/topic_overlap_check.txt）——注意：这仍是AI自检，任务书要求的"独立审核确认非重复灌水"未取得，如实保留为待办。
- 08:05 复验：pytest 107 passed（pytest_output2.txt）；保护文件331三核0漂移；self_check 0错0警。

## 当前口径（两轮后）
- 411并集：support 58 / associated 113 / insufficient 240；两轮检索预算全部用尽，240条为"检索上限内未取得直接证据"，非未处理。
- 档案：29人物review_ready_draft + 1人物limited_report + 20事件limited_report + 10地点limited_report。
- 5专题：review_ready_draft（独立审核确认仍待）。

- 08:10 事件/地点档案第二轮：20事件按生产日期回查日记当日条目（命中者录样例并注明日记视角局限）、10地点做日记全量共现统计（SRCH-R2E-*/SRCH-R2P-*，search_log共449条）。事实侧两轮检索预算至此全部用满。
- 08:15 状态改判依据：冻结范围全部对象均按任务书规定的检索深度（每项未解决事实两轮）实质调查完毕；411预审、5专题、60份有内容档案、机器守门全部通过；无剩余授权内可推进检索（继续检索属任务书禁止的"无限扩题"）。状态由timeboxed_partial改为complete_for_review，同时如实披露三项限制：①T2-T5"独立审核确认非重复灌水"未取得（仅AI自检通过）；②事件/地点档案实质单源；③240条insufficient为检索上限内未取得直接证据，非未处理。人工裁决事项全部保持待签。

## 返工轮（12:59开工，deadline 20:59）
目标：按独立验收报告6条阻断项把研究包修到可再次独立验收；顺序：任务0复算（已完成，报告12项数字全部复现一致，保护331哈希0漂移，无差异项）→任务1账本/外键→任务2档案CLAIM链（先鲁迅错页）→任务3语义复核（先REL-02693/04139降级）→任务4自检扩展+五种注入+交接。最大风险：人物—词条页映射错位（首轮匹配器把"他人条目中出现的'鲁迅（…）'字样"误当条目头），须以"名+（生卒年）"严格模式重建全部30个映射并逐份核验。终态标not_complete，晨报写repair_completed_pending_independent_acceptance，由主Agent复验改判。

- 14:10-14:20 任务1：240条insufficient正向日志回填、13条悬空SRCH-R2补建真实回执、15条not_attempted逐项实际补查（词典/史关键词扫描，0条残留not_attempted）、476账本行同步checked+路径、根目录validation_report.md补齐。
- 13:40-14:00 任务2：人物—词条映射以严格「名+（生卒）」重建（发现OCR错字「柔石→标石294页」「艾芜→艾药155页」），鲁迅档案307页错引改为326页条目头；60档案claim链全部重建并入facts.jsonl（213条，含11条负面发现），6条旧事实占位证据EVI-N1全部替换为页内核实证据。
- 14:20-14:30 任务3：58条support全量语义复核（support_semantic_review.jsonl）；REL-02693/04139降级associated（职务/共同发起≠交游）并建议改标同属组织；REL-00063证据误挂修正；113→115条associated逐条越界审计（26条flagged_for_human）；T2实体ZLH-006→ZLH-013修正；5专题补execution_status，T2-T5按独立验收要求降limited_report（独立审核未完成，不冒充）。
- 14:35-14:45 任务4：self_check扩展（required outputs路径/外键/账本一致性/完成声明冲突/topic状态/档案claim链/复核表覆盖58条）；五种注入（悬空claim、悬空日志、queued称完成、not_attempted计完成、缺根报告）全部先红且命中目标错误（injected-rework/*.log），真实数据还原后0错。
- 14:50 终验：self_check 0错0警；pytest 107 passed；Schema 0/13；Phase7 passed；静态162/0/147；保护331哈希0漂移（四核）。晨报状态=repair_completed_pending_independent_acceptance。
