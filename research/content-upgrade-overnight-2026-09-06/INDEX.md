# INDEX — 通宵内容建设 2026-09-06（night-20260906-a）

- 运行状态: not_complete → repair_completed_pending_independent_acceptance（14:50返工完成，：全部对象按两轮检索深度实质调查完毕、四项完成条件满足；三项限制见morning-report.md——独立审核确认未取得、事件/地点单源、240条insufficient为检索上限内未得证）
- night-state.json / PROGRESS.md / BLOCKED.md / morning-report.md — 状态与交接
- selection.json — 冻结范围（30人物/20事件/10地点/400样本+12候选=并集411，交集REL-00097）
- scope_ledger.jsonl — 476对象执行账本（30+20+10+411+5）
- relations.jsonl — 411条预审（两轮检索后：support 58 / associated 113 / insufficient 240 / conflict 0；execution_status全checked；人工字段全部为空）
- evidence.jsonl — 81条去重证据（local_checked；hash_basis=local_file_bytes；引文quote_sha256逐条）
- search_log.jsonl — 449条检索回执（含全部两轮）（每条关系至少1条本地核对回执）
- facts.jsonl / dossiers.jsonl（60档案索引，正文在 dossiers/）/ topics.json（5专题，正文在 topics/）
- change_proposals.csv — 18条关系类型改标建议（全部待人工裁决）
- tools/ — 本轮脚本（corpus索引、批量预筛、判定记录、self_check）
- work/ — 各批次摘录回执 auto_*.md 与判定文件 verdicts_*.txt
- verification/ — pytest/schema/phase7/static 真实输出、self_check 绿绿红红回执、second_pass_support.md、static/（正式静态构建副本）
- night-state.json 中 ai_executor 与各jsonl内 ai_executor 字段为 AI 执行者；全部记录 review_status=pending_human_review，人工签核字段为空。

## 统计口径（三个集合分开计）
- sample400：support 50 / associated 98 / insufficient 252
- phase7-12：support 9 / associated 2 / insufficient 1（REL-00523、REL-01219、REL-01743 由insufficient升级，依据见relations.jsonl）
- 并集411：support 58 / associated 100 / insufficient 253
以上为预审分布计数，不是准确率，不能外推总体。
