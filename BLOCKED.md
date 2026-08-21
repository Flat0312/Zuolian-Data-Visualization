# BLOCKED - 合并管线幂等化与来源挂接任务

无阻塞事项。以下为观察记录（超出白名单，未处理）：

1. 【已还原的瞬态改动】2026-08-21 修复任务期间，`data/publish/publish_manifest.json` 曾出现一次仅末尾换行差异的未暂存改动（内容零变化）。`git checkout` 还原后经隔离测试组与全量 pytest 各复跑一次均无法复现，疑与 git autocrlf 换行归一化或某次进程触碰有关。生产数据未受影响。
2. 【口径提示】`data/publish/publish_manifest.json` 内记录的 `schema_warnings: 37` 与 `kb_schema.validate_data_dir` 实报的 13 不一致，系生成脚本统计口径不同（可能含逐行 issue 计数），属历史生成产物。`data/publish/` 在白名单外，本轮未核实与修正；建议后续维护生成脚本时统一。
