# validation_report.md（返工轮根目录索引，2026-09-06）

本文件为任务书要求的根目录验收索引；详细机器输出在 verification/ 下：

- 机器守门：verification/pytest_output2.txt（107 passed）、schema_output.txt（0/13）、phase7_validator_output.txt（passed）、static_build_output.txt（162/0/147）
- 保护文件：verification/protected_recheck.txt（331/331，三次核对0漂移）
- 自检与注入：python tools/self_check.py（返工扩展版，见 tools/self_check.py）；注入副本 verification/injected-rework/（五种注入各存失败输出）
- 语义复核：verification/support_semantic_review.jsonl（58条support全量）；verification/associated_review.jsonl（113条associated复核）
- 专题重叠：verification/topic_overlap_check.txt
- 当前状态：not_complete → 修复完成后 repair_completed_pending_independent_acceptance（见 night-state.json 与 morning-report.md）
