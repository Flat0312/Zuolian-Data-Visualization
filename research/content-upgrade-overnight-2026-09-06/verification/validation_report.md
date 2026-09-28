# validation_report.md — 验收与反向验证（2026-09-06）

## 机器守门（真实输出存 verification/）
1. `python -m pytest -v` → **107 passed**，0 failed/0 skipped/0 xfailed（pytest_output.txt，exit 0）。basetemp使用默认目录，未复用旧残留路径。
2. Schema `validate_data_dir('data/processed')` → **0 errors / 13 warnings**（schema_output.txt）。
3. `python research/analysis/validate_phase7_relation_candidates.py` → **passed**（phase7_validator_output.txt，exit 0）。
4. `python build_static_site.py --output-dir research/content-upgrade-overnight-2026-09-06/verification/static` → **162 people, 0 relation cards, 147 events**（static_build_output.txt，exit 0）。

## 保护文件
- 开工（02:10）与收尾（06:50）两次核对 baseline.protected_sha256：**331/331 一致，0漂移**（protected_recheck.txt）。

## R内自检（tools/self_check.py）
- 真实数据：**errors=0 warnings=0**（校验：冻结411并集全覆盖、无重复/越界ID、support必有非空引文+理由、引文回定位逐字命中、review_status/人工字段合法、枚举合法）。
- 反向验证（副本 verification/injected/）：注入①support空引文 ②定位日期与引文错配 ③复制ID并漏掉冻结ID REL-00097 → **5 ERROR / 1 WARN**（含 duplicate relation_id、outside union、empty quote、quote-date mismatch、union missing），证明自检能捕获三类注入。注入仅作用于副本，真实资料未改动。
- 复现性：self_check 二跑一致（纯读取）；corpus索引、批次摘录 auto_*.md 为固定输入输出，可重放。

## 第二轮质量复核（AI自我复核，非人工/独立验收）
- support抽样：seed=20260906，抽20/58，回原文逐字核对 20/20 通过（4条首查因定位前缀解析未命中，修正解析后通过；详见 second_pass_support.md）。
- 完整专题关键事实复查：T2成立大会（左联史20页与词典163页双源互证）、T3殉难日期精度（词典131页“2月?日”与通行2月7日之差，如实标注）、T5酒会名单（左联史191页）。
- 发现并修复：relations/evidence 重复ID（批次区间重叠），去重后 self_check 转绿。

## 统计口径声明
sample400 / phase7-12 / 并集411 三集合分别计数（见INDEX.md）；预审分布不是准确率。
