# Phase 3 研究层与展示发布层分离实施计划

**目标：** 建立可重复、可审计的发布数据生成流程，避免研究层中的候选和争议结论被前端当成确定史实。

## 文件范围

- 新增 `research/analysis/build_publish_data.py`
- 新增 `data/publish/`
- 新增 `data/publish/publish_manifest.json`
- 新增 `tests/test_publish_data.py`
- 修改前端数据目录解析和模式选择
- 新增 `research/drafts/reports/phase3_publish_gate_report.md`

## Task 1：定义发布规则与清单

- [ ] 编写发布过滤失败测试。
- [ ] 定义各表的公开规则和保留原因。
- [ ] 生成包含输入数量、输出数量和过滤数量的发布清单。

## Task 2：生成引用闭合的发布数据

- [ ] 过滤候选与争议组织身份。
- [ ] 仅保留公开结论引用的组织身份证据。
- [ ] 过滤非公开关系状态。
- [ ] 裁剪引用闭合的事实证据与来源。
- [ ] 运行发布目录 Schema 校验。

## Task 3：前端模式切换

- [ ] 前端支持“公开模式”和“研究模式”。
- [ ] 公开模式读取 `data/publish/`。
- [ ] 研究模式读取 `data/processed/`，并清楚展示候选与争议标记。

## Task 4：报告与验收

- [ ] 输出发布门禁报告。
- [ ] 验证删除 `data/publish/` 后可一条命令重建。
- [ ] 运行全量测试、Schema、静态构建和定向 lint。

## 验收标准

- 发布层不包含候选或争议组织身份。
- 发布清单记录每张表的输入、输出和过滤数量。
- 发布层所有引用闭合且 Schema 0 严重错误。
- 研究层数据未被删除或覆盖。
