# AGENTS.md

## 项目定位

本项目是大学生创新创业训练计划项目，不是教学产品。目标是以可追溯历史证据研究左联人物、组织、事件与社会关系网络，并完成研究展示和答辩交付。

## 工作边界

- `data/processed/` 是研究主数据；`data/publish/` 是脚本生成的展示发布层。
- 不得用 `persons.role` 自动确认正式成员身份；组织身份以证据台账和判定规则为准。
- 候选、争议和低置信度记录不得在展示文案中写成确定史实。
- `research/archive/` 是历史归档，不作为当前实现真值。
- 修改应保持最小范围，不顺手清理全仓历史 lint 或重写研究资产。

## 常用命令

```powershell
pwsh ./tasks.ps1 run
pwsh ./tasks.ps1 test
pwsh ./tasks.ps1 lint
pwsh ./tasks.ps1 build-static
python research/analysis/build_publish_data.py
```

## 验收要求

涉及数据或发布逻辑的修改至少验证：

```powershell
python -m pytest -v
python -c "from kb_schema import validate_data_dir; r=validate_data_dir('data/processed'); print(len(r.errors), len(r.warnings))"
python build_static_site.py
```

详细设计与阶段状态见 `docs/superpowers/specs/2026-06-05-research-and-presentation-dual-layer-design.md` 和 `task_plan.md`。
