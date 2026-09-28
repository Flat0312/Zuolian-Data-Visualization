# 设计与方案文档

本目录存放项目的架构设计与各阶段实施方案，**受版本管理**。

## 为什么不在 `docs/` 下

`docs/` 被 `.gitignore` 忽略（它是 `build_static_site.py` 的静态站输出目录）。
这些文档原先位于 `docs/superpowers/specs|plans/`，本地可读，但在远端仓库与 GitHub 上是死链，
评审者点不开——而双层架构设计与实施方案正是本项目方法论的主干。2026-09-28 迁至此处。

## 文件

| 文件 | 内容 |
| --- | --- |
| `specs/2026-06-05-research-and-presentation-dual-layer-design.md` | 研究层 / 发布层双层架构设计（AGENTS.md 与 task_plan.md 引用的规范入口） |
| `plans/2026-06-05-full-project-implementation-roadmap.md` | 完整实施方案与阶段划分 |
| `plans/2026-06-05-phase-1-org-membership-evidence-plan.md` | Phase 1 组织身份证据台账方案 |
| `plans/2026-06-05-phase-2-fact-evidence-plan.md` | Phase 2 事实级证据层方案 |
| `plans/2026-06-05-phase-3-research-publish-separation-plan.md` | Phase 3 研究层与发布层分离方案 |

## 迁移溯源

- 迁移日期：2026-09-28；方式：文件移动，内容未改写。
- 原路径：`docs/superpowers/specs/` 与 `docs/superpowers/plans/`。
- `research/briefs/content-upgrade-2026-09-06/baseline.json` 与
  `research/briefs/content-upgrade-overnight-2026-09-06/baseline.json` 中以旧路径登记了
  `2026-06-05-research-and-presentation-dual-layer-design.md` 的 SHA-256（`dc04a008…`）。
  这两份是历史批次的审计基线，按「历史不可变」原则未改写；核对该哈希时请按本目录的新路径取文件。
- 仍留在 `docs/superpowers/` 下的只有答辩相关存根与旧版 PPTX，不属于本目录范围。
