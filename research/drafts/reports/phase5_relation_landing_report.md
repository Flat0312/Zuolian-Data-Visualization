# Phase 5 关系裁决生产层落地报告（Tier 1：证据派生 supported）

> 执行脚本：`research/analysis/apply_phase5_relation_landing.py`（幂等）  
> 批次标记：`P5-LANDING-2026-09-28`  落地日期：2026-09-28

## 0. 授权与口径声明

- 授权人：用户（会话授权）；授权时间：2026-09-20；签核路径：C（全量按 AI 建议执行）。
- 授权语（逐字）：「我都没问题，你自己看着办吧」，登记于 `phase5_review_adjudication_record.md`。
- 裁决口径为**授权按建议执行**，非逐条独立人工复核；公开层展示与答辩引用均须带此口径。
- 本批公开层只接受 **derived `supported`**（证据派生），未使用 `human_adjudication` 通道，未把 critical/high 风险关系推入公开层，未改写 `relation_risk_level`。

## 1. 落地范围

- 候选关系：48 条（夜间轮判定 `evidence_support=support` ∧ 人工裁决 correct/wrong_type）。
- **双方佐证门通过：20 条**（引文空白归一后同时含两端姓名/别名；鲁迅日记按日记体免甲方自名）。只有通过的才落地。
- 双方佐证门未通过：28 条 → 全部转入 `phase5_quote_recapture_queue.csv`（`pending_human_review`，生产层零改动）。
- 引文逐字复核：23/23 通过（quote_sha256 与原文 content_hash 一致；OCR 全文字间带空格，须空白归一后比对）。
- 新增关系证据行：20 条。
- 新注册来源：1 条（鲁迅日记 1928年7月1日）。
- 关系类型更正：3 条。

## 2. 结果分布（仅本批 48 条）

| 落地后 publish_status | 条数 |
| --- | ---: |
| supported | 5 |
| pending_review | 14 |
| inferred | 1 |

全表发布状态：supported 5 / pending_review 2451 / inferred 1782（合计 4238，origin 全为 derived）。

## 3. 进入公开层的关系

| relation_id | 人物对 | 类型（落地后） | 证据 locator | 新证据行 |
| --- | --- | --- | --- | --- |
| REL-00046 | ZLH-001 → ZLH-048 | 交游 | 鲁迅日记 1928年7月1日 | RELE-10251 |
| REL-00059 | ZLH-001 → ZLH-071 | 交往 | 鲁迅日记 1929年5月20日 | RELE-10252 |
| REL-00097 | ZLH-001 → ZLH-128 | 通信 | 鲁迅日记 1928年11月24日 | RELE-10253 |
| REL-03289 | ZLH-058 → ZLH-120 | 签名联署 | 左联词典 第506页 | RELE-10266 |
| REL-03518 | ZLH-069 → ZLH-081 | 签名联署 | 左联史 第411页 | RELE-10269 |

## 4. 类型更正明细

| relation_id | 更正前 | 更正后 | 落地后状态 |
| --- | --- | --- | --- |
| REL-00059 | 时空共现 | 交往 | supported |
| REL-00309 | 待核验 | 同属组织 | pending_review |
| REL-03518 | 创作合作 | 签名联署 | supported |

## 5. 本批发现的系统性缺陷：引文窗口捕错

夜间轮记录的 `quote` 多数是按页码定位截取的固定窗口，而不是真正记载该关系的那一句，
导致 `reason` 字段描述的史实与 `quote` 内容不一致。48 条候选中：

- 20 条引文确实同时记载双方当事人 → 已落地；
- 28 条引文里找不到当事人（16 条两端都缺、11 条只缺乙方、1 条只缺甲方）。

未通过的 28 条按下述三类进入重捕队列，**生产层零改动**：

| proposed_action | 含义 | 条数 |
| --- | --- | ---: |
| `recapture_quote_then_regrade` | 同一本地原文中存在同时记载两人的叙述性段落，可重捕引文后再评级 | 9 |
| `cooccurrence_only_keep_associated` | 同窗共现仅为顿号人名罗列，属关联级，不得升为 support | 16 |
| `no_local_support_mark_insufficient` | 本地原文中找不到任何同窗共现，应判证据不足 | 3 |

结论：`evidence_support=support` 不能单独作为公开依据，必须再过双方佐证门。
该门已固化在 `research/analysis/quote_attestation.py`，可复用于后续所有关系补证批次。

## 6. 明确未做的事（避免夸大结论）

- **未改判既有 10249 条 `associated` 证据**：其 reviewer_note 写明「仅表示来源关联，不断言支持强度」，改判为 support 等于凭空提升证据等级。本批改为新增经逐字复核的 support 行，旧行原样保留。
- **未落地 associated 层 113 条**（人工裁决成立但证据仅为关联级）：证据等级不足以进入公开层，保留在研究层。
- **未把 239 条 `not_supported` 写为 `rejected`**：五态中 `rejected` 语义是「人工否定」（对应 contradicted，本批 0 条），`not_supported` 只表示证据不足；映射会把「未证实」夸大成「已证伪」。这些关系仍为 pending_review/inferred，不进入公开层。
- **未使用 `human_adjudication` 通道**：通过佐证门但被机器保守标记挡住的关系仍留在 pending_review，要公开它们需要逐条独立人工复核，不能由一次概括授权代行。
- **未对未通过佐证门的关系做任何类型更正或降级**：引文捕错不等于关系不成立，改类型或判不足都需要重捕引文后另行裁决。
- **未改写 `relation_risk_level`**：critical 计数保持 1974。

## 7. 复核方式

```powershell
python research/analysis/apply_phase5_relation_landing.py --dry-run   # 只校验不写
python research/analysis/apply_phase5_relation_landing.py             # 落地（二跑输出「无新增/已完成」）
python research/analysis/build_publish_data.py
python build_static_site.py
python -m pytest -q
```

逐条痕迹见 `phase5_relation_landing_ledger.csv`（20 行，每行带 quote_sha256、检索凭据与授权溯源）。
