# 时空与枚举审计报告（Agent B，只读审计）

- 数据快照：`D:\1大创\左联知识库项目\data\processed`（读取日期 2026-09-04，只读；未修改任何生产数据）
- events 时间列（start_date/end_date/date_certainty）：缺失
- 事件级候选：104 条；分布：{'generic_city_attachment': 83, 'same_name_different_date': 8, 'duplicate_canonical_event_key': 7, 'missing_place_id': 5, 'missing_time_columns': 1}
- 地点级候选：39 条；分布：{'coordinate_quality': 23, 'merge_candidate': 8, 'same_coordinate_different_name': 8}
- participant_role 现值分布：{'直接参与者': 111, 'unclear': 90, '待核': 17, '关联人物': 4}
- 泛化挂接到城市级地点（上海）的事件：83 个
- 无 place_id 但有历史地点名的事件：5 个

## 审计口径

1. **时间字段**：缺列或缺值均登记为候选，建议按 event_date/date_precision 派生（日→exact、月→approximate_month、年→approximate_year、空→uncertain）。
2. **canonical_event_key**：重复 key 与『同名不同日期』分别登记；同日同名疑似重复条目只提示人工合并流程，不自动合并。
3. **participant_role**：以 直接参与者/关联人物/待核/不明 为规范枚举；legacy 值（unclear/发起人 等）给出映射建议；非预设值一律人工归类。
4. **地点合并**：仅规则式候选（名称组 + 同坐标异名检测），全部 pending_human_review；涉及历史沿革（如内山书店 1929 年迁入四川北路2050号）须人工核对后再执行。
5. **城市级泛化**：挂接到『上海』等城市级地点的事件逐条列出，建议按证据细化到街区/建筑级，无法细化者保留并在展示层标注精度。
6. **坐标质量**：confidence=low、精度未知、缺坐标、城市级概略坐标逐条列出，建议补坐标来源后人工定级。

## 证据局限与边界

- 本审计只描述当前数据的结构与口径问题，**不构成对任何历史事实的判定**。
- 所有 decision=pending_human_review；合并/改名/改期/挂接调整均需人工裁决后另行执行。
- 『上海东方旅社、中山旅社秘密会议』这类复合名地点，其与『东方旅社』『中山旅社』的层级关系（事件名 vs 地点名）需人工先厘清语义再谈合并。
