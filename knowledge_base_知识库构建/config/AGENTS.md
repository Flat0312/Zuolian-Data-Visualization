# 项目目标

本项目的目标不是简单修正 Excel，而是生成“可作为知识库入库前最终数据源”的规范化数据包。

输入包括：
1. 一个候选结构化表：`input/raw_excel/《左联相关档案资源目录》.xlsx`
2. 三份原始文本证据：
   - `input/raw_texts/日记全编：全2册 (鲁迅 著) (Z-Library).txt`
   - `input/raw_texts/左联词典.txt`
   - `input/raw_texts/左联史.txt`

输出必须是结构化、可追溯、可校验、可人工复核的数据包，而不是仅仅一个“修过的 Excel”。

# 数据源定位

请将不同输入源按以下角色处理：

## A. Excel
- 视为“候选事实表”
- 不视为最终事实
- 其中的人物关系、事件时间、地点、人物字段都需要重新校验

## B. 原始文本
- 视为“本地证据层”
- 用于抽取、核对、反证和补证
- 不可直接整段入库
- 不可把 OCR 结果未经清洗直接当最终事实

## C. 联网结果
- 视为“公开交叉核验层”
- 用于确认关键事实、消解冲突、补充标准事件名和标准地点
- 若公开来源冲突，则不得强行定稿，必须进入人工复核队列

# 证据优先级

当不同来源冲突时，按以下顺序处理：

1. 明确、可定位的原始文本证据
2. 权威公开机构页面或可核查史料整理页面
3. 研究型整理文本
4. Excel 原始候选表的已有字段
5. 基于文本共现的自动推断

注意：
- “自动推断”不能直接变成最终事实
- 仅凭共现不得认定人物关系类型
- 不能把“同属左联”直接写成“人物-人物的组织隶属”

# 最终必须生成的文件

所有输出放入 `output/final_dataset/`、`output/logs/`、`output/reports/`。

## 1. persons.csv
字段至少包括：
- person_id
- canonical_name
- aliases
- gender_if_known
- birth_date_if_known
- death_date_if_known
- description
- source_url
- local_evidence
- confidence
- review_status

## 2. organizations.csv
字段至少包括：
- org_id
- canonical_name
- aliases
- org_type
- description
- source_url
- local_evidence
- confidence
- review_status

## 3. places.csv
字段至少包括：
- place_id
- canonical_name
- historical_name
- current_address
- city
- province
- country
- lat_if_known
- lng_if_known
- description
- source_url
- local_evidence
- confidence
- review_status

## 4. events.csv
字段至少包括：
- event_id
- canonical_name
- aliases
- event_type
- start_date
- end_date
- date_precision
- historical_location_id
- current_place_note
- summary
- source_url
- local_evidence
- confidence
- review_status

## 5. person_relations.csv
字段至少包括：
- relation_id
- source_person_id
- target_person_id
- relation_type
- start_date_if_known
- end_date_if_known
- evidence_text
- source_url
- local_evidence
- confidence
- review_status

## 6. org_memberships.csv
字段至少包括：
- membership_id
- person_id
- org_id
- role_if_known
- start_date_if_known
- end_date_if_known
- evidence_text
- source_url
- local_evidence
- confidence
- review_status

## 7. event_participants.csv
字段至少包括：
- ep_id
- event_id
- entity_type
- entity_id
- role_in_event
- evidence_text
- source_url
- local_evidence
- confidence
- review_status

## 8. sources.csv
字段至少包括：
- source_id
- source_title
- source_type
- source_path_or_url
- publisher
- access_date
- reliability_note

## 9. review_queue.csv
收纳以下内容：
- 来源冲突
- 证据不足
- OCR 严重污染
- 实体未对齐
- 关系类型不确定
- 事件边界不清
- 时间精度不足但原表伪装成完整日期

## 10. correction_log.csv
记录所有自动修改：
- source_sheet
- row_number
- primary_key
- field_name
- original_value
- new_value
- issue_type
- correction_reason
- source_url
- local_evidence
- confidence
- needs_manual_review

## 11. validation_report.md
必须包含：
- 输入文件清单
- 总记录数
- 自动修正数
- 自动核验通过数
- 待人工复核数
- 冲突记录数
- 缺失来源数
- 缺失 ID 数
- 重名实体数
- 重复事件簇数
- 未解决实体对齐数
- 关键高风险问题摘要

## 12. workbook_corrected.xlsx
- 保留原始 workbook
- 在不覆盖原文件的前提下生成修正版 workbook
- 至少新增 corrected sheet、review sheet、summary sheet

# 建模规则

## 实体拆分
必须拆分以下实体：
- 人物
- 组织
- 地点
- 事件
- 来源

## 关系拆分
必须拆分以下关系：
- 人物-人物关系
- 人物-组织隶属关系
- 事件-参与者关系

禁止：
- 把人物-人物关系写成组织隶属
- 把事件参与人物塞进事件主表的一个长文本字段
- 把历史地点和今址混写为同一字段

# 日期规则

日期必须显式记录精度。

允许值仅为：
- 年
- 月
- 日
- 区间
- 未知

规则：
- 只能确认年份时，不得补成 YYYY-01-01
- 只能确认月份时，不得补成 YYYY-MM-01
- 区间事件必须填写 start_date 和 end_date
- 无法确认时写未知，并进入 review_queue

# 人物关系规则

仅在证据足够时，才能写入明确关系类型，例如：
- 亲属
- 师生
- 通信
- 合作
- 交往
- 论战
- 同属组织
- 共同活动
- 悼念/纪念关联

如果只是共现或弱证据：
- 不得强行定类
- 写入 `needs_manual_review`
- 可暂记为 `待核验`

特别规则：
- 如果 Source 和 Target 都是人物，则不得使用“组织隶属”作为人物关系类型
- 组织隶属必须落在 `org_memberships.csv`

# 事件规则

事件必须区分：
- 标准事件名
- 时间
- 历史发生地点
- 今址说明
- 事件参与者
- 参与角色

必须重点检查：
- 左联成立大会
- 左联五烈士遇难相关事件
- 鲁迅与柔石会面相关事件
- 内山书店相关事件
- 其他重复出现且时间地点冲突的事件簇

# 本地文本使用规则

对三份 txt 必须执行以下处理：
1. 清洗编码、去除明显 OCR 噪声
2. 按段落或句子切分
3. 建立可追溯的本地证据片段
4. 抽取人物、组织、地点、事件候选
5. 将本地证据片段写入 `local_evidence` 字段或日志中

注意：
- local_evidence 应是简短可追溯片段，不要整段复制大文本
- OCR 明显错误的片段不得直接作为高置信证据
- 若本地文本与公开来源冲突，必须进入 review_queue

# 执行顺序

必须按以下顺序执行：

1. 读取并分析 Excel 结构
2. 读取并清洗三份 txt
3. 从 Excel 和 txt 中抽取候选实体、候选关系、候选事件
4. 做实体对齐与去重
5. 对高风险记录进行联网核验
6. 重构为规范化数据表
7. 生成 corrected workbook
8. 生成日志和报告
9. 运行校验
10. 打印实际输出文件列表

# 质量底线

- 不得覆盖原始 Excel
- 不得伪造来源
- 不得把推测写成事实
- 宁可少改，不可乱改
- 证据不足时必须进入人工复核队列
- 若缺失必需输出文件，则任务不算完成

# 完成判定

只有在以下条件同时满足时，任务才算完成：
1. 目标输出文件全部生成
2. validation_report.md 已生成
3. correction_log.csv 已生成
4. review_queue.csv 已生成
5. workbook_corrected.xlsx 已生成
6. 已打印 output 目录下的全部文件路径