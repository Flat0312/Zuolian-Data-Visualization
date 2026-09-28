# Phase 5 · 关系审核准确率报告

> 生成脚本：``research/analysis/analyze_relation_review.py``（幂等）
> 输入：``phase5_relation_review_adjudicated.csv`` 的 ``human_verdict`` 人工裁决列

## 0. 口径声明

本报告只统计人工已裁决行；AI 建议列不进入任何准确率分子分母。
关系成立率把 wrong_type 计为成立（关系存在、类型需改），类型准确率只认 correct。

## 1. 总体（仅已裁决行）

- 已裁决 400 / 400，待裁决 0
- 裁决分布：correct 135 / wrong_type 26 / not_supported 239 / contradicted 0
- 关系成立率（correct+wrong_type）：**40.2%**
- 类型准确率（correct）：**33.8%**
- not_supported 率：59.8%；contradicted 率：0.0%
- 人工与 AI 建议一致率：100.0%（一致性仅供交叉参考，不替代人工判定）

## 分层 · 风险层

| 分层 | 已裁决 | 成立率 | 类型准确率 |
| --- | ---: | ---: | ---: |
| critical | 184 | 47.8% | 44.0% |
| high | 46 | 19.6% | 13.0% |
| low | 160 | 39.4% | 29.4% |
| medium | 10 | 10.0% | 10.0% |

## 分层 · 证据分诊层

| 分层 | 已裁决 | 成立率 | 类型准确率 |
| --- | ---: | ---: | ---: |
| associated | 113 | 100.0% | 80.5% |
| insufficient | 239 | 0.0% | 0.0% |
| support | 48 | 100.0% | 91.7% |

## 修订规则（按错误阈值确定性生成）

| 维度 | 分层 | 已裁决 | 错误数 | 错误率 | 主导错误 | 建议 |
| --- | --- | ---: | ---: | ---: | --- | --- |
| risk | critical | 184 | 103 | 56.0% | not_supported | critical 层以 not_supported 为主，建议整层按裁决模式修订（降级待证 / 改类型）后复算 |
| risk | high | 46 | 40 | 87.0% | not_supported | high 层以 not_supported 为主，建议整层按裁决模式修订（降级待证 / 改类型）后复算 |
| risk | low | 160 | 113 | 70.6% | not_supported | low 层以 not_supported 为主，建议整层按裁决模式修订（降级待证 / 改类型）后复算 |
| risk | medium | 10 | 9 | 90.0% | not_supported | medium 层以 not_supported 为主，建议整层按裁决模式修订（降级待证 / 改类型）后复算 |
| confidence | low | 170 | 98 | 57.6% | not_supported | low 层以 not_supported 为主，建议整层按裁决模式修订（降级待证 / 改类型）后复算 |
| confidence | medium | 230 | 167 | 72.6% | not_supported | medium 层以 not_supported 为主，建议整层按裁决模式修订（降级待证 / 改类型）后复算 |
| support | associated | 113 | 22 | 19.5% | wrong_type | associated 层以 wrong_type 为主，建议整层按裁决模式修订（降级待证 / 改类型）后复算 |
| support | insufficient | 239 | 239 | 100.0% | not_supported | insufficient 层以 not_supported 为主，建议整层按裁决模式修订（降级待证 / 改类型）后复算 |
| type | 交游 | 112 | 77 | 68.8% | not_supported | 交游 层以 not_supported 为主，建议整层按裁决模式修订（降级待证 / 改类型）后复算 |
| type | 亲属关系 | 5 | 5 | 100.0% | not_supported | 亲属关系 层以 not_supported 为主，建议整层按裁决模式修订（降级待证 / 改类型）后复算 |
| type | 创作合作 | 13 | 11 | 84.6% | not_supported | 创作合作 层以 not_supported 为主，建议整层按裁决模式修订（降级待证 / 改类型）后复算 |
| type | 合作 | 31 | 21 | 67.7% | not_supported | 合作 层以 not_supported 为主，建议整层按裁决模式修订（降级待证 / 改类型）后复算 |
| type | 同属组织 | 67 | 30 | 44.8% | not_supported | 同属组织 层以 not_supported 为主，建议整层按裁决模式修订（降级待证 / 改类型）后复算 |
| type | 地下通讯 | 10 | 9 | 90.0% | not_supported | 地下通讯 层以 not_supported 为主，建议整层按裁决模式修订（降级待证 / 改类型）后复算 |
| type | 师生关系 | 4 | 4 | 100.0% | not_supported | 师生关系 层以 not_supported 为主，建议整层按裁决模式修订（降级待证 / 改类型）后复算 |
| type | 待核验 | 103 | 68 | 66.0% | not_supported | 待核验 层以 not_supported 为主，建议整层按裁决模式修订（降级待证 / 改类型）后复算 |
| type | 文学论战 | 9 | 8 | 88.9% | not_supported | 文学论战 层以 not_supported 为主，建议整层按裁决模式修订（降级待证 / 改类型）后复算 |
| type | 签名联署 | 19 | 13 | 68.4% | not_supported | 签名联署 层以 not_supported 为主，建议整层按裁决模式修订（降级待证 / 改类型）后复算 |
| type | 纪念/悼念 | 15 | 12 | 80.0% | not_supported | 纪念/悼念 层以 not_supported 为主，建议整层按裁决模式修订（降级待证 / 改类型）后复算 |
| type | 论战 | 4 | 4 | 100.0% | not_supported | 论战 层以 not_supported 为主，建议整层按裁决模式修订（降级待证 / 改类型）后复算 |
