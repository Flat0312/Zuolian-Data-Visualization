"""
左联知识图谱数据可靠性验证脚本
通过 GPT-4 API 对 Sheet2（关系）和 Sheet3（时空）数据进行事实核查
"""
import json
import os
import re
import time

import pandas as pd
from openai import OpenAI

# ══════════════════════════════════════════════════════════
# API 配置
# ══════════════════════════════════════════════════════════
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "").strip()
OPENAI_BASE_URL = os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1").strip()
MODEL = os.getenv("OPENAI_MODEL", "gpt-4o").strip()
client = None

# ══════════════════════════════════════════════════════════
# 文件路径
# ══════════════════════════════════════════════════════════
INPUT = '数据/输出结果/modified_zuolian.xlsx'
OUTPUT_REPORT = '数据/输出结果/验证报告.xlsx'

# ══════════════════════════════════════════════════════════
# 抽样策略：不对全部4000+行调用API，按策略抽样
# ══════════════════════════════════════════════════════════
SHEET2_SAMPLE_SIZE = 50   # Sheet2 抽样数
SHEET3_SAMPLE_SIZE = 30   # Sheet3 抽样数
BATCH_SIZE = 5            # 每次发给API的条目数
API_DELAY = 2             # 请求间隔(秒)，防限流


def call_gpt(system_prompt, user_prompt, max_retries=3):
    """调用 GPT-4，带重试"""
    if client is None:
        raise RuntimeError("OpenAI 客户端未初始化。")
    for attempt in range(max_retries):
        try:
            resp = client.chat.completions.create(
                model=MODEL,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt},
                ],
                temperature=0.1,
                max_tokens=2000,
            )
            return resp.choices[0].message.content.strip()
        except Exception as e:
            print(f"  API 调用失败 (第{attempt+1}次): {e}")
            if attempt < max_retries - 1:
                time.sleep(5 * (attempt + 1))
    return None


# ══════════════════════════════════════════════════════════
# Sheet2 关系数据验证
# ══════════════════════════════════════════════════════════
SYSTEM_PROMPT_SHEET2 = """你是一位中国现代文学史专家，专精左翼文学运动（1930年代）。
请对以下历史人物关系数据逐条进行事实核查，以JSON数组格式返回结果。

每条数据包含: 序号、人物A、人物B、关系类型、上下文、出处。
你需要判断:
1. reliability_score: 可靠性评分(1-5)，5=确凿史实，4=高度可信，3=基本可信，2=存疑，1=可能错误
2. verdict: "confirmed"(确认)/"plausible"(合理)/"uncertain"(存疑)/"incorrect"(错误)
3. reason: 简要说明判断依据(中文，30字以内)
4. correction: 如有错误，给出修正建议；无则为null

严格以JSON数组输出，不要添加其他文字。格式:
[{"序号":1,"reliability_score":4,"verdict":"confirmed","reason":"...","correction":null}, ...]"""

SYSTEM_PROMPT_SHEET3 = """你是一位中国现代文学史和上海历史地理专家。
请对以下左翼作家时空活动数据逐条进行事实核查，以JSON数组格式返回结果。

每条数据包含: 序号、人物、时间、历史地点、事件。
你需要判断:
1. reliability_score: 可靠性评分(1-5)
2. verdict: "confirmed"/"plausible"/"uncertain"/"incorrect"
3. reason: 判断依据(中文，30字以内)
4. location_correct: 地点是否准确(true/false)
5. date_correct: 日期是否准确(true/false)
6. correction: 修正建议(如有)

严格以JSON数组输出。格式:
[{"序号":1,"reliability_score":4,"verdict":"confirmed","reason":"...","location_correct":true,"date_correct":true,"correction":null}, ...]"""


def parse_json_response(text):
    """从GPT回复中提取JSON"""
    if not text:
        return []
    # 尝试直接解析
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    # 提取```json...```块
    m = re.search(r'```json?\s*(.*?)\s*```', text, re.DOTALL)
    if m:
        try:
            return json.loads(m.group(1))
        except json.JSONDecodeError:
            pass
    # 提取[...]
    m = re.search(r'\[.*\]', text, re.DOTALL)
    if m:
        try:
            return json.loads(m.group(0))
        except json.JSONDecodeError:
            pass
    return []


def verify_sheet2(df2, df1):
    """验证 Sheet2 关系数据"""
    id_to_name = dict(zip(df1['Entity_ID'], df1['True_Name']))

    # 分层抽样: 每种 Relation_Type 至少取几条
    samples = []
    for rt in df2['Relation_Type'].unique():
        subset = df2[df2['Relation_Type'] == rt]
        n = max(3, int(SHEET2_SAMPLE_SIZE * len(subset) / len(df2)))
        n = min(n, len(subset))
        samples.append(subset.sample(n=n, random_state=42))
    sample_df = pd.concat(samples).drop_duplicates().head(SHEET2_SAMPLE_SIZE)

    print(f"\n{'='*60}")
    print(f"Sheet2 关系验证: 抽样 {len(sample_df)}/{len(df2)} 条")
    print(f"{'='*60}")

    all_results = []

    for batch_start in range(0, len(sample_df), BATCH_SIZE):
        batch = sample_df.iloc[batch_start:batch_start + BATCH_SIZE]
        items = []
        for _, row in batch.iterrows():
            src_name = id_to_name.get(row['Source_ID'], row['Source_ID'])
            tgt_name = id_to_name.get(row['Target_ID'], row['Target_ID'])
            ctx = str(row['Context'])[:150] if pd.notna(row['Context']) else ''
            ref = str(row['Evidence_Ref'])[:100] if pd.notna(row['Evidence_Ref']) else ''
            items.append(
                f"序号{row['序号']}: 人物A={src_name}, 人物B={tgt_name}, "
                f"关系={row['Relation_Type']}, 上下文={ctx}, 出处={ref}"
            )

        prompt = "请核查以下历史关系数据:\n\n" + "\n\n".join(items)
        print(f"  验证第 {batch_start+1}-{batch_start+len(batch)} 条...")

        response = call_gpt(SYSTEM_PROMPT_SHEET2, prompt)
        results = parse_json_response(response)

        if results:
            for r in results:
                r['_batch'] = batch_start
            all_results.extend(results)
            scores = [r.get('reliability_score', 0) for r in results]
            print(f"    → 得到 {len(results)} 条结果, 平均分: {sum(scores)/len(scores):.1f}")
        else:
            print(f"    ⚠ 解析失败, 原始回复: {response[:200] if response else 'None'}")

        time.sleep(API_DELAY)

    return all_results, sample_df


def verify_sheet3(df3, df1):
    """验证 Sheet3 时空数据"""
    id_to_name = dict(zip(df1['Entity_ID'], df1['True_Name']))

    sample_df = df3.sample(n=min(SHEET3_SAMPLE_SIZE, len(df3)), random_state=42)

    print(f"\n{'='*60}")
    print(f"Sheet3 时空验证: 抽样 {len(sample_df)}/{len(df3)} 条")
    print(f"{'='*60}")

    all_results = []

    for batch_start in range(0, len(sample_df), BATCH_SIZE):
        batch = sample_df.iloc[batch_start:batch_start + BATCH_SIZE]
        items = []
        for _, row in batch.iterrows():
            name = id_to_name.get(row['Entity_ID'], row['Entity_ID'])
            items.append(
                f"序号{row['序号']}: 人物={name}, 时间={row['Timestamp']}, "
                f"地点={row['Hist_Loc']}, 事件={row['Event']}"
            )

        prompt = "请核查以下历史时空活动数据:\n\n" + "\n\n".join(items)
        print(f"  验证第 {batch_start+1}-{batch_start+len(batch)} 条...")

        response = call_gpt(SYSTEM_PROMPT_SHEET3, prompt)
        results = parse_json_response(response)

        if results:
            all_results.extend(results)
            scores = [r.get('reliability_score', 0) for r in results]
            print(f"    → 得到 {len(results)} 条结果, 平均分: {sum(scores)/len(scores):.1f}")
        else:
            print(f"    ⚠ 解析失败, 原始回复: {response[:200] if response else 'None'}")

        time.sleep(API_DELAY)

    return all_results, sample_df


# ══════════════════════════════════════════════════════════
# 主流程
# ══════════════════════════════════════════════════════════
if __name__ == '__main__':
    if not OPENAI_API_KEY:
        raise RuntimeError("缺少 OPENAI_API_KEY。请先设置环境变量后再运行。")

    client = OpenAI(
        api_key=OPENAI_API_KEY,
        base_url=OPENAI_BASE_URL,
    )

    print("加载数据...")
    df1 = pd.read_excel(INPUT, sheet_name='Sheet1')
    df2 = pd.read_excel(INPUT, sheet_name='Sheet2')
    df3 = pd.read_excel(INPUT, sheet_name='Sheet3')

    # 先测试API连通性
    print("测试 API 连接...")
    test = call_gpt("你是助手", "请回复'OK'两个字母")
    if test:
        print(f"  ✓ API 连接成功: {test[:50]}")
    else:
        print("  ✗ API 连接失败，请检查密钥和URL")
        exit(1)

    # 验证 Sheet2
    s2_results, s2_sample = verify_sheet2(df2, df1)

    # 验证 Sheet3
    s3_results, s3_sample = verify_sheet3(df3, df1)

    # ── 汇总报告 ──
    print(f"\n{'='*60}")
    print("汇总报告")
    print(f"{'='*60}")

    if s2_results:
        s2_scores = [r.get('reliability_score', 0) for r in s2_results]
        s2_verdicts = [r.get('verdict', '') for r in s2_results]
        print(f"\nSheet2 关系验证 ({len(s2_results)} 条):")
        print(f"  平均可靠性: {sum(s2_scores)/len(s2_scores):.2f}/5")
        for v in ['confirmed', 'plausible', 'uncertain', 'incorrect']:
            cnt = s2_verdicts.count(v)
            if cnt:
                print(f"  {v}: {cnt} ({cnt/len(s2_verdicts)*100:.0f}%)")

    if s3_results:
        s3_scores = [r.get('reliability_score', 0) for r in s3_results]
        s3_verdicts = [r.get('verdict', '') for r in s3_results]
        print(f"\nSheet3 时空验证 ({len(s3_results)} 条):")
        print(f"  平均可靠性: {sum(s3_scores)/len(s3_scores):.2f}/5")
        for v in ['confirmed', 'plausible', 'uncertain', 'incorrect']:
            cnt = s3_verdicts.count(v)
            if cnt:
                print(f"  {v}: {cnt} ({cnt/len(s3_verdicts)*100:.0f}%)")

    # ── 保存验证报告 ──
    s2_report = pd.DataFrame(s2_results) if s2_results else pd.DataFrame()
    s3_report = pd.DataFrame(s3_results) if s3_results else pd.DataFrame()

    # 标记存疑/错误条目
    problems = []
    for r in s2_results + s3_results:
        if r.get('verdict') in ('uncertain', 'incorrect'):
            problems.append(r)
    problems_df = pd.DataFrame(problems) if problems else pd.DataFrame()

    with pd.ExcelWriter(OUTPUT_REPORT, engine='openpyxl') as writer:
        if not s2_report.empty:
            s2_report.to_excel(writer, sheet_name='Sheet2验证', index=False)
        if not s3_report.empty:
            s3_report.to_excel(writer, sheet_name='Sheet3验证', index=False)
        if not problems_df.empty:
            problems_df.to_excel(writer, sheet_name='问题条目', index=False)

    print(f"\n✓ 验证报告已保存至: {OUTPUT_REPORT}")
    if problems:
        print(f"⚠ 发现 {len(problems)} 条存疑/错误条目，详见「问题条目」sheet")
