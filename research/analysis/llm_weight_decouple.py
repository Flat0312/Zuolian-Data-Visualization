"""
llm_weight_decouple.py
════════════════════════════════════════════════════════════
左联知识图谱 — 三阶段流水线脚本
  阶段一：LLM API 纯史料权重盲测（Weight 重评）
  阶段二：本地标签解耦（Relation_Type 归一化）
  阶段三：维度重组与最终导出

依赖：
  pip install pandas openai
════════════════════════════════════════════════════════════
"""

import re
import time
import json
import os
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

import pandas as pd
from openai import OpenAI

# ════════════════════════════════════════════════════════════
# 0. 全局配置
# ════════════════════════════════════════════════════════════

# ---------- 文件路径 ----------
XLSX_PATH       = Path(r"d:\1大创\《左联相关档案资源目录》.xlsx")
SHEET2_CSV      = Path(r"d:\1大创\cleaned_Sheet2.csv")
OUTPUT_CSV      = Path(r"d:\1大创\Decoupled_Weighted_Sheet2.csv")
INTERIM_CSV     = Path(r"d:\1大创\_llm_weights_interim.csv")   # 中间缓存，断点续跑用

# ---------- OpenAI 配置 ----------
# 通过环境变量注入，不在代码中保存密钥：
#   OPENAI_API_KEY=...
#   OPENAI_BASE_URL=https://api.openai.com/v1（可选）
#   OPENAI_MODEL=gpt-4o（可选）
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "").strip()
OPENAI_BASE_URL = os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1").strip()
MODEL_NAME = os.getenv("OPENAI_MODEL", "gpt-4o").strip()

# ---------- 调用策略 ----------
CONTEXT_MAX_CHARS = 2000    # 每对人物合并 Context 的最大字符数
SLEEP_BETWEEN_API = 1.0     # 每次 API 调用后的间隔（秒），避免触发限速
MAX_RETRIES       = 3       # API 失败后的最大重试次数

# ---------- 真实亲属白名单 ----------
# 若某对 ID 确为亲属关系，在此列表中添加 frozenset，防止被脚本自动改标
# 示例：frozenset(["ZLH-001", "ZLH-002"])
KINSHIP_WHITELIST: list[frozenset] = [
    # frozenset(["ZLH-XXX", "ZLH-YYY"]),
]

# ════════════════════════════════════════════════════════════
# 1. 工具函数：读取 xlsx（XML 直解，兼容 WPS 缺少 sharedStrings.xml）
# ════════════════════════════════════════════════════════════

_NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def _cell_value(cell_elem) -> str:
    """解析单元格值：优先内联字符串(is)，其次普通值(v)"""
    is_t = cell_elem.findall(".//main:is//main:t", _NS)
    if is_t:
        return "".join(e.text or "" for e in is_t)
    v = cell_elem.find("main:v", _NS)
    return v.text if v is not None else ""


def read_xlsx_sheet(xlsx_path: Path, sheet_xml: str) -> pd.DataFrame:
    """
    用 zipfile + ElementTree 直接解析 xlsx 内部 sheet XML。
    完全绕过 openpyxl 对 sharedStrings.xml 的依赖。
    """
    with zipfile.ZipFile(xlsx_path, "r") as z:
        with z.open(sheet_xml) as f:
            root = ET.parse(f).getroot()

    rows_data = []
    for row_elem in root.findall(".//main:sheetData/main:row", _NS):
        rows_data.append([_cell_value(c) for c in row_elem.findall("main:c", _NS)])

    if not rows_data:
        return pd.DataFrame()

    headers  = rows_data[0]
    data     = rows_data[1:]
    ncols    = len(headers)
    data     = [r + [""] * (ncols - len(r)) if len(r) < ncols else r[:ncols] for r in data]
    return pd.DataFrame(data, columns=headers)


# ════════════════════════════════════════════════════════════
# 2. 阶段一：LLM 权重盲测
# ════════════════════════════════════════════════════════════

def make_pair_key(a: str, b: str) -> tuple[str, str]:
    """双向归一：始终将较小的 ID 排在前，使 (A,B) == (B,A)"""
    return (min(a, b), max(a, b))


def aggregate_pairs(df: pd.DataFrame, name_map: dict) -> pd.DataFrame:
    """
    将关系表按人物对（双向视为同一对）GroupBy 聚合：
    - 拼接所有 Context（截断至 CONTEXT_MAX_CHARS 字）
    - 映射真实姓名
    返回一个以 (id_a, id_b) 为主键的聚合表。
    """
    # 生成规范化的 pair_key
    df = df.copy()
    df["pair_key"] = df.apply(
        lambda r: make_pair_key(str(r["Source_ID"]), str(r["Target_ID"])), axis=1
    )
    df["id_a"] = df["pair_key"].apply(lambda x: x[0])
    df["id_b"] = df["pair_key"].apply(lambda x: x[1])

    # 聚合：拼接 Context，用「/」分隔
    agg = (
        df.groupby(["id_a", "id_b"])["Context"]
        .apply(lambda texts: " / ".join(str(t) for t in texts if str(t).strip()))
        .reset_index()
    )
    agg.columns = ["id_a", "id_b", "combined_context"]

    # 截断至 CONTEXT_MAX_CHARS
    agg["combined_context"] = agg["combined_context"].str[:CONTEXT_MAX_CHARS]

    # 映射真实姓名（无匹配则用 ID 代替）
    agg["name_a"] = agg["id_a"].map(name_map).fillna(agg["id_a"])
    agg["name_b"] = agg["id_b"].map(name_map).fillna(agg["id_b"])

    return agg


def build_prompt(name_a: str, name_b: str, context: str) -> str:
    """
    构建防偏见 Prompt：
    - 不传 Relation_Type，避免先入为主的误导
    - 要求 LLM 只依据史料原文给出 1-5 分的权重
    """
    return f"""你是一位专业的近代中国文学史与左翼文化运动研究专家。
请仅根据以下史料原文（勿参考任何外部知识），评估人物【{name_a}】与人物【{name_b}】在历史档案中的关联强度。

【史料片段】
{context}

【评分标准】
5分：两人在史料中高频共同出现、深度共事（如并肩领导组织、长期通信往来、紧密协作活动）。
3-4分：有明确的历史交集与合作，但深度或频率较有限。
1-2分：史料中仅偶尔提及、边缘交游、或仅为家乡亲眷等弱关联。

【强制要求】
请只输出以下 JSON，不要包含任何其他文字：
{{"weight": 整数(1-5), "reason": "简明历史依据（50字以内）"}}"""


def call_llm_with_retry(client: OpenAI, prompt: str) -> dict:
    """
    调用 LLM，强制解析 JSON 返回。
    内置重试机制（最多 MAX_RETRIES 次），每次失败后指数退避。
    返回 {"weight": int, "reason": str}；全部失败则返回 {"weight": -1, "reason": "API调用失败"}。
    """
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = client.chat.completions.create(
                model=MODEL_NAME,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.1,       # 低温确保稳定输出
                max_tokens=150,
                response_format={"type": "json_object"},   # 强制 JSON 模式
            )
            raw = response.choices[0].message.content.strip()
            parsed = json.loads(raw)

            # 校验 weight 范围
            weight = int(parsed.get("weight", -1))
            if not (1 <= weight <= 5):
                raise ValueError(f"weight 超出 1-5 范围：{weight}")

            return {"weight": weight, "reason": parsed.get("reason", "")}

        except Exception as e:
            wait = 2 ** attempt   # 指数退避：2s, 4s, 8s
            print(f"    [重试 {attempt}/{MAX_RETRIES}] 错误：{e}，{wait}s 后重试...")
            time.sleep(wait)

    return {"weight": -1, "reason": "API调用失败（已重试）"}


def run_phase1_llm(df_pairs: pd.DataFrame, client: OpenAI) -> pd.DataFrame:
    """
    阶段一主函数：遍历所有人物对，调用 LLM 进行权重盲测。
    支持断点续跑：若中间缓存文件 _llm_weights_interim.csv 存在，跳过已完成的对。
    """
    print("\n[阶段一] 开始 LLM 权重盲测...")
    total = len(df_pairs)

    # 加载已有中间结果（断点续跑）
    done_keys: set = set()
    interim_rows: list = []
    if INTERIM_CSV.exists():
        df_interim = pd.read_csv(INTERIM_CSV, encoding="utf-8-sig")
        for _, r in df_interim.iterrows():
            done_keys.add((r["id_a"], r["id_b"]))
            interim_rows.append(r.to_dict())
        print(f"    检测到缓存：已完成 {len(done_keys)} / {total} 对，跳过...")

    results = list(interim_rows)

    for idx, row in df_pairs.iterrows():
        pair = (row["id_a"], row["id_b"])
        if pair in done_keys:
            continue   # 断点续跑：跳过已处理

        progress = idx + 1
        print(f"    [{progress}/{total}] {row['name_a']} × {row['name_b']} ...", end=" ")

        prompt = build_prompt(row["name_a"], row["name_b"], row["combined_context"])
        result = call_llm_with_retry(client, prompt)

        print(f"weight={result['weight']}  ｜ {result['reason'][:40]}")

        results.append({
            "id_a":    row["id_a"],
            "id_b":    row["id_b"],
            "name_a":  row["name_a"],
            "name_b":  row["name_b"],
            "llm_weight":  result["weight"],
            "llm_reason":  result["reason"],
        })

        # 每处理完一对立即写入中间缓存（防止意外中断丢失进度）
        pd.DataFrame(results).to_csv(INTERIM_CSV, index=False, encoding="utf-8-sig")

        time.sleep(SLEEP_BETWEEN_API)

    df_weights = pd.DataFrame(results)
    print(f"[阶段一] 完成！共处理 {len(df_weights)} 对人物。")
    return df_weights


# ════════════════════════════════════════════════════════════
# 3. 阶段二：本地标签解耦与亲属修正
# ════════════════════════════════════════════════════════════

def run_phase2_decouple(df: pd.DataFrame, role_map: dict) -> pd.DataFrame:
    """
    阶段二主函数：
    (A) 正则剥离所有「强关联-」「弱关联-」等强弱修饰前缀
    (B) 对剩余裸「亲属」标签，结合双方 Role 进行逻辑修正
    """
    print("\n[阶段二] 开始标签解耦...")

    df = df.copy()

    # ── (A) 剥离强弱修饰前缀 ─────────────────────────────────────────────
    # 覆盖所有可能的前缀写法：强关联-、弱关联-、强关系-、弱关系- 等
    prefix_pattern = re.compile(r"^(强|弱)(关联|关系)\s*[-—–]\s*")
    df["Relation_Type"] = df["Relation_Type"].astype(str).apply(
        lambda x: prefix_pattern.sub("", x).strip()
    )

    before_kinship = (df["Relation_Type"] == "亲属").sum()
    print(f"    前缀剥离完成，剩余裸「亲属」标签：{before_kinship} 条")

    # 定义左联角色圈层
    INTERNAL_ROLES = {"核心领导", "普通成员"}
    EXTERNAL_ROLES = {"外围联络人", "相关人士"}

    # ── (B) 亲属标签逻辑修正 ─────────────────────────────────────────────
    def fix_kinship(row) -> str:
        if row["Relation_Type"] != "亲属":
            return row["Relation_Type"]

        src, tgt = str(row["Source_ID"]), str(row["Target_ID"])

        # 白名单保护：确认为真实亲属的跳过修正
        if frozenset([src, tgt]) in KINSHIP_WHITELIST:
            return "亲属"

        src_role = role_map.get(src, "未知")
        tgt_role = role_map.get(tgt, "未知")

        # 规则 A：双方均为内部成员 -> 组织隶属（可能被误判为亲属）
        if src_role in INTERNAL_ROLES and tgt_role in INTERNAL_ROLES:
            return "组织隶属"
        # 规则 B：含外围/相关人士 -> 交游
        elif src_role in EXTERNAL_ROLES or tgt_role in EXTERNAL_ROLES:
            return "交游"
        # 默认兜底
        else:
            return "交游"

    df["Relation_Type"] = df.apply(fix_kinship, axis=1)

    after_kinship = (df["Relation_Type"] == "亲属").sum()
    print(f"    亲属修正完成：{before_kinship} 条 -> {after_kinship} 条保留为亲属")

    # ── 统计最终分布 ─────────────────────────────────────────────────────
    print("    解耦后 Relation_Type 分布：")
    for rel, cnt in df["Relation_Type"].value_counts().items():
        print(f"      {rel:20s}  {cnt}")

    return df


# ════════════════════════════════════════════════════════════
# 4. 阶段三：权重映射回原表 & 导出
# ════════════════════════════════════════════════════════════

def run_phase3_export(df_relations: pd.DataFrame, df_weights: pd.DataFrame) -> pd.DataFrame:
    """
    阶段三主函数：
    (A) 将 LLM 权重通过 (id_a, id_b) 映射回原始行级 DataFrame
    (B) 替换旧 Weight 列，新增 LLM_Reason 列
    (C) 导出最终 CSV
    """
    print("\n[阶段三] 开始权重映射与导出...")

    df = df_relations.copy()

    # 为原始 DataFrame 生成双向归一的 pair_key
    df["id_a"] = df.apply(lambda r: min(str(r["Source_ID"]), str(r["Target_ID"])), axis=1)
    df["id_b"] = df.apply(lambda r: max(str(r["Source_ID"]), str(r["Target_ID"])), axis=1)

    # 构建 (id_a, id_b) -> {weight, reason} 的查找表
    weight_lookup = df_weights.set_index(["id_a", "id_b"])[["llm_weight", "llm_reason"]].to_dict("index")

    def get_weight(row):
        key = (row["id_a"], row["id_b"])
        return weight_lookup.get(key, {}).get("llm_weight", None)

    def get_reason(row):
        key = (row["id_a"], row["id_b"])
        return weight_lookup.get(key, {}).get("llm_reason", "")

    # 覆盖旧 Weight 列（完全作废原有权重）
    df["Weight"]     = df.apply(get_weight, axis=1)
    df["LLM_Reason"] = df.apply(get_reason, axis=1)

    # 清理辅助列
    df.drop(columns=["id_a", "id_b"], inplace=True)

    # 统计覆盖情况
    missing = df["Weight"].isna().sum()
    if missing:
        print(f"    警告：{missing} 行未能匹配到 LLM 权重（可能 API 调用失败），权重为 NaN")

    # 导出
    df.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")
    print(f"    导出成功：{OUTPUT_CSV}  共 {len(df)} 行")

    return df


# ════════════════════════════════════════════════════════════
# 5. 主流程
# ════════════════════════════════════════════════════════════

def main():
    print("=" * 60)
    print("  左联知识图谱 — LLM权重评估 & 标签解耦 流水线")
    print("=" * 60)

    if not OPENAI_API_KEY:
        raise RuntimeError("缺少 OPENAI_API_KEY。请先设置环境变量后再运行。")

    # ── 初始化 OpenAI 客户端 ─────────────────────────────────
    client = OpenAI(api_key=OPENAI_API_KEY, base_url=OPENAI_BASE_URL)

    # ── 读取 Sheet1（实体表） ────────────────────────────────
    print("\n[读取] Sheet1（实体表）...")
    df_entities = read_xlsx_sheet(XLSX_PATH, "xl/worksheets/sheet1.xml")
    print(f"    {len(df_entities)} 行，列名：{list(df_entities.columns)}")

    # 构建 ID->姓名 和 ID->角色 映射字典
    name_map = dict(zip(df_entities["Entity_ID"], df_entities["True_Name"]))
    role_map = dict(zip(df_entities["Entity_ID"], df_entities["Role"]))
    print(f"    name_map 样例：{list(name_map.items())[:3]}")

    # ── 读取 Sheet2（关系表） ────────────────────────────────
    print("\n[读取] 关系表 cleaned_Sheet2.csv...")
    df_relations = pd.read_csv(SHEET2_CSV, encoding="utf-8-sig")
    print(f"    {len(df_relations)} 行")

    # ════════════════════════════════════════════════════════
    # 阶段一：LLM 权重盲测
    # ════════════════════════════════════════════════════════
    df_pairs   = aggregate_pairs(df_relations, name_map)
    print(f"\n[聚合] 双向归一后共 {len(df_pairs)} 对独立人物对")

    df_weights = run_phase1_llm(df_pairs, client)

    # ════════════════════════════════════════════════════════
    # 阶段二：标签解耦
    # ════════════════════════════════════════════════════════
    df_relations = run_phase2_decouple(df_relations, role_map)

    # ════════════════════════════════════════════════════════
    # 阶段三：权重映射 & 导出
    # ════════════════════════════════════════════════════════
    run_phase3_export(df_relations, df_weights)

    print("\n" + "=" * 60)
    print("  全部完成！")
    print(f"  最终输出：{OUTPUT_CSV}")
    print(f"  LLM中间缓存：{INTERIM_CSV}")
    print("=" * 60)


if __name__ == "__main__":
    main()
