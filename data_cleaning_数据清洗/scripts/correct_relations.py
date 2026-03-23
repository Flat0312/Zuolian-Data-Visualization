"""
correct_relations.py
────────────────────────────────────────────────────────────
根据实体角色，修正 Sheet2 中被误标为「亲属」的关系类型。
- Sheet1（实体表）：直接从 xlsx 读取
- Sheet2（关系表）：读取已清洗好的 cleaned_Sheet2.csv
输出：Cleaned_Sheet2_Relations.csv
────────────────────────────────────────────────────────────
"""

import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
from pathlib import Path

# ─────────────────────────────────────────────
# 路径配置
# ─────────────────────────────────────────────
XLSX_PATH    = Path(r"d:\1大创\《左联相关档案资源目录》.xlsx")
SHEET2_CSV   = Path(r"d:\1大创\cleaned_Sheet2.csv")       # 使用已清洗的版本
OUTPUT_CSV   = Path(r"d:\1大创\Cleaned_Sheet2_Relations.csv")

NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


# ─────────────────────────────────────────────
# 1. 用 XML 直解读取 Sheet1（兼容 WPS xlsx）
# ─────────────────────────────────────────────

def read_sheet_xml(xlsx_path: Path, sheet_xml_path: str) -> pd.DataFrame:
    """直接解析 xlsx 内部 sheet XML，支持内联字符串单元格。"""
    def cell_value(cell_elem) -> str:
        is_t = cell_elem.findall(".//main:is//main:t", NS)
        if is_t:
            return "".join(e.text or "" for e in is_t)
        v = cell_elem.find("main:v", NS)
        return v.text if v is not None else ""

    with zipfile.ZipFile(xlsx_path, "r") as z:
        with z.open(sheet_xml_path) as f:
            root = ET.parse(f).getroot()

    rows_data = []
    for row_elem in root.findall(".//main:sheetData/main:row", NS):
        rows_data.append([cell_value(c) for c in row_elem.findall("main:c", NS)])

    if not rows_data:
        return pd.DataFrame()

    headers = rows_data[0]
    data    = rows_data[1:]
    ncols   = len(headers)
    data    = [r + [""] * (ncols - len(r)) if len(r) < ncols else r[:ncols] for r in data]
    return pd.DataFrame(data, columns=headers)


# ─────────────────────────────────────────────
# 2. 读取数据
# ─────────────────────────────────────────────

print("[读取] 正在读取 Sheet1（实体表）...")
df_entities = read_sheet_xml(XLSX_PATH, "xl/worksheets/sheet1.xml")
print(f"   Sheet1：{len(df_entities)} 行，列名：{list(df_entities.columns)}")

print("[读取] 正在读取 cleaned_Sheet2.csv（关系表）...")
df_relations = pd.read_csv(SHEET2_CSV, encoding="utf-8-sig")
print(f"   Sheet2：{len(df_relations)} 行")

# ─────────────────────────────────────────────
# 3. 构建 Entity_ID -> Role 字典
# ─────────────────────────────────────────────

role_dict = dict(zip(df_entities["Entity_ID"], df_entities["Role"]))
print(f"\n[构建] Role 字典共 {len(role_dict)} 条条目")
# 打印几个样本确认
for k, v in list(role_dict.items())[:3]:
    print(f"   {k} -> {v}")

# ─────────────────────────────────────────────
# 4. 真实亲属白名单（确定无误的亲属关系，防止误改）
# ─────────────────────────────────────────────

kinship_whitelist = [
    # 示例：frozenset(["ZLH-XXX", "ZLH-YYY"])
    # 请根据实际情况填写
]


# ─────────────────────────────────────────────
# 5. 关系修正函数
# ─────────────────────────────────────────────

# 左联内部核心圈层角色定义
INTERNAL_ROLES  = {"核心领导", "普通成员"}

# 关系类型命名规范（强关联/弱关联 + 子类型）
# 交游：社交往来，归为弱关联（可根据实际情况修改为「强关联-交游」）
REL_KINSHIP_FIX_INTERNAL = "强关联-组织隶属"   # 双方均为内部成员时的修正值
REL_KINSHIP_FIX_EXTERNAL = "弱关联-交游"        # 含外围/相关人士时的修正值
EXTERNAL_ROLES  = {"外围联络人", "相关人士"}


def correct_relation(row) -> str:
    """
    仅针对被误标为「亲属」的关系，根据双方角色判断并修正。
    白名单内的真实亲属关系不做修改。
    """
    source      = row["Source_ID"]
    target      = row["Target_ID"]
    current_rel = str(row["Relation_Type"])

    # 精确匹配：只修正裸写的「亲属」，保留「弱关联-亲属」等已有前缀的真实亲属关系
    # 使用 strip() 后完全等于「亲属」作为判断，不影响「弱关联-亲属」
    if current_rel.strip() != "亲属":
        return current_rel

    # 白名单保护：确定是真实亲属的直接保留
    if frozenset([source, target]) in kinship_whitelist:
        return current_rel

    # 获取双方角色（默认「未知」）
    source_role = role_dict.get(source, "未知")
    target_role = role_dict.get(target, "未知")

    # 规则 A：双方均为左联内部成员 -> 强关联-组织隶属
    if source_role in INTERNAL_ROLES and target_role in INTERNAL_ROLES:
        return REL_KINSHIP_FIX_INTERNAL

    # 规则 B：任意一方为外围/相关人士 -> 弱关联-交游
    elif source_role in EXTERNAL_ROLES or target_role in EXTERNAL_ROLES:
        return REL_KINSHIP_FIX_EXTERNAL

    # 默认兜底 -> 弱关联-交游
    else:
        return REL_KINSHIP_FIX_EXTERNAL


# ─────────────────────────────────────────────
# 6. 应用修正
# ─────────────────────────────────────────────

# 统计修正前各类型数量
before_bare_kinship = (df_relations["Relation_Type"].str.strip() == "亲属").sum()
before_weak_kinship = (df_relations["Relation_Type"].str.strip() == "弱关联-亲属").sum()
before_bare_org     = (df_relations["Relation_Type"].str.strip() == "组织隶属").sum()

# 步骤 A：修正裸写「亲属」
df_relations["Relation_Type"] = df_relations.apply(correct_relation, axis=1)

# 步骤 B：将历史数据中裸写的「组织隶属」统一补全为「强关联-组织隶属」
df_relations["Relation_Type"] = df_relations["Relation_Type"].apply(
    lambda x: "强关联-组织隶属" if str(x).strip() == "组织隶属" else x
)

# 步骤 C：裸写「交游」补全为「弱关联-交游」
df_relations["Relation_Type"] = df_relations["Relation_Type"].apply(
    lambda x: "弱关联-交游" if str(x).strip() == "交游" else x
)

after_bare_kinship  = (df_relations["Relation_Type"].str.strip() == "亲属").sum()
after_weak_kinship  = (df_relations["Relation_Type"].str.strip() == "弱关联-亲属").sum()

print(f"\n[修正统计]")
print(f"  裸写「亲属」修正：{before_bare_kinship} -> {after_bare_kinship} 条")
print(f"  「弱关联-亲属」保留：{before_weak_kinship} -> {after_weak_kinship} 条")
print(f"  裸写「组织隶属」补全为「强关联-组织隶属」：{before_bare_org} 条")

# 展示修正后 Relation_Type 的分布
print("\n[统计] 修正后 Relation_Type 分布：")
print(df_relations["Relation_Type"].value_counts().to_string())

# ─────────────────────────────────────────────
# 7. 导出
# ─────────────────────────────────────────────

df_relations.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")
print(f"\n[完成] 文件已保存：{OUTPUT_CSV}  共 {len(df_relations)} 行")
