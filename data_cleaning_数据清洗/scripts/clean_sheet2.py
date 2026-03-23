"""
clean_sheet2.py
────────────────────────────────────────────────────────────
左联相关档案资源目录 Sheet2 — Context 列数据清洗脚本
作者：Antigravity（根据用户需求生成）
日期：2026-02-25

功能概要：
  1. 用 XML 直接解析 xlsx（兼容缺少 sharedStrings.xml 的 WPS 文件）
  2. 多轮正则清洗（OCR 乱码、索引残留、冗余标点）
  3. 修补不自然内部换行，平滑句子
  4. 基于规则的简单 NER 人名/地名预提取框架
  5. 导出 cleaned_Sheet2.csv
────────────────────────────────────────────────────────────
"""

import re
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
from pathlib import Path

# ─────────────────────────────────────────────
# 0. 路径配置
# ─────────────────────────────────────────────
XLSX_PATH  = Path(r"d:\1大创\《左联相关档案资源目录》.xlsx")
SHEET_XML  = "xl/worksheets/sheet2.xml"   # xlsx 内部路径
OUTPUT_CSV = Path(r"d:\1大创\cleaned_Sheet2.csv")

# xlsx 命名空间
NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


# ─────────────────────────────────────────────
# 1. 读取 xlsx（XML 直解）
# ─────────────────────────────────────────────

def read_sheet_xml(xlsx_path: Path, sheet_xml_path: str) -> pd.DataFrame:
    """
    直接解析 xlsx 内部的 sheet XML，兼容缺少 sharedStrings.xml 的文件。
    支持内联字符串（inline strings）和普通数值单元格。
    """
    def cell_value(cell_elem) -> str:
        """解析单元格值：优先读内联字符串(is)，其次读普通值(v)"""
        is_elems = cell_elem.findall(".//main:is//main:t", NS)
        if is_elems:
            return "".join(e.text or "" for e in is_elems)
        v = cell_elem.find("main:v", NS)
        return v.text if v is not None else ""

    with zipfile.ZipFile(xlsx_path, "r") as z:
        with z.open(sheet_xml_path) as f:
            root = ET.parse(f).getroot()

    rows_data = []
    for row_elem in root.findall(".//main:sheetData/main:row", NS):
        row = [cell_value(c) for c in row_elem.findall("main:c", NS)]
        rows_data.append(row)

    if not rows_data:
        return pd.DataFrame()

    # 第一行作为表头
    headers = rows_data[0]
    data    = rows_data[1:]

    # 列数对齐（防止某些行列数不一致）
    ncols = len(headers)
    data  = [r + [""] * (ncols - len(r)) if len(r) < ncols else r[:ncols] for r in data]

    df = pd.DataFrame(data, columns=headers)
    return df


# ─────────────────────────────────────────────
# 2. 正则规则定义
# ─────────────────────────────────────────────

CLEAN_RULES = [

    # ── 2.1 修复内部换行符 ──────────────────────────────────────
    # 将文本中的换行符替换为空格，使断开的句子平滑连接
    (
        "修复内部换行符",
        r"[ \t]*[\r\n]+[ \t]*",
        " "
    ),

    # ── 2.2 压缩连续重复引号/书名号 ─────────────────────────────
    # 如 """""  《《《 等 OCR 对引号识别混乱产生的连续同类符号
    (
        "压缩连续重复引号",
        r'(["""《〔〕】【『』「」]{2,})',
        lambda m: m.group(0)[0]   # 只保留第一个
    ),

    # ── 2.3 删除书籍笔画索引残留 ────────────────────────────────
    # 针对：(苏联)〉S,狄纳莫夫_十八画 之类的索引噪音
    (
        "删除书籍笔画索引残留",
        r"[_＿]\s*[一二三四五六七八九十百]+\s*画",
        ""
    ),
    (
        "删除OCR尖括号字母索引",
        r"[〉〈>]\s*[A-Za-z]+\s*[,，]",
        ""
    ),

    # ── 2.4 删除汉字间夹杂的孤立 ASCII 短串 ─────────────────────
    # 两侧均为汉字/标点时，夹在中间的纯ASCII短串（≤5字符）视为OCR噪音
    (
        "删除汉字间夹杂的孤立ASCII短串",
        r"(?<=[\u4e00-\u9fff，。！？；：])([A-Za-z0-9]{1,5})(?=[\u4e00-\u9fff，。！？；：])",
        ""
    ),

    # ── 2.5 压缩连续省略号/句号 ─────────────────────────────────
    (
        "压缩连续省略号",
        r"[.．]{3,}",
        "……"
    ),
    (
        "压缩连续中文句号",
        r"[。]{2,}",
        "。"
    ),

    # ── 2.6 压缩多余空白 ────────────────────────────────────────
    (
        "压缩多余空白",
        r"[ \t　]{2,}",
        " "
    ),

    # ── 2.7 去除首尾空白 ────────────────────────────────────────
    (
        "去除首尾空白",
        r"^\s+|\s+$",
        ""
    ),

    # ── 2.8 删除孤立版面噪音符号 ────────────────────────────────
    # 仅当前后都是空白/行首行尾时才删除，保守处理
    (
        "删除孤立版面噪音符号",
        r"(?<!\S)[〉〈◎□■◆▲※＊]{1,3}(?!\S)",
        ""
    ),
]


def clean_context(text: str) -> str:
    """对单条 Context 文本依次应用所有清洗规则，返回清洗后的字符串。"""
    if not isinstance(text, str) or not text.strip():
        return ""

    for _desc, pattern, replacement in CLEAN_RULES:
        if callable(replacement):
            text = re.sub(pattern, replacement, text)
        else:
            text = re.sub(pattern, replacement, text)

    return text


# ─────────────────────────────────────────────
# 3. 数据质量评估
# ─────────────────────────────────────────────

def assess_quality(series: pd.Series, label: str = ""):
    """打印 Context 列的关键统计信息，用于清洗前后对比。"""
    total   = len(series)
    empty   = (series.isna() | (series == "")).sum()
    nonempty = series[~(series.isna() | (series == ""))]
    avg_len = nonempty.str.len().mean() if len(nonempty) else 0
    print(f"\n【{label}】")
    print(f"  总行数：{total}")
    print(f"  空值/空串数：{empty}（{empty/total*100:.1f}%）")
    print(f"  非空行平均文本长度：{avg_len:.1f} 字符")


# ─────────────────────────────────────────────
# 4. NER 预处理（基于规则）
# ─────────────────────────────────────────────

# 左联核心人名词表（可继续扩充）
KNOWN_PERSONS = [
    "鲁迅", "茅盾", "郭沫若", "冯雪峰", "柔石", "丁玲", "胡也频", "殷夫",
    "洪灵菲", "冯铿", "李伟森", "夏衍", "阳翰笙", "田汉", "欧阳山",
    "周扬", "艾思奇", "何干之", "任白戈", "张天翼",
]

# 左联相关地名词表
KNOWN_PLACES = [
    "上海", "北京", "南京", "广州", "武汉", "延安", "莫斯科",
    "虹口", "闸北", "法租界", "英租界", "多伦路", "内山书店",
]


def extract_entities_by_rule(text: str) -> dict:
    """基于词表提取人名和地名，返回 {"persons": [...], "places": [...]}"""
    if not isinstance(text, str) or not text.strip():
        return {"persons": [], "places": []}
    return {
        "persons": [p for p in KNOWN_PERSONS if p in text],
        "places":  [p for p in KNOWN_PLACES  if p in text],
    }


# ─────────────────────────────────────────────
# 5. 主流程
# ─────────────────────────────────────────────

def main():
    print("=" * 60)
    print("  左联档案 Sheet2 — Context 列数据清洗脚本")
    print("=" * 60)

    # ── 5.1 读取数据（XML 直解，兼容 WPS） ───────────────────
    print(f"\n[读取] 正在读取：{XLSX_PATH}")
    df = read_sheet_xml(XLSX_PATH, SHEET_XML)
    print(f"   成功读取 {len(df)} 行，列名：{list(df.columns)}")

    if "Context" not in df.columns:
        raise ValueError(f"❌ 未找到 'Context' 列，当前列名：{list(df.columns)}")

    # ── 5.2 清洗前质量评估 ────────────────────────────────────
    assess_quality(df["Context"].fillna(""), label="清洗前")

    # 保留原始列以便对比审核
    df["Context_raw"] = df["Context"]

    # ── 5.3 执行清洗 ──────────────────────────────────────────
    print("\n[清洗] 正在应用清洗规则...")
    df["Context"] = df["Context"].apply(clean_context)
    print("   清洗完成！")

    # ── 5.4 清洗后质量评估 ────────────────────────────────────
    assess_quality(df["Context"].fillna(""), label="清洗后")

    # ── 5.5 前5行对比打印 ────────────────────────────────────
    print("\n[对比] 前5行 Context 清洗对比：")
    for i, row in df.head(5).iterrows():
        raw     = str(row["Context_raw"])[:100].replace("\n", "[LF]")
        cleaned = str(row["Context"])[:100]
        print(f"  [行{i+1}] 原文: {raw}")
        print(f"        清洗: {cleaned}")
        print()

    # ── 5.6 NER 预提取 ────────────────────────────────────────
    print("[NER] 正在进行 NER 预提取（基于规则词表）...")
    ner            = df["Context"].apply(extract_entities_by_rule)
    df["NER_persons"] = ner.apply(lambda x: "、".join(x["persons"]))
    df["NER_places"]  = ner.apply(lambda x: "、".join(x["places"]))
    print("   NER 提取完成！新增列：NER_persons, NER_places")

    # ── 5.7 导出 CSV ──────────────────────────────────────────
    print(f"\n[导出] 正在导出：{OUTPUT_CSV}")
    df.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")
    print(f"   导出成功！共 {len(df)} 行")
    print("=" * 60)


if __name__ == "__main__":
    main()
