"""
extract_context.py
════════════════════════════════════════════════════════════
从三个原始史料文件重新提取 Context，填回 Sheet2：
  1. 鲁迅日记（txt）
  2. 左联词典（txt）
  3. 左联史（txt 或 json）

策略：
  - OCR 去空格预处理：消除字间多余空格
  - 段落/行分割
  - 双人共现命中：在同一段落内两人任一称谓同时出现
  - 提取命中段落（含前后扩展），限 400 字
  - 未命中行保留原 Context
════════════════════════════════════════════════════════════
"""

import re
import json
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
from pathlib import Path

# ────────────────────────────────────────────────────────────
# 路径配置
# ────────────────────────────────────────────────────────────
XLSX_BACKUP   = Path(r"d:\1大创\《左联相关档案资源目录》_备份_20260225_111914.xlsx")
CLEANED_CSV   = Path(r"d:\1大创\cleaned_Sheet2.csv")
OUTPUT_CSV    = Path(r"d:\1大创\context_extracted.csv")

# 三个文本文件
DIARY_TXT     = Path(r"d:\1大创\日记全编：全2册 (鲁迅 著) (Z-Library).txt")
CIDIAN_TXT    = Path(r"d:\1大创\左联词典.txt")
ZUOLIAN_TXT   = Path(r"d:\1大创\左联史.txt")
ZUOLIAN_JSON  = Path(r"d:\1大创\左联史_ocr_text.json")   # 若存在优先用 json

MAX_CTX_LEN   = 400   # Context 最大字符数
WINDOW_CHARS  = 80    # 命中关键词前后扩展字符数

# ────────────────────────────────────────────────────────────
# 工具函数
# ────────────────────────────────────────────────────────────

def read_text_file(path: Path) -> str:
    """读取文本文件，自动检测编码。"""
    for enc in ("utf-8", "utf-8-sig", "gbk", "gb2312"):
        try:
            return path.read_text(encoding=enc, errors="strict")
        except Exception:
            continue
    return path.read_text(encoding="utf-8", errors="replace")


def remove_ocr_spaces(text: str) -> str:
    """
    消除 OCR 扫描产生的字间空格：
    仅删除两个 CJK 字符之间（或 CJK 与标点之间）的单个/多个空格。
    保留英文、数字之间的空格。
    """
    # CJK 字符范围
    cjk = r"[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]"
    # 去除 CJK 字符之间的空格
    text = re.sub(rf"({cjk}) +({cjk})", r"\1\2", text)
    text = re.sub(rf"({cjk}) +({cjk})", r"\1\2", text)  # 再过一遍保险
    return text


def split_paragraphs(text: str, min_len: int = 10) -> list[str]:
    """
    将文本分割为段落列表。
    连续非空行合并为一段，空行分隔。
    """
    paragraphs = []
    current = []
    for line in text.splitlines():
        line = line.strip()
        if line:
            current.append(line)
        else:
            if current:
                para = "".join(current)
                if len(para) >= min_len:
                    paragraphs.append(para)
                current = []
    if current:
        para = "".join(current)
        if len(para) >= min_len:
            paragraphs.append(para)
    return paragraphs


def split_sentences(text: str) -> list[str]:
    """按句号/换行将文本分割为句子列表（用于日记）。"""
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    return lines


# ────────────────────────────────────────────────────────────
# 读取数据源
# ────────────────────────────────────────────────────────────

_NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

def cell_value(elem) -> str:
    is_t = elem.findall(".//main:is//main:t", _NS)
    if is_t:
        return "".join(e.text or "" for e in is_t)
    v = elem.find("main:v", _NS)
    return v.text if v is not None else ""


def read_sheet1(xlsx_path: Path) -> pd.DataFrame:
    with zipfile.ZipFile(xlsx_path) as z:
        with z.open("xl/worksheets/sheet1.xml") as f:
            root = ET.parse(f).getroot()
    rows = root.findall(".//main:sheetData/main:row", _NS)
    headers = [cell_value(c) for c in rows[0].findall("main:c", _NS)]
    data = []
    for r in rows[1:]:
        vals = [cell_value(c) for c in r.findall("main:c", _NS)]
        ncols = len(headers)
        vals = vals + [""] * (ncols - len(vals)) if len(vals) < ncols else vals[:ncols]
        data.append(dict(zip(headers, vals)))
    return pd.DataFrame(data)


def build_alias_map(df_s1: pd.DataFrame) -> dict[str, list[str]]:
    """
    构建 {Entity_ID: [称谓列表]} 字典。
    称谓 = True_Name + Alias 中所有非空别名（已做 OCR 去空格处理）。
    """
    alias_map = {}
    for _, row in df_s1.iterrows():
        eid = str(row.get("Entity_ID", "")).strip()
        names = []
        true_name = str(row.get("True_Name", "")).strip()
        if true_name:
            names.append(remove_ocr_spaces(true_name))
        alias_str = str(row.get("Alias", "")).strip()
        if alias_str:
            # 别名可能以 /、、、, 等分隔
            for sep in ("、", "/", "，", ",", "；", ";", " "):
                alias_str = alias_str.replace(sep, "｜")
            for a in alias_str.split("｜"):
                a = remove_ocr_spaces(a.strip())
                if a and a not in names:
                    names.append(a)
        # 过滤长度 < 2 的称谓（太短容易误匹配）
        names = [n for n in names if len(n) >= 2]
        if eid and names:
            alias_map[eid] = names
    return alias_map


def load_corpus(
    diary_txt: Path,
    cidian_txt: Path,
    zuolian_txt: Path,
    zuolian_json: Path,
) -> list[tuple[str, str]]:
    """
    加载并预处理所有语料，返回 [(来源标签, 段落文本), ...]。
    优先级：鲁迅日记 > 左联史 > 左联词典
    """
    corpus: list[tuple[str, str]] = []

    # 1. 鲁迅日记（按行处理，每行为一条记录）
    print("  加载鲁迅日记...")
    raw_diary = remove_ocr_spaces(read_text_file(diary_txt))
    for line in split_sentences(raw_diary):
        if len(line) >= 10:
            corpus.append(("鲁迅日记", line))
    print(f"    → {len(corpus)} 行")

    # 2. 左联史（优先用 JSON）
    print("  加载左联史...")
    n_before = len(corpus)
    if zuolian_json.exists():
        with open(zuolian_json, encoding="utf-8") as f:
            pages = json.load(f)
        raw_zuolian = " ".join(pages.values())
    else:
        raw_zuolian = read_text_file(zuolian_txt)
    raw_zuolian = remove_ocr_spaces(raw_zuolian)
    for para in split_paragraphs(raw_zuolian, min_len=20):
        corpus.append(("左联史", para))
    print(f"    → {len(corpus) - n_before} 段")

    # 3. 左联词典
    print("  加载左联词典...")
    n_before2 = len(corpus)
    raw_cidian = remove_ocr_spaces(read_text_file(cidian_txt))
    for para in split_paragraphs(raw_cidian, min_len=20):
        corpus.append(("左联词典", para))
    print(f"    → {len(corpus) - n_before2} 段")

    print(f"  语料总计：{len(corpus)} 条")
    return corpus


# ────────────────────────────────────────────────────────────
# 双人共现命中
# ────────────────────────────────────────────────────────────

def names_hit(text: str, names: list[str]) -> str | None:
    """返回文本中首个命中的称谓，未命中返回 None。"""
    for n in names:
        if n in text:
            return n
    return None


def extract_snippet(text: str, hit_a: str, hit_b: str) -> str:
    """
    从 text 中提取包含 hit_a 和 hit_b 的片段，
    取两处命中位置的并集区间，前后各扩展 WINDOW_CHARS 字符。
    """
    pos_a = text.find(hit_a)
    pos_b = text.find(hit_b)
    start = max(0, min(pos_a, pos_b) - WINDOW_CHARS)
    end   = min(len(text), max(pos_a + len(hit_a), pos_b + len(hit_b)) + WINDOW_CHARS)
    snippet = text[start:end].strip()
    # 清理行首行尾多余空白
    snippet = re.sub(r"\s+", " ", snippet)
    return snippet[:MAX_CTX_LEN]


def search_pair(
    names_a: list[str],
    names_b: list[str],
    corpus: list[tuple[str, str]],
) -> str:
    """
    在 corpus 中搜索两人共现的段落，
    按来源优先级（鲁迅日记>左联史>左联词典）返回最优片段。
    最多拼接 2 个来源的片段，用「/」分隔。
    """
    priority = {"鲁迅日记": 0, "左联史": 1, "左联词典": 2}
    hits: dict[str, list[str]] = {}  # {来源: [snippet, ...]}

    for source, para in corpus:
        hit_a = names_hit(para, names_a)
        hit_b = names_hit(para, names_b)
        if hit_a and hit_b:
            snippet = extract_snippet(para, hit_a, hit_b)
            hits.setdefault(source, []).append(snippet)

    if not hits:
        return ""

    # 按优先级拼接，最多取 2 个来源各 1 条
    result_parts = []
    for src in sorted(hits.keys(), key=lambda s: priority.get(s, 99)):
        result_parts.append(hits[src][0])
        if len(result_parts) >= 2:
            break

    combined = " / ".join(result_parts)
    return combined[:MAX_CTX_LEN]


# ────────────────────────────────────────────────────────────
# 主流程
# ────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("  从原始史料重提取 Context")
    print("=" * 60)

    # 读 Sheet1 构建别名字典
    print("\n[1/4] 读取 Sheet1 别名表...")
    df_s1    = read_sheet1(XLSX_BACKUP)
    alias_map = build_alias_map(df_s1)
    print(f"      实体数：{len(alias_map)}")
    # 示例
    sample_ids = list(alias_map.keys())[:3]
    for sid in sample_ids:
        print(f"      {sid}: {alias_map[sid]}")

    # 读 Sheet2
    print("\n[2/4] 读取 cleaned_Sheet2.csv...")
    df_s2 = pd.read_csv(CLEANED_CSV, encoding="utf-8-sig")
    drop_cols = [c for c in ["Context_raw", "NER_persons", "NER_places"] if c in df_s2.columns]
    if drop_cols:
        df_s2.drop(columns=drop_cols, inplace=True)
    print(f"      {len(df_s2)} 行，{len(df_s2.columns)} 列")
    # 获取唯一人物对
    pairs = df_s2[["Source_ID", "Target_ID"]].drop_duplicates()
    print(f"      唯一人物对：{len(pairs)}")

    # 加载语料
    print("\n[3/4] 加载并预处理原始史料...")
    corpus = load_corpus(DIARY_TXT, CIDIAN_TXT, ZUOLIAN_TXT, ZUOLIAN_JSON)

    # 逐对提取
    print(f"\n[4/4] 双人共现提取（{len(pairs)} 对）...")
    pair_ctx: dict[tuple[str, str], str] = {}
    hit_count = 0

    for i, (_, row) in enumerate(pairs.iterrows()):
        src_id = str(row["Source_ID"]).strip()
        tgt_id = str(row["Target_ID"]).strip()
        names_a = alias_map.get(src_id, [])
        names_b = alias_map.get(tgt_id, [])

        if not names_a or not names_b:
            pair_ctx[(src_id, tgt_id)] = ""
            continue

        ctx = search_pair(names_a, names_b, corpus)
        pair_ctx[(src_id, tgt_id)] = ctx
        if ctx:
            hit_count += 1

        if (i + 1) % 50 == 0 or (i + 1) == len(pairs):
            print(f"  [{i+1}/{len(pairs)}] 已命中 {hit_count} 对")

    print(f"\n  命中率：{hit_count}/{len(pairs)} = {hit_count/max(len(pairs),1)*100:.1f}%")

    # 把提取到的 Context 映射回行级数据
    def get_new_ctx(row) -> str:
        key = (str(row["Source_ID"]).strip(), str(row["Target_ID"]).strip())
        # 双向查找
        new_ctx = pair_ctx.get(key, "") or pair_ctx.get((key[1], key[0]), "")
        # 未命中则保留原 Context
        if not new_ctx:
            return str(row.get("Context", ""))
        return new_ctx

    df_s2["Context"] = df_s2.apply(get_new_ctx, axis=1)

    # 输出
    df_s2.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")
    print(f"\n  输出：{OUTPUT_CSV}")
    print(f"  行数：{len(df_s2)}")
    print("\n  前5条样例：")
    for _, r in df_s2.head(5).iterrows():
        print(f"    {r['Source_ID']} → {r['Target_ID']}: {repr(r['Context'][:80])}")

    print("\n" + "=" * 60)
    print("  完成！下一步：运行 write_back_to_xlsx.py 将结果写回 xlsx")
    print("  (将 SHEET2_CSV 改为 context_extracted.csv)")
    print("=" * 60)


if __name__ == "__main__":
    main()
