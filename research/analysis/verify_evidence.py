import re

import pandas as pd

df = pd.read_excel('数据/输出结果/《左联相关档案资源目录》.xlsx', sheet_name='Sheet2')
df1 = pd.read_excel('数据/输出结果/《左联相关档案资源目录》.xlsx', sheet_name='Sheet1')
id_to_name = dict(zip(df1['Entity_ID'], df1['True_Name']))

with open('数据/原始文本/左联词典.txt', encoding='utf-8') as f:
    zuolian_cidian = f.read()
with open('数据/原始文本/左联史.txt', encoding='utf-8') as f:
    zuolian_shi = f.read()

def name_to_ocr_pattern(name):
    """将汉字姓名转为可匹配 OCR 空格的正则: '丁玲' -> '丁\\s*玲'"""
    chars = list(name)
    return r'\s*'.join(re.escape(c) for c in chars)

def find_page_content(text, page_num, context_chars=2000):
    """定位页码附近的文本"""
    # OCR文本页码格式: "第 103 页" 或 "── 103 ──"
    patterns = [
        rf'第\s*{page_num}\s*页',
        rf'[─—\-]+\s*{page_num}\s*[─—\-]+',
    ]
    for pat in patterns:
        m = re.search(pat, text)
        if m:
            # 取页码后面到下一页的内容
            start = m.end()
            # 找下一页
            next_page = page_num + 1
            next_pats = [
                rf'第\s*{next_page}\s*页',
                rf'[─—\-]+\s*{next_page}\s*[─—\-]+',
            ]
            end = min(start + context_chars, len(text))
            for np in next_pats:
                nm = re.search(np, text[start:])
                if nm:
                    end = start + nm.start()
                    break
            return text[start:end]
    return None

def parse_ref(ref):
    m = re.match(r'(.+?)\s*第(\d+)页', ref)
    if m:
        return m.group(1).strip(), int(m.group(2))
    return None, None

def check_name_in_text(name, text):
    """检查名字是否在OCR文本中（考虑空格）"""
    pat = name_to_ocr_pattern(name)
    return bool(re.search(pat, text))

top_refs = df['Evidence_Ref'].value_counts().head(5)
output = []

for ref_str, count in top_refs.items():
    book, page = parse_ref(ref_str)
    if not book:
        continue

    if '词典' in book:
        text = zuolian_cidian
    elif '史' in book:
        text = zuolian_shi
    else:
        text = None

    subset = df[df['Evidence_Ref'] == ref_str]
    all_ids = set(subset['Source_ID'].unique()) | set(subset['Target_ID'].unique())
    all_names = {id_to_name.get(eid, eid) for eid in all_ids}

    output.append("=" * 80)
    output.append(f"验证: {ref_str} （共 {count} 条关系）")
    output.append(f"关系类型: {subset['Relation_Type'].value_counts().to_dict()}")
    output.append(f"涉及人物 ({len(all_names)} 人): {', '.join(sorted(all_names))}")
    output.append("--- 前10条关系 ---")
    for i, (_, row) in enumerate(subset.head(10).iterrows()):
        s = id_to_name.get(row['Source_ID'], row['Source_ID'])
        t = id_to_name.get(row['Target_ID'], row['Target_ID'])
        ctx = str(row.get('Context', ''))[:60] if pd.notna(row.get('Context')) else '(无)'
        output.append(f"  {i+1}. {s} -> {t} [{row['Relation_Type']}] Context: {ctx}")

    if text is None:
        output.append(f"⚠ 无对应原始文本文件（{book}.txt），无法验证")
        output.append("")
        continue

    page_content = find_page_content(text, page)

    if not page_content:
        output.append(f"⚠ 在原始文本中未找到第{page}页")
    else:
        output.append(f"✓ 找到第{page}页内容（{len(page_content)} 字符）")
        output.append("--- 原文截取（前800字符）---")
        output.append(page_content[:800])

        found_names = [n for n in all_names if check_name_in_text(n, page_content)]
        not_found_names = [n for n in all_names if not check_name_in_text(n, page_content)]

        output.append("\n--- 人名验证（OCR空格已处理）---")
        output.append(f"在原文中找到 {len(found_names)}/{len(all_names)} 个人名 ({len(found_names)/len(all_names)*100:.0f}%)")
        if found_names:
            output.append(f"  ✓ 找到: {', '.join(sorted(found_names))}")
        if not_found_names:
            output.append(f"  ✗ 未找到: {', '.join(sorted(not_found_names))}")

        # 对未找到的名字，扩大搜索到整个页面前后5000字符
        if not_found_names:
            # 在全文中搜索这些人名出现的最近页码
            output.append("\n--- 对未找到人名的额外检查 ---")
            for name in sorted(not_found_names)[:10]:
                pat = name_to_ocr_pattern(name)
                matches = list(re.finditer(pat, text))
                if matches:
                    # 找出这些匹配最靠近 page_num 的位置
                    output.append(f"  {name}: 在全文中出现{len(matches)}次，但不在第{page}页")
                else:
                    output.append(f"  {name}: 在全文中完全未出现")

    output.append("")

with open('verify_result.txt', 'w', encoding='utf-8') as f:
    f.write('\n'.join(output))

print("验证结果已写入 verify_result.txt")
