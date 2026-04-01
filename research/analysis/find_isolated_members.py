# -*- coding: utf-8 -*-
"""
孤岛成员识别与词典检索
"""

import pandas as pd
import pdfplumber
import re
from collections import defaultdict

EXCEL_PATH = r'd:\1大创\大创数据收集1.xlsx'
PDF_PATH = r'd:\1大创\左联词典 (姚辛) (z-library.sk, 1lib.sk, z-lib.sk).pdf'

def identify_isolated_members():
    """识别孤岛成员"""
    sheet1 = pd.read_excel(EXCEL_PATH, sheet_name='Sheet1')
    sheet2 = pd.read_excel(EXCEL_PATH, sheet_name='Sheet2')
    
    # Sheet1中的所有成员ID
    all_members = set(sheet1['Entity_ID'].dropna().unique())
    
    # Sheet2中出现过的所有ID
    appeared_ids = set(sheet2['Source_ID'].dropna().unique()) | set(sheet2['Target_ID'].dropna().unique())
    
    # 孤岛成员
    isolated_ids = all_members - appeared_ids
    
    # 获取孤岛成员详情
    isolated_members = sheet1[sheet1['Entity_ID'].isin(isolated_ids)][['Entity_ID', 'True_Name', 'Alias']].copy()
    
    return isolated_members, len(all_members), len(appeared_ids)

def search_in_pdf(members_df):
    """在PDF中检索孤岛成员"""
    results = defaultdict(lambda: {'count': 0, 'pages': []})
    
    # 构建搜索词列表
    search_terms = {}
    for _, row in members_df.iterrows():
        member_id = row['Entity_ID']
        names = [row['True_Name']]
        if pd.notna(row['Alias']) and row['Alias']:
            aliases = str(row['Alias']).split('、')
            names.extend([a.strip() for a in aliases if a.strip()])
        search_terms[member_id] = names
    
    print(f"正在检索PDF中的{len(search_terms)}名孤岛成员...")
    
    with pdfplumber.open(PDF_PATH) as pdf:
        total_pages = len(pdf.pages)
        print(f"PDF共{total_pages}页")
        
        for i, page in enumerate(pdf.pages):
            if (i + 1) % 100 == 0:
                print(f"  处理第{i+1}/{total_pages}页...")
            
            text = page.extract_text()
            if not text:
                continue
            
            for member_id, names in search_terms.items():
                for name in names:
                    if len(name) >= 2 and name in text:
                        count = text.count(name)
                        results[member_id]['count'] += count
                        if i + 1 not in results[member_id]['pages']:
                            results[member_id]['pages'].append(i + 1)
    
    return results

def main():
    print("=" * 60)
    print("第一阶段：孤岛识别与定位")
    print("=" * 60)
    
    # 1. 识别孤岛成员
    print("\n[1] 识别孤岛成员...")
    isolated_members, total, appeared = identify_isolated_members()
    print(f"  Sheet1总成员数: {total}")
    print(f"  Sheet2出现过的ID数: {appeared}")
    print(f"  孤岛成员数: {len(isolated_members)}")
    
    print("\n孤岛成员列表:")
    for _, row in isolated_members.iterrows():
        alias = row['Alias'] if pd.notna(row['Alias']) else ''
        print(f"  {row['Entity_ID']}: {row['True_Name']} ({alias})")
    
    # 2. 在PDF中检索
    print("\n[2] 在《左联词典》中检索...")
    pdf_results = search_in_pdf(isolated_members)
    
    # 3. 按出现频率排序
    print("\n[3] 词典中信息最丰富的孤岛成员（按出现频率）:")
    sorted_results = sorted(
        [(mid, data) for mid, data in pdf_results.items() if data['count'] > 0],
        key=lambda x: x[1]['count'],
        reverse=True
    )
    
    # 创建结果表
    name_map = dict(zip(isolated_members['Entity_ID'], isolated_members['True_Name']))
    
    print("\n优先处理名单:")
    print("-" * 60)
    print(f"{'ID':<12} {'姓名':<10} {'出现次数':<10} {'涉及页码数':<10}")
    print("-" * 60)
    
    for member_id, data in sorted_results[:20]:
        name = name_map.get(member_id, member_id)
        print(f"{member_id:<12} {name:<10} {data['count']:<10} {len(data['pages']):<10}")
    
    # 保存完整结果
    result_df = pd.DataFrame([
        {
            'ID': mid,
            'Name': name_map.get(mid, ''),
            'Count': data['count'],
            'Pages': ', '.join(map(str, sorted(data['pages'])[:10]))
        }
        for mid, data in sorted_results
    ])
    result_df.to_csv('isolated_members_pdf_search.csv', encoding='utf-8-sig', index=False)
    print(f"\n完整结果已保存到: isolated_members_pdf_search.csv")

if __name__ == '__main__':
    main()
