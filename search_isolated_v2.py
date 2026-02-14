# -*- coding: utf-8 -*-
"""优化版：孤岛成员词典检索"""

import pandas as pd
import pdfplumber
import re
from collections import defaultdict

EXCEL_PATH = r'd:\1大创\大创数据收集1.xlsx'
PDF_PATH = r'd:\1大创\左联词典 (姚辛) (z-library.sk, 1lib.sk, z-lib.sk).pdf'

def get_isolated_members():
    """获取孤岛成员"""
    sheet1 = pd.read_excel(EXCEL_PATH, sheet_name='Sheet1')
    sheet2 = pd.read_excel(EXCEL_PATH, sheet_name='Sheet2')
    
    all_members = set(sheet1['Entity_ID'].dropna().unique())
    appeared_ids = set(sheet2['Source_ID'].dropna().unique()) | set(sheet2['Target_ID'].dropna().unique())
    isolated_ids = all_members - appeared_ids
    
    isolated = sheet1[sheet1['Entity_ID'].isin(isolated_ids)][['Entity_ID', 'True_Name', 'Alias']].copy()
    return isolated

def search_pdf(members_df):
    """检索PDF - 优化版，只搜索真实姓名(2字以上)"""
    results = defaultdict(lambda: {'count': 0, 'pages': set()})
    
    # 构建搜索词：只用真名和长度>=2的别名
    search_map = {}  # name -> member_id
    for _, row in members_df.iterrows():
        mid = row['Entity_ID']
        name = row['True_Name']
        if pd.notna(name) and len(name) >= 2:
            search_map[name] = mid
        
        if pd.notna(row['Alias']):
            for alias in str(row['Alias']).split('、'):
                alias = alias.strip()
                if len(alias) >= 2:  # 只用2字以上的别名
                    search_map[alias] = mid
    
    print(f"构建了{len(search_map)}个搜索词")
    
    # 先测试PDF是否可读
    with pdfplumber.open(PDF_PATH) as pdf:
        # 测试第一页
        test_text = pdf.pages[0].extract_text()
        print(f"PDF测试页文本长度: {len(test_text) if test_text else 0}")
        if test_text:
            print(f"测试页前100字: {test_text[:100]}")
        
        # 检索所有页面
        print(f"\n正在检索{len(pdf.pages)}页...")
        for i, page in enumerate(pdf.pages):
            if (i + 1) % 100 == 0:
                print(f"  {i+1}/{len(pdf.pages)}页...")
            
            text = page.extract_text()
            if not text:
                continue
            
            for name, mid in search_map.items():
                if name in text:
                    count = text.count(name)
                    results[mid]['count'] += count
                    results[mid]['pages'].add(i + 1)
    
    return results, search_map

def main():
    print("="*60)
    print("孤岛成员词典检索（优化版）")
    print("="*60)
    
    # 获取孤岛成员
    isolated = get_isolated_members()
    print(f"\n孤岛成员数: {len(isolated)}")
    
    # 检索PDF
    results, search_map = search_pdf(isolated)
    
    # 统计
    found = [(mid, data) for mid, data in results.items() if data['count'] > 0]
    found.sort(key=lambda x: x[1]['count'], reverse=True)
    
    name_map = dict(zip(isolated['Entity_ID'], isolated['True_Name']))
    
    print(f"\n在词典中找到{len(found)}名成员")
    print("\n" + "="*60)
    print("优先处理名单（按出现频率Top20）:")
    print("="*60)
    print(f"{'排名':<4} {'ID':<12} {'姓名':<10} {'出现次数':<8} {'涉及页数':<8}")
    print("-"*60)
    
    for rank, (mid, data) in enumerate(found[:20], 1):
        name = name_map.get(mid, mid)
        pages = sorted(data['pages'])
        print(f"{rank:<4} {mid:<12} {name:<10} {data['count']:<8} {len(pages):<8}")
    
    # 保存结果
    result_data = []
    for mid, data in found:
        pages = sorted(data['pages'])
        result_data.append({
            'ID': mid,
            'Name': name_map.get(mid, ''),
            'Count': data['count'],
            'PageCount': len(pages),
            'SamplePages': ', '.join(map(str, pages[:10]))
        })
    
    pd.DataFrame(result_data).to_excel('isolated_members_found.xlsx', index=False)
    print(f"\n结果已保存到: isolated_members_found.xlsx")

if __name__ == '__main__':
    main()
