# -*- coding: utf-8 -*-
"""导出孤岛成员名单"""

import pandas as pd

EXCEL_PATH = r'd:\1大创\大创数据收集1.xlsx'

def main():
    sheet1 = pd.read_excel(EXCEL_PATH, sheet_name='Sheet1')
    sheet2 = pd.read_excel(EXCEL_PATH, sheet_name='Sheet2')
    
    all_ids = set(sheet1['Entity_ID'].dropna())
    appeared = set(sheet2['Source_ID'].dropna()) | set(sheet2['Target_ID'].dropna())
    isolated = all_ids - appeared
    
    result = sheet1[sheet1['Entity_ID'].isin(isolated)][['Entity_ID', 'True_Name', 'Alias', 'Role', 'Birth_Death']].copy()
    result = result.sort_values('Entity_ID')
    
    # 导出Excel
    result.to_excel('孤岛成员名单.xlsx', index=False)
    
    print(f'已导出{len(result)}名孤岛成员')
    print(f'输出文件: 孤岛成员名单.xlsx')
    print()
    print('='*60)
    print('孤岛成员列表:')
    print('='*60)
    
    for i, (_, row) in enumerate(result.iterrows(), 1):
        name = row['True_Name']
        alias = str(row['Alias'])[:20] if pd.notna(row['Alias']) else ''
        role = row['Role'] if pd.notna(row['Role']) else ''
        print(f"{i:3}. {row['Entity_ID']}: {name} | {alias} | {role}")

if __name__ == '__main__':
    main()
