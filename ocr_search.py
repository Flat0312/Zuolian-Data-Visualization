# -*- coding: utf-8 -*-
"""
使用OCR识别《左联词典》PDF并检索孤岛成员
需要安装: pip install pytesseract pdf2image pillow
需要安装Tesseract-OCR: https://github.com/tesseract-ocr/tesseract
"""

import os
import pandas as pd
from pdf2image import convert_from_path
import pytesseract
from collections import defaultdict
import json

# 配置路径
EXCEL_PATH = r'd:\1大创\大创数据收集1.xlsx'
PDF_PATH = r'd:\1大创\左联词典 (姚辛) (z-library.sk, 1lib.sk, z-lib.sk).pdf'
OCR_CACHE = r'd:\1大创\ocr_cache'  # OCR结果缓存目录
OCR_TEXT_FILE = r'd:\1大创\左联词典_ocr_text.json'  # 保存OCR结果

# Tesseract路径（Windows）
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
# 使用本地语言包目录
os.environ['TESSDATA_PREFIX'] = r'd:\1大创'

def get_isolated_members():
    """获取孤岛成员"""
    sheet1 = pd.read_excel(EXCEL_PATH, sheet_name='Sheet1')
    sheet2 = pd.read_excel(EXCEL_PATH, sheet_name='Sheet2')
    
    all_members = set(sheet1['Entity_ID'].dropna().unique())
    appeared_ids = set(sheet2['Source_ID'].dropna().unique()) | set(sheet2['Target_ID'].dropna().unique())
    isolated_ids = all_members - appeared_ids
    
    isolated = sheet1[sheet1['Entity_ID'].isin(isolated_ids)][['Entity_ID', 'True_Name', 'Alias']].copy()
    return isolated

def ocr_pdf():
    """OCR识别PDF并保存结果"""
    # 如果已有缓存，直接加载
    if os.path.exists(OCR_TEXT_FILE):
        print(f"加载已有OCR结果: {OCR_TEXT_FILE}")
        with open(OCR_TEXT_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    
    print("开始OCR识别PDF（这可能需要30-60分钟）...")
    os.makedirs(OCR_CACHE, exist_ok=True)
    
    page_texts = {}
    
    # 分批处理避免内存溢出
    batch_size = 10
    total_pages = None
    
    # 先获取总页数
    from pdf2image.pdf2image import pdfinfo_from_path
    try:
        info = pdfinfo_from_path(PDF_PATH)
        total_pages = info['Pages']
        print(f"PDF共{total_pages}页")
    except:
        total_pages = 708  # 默认值
    
    for start in range(1, total_pages + 1, batch_size):
        end = min(start + batch_size - 1, total_pages)
        print(f"处理第{start}-{end}页...")
        
        try:
            images = convert_from_path(
                PDF_PATH, 
                first_page=start, 
                last_page=end,
                dpi=200  # 降低DPI以加快速度
            )
            
            for i, img in enumerate(images):
                page_num = start + i
                # 中文OCR
                text = pytesseract.image_to_string(img, lang='chi_sim')
                page_texts[str(page_num)] = text
                
        except Exception as e:
            print(f"  第{start}-{end}页处理失败: {e}")
    
    # 保存OCR结果
    with open(OCR_TEXT_FILE, 'w', encoding='utf-8') as f:
        json.dump(page_texts, f, ensure_ascii=False, indent=2)
    print(f"OCR结果已保存到: {OCR_TEXT_FILE}")
    
    return page_texts

def search_members(page_texts, members_df):
    """在OCR文本中检索成员"""
    results = defaultdict(lambda: {'count': 0, 'pages': set()})
    
    # 构建搜索词
    search_map = {}
    for _, row in members_df.iterrows():
        mid = row['Entity_ID']
        name = row['True_Name']
        if pd.notna(name) and len(name) >= 2:
            search_map[name] = mid
        
        if pd.notna(row['Alias']):
            for alias in str(row['Alias']).split('、'):
                alias = alias.strip()
                if len(alias) >= 2:
                    search_map[alias] = mid
    
    print(f"搜索词数量: {len(search_map)}")
    
    for page_str, text in page_texts.items():
        if not text:
            continue
        page = int(page_str)
        
        for name, mid in search_map.items():
            if name in text:
                count = text.count(name)
                results[mid]['count'] += count
                results[mid]['pages'].add(page)
    
    return results

def main():
    print("="*60)
    print("OCR识别《左联词典》并检索孤岛成员")
    print("="*60)
    
    # 1. 获取孤岛成员
    print("\n[1] 获取孤岛成员...")
    isolated = get_isolated_members()
    print(f"  孤岛成员数: {len(isolated)}")
    
    # 2. OCR识别PDF
    print("\n[2] OCR识别PDF...")
    page_texts = ocr_pdf()
    print(f"  成功识别{len(page_texts)}页")
    
    # 3. 检索成员
    print("\n[3] 检索孤岛成员...")
    results = search_members(page_texts, isolated)
    
    # 4. 输出结果
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
        print(f"{rank:<4} {mid:<12} {name:<10} {data['count']:<8} {len(data['pages']):<8}")
    
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
    
    pd.DataFrame(result_data).to_excel('isolated_members_ocr_result.xlsx', index=False)
    print(f"\n结果已保存到: isolated_members_ocr_result.xlsx")

if __name__ == '__main__':
    main()
