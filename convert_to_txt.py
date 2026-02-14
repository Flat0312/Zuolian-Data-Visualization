# -*- coding: utf-8 -*-
"""
OCR识别《左联史》PDF并将《左联词典》和《左联史》转换为TXT文件
需要安装: pip install pytesseract pdf2image pillow
"""

import os
import json
from pdf2image import convert_from_path
import pytesseract

# 配置路径
PDF_ZUOLIAN_CIDIAN = r'd:\1大创\左联词典 (姚辛) (z-library.sk, 1lib.sk, z-lib.sk).pdf'
PDF_ZUOLIAN_SHI = r'd:\1大创\左联史 (姚辛著, Yao, Xin. etc.) (z-library.sk, 1lib.sk, z-lib.sk).pdf'

OCR_CIDIAN_JSON = r'd:\1大创\左联词典_ocr_text.json'
OCR_SHI_JSON = r'd:\1大创\左联史_ocr_text.json'

TXT_CIDIAN = r'd:\1大创\左联词典.txt'
TXT_SHI = r'd:\1大创\左联史.txt'

# Tesseract路径（Windows）
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
# 使用本地语言包目录
os.environ['TESSDATA_PREFIX'] = r'd:\1大创'


def ocr_pdf(pdf_path, json_path):
    """OCR识别PDF并保存结果到JSON"""
    # 如果已有缓存，直接加载
    if os.path.exists(json_path):
        print(f"加载已有OCR结果: {json_path}")
        with open(json_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    
    print(f"开始OCR识别: {os.path.basename(pdf_path)}")
    print("（这可能需要较长时间，请耐心等待...）")
    
    page_texts = {}
    
    # 分批处理避免内存溢出
    batch_size = 10
    total_pages = None
    
    # 先获取总页数
    from pdf2image.pdf2image import pdfinfo_from_path
    try:
        info = pdfinfo_from_path(pdf_path)
        total_pages = info['Pages']
        print(f"PDF共{total_pages}页")
    except Exception as e:
        print(f"获取页数失败: {e}")
        return {}
    
    for start in range(1, total_pages + 1, batch_size):
        end = min(start + batch_size - 1, total_pages)
        print(f"处理第{start}-{end}页（共{total_pages}页）...")
        
        try:
            images = convert_from_path(
                pdf_path, 
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
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(page_texts, f, ensure_ascii=False, indent=2)
    print(f"OCR结果已保存到: {json_path}")
    
    return page_texts


def json_to_txt(json_path, txt_path, title):
    """将JSON格式的OCR结果转换为TXT文件"""
    if not os.path.exists(json_path):
        print(f"未找到OCR结果文件: {json_path}")
        return False
    
    print(f"转换 {title} 到TXT格式...")
    
    with open(json_path, 'r', encoding='utf-8') as f:
        page_texts = json.load(f)
    
    # 按页码排序并合并
    sorted_pages = sorted(page_texts.keys(), key=lambda x: int(x))
    
    with open(txt_path, 'w', encoding='utf-8') as f:
        f.write(f"{'='*60}\n")
        f.write(f"{title}\n")
        f.write(f"OCR识别结果 - 共{len(sorted_pages)}页\n")
        f.write(f"{'='*60}\n\n")
        
        for page_num in sorted_pages:
            text = page_texts[page_num]
            f.write(f"\n{'─'*40}\n")
            f.write(f"第 {page_num} 页\n")
            f.write(f"{'─'*40}\n\n")
            f.write(text)
            f.write("\n")
    
    print(f"  已保存到: {txt_path}")
    return True


def main():
    print("="*60)
    print("《左联史》OCR识别 & TXT转换工具")
    print("="*60)
    
    # Step 1: OCR《左联史》
    print("\n[1] 处理《左联史》OCR...")
    ocr_pdf(PDF_ZUOLIAN_SHI, OCR_SHI_JSON)
    
    # Step 2: 转换《左联词典》为TXT
    print("\n[2] 转换《左联词典》为TXT...")
    json_to_txt(OCR_CIDIAN_JSON, TXT_CIDIAN, "《左联词典》")
    
    # Step 3: 转换《左联史》为TXT
    print("\n[3] 转换《左联史》为TXT...")
    json_to_txt(OCR_SHI_JSON, TXT_SHI, "《左联史》")
    
    print("\n" + "="*60)
    print("处理完成!")
    print("="*60)
    print(f"生成的文件:")
    print(f"  - {TXT_CIDIAN}")
    print(f"  - {TXT_SHI}")


if __name__ == '__main__':
    main()
