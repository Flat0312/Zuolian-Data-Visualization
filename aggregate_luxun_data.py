# -*- coding: utf-8 -*-
"""
鲁迅节点数据清洗与聚合
- 保留强关联记录
- 按月聚合弱关联-通信记录
- 剔除无关记录
"""

import pandas as pd
import re
import shutil
from datetime import datetime
from collections import defaultdict

EXCEL_PATH = r'd:\1大创\大创数据收集1.xlsx'
BACKUP_PATH = r'd:\1大创\大创数据收集1_备份_{}.xlsx'.format(datetime.now().strftime('%Y%m%d_%H%M%S'))

# ID到姓名映射
ID_TO_NAME = {
    "ZLH-001": "鲁迅", "ZLH-002": "茅盾", "ZLH-003": "瞿秋白", "ZLH-004": "夏衍",
    "ZLH-005": "冯雪峰", "ZLH-006": "潘汉年", "ZLH-007": "周扬", "ZLH-008": "田汉",
    "ZLH-009": "郑伯奇", "ZLH-010": "洪深", "ZLH-011": "钱杏邨", "ZLH-012": "蒋光慈",
    "ZLH-013": "阳翰笙", "ZLH-014": "冯乃超", "ZLH-015": "郁达夫", "ZLH-016": "柔石",
    "ZLH-017": "胡也频", "ZLH-018": "冯铿", "ZLH-019": "殷夫", "ZLH-020": "李求实",
    "ZLH-021": "丁玲", "ZLH-060": "鲁彦", "ZLH-071": "李霁野", "ZLH-121": "内山完造",
    "ZLH-122": "许广平", "ZLH-124": "李小峰", "ZLH-126": "沈兼士", "ZLH-128": "林语堂",
    "ZLH-139": "陈望道",
}

def extract_month(evidence_ref):
    """从Evidence_Ref中提取年月，如'1930年10月'"""
    match = re.search(r'(\d{4})年(\d{1,2})月', str(evidence_ref))
    if match:
        return f"{match.group(1)}年{match.group(2)}月"
    return None

def aggregate_communications(df):
    """按月聚合弱关联-通信记录"""
    # 筛选鲁迅的弱关联-通信记录
    luxun_comm = df[
        (df['Source_ID'] == 'ZLH-001') & 
        (df['Relation_Type'] == '弱关联-通信')
    ].copy()
    
    # 提取月份
    luxun_comm['Month'] = luxun_comm['Evidence_Ref'].apply(extract_month)
    
    # 按(Target_ID, Month)分组聚合
    aggregated = []
    grouped = luxun_comm.groupby(['Target_ID', 'Month'])
    
    for (target_id, month), group in grouped:
        if month is None:
            # 无法解析月份的记录保持原样
            for _, row in group.iterrows():
                aggregated.append(row.to_dict())
            continue
        
        # 获取对方姓名
        target_name = ID_TO_NAME.get(target_id, target_id)
        
        # 提取原有Context的关键词
        contexts = group['Context'].dropna().tolist()
        keywords = set()
        for ctx in contexts:
            if '寄' in str(ctx) or '复' in str(ctx):
                keywords.add('寄信')
            if '收' in str(ctx):
                keywords.add('收信')
        
        keyword_str = '、'.join(keywords) if keywords else '书信往来'
        
        # 创建聚合记录
        new_context = f"{month}间，鲁迅与{target_name}有频繁书信往来，涉及{keyword_str}"
        evidence_refs = '; '.join(group['Evidence_Ref'].dropna().unique()[:3])
        if len(group) > 3:
            evidence_refs += f" 等{len(group)}条"
        
        aggregated.append({
            'Source_ID': 'ZLH-001',
            'Target_ID': target_id,
            'Relation_Type': '弱关联-通信',
            'Context': new_context,
            'Evidence_Ref': evidence_refs,
            'Weight': 5  # 提升权重
        })
    
    return pd.DataFrame(aggregated)

def main():
    print("=" * 60)
    print("鲁迅节点数据清洗与聚合")
    print("=" * 60)
    
    # 1. 备份
    print(f"\n[1/5] 备份原始文件...")
    shutil.copy2(EXCEL_PATH, BACKUP_PATH)
    print(f"  备份到: {BACKUP_PATH}")
    
    # 2. 读取数据
    print(f"\n[2/5] 读取数据...")
    df = pd.read_excel(EXCEL_PATH, sheet_name='Sheet2')
    original_count = len(df)
    luxun_original = len(df[(df['Source_ID'] == 'ZLH-001') | (df['Target_ID'] == 'ZLH-001')])
    print(f"  原始记录: {original_count}")
    print(f"  鲁迅相关: {luxun_original} ({luxun_original/original_count*100:.1f}%)")
    
    # 3. 分离记录
    print(f"\n[3/5] 分离记录...")
    
    # 非鲁迅记录（保留）
    non_luxun = df[
        (df['Source_ID'] != 'ZLH-001') & (df['Target_ID'] != 'ZLH-001')
    ].copy()
    print(f"  非鲁迅记录: {len(non_luxun)}")
    
    # 鲁迅强关联记录（保留）
    luxun_strong = df[
        ((df['Source_ID'] == 'ZLH-001') | (df['Target_ID'] == 'ZLH-001')) &
        (df['Relation_Type'].str.contains('强关联', na=False))
    ].copy()
    print(f"  鲁迅强关联: {len(luxun_strong)} (保留)")
    
    # 鲁迅弱关联-通信（需聚合）
    luxun_comm = df[
        (df['Source_ID'] == 'ZLH-001') & 
        (df['Relation_Type'] == '弱关联-通信')
    ]
    print(f"  鲁迅弱关联-通信: {len(luxun_comm)} (待聚合)")
    
    # 鲁迅其他弱关联（如时空共现，保留部分）
    luxun_other = df[
        ((df['Source_ID'] == 'ZLH-001') | (df['Target_ID'] == 'ZLH-001')) &
        (~df['Relation_Type'].str.contains('强关联', na=False)) &
        (df['Relation_Type'] != '弱关联-通信')
    ].copy()
    print(f"  鲁迅其他弱关联: {len(luxun_other)}")
    
    # 4. 聚合通信记录
    print(f"\n[4/5] 按月聚合通信记录...")
    aggregated_comm = aggregate_communications(df)
    print(f"  聚合后: {len(aggregated_comm)} (从{len(luxun_comm)}条合并)")
    
    # 5. 合并并写回
    print(f"\n[5/5] 合并数据并写回...")
    
    # 合并所有记录
    final_df = pd.concat([
        non_luxun,
        luxun_strong,
        aggregated_comm,
        luxun_other
    ], ignore_index=True)
    
    # 删除临时列
    if 'Month' in final_df.columns:
        final_df = final_df.drop(columns=['Month'])
    
    # 重新编号
    final_df['序号'] = range(1, len(final_df) + 1)
    
    # 确保列顺序
    columns = ['序号', 'Source_ID', 'Target_ID', 'Relation_Type', 'Context', 'Evidence_Ref', 'Weight']
    final_df = final_df[[c for c in columns if c in final_df.columns]]
    
    # 写入Excel
    with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        final_df.to_excel(writer, sheet_name='Sheet2', index=False)
    
    # 统计结果
    luxun_final = len(final_df[(final_df['Source_ID'] == 'ZLH-001') | (final_df['Target_ID'] == 'ZLH-001')])
    
    print("\n" + "=" * 60)
    print("处理完成!")
    print("=" * 60)
    print(f"\n记录数量变化:")
    print(f"  总记录: {original_count} → {len(final_df)}")
    print(f"  鲁迅相关: {luxun_original} → {luxun_final}")
    print(f"  鲁迅占比: {luxun_original/original_count*100:.1f}% → {luxun_final/len(final_df)*100:.1f}%")
    print(f"\n备份文件: {BACKUP_PATH}")

if __name__ == '__main__':
    main()
