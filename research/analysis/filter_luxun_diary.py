# -*- coding: utf-8 -*-
"""
筛选《鲁迅日记》中与左联关系密切的记录
保留：左联核心成员(ZLH-001~ZLH-050) + 内山完造(ZLH-121) + 许广平(ZLH-122)
删除：其他非核心人物的弱关联记录
"""

import pandas as pd
import shutil
from datetime import datetime

# 配置
EXCEL_PATH = r'd:\1大创\大创数据收集1.xlsx'
BACKUP_PATH = r'd:\1大创\大创数据收集1_备份_{}.xlsx'.format(datetime.now().strftime('%Y%m%d_%H%M%S'))

# 需要保留的人物ID
KEEP_IDS = set()
# 左联核心成员 ZLH-001 ~ ZLH-050
for i in range(1, 51):
    KEEP_IDS.add(f"ZLH-{i:03d}")
# 额外保留：内山完造、许广平
KEEP_IDS.add("ZLH-121")  # 内山完造
KEEP_IDS.add("ZLH-122")  # 许广平

def main():
    print("=" * 60)
    print("筛选《鲁迅日记》左联相关内容")
    print("=" * 60)
    
    # 1. 备份原始文件
    print(f"\n[1/4] 备份原始文件...")
    shutil.copy2(EXCEL_PATH, BACKUP_PATH)
    print(f"  已备份到: {BACKUP_PATH}")
    
    # 2. 读取数据
    print(f"\n[2/4] 读取Sheet2数据...")
    df = pd.read_excel(EXCEL_PATH, sheet_name='Sheet2')
    original_count = len(df)
    print(f"  原始记录数: {original_count}")
    
    # 3. 筛选鲁迅日记数据
    print(f"\n[3/4] 执行筛选...")
    
    # 识别来源类型
    df['来源类型'] = df['Evidence_Ref'].apply(
        lambda x: '鲁迅日记' if '鲁迅日记' in str(x) else '左联回忆录'
    )
    
    # 分离鲁迅日记和左联回忆录数据
    luxun_df = df[df['来源类型'] == '鲁迅日记'].copy()
    zuolian_df = df[df['来源类型'] == '左联回忆录'].copy()
    
    print(f"  鲁迅日记记录: {len(luxun_df)}")
    print(f"  左联回忆录记录: {len(zuolian_df)}")
    
    # 筛选鲁迅日记：保留Target_ID在KEEP_IDS中的记录
    # 或者Relation_Type为强关联的记录
    filtered_luxun = luxun_df[
        (luxun_df['Target_ID'].isin(KEEP_IDS)) |
        (luxun_df['Relation_Type'].str.contains('强关联', na=False))
    ].copy()
    
    deleted_count = len(luxun_df) - len(filtered_luxun)
    print(f"  筛选后鲁迅日记: {len(filtered_luxun)} (删除 {deleted_count} 条)")
    
    # 4. 合并并写回
    print(f"\n[4/4] 写回Excel...")
    
    # 合并筛选后的鲁迅日记和完整的左联回忆录
    final_df = pd.concat([filtered_luxun, zuolian_df], ignore_index=True)
    
    # 删除临时列
    final_df = final_df.drop(columns=['来源类型'], errors='ignore')
    
    # 重新编号
    final_df['序号'] = range(1, len(final_df) + 1)
    
    # 写入Excel
    with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        final_df.to_excel(writer, sheet_name='Sheet2', index=False)
    
    print(f"\n" + "=" * 60)
    print("筛选完成!")
    print(f"  原始记录: {original_count}")
    print(f"  筛选后记录: {len(final_df)}")
    print(f"  删除记录: {original_count - len(final_df)}")
    print(f"  备份文件: {BACKUP_PATH}")
    print("=" * 60)
    
    # 统计删除的记录
    print("\n删除记录的人物分布:")
    deleted_df = luxun_df[~luxun_df.index.isin(filtered_luxun.index)]
    if len(deleted_df) > 0:
        target_counts = deleted_df['Target_ID'].value_counts()
        for tid, count in target_counts.head(10).items():
            print(f"  {tid}: {count}条")

if __name__ == '__main__':
    main()
