import pandas as pd

# Read Excel file
df = pd.read_excel(r'c:\Users\33158\Desktop\大创\大创数据收集1.xlsx', sheet_name='Sheet1')

print("Original data shape:", df.shape)

# 1. 角色调整：将特定作家的 Role 改为 "相关人士"
role_change_names = ["沈从文", "戴望舒", "施蛰存", "刘呐鸥", "穆时英", "叶灵凤"]
mask = df['True_Name'].isin(role_change_names)
df.loc[mask, 'Role'] = "相关人士"
print(f"Updated Role for {mask.sum()} rows: {role_change_names}")

# 2. 笔名去重与修正
# 从钱杏邨的 Alias 中删除"钱大昕"
qian_mask = df['True_Name'] == '钱杏邨'
if qian_mask.any():
    alias = df.loc[qian_mask, 'Alias'].values[0]
    if alias and '钱大昕' in str(alias):
        new_alias = str(alias).replace('钱大昕', '').replace('、、', '、').strip('、')
        df.loc[qian_mask, 'Alias'] = new_alias
        print(f"Removed '钱大昕' from 钱杏邨's Alias")

# 从鲁迅的 Alias 中删除"巴人"
luxun_mask = df['True_Name'] == '鲁迅'
if luxun_mask.any():
    alias = df.loc[luxun_mask, 'Alias'].values[0]
    if alias and '巴人' in str(alias):
        new_alias = str(alias).replace('巴人', '').replace('、、', '、').strip('、')
        df.loc[luxun_mask, 'Alias'] = new_alias
        print(f"Removed '巴人' from 鲁迅's Alias")

# 3. 身份信息更正：将"蒲风"改为"黄其奎"
pufeng_mask = df['True_Name'] == '蒲风'
if pufeng_mask.any():
    df.loc[pufeng_mask, 'True_Name'] = '黄其奎'
    print(f"Changed True_Name from '蒲风' to '黄其奎'")

# 4. 保存为新文件
output_path = r'c:\Users\33158\Desktop\大创\大创数据收集_修正版.xlsx'
df.to_excel(output_path, sheet_name='Sheet1', index=False)
print(f"\nSaved to: {output_path}")

# Verify changes
print("\n=== Verification ===")
for name in role_change_names:
    row = df[df['True_Name'] == name]
    if not row.empty:
        print(f"{name}: Role = {row['Role'].values[0]}")

qian = df[df['True_Name'] == '钱杏邨']
if not qian.empty:
    print(f"钱杏邨 Alias: {qian['Alias'].values[0]}")

luxun = df[df['True_Name'] == '鲁迅']
if not luxun.empty:
    print(f"鲁迅 Alias: {luxun['Alias'].values[0]}")

huangqk = df[df['True_Name'] == '黄其奎']
if not huangqk.empty:
    print(f"黄其奎 (原蒲风) found with ID: {huangqk['Entity_ID'].values[0]}")
