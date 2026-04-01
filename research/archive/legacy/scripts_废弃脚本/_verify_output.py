import pandas as pd
import re

df = pd.read_csv(r'd:\1大创\Decoupled_Weighted_Sheet2.csv', encoding='utf-8-sig')
print('行数:', len(df))
print('列名:', list(df.columns))
print()
print('Relation_Type分布:')
for k,v in df['Relation_Type'].value_counts().items():
    print(' ', k, ':', v)
print()
w = pd.to_numeric(df['Weight'], errors='coerce')
print('Weight均值:', round(w.mean(),2), '范围:', w.min(), '~', w.max())
ctx = df['Context'].fillna('').astype(str)
print('平均Context:', round(ctx.str.len().mean()), '字符，最大:', ctx.str.len().max())
dup = int(df.duplicated(subset=['Source_ID','Target_ID','Relation_Type']).sum())
print('重复行:', dup)
CJK = r'[\u4e00-\u9fff]'
ocr = int(ctx.str.contains(rf'{CJK} {CJK}', regex=True).sum())
print('OCR空格残留:', ocr)
multi = int(ctx.str.contains(' / ').sum())
print('多来源拼接:', multi)
prefix = df['Relation_Type'].astype(str).str.contains('^(强|弱)(关联|关系)', regex=True).sum()
print('含强/弱前缀行:', int(prefix))
print()
print('前3行样例:')
for i, r in df.head(3).iterrows():
    src = str(r['Source_ID'])
    tgt = str(r['Target_ID'])
    rel = str(r['Relation_Type'])
    w_  = str(r['Weight'])
    ctx_ = str(r['Context'])[:60]
    print(f'  {src} -> {tgt} | {rel} | W={w_} | {repr(ctx_)}')
