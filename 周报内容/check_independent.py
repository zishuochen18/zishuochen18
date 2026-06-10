import sys
sys.stdout.reconfigure(encoding="utf-8")
import pandas as pd

df = pd.read_excel('港澳流速/sample/海外港澳商务_各渠道主辅投数据 (2).xlsx', header=5)
df['渠道组'] = df['渠道组'].ffill()
df['供应商'] = df['供应商'].ffill()

# 过滤
df = df[df['渠道组'].isin(['KOLHK', '主页HK', '独立站HK', '社群HK', '线下HK', '总计'])]

# 找独立站HK的总计行
supplier_totals = df[df['渠道名称'].str.contains('总计', na=False)]
ind_rows = supplier_totals[supplier_totals['渠道组'] == '独立站HK']

print("独立站HK 供应商:")
for name in ind_rows['供应商'].values:
    print(f"  {name}")

print(f"\n独立站HK 总计数据:")
if len(ind_rows) > 0:
    row = ind_rows.iloc[0]
    print(f"  例子数: {row.get('例子数', 0)}")
    print(f"  约课数: {row.get('约课数', 0)}")
    print(f"  滚动消耗: {row.get('滚动消耗', 0)}")
