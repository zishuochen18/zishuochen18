import sys
sys.stdout.reconfigure(encoding="utf-8")
import pandas as pd

df = pd.read_excel('港澳流速/sample/海外港澳商务_各渠道主辅投数据 (2).xlsx', header=5)
df['渠道组'] = df['渠道组'].ffill()
df['供应商'] = df['供应商'].ffill()

# 过滤
df = df[df['渠道组'].isin(['KOLHK', '主页HK', '独立站HK', '社群HK', '线下HK', '总计'])]

# 找总计行
total_rows = df[df['渠道组'] == '总计']
print("数据底表 - '总计' 行数据:")
for _, row in total_rows.iterrows():
    print(f"  供应商: {row['供应商']}")
    print(f"  例子数: {row.get('例子数', 0)}")
    print(f"  约课数: {row.get('约课数', 0)}")
    print(f"  滚动消耗: {row.get('滚动消耗', 0)}")
