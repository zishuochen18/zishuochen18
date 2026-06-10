import sys
sys.stdout.reconfigure(encoding="utf-8")
import pandas as pd

df = pd.read_excel('港澳流速/sample/海外港澳商务_各渠道主辅投数据 (2).xlsx', header=5)
df['渠道组'] = df['渠道组'].ffill()
print("数据底表中的所有渠道组:")
print(df['渠道组'].unique())
