"""确认数据底表中每个供应商的总计行提取方式"""
import sys
import pandas as pd
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

SAMPLE_DIR = Path(r"c:\Users\chenzishuo\Desktop\新建文件夹 (2)\周报内容\港澳流速\sample")
DATA_FILE = SAMPLE_DIR / "海外港澳商务_各渠道主辅投数据.xlsx"

df = pd.read_excel(DATA_FILE, header=5)

# forward fill 渠道组
df["渠道组"] = df["渠道组"].ffill()

# 过滤掉口径说明行
df = df[~df["渠道组"].str.contains("口径|说明", na=False)]

# 找到每个供应商的总计行（渠道名称含"总计"）
print("📌 KOLHK 渠道组 - 各供应商总计行:")
kol_df = df[df["渠道组"] == "KOLHK"]
kol_totals = kol_df[kol_df["渠道名称"].str.contains("总计", na=False)]
for _, row in kol_totals.iterrows():
    print(f"  {row['供应商']:20} | 例子={row['例子数']}, 约课={row['约课数']}, 滚动消耗={row['滚动消耗']}")

print(f"\n📌 社群HK 渠道组 - 各供应商总计行:")
sq_df = df[df["渠道组"] == "社群HK"]
sq_totals = sq_df[sq_df["渠道名称"].str.contains("总计", na=False)]
for _, row in sq_totals.iterrows():
    print(f"  {row['供应商']:20} | 例子={row['例子数']}, 约课={row['约课数']}, 滚动消耗={row['滚动消耗']}")

print(f"\n📌 线下HK 渠道组 - 各供应商总计行:")
xx_df = df[df["渠道组"] == "线下HK"]
xx_totals = xx_df[xx_df["渠道名称"].str.contains("总计", na=False)]
for _, row in xx_totals.iterrows():
    print(f"  {row['供应商']:20} | 例子={row['例子数']}, 约课={row['约课数']}, 滚动消耗={row['滚动消耗']}")

# 验证：KOLHK 总计
print(f"\n📌 KOLHK 渠道组总计（供应商='总计'的行）:")
kol_grand = kol_df[kol_df["供应商"] == "总计"]
for _, row in kol_grand.iterrows():
    print(f"  例子={row['例子数']}, 约课={row['约课数']}, 滚动消耗={row['滚动消耗']}")
