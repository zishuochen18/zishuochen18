"""查看数据底表的渠道组分类和总计行结构"""
import sys
import pandas as pd
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

SAMPLE_DIR = Path(r"c:\Users\chenzishuo\Desktop\新建文件夹 (2)\周报内容\港澳流速\sample")
DATA_FILE = SAMPLE_DIR / "海外港澳商务_各渠道主辅投数据.xlsx"

df = pd.read_excel(DATA_FILE, header=5)

# 渠道组唯一值
print("📌 渠道组列的唯一值:")
for v in df["渠道组"].dropna().unique():
    print(f"  - {v}")

# 看每个供应商对应的渠道组
print("\n📌 供应商 → 渠道组 映射:")
for _, row in df[df["供应商"].notna() & (df["供应商"] != "总计")].drop_duplicates(subset=["供应商"]).iterrows():
    print(f"  {row['渠道组']:12} | {row['供应商']}")

# 看总计行的结构
print("\n📌 总计行（供应商='总计'）:")
totals = df[df["供应商"] == "总计"]
for _, row in totals.iterrows():
    print(f"  渠道组={row['渠道组']}, 例子数={row['例子数']}, 约课数={row['约课数']}, 滚动消耗={row['滚动消耗']}")

# 看每个供应商的总计行（渠道名称='总计'的行）
print("\n📌 各供应商总计行（前20个）:")
supplier_totals = df[df["渠道名称"] == "总计"]
for i, (_, row) in enumerate(supplier_totals.iterrows()):
    if i >= 25:
        break
    print(f"  {row['渠道组']:12} | {row['供应商']:20} | 例子={row['例子数']}, 约课={row['约课数']}, 滚动消耗={row['滚动消耗']}")
