"""正确提取数据底表中每个供应商的总计行"""
import sys
import pandas as pd
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

SAMPLE_DIR = Path(r"c:\Users\chenzishuo\Desktop\新建文件夹 (2)\周报内容\港澳流速\sample")
DATA_FILE = SAMPLE_DIR / "海外港澳商务_各渠道主辅投数据.xlsx"

df = pd.read_excel(DATA_FILE, header=5)

# forward fill 渠道组和供应商（合并单元格）
df["渠道组"] = df["渠道组"].ffill()
df["供应商"] = df["供应商"].ffill()

# 过滤掉口径说明行和总计渠道组
df_clean = df[~df["渠道组"].str.contains("口径|说明|总计", na=False)].copy()
df_clean = df_clean[df_clean["渠道组"].isin(["KOLHK", "主页HK", "独立站HK", "社群HK", "线下HK"])]

# 找到每个供应商的总计行
supplier_totals = df_clean[df_clean["渠道名称"].str.contains("总计", na=False)].copy()

print("📌 KOLHK 渠道组 - 各供应商总计行:")
kol_totals = supplier_totals[supplier_totals["渠道组"] == "KOLHK"]
for _, row in kol_totals.iterrows():
    if row["供应商"] != "总计":
        print(f"  {row['供应商']:20} | 例子={row['例子数']}, 约课={row['约课数']}, 滚动消耗={row['滚动消耗']}")
print(f"  --- KOLHK 总计: 例子={kol_totals[kol_totals['供应商']=='总计']['例子数'].sum()}, 约课={kol_totals[kol_totals['供应商']=='总计']['约课数'].sum()}")

print(f"\n📌 社群HK 渠道组 - 各供应商总计行:")
sq_totals = supplier_totals[supplier_totals["渠道组"] == "社群HK"]
for _, row in sq_totals.iterrows():
    if row["供应商"] != "总计":
        print(f"  {row['供应商']:20} | 例子={row['例子数']}, 约课={row['约课数']}, 滚动消耗={row['滚动消耗']}")
print(f"  --- 社群HK 总计: 例子={sq_totals[sq_totals['供应商']=='总计']['例子数'].sum()}, 约课={sq_totals[sq_totals['供应商']=='总计']['约课数'].sum()}")

print(f"\n📌 线下HK 渠道组 - 各供应商总计行:")
xx_totals = supplier_totals[supplier_totals["渠道组"] == "线下HK"]
for _, row in xx_totals.iterrows():
    if row["供应商"] != "总计":
        print(f"  {row['供应商']:20} | 例子={row['例子数']}, 约课={row['约课数']}, 滚动消耗={row['滚动消耗']}")
print(f"  --- 线下HK 总计: 例子={xx_totals[xx_totals['供应商']=='总计']['例子数'].sum()}, 约课={xx_totals[xx_totals['供应商']=='总计']['约课数'].sum()}")

# 验证：流速表中的供应商能否在数据底表中找到
print(f"\n\n📌 验证流速表供应商匹配:")
flow_suppliers = ["钟嘉欣图片", "周家蔚", "钟嘉欣视频1"]
all_suppliers = supplier_totals["供应商"].unique()
for s in flow_suppliers:
    found = s in all_suppliers
    if found:
        row = supplier_totals[supplier_totals["供应商"] == s].iloc[0]
        print(f"  ✅ {s:20} → 例子={row['例子数']}, 约课={row['约课数']}")
    else:
        print(f"  ❌ {s:20} → 未找到")
