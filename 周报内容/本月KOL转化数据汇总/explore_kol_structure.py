"""探查 KOL 汇总表和转化链路表的结构"""
import sys
import pandas as pd
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

SAMPLE_DIR = Path(r"c:\Users\chenzishuo\Desktop\新建文件夹 (2)\周报内容\本月KOL转化数据汇总\sample")
KOL_SUMMARY_FILE = SAMPLE_DIR / "海外港澳商务_各渠道主辅投数据-KOL汇总.xlsx"
CONVERSION_FILE = SAMPLE_DIR / "本月KOL转化链路数据汇总.xlsx"

print("=" * 80)
print("📊 KOL汇总表 sheets")
print("=" * 80)
xl1 = pd.ExcelFile(KOL_SUMMARY_FILE)
for sheet in xl1.sheet_names:
    print(f"  - {sheet}")

print("\n" + "=" * 80)
print("📊 转化链路表 sheets")
print("=" * 80)
xl2 = pd.ExcelFile(CONVERSION_FILE)
for sheet in xl2.sheet_names:
    print(f"  - {sheet}")

# 读取【钟嘉欣--求和】sheet 的结构
print("\n" + "=" * 80)
print("📊 【钟嘉欣--求和】sheet 结构（前10行）")
print("=" * 80)
df_zjx_sum = pd.read_excel(KOL_SUMMARY_FILE, sheet_name="钟嘉欣--求和", header=None, nrows=10)
print(f"形状: {df_zjx_sum.shape}")
for i in range(min(10, len(df_zjx_sum))):
    row_vals = []
    for j in range(min(20, df_zjx_sum.shape[1])):
        v = df_zjx_sum.iloc[i, j]
        if pd.notna(v):
            row_vals.append(f"[{j}]{v}")
    print(f"  行{i}: {', '.join(row_vals[:15])}")

# 读取【钟嘉欣】sheet 的结构
print("\n" + "=" * 80)
print("📊 【钟嘉欣】sheet 结构（前15行）")
print("=" * 80)
df_zjx = pd.read_excel(KOL_SUMMARY_FILE, sheet_name="钟嘉欣", header=None, nrows=15)
print(f"形状: {df_zjx.shape}")
for i in range(min(15, len(df_zjx))):
    row_vals = []
    for j in range(min(20, df_zjx.shape[1])):
        v = df_zjx.iloc[i, j]
        if pd.notna(v):
            row_vals.append(f"[{j}]{v}")
    print(f"  行{i}: {', '.join(row_vals[:15])}")

# 读取【周家蔚】sheet 的结构
print("\n" + "=" * 80)
print("📊 【周家蔚】sheet 结构（前15行）")
print("=" * 80)
df_zjw = pd.read_excel(KOL_SUMMARY_FILE, sheet_name="周家蔚", header=None, nrows=15)
print(f"形状: {df_zjw.shape}")
for i in range(min(15, len(df_zjw))):
    row_vals = []
    for j in range(min(20, df_zjw.shape[1])):
        v = df_zjw.iloc[i, j]
        if pd.notna(v):
            row_vals.append(f"[{j}]{v}")
    print(f"  行{i}: {', '.join(row_vals[:15])}")

# 读取【本月汇总数据】sheet 的结构
print("\n" + "=" * 80)
print("📊 【本月汇总数据】sheet 结构（前20行）")
print("=" * 80)
df_month = pd.read_excel(CONVERSION_FILE, sheet_name="本月汇总数据", header=None, nrows=20)
print(f"形状: {df_month.shape}")
for i in range(min(20, len(df_month))):
    row_vals = []
    for j in range(min(20, df_month.shape[1])):
        v = df_month.iloc[i, j]
        if pd.notna(v):
            row_vals.append(f"[{j}]{v}")
    print(f"  行{i}: {', '.join(row_vals[:15])}")
