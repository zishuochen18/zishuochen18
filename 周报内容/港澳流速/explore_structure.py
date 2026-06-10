"""读取港澳流速表和数据底表的结构"""
import sys
import pandas as pd
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

SAMPLE_DIR = Path(r"c:\Users\chenzishuo\Desktop\新建文件夹 (2)\周报内容\港澳流速\sample")
FLOW_FILE = SAMPLE_DIR / "2026年5月港澳市场流速-初稿4.21.xlsx"
DATA_FILE = SAMPLE_DIR / "海外港澳商务_各渠道主辅投数据.xlsx"

# 1. 流速表 sheets
print("=" * 80)
print("📊 流速表 sheets")
print("=" * 80)
xl = pd.ExcelFile(FLOW_FILE)
for sheet in xl.sheet_names:
    print(f"  - {sheet}")

# 2. 读取【香港市场目标】sheet 的表头
print("\n" + "=" * 80)
print("📊 流速表【香港市场目标】sheet 结构")
print("=" * 80)
df_flow = pd.read_excel(FLOW_FILE, sheet_name="香港市场目标", header=None, nrows=5)
print(f"前5行 × {df_flow.shape[1]} 列")
# 打印前几行看表头结构
for i in range(min(5, len(df_flow))):
    row_vals = []
    for j in range(min(50, df_flow.shape[1])):
        v = df_flow.iloc[i, j]
        if pd.notna(v):
            row_vals.append(f"[{j}]{v}")
    print(f"  行{i}: {', '.join(row_vals[:20])}")

# 重点看 AO-AU 列（索引 40-46）
print(f"\n📌 AO-AU 列区域（索引 40-46）：")
for i in range(min(5, len(df_flow))):
    row_vals = []
    for j in range(40, min(47, df_flow.shape[1])):
        v = df_flow.iloc[i, j]
        if pd.notna(v):
            row_vals.append(f"[{j}]{v}")
    if row_vals:
        print(f"  行{i}: {', '.join(row_vals)}")

# 看 B 列（供应商列）
print(f"\n📌 B 列（供应商）前 30 行：")
df_flow_full = pd.read_excel(FLOW_FILE, sheet_name="香港市场目标", header=None, nrows=50)
for i in range(min(50, len(df_flow_full))):
    v = df_flow_full.iloc[i, 1]  # B列 = 索引1
    if pd.notna(v):
        print(f"  行{i}: {v}")

print("\n" + "=" * 80)
print("📊 数据底表 sheets")
print("=" * 80)
xl2 = pd.ExcelFile(DATA_FILE)
for sheet in xl2.sheet_names:
    print(f"  - {sheet}")

# 读取数据底表表头
print("\n📊 数据底表第一个 sheet 结构")
df_data = pd.read_excel(DATA_FILE, header=None, nrows=5)
print(f"前5行 × {df_data.shape[1]} 列")
for i in range(min(5, len(df_data))):
    row_vals = []
    for j in range(min(30, df_data.shape[1])):
        v = df_data.iloc[i, j]
        if pd.notna(v):
            row_vals.append(f"[{j}]{v}")
    print(f"  行{i}: {', '.join(row_vals[:20])}")
