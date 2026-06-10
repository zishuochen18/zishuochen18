"""探查书展数据汇总表结构"""
import sys
import pandas as pd
from openpyxl import load_workbook
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

FILE = Path(r"c:\Users\chenzishuo\Desktop\新建文件夹 (2)\周报内容\书展内容\sample\书展数据汇总.xlsx")

print("📊 sheet 列表：")
xl = pd.ExcelFile(FILE)
for s in xl.sheet_names:
    print(f"  - {s}")

# 读取第一个 sheet
print(f"\n=== Sheet 1 ({xl.sheet_names[0]}) 前20行结构 ===")
df = pd.read_excel(FILE, sheet_name=xl.sheet_names[0], header=None, nrows=30)
print(f"形状: {df.shape}")
for i in range(min(30, len(df))):
    row_vals = []
    for j in range(min(40, df.shape[1])):
        v = df.iloc[i, j]
        if pd.notna(v):
            row_vals.append(f"[{j}]{v}")
    if row_vals:
        print(f"  行{i}: {', '.join(row_vals[:30])}")
