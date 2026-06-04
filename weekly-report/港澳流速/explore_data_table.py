"""读取数据底表的详细表头和供应商列表"""
import sys
import pandas as pd
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

SAMPLE_DIR = Path(r"c:\Users\chenzishuo\Desktop\新建文件夹 (2)\周报内容\港澳流速\sample")
DATA_FILE = SAMPLE_DIR / "海外港澳商务_各渠道主辅投数据.xlsx"

# 读取数据底表，跳过前面的筛选信息行，找到真正的表头
df_raw = pd.read_excel(DATA_FILE, header=None, nrows=10)
print("前10行结构（找表头）：")
for i in range(min(10, len(df_raw))):
    row_vals = []
    for j in range(min(20, df_raw.shape[1])):
        v = df_raw.iloc[i, j]
        if pd.notna(v):
            row_vals.append(f"[{j}]{v}")
    print(f"  行{i}: {', '.join(row_vals[:20])}")

# 尝试用第5行或第6行作为表头
print("\n\n尝试 header=5:")
df5 = pd.read_excel(DATA_FILE, header=5, nrows=3)
print(f"  列名: {list(df5.columns[:20])}")

print("\n尝试 header=4:")
df4 = pd.read_excel(DATA_FILE, header=4, nrows=3)
print(f"  列名: {list(df4.columns[:20])}")

print("\n尝试 header=6:")
df6 = pd.read_excel(DATA_FILE, header=6, nrows=3)
print(f"  列名: {list(df6.columns[:20])}")

# 找到正确表头后，看供应商列
print("\n\n" + "=" * 80)
print("📊 数据底表完整表头（header=5）")
print("=" * 80)
df = pd.read_excel(DATA_FILE, header=5)
print(f"形状: {df.shape}")
print(f"\n全部列名:")
for i, col in enumerate(df.columns):
    print(f"  [{i}] {col}")

# 看供应商列的唯一值
print(f"\n\n📌 供应商列的唯一值:")
# 找到含"供应商"的列
for col in df.columns:
    if "供应商" in str(col):
        vals = df[col].dropna().unique()
        print(f"  列名: {col}")
        print(f"  唯一值 ({len(vals)} 个):")
        for v in sorted(vals, key=str):
            print(f"    - {v}")
        break

# 找例子数、约课数列
print(f"\n📌 关键指标列:")
for col in df.columns:
    if "例子" in str(col) or "约课" in str(col) or "消耗" in str(col):
        print(f"  - {col}")
