import sys
import os
from pathlib import Path
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

SAMPLE_DIR = Path("sample")
OUTPUT_DIR = Path("output")

# 找到导出的文件
files = list(SAMPLE_DIR.glob("*.xlsx"))
if files:
    excel_file = files[0]
    print(f"使用文件: {excel_file.name}\n")

    # 读取数据
    df = pd.read_excel(excel_file, sheet_name=0)
    print(f"[步骤 4] 计算曝光裂变率...")
    print(f"  读取 {len(df)} 行数据")

    # 查找列名
    poster_id_col = None
    exposure_col = None  # [A]
    auth_col = None      # [F]

    for col in df.columns:
        if "海报ID" in col:
            poster_id_col = col
        if "[A]" in col and "曝光" in col:
            exposure_col = col
        if "[F]" in col and "授权" in col:
            auth_col = col

    print(f"  找到必要的列:")
    print(f"    海报ID: {poster_id_col}")
    print(f"    [A]曝光: {exposure_col}")
    print(f"    [F]授权: {auth_col}")

    # 计算裂变率
    df_calc = df.copy()
    df_calc['裂变率'] = df_calc[auth_col] / df_calc[exposure_col]
    df_calc['裂变率'] = df_calc['裂变率'].replace([float('inf'), float('-inf')], 0)
    df_calc['裂变率'] = df_calc['裂变率'].fillna(0)

    # 筛选和排序
    df_filtered = df_calc[df_calc['裂变率'] > 0].copy()
    df_sorted = df_filtered.sort_values('裂变率', ascending=False).reset_index(drop=True)

    print(f"  筛选结果: {len(df_sorted)} / {len(df_calc)} (裂变率 > 0)")

    # 保存结果
    OUTPUT_DIR.mkdir(exist_ok=True)
    output_name = f"海报裂变率排序_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx"
    output_path = OUTPUT_DIR / output_name

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_sorted.to_excel(writer, sheet_name='裂变率排序', index=False)

    print(f"  结果已保存: {output_path.name}")
    print(f"\n  排序 TOP 10:")
    for idx, row in df_sorted.head(10).iterrows():
        poster_id = row[poster_id_col]
        exposure = row[exposure_col]
        auth = row[auth_col]
        rate = row['裂变率']
        print(f"    [{idx+1}] 海报ID: {poster_id}, 曝光: {exposure}, 授权: {auth}, 裂变率: {rate:.4f}")
else:
    print("没有找到 Excel 文件")
