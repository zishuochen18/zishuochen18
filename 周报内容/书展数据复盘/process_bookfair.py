"""
书展数据整合脚本 - 跨月追踪（含自动下载）

逻辑：
1. 【Phase 3-Δ】自动下载最新数据（截至昨天）
   - 调用 download_bookfair.py，获取上月完整月 + 上月至今数据
2. 读取3个报表进行整合：
   - 5.1-5.31（5月完整月）
   - 5.1-yesterday（5月至今）
   - 6.1-yesterday（6月至今，共享底表）
3. 对书展供应商进行数据整合：
   - 滚动消耗_整合 = 5月完整.滚动消耗 + 6月至今.滚动消耗
   - 滚动ROI2总成本_整合 = 5月完整.滚动ROI2总成本 + 6月至今.滚动ROI2总成本
4. 重新计算 滚动ROI2 = 滚动GMV / 滚动ROI2总成本_整合
"""
import sys
import os
from pathlib import Path
from datetime import datetime, timedelta
from calendar import monthrange
import pandas as pd
from openpyxl import load_workbook

sys.stdout.reconfigure(encoding="utf-8")

# 路径配置
BASE_DIR = Path(__file__).parent
SAMPLE_DIR = BASE_DIR / "sample"
OUTPUT_DIR = BASE_DIR / "output"

# 书展供应商配置
BOOKFAIR_SUPPLIERS = [
    "26年6月",
]


def get_bookfair_periods():
    """
    计算书展数据所需的日期范围

    返回:
        dict: {
            'today': date,
            'yesterday': date,
            'last_month': {'year': 2026, 'month': 5, 'prefix': '26年5月'},
            'current_month': {'year': 2026, 'month': 6, 'prefix': '26年6月'},
            'last_month_full': (start, end),
            'last_month_to_now': (start, end),
            'current_month_to_now': (start, end),
            'is_month_start': bool,
        }
    """
    today = datetime.now().date()
    yesterday = today - timedelta(days=1)

    current_year = today.year
    current_month = today.month

    # 上月计算
    if current_month == 1:
        last_year = current_year - 1
        last_month = 12
    else:
        last_year = current_year
        last_month = current_month - 1

    # 日期范围
    last_month_start = datetime(last_year, last_month, 1).date()
    last_month_end = datetime(last_year, last_month, monthrange(last_year, last_month)[1]).date()
    current_month_start = today.replace(day=1)

    return {
        'today': today,
        'yesterday': yesterday,
        'last_month': {
            'year': last_year,
            'month': last_month,
            'year_short': str(last_year)[2:],
            'prefix': f"{str(last_year)[2:]}年{last_month}月",
        },
        'current_month': {
            'year': current_year,
            'month': current_month,
            'year_short': str(current_year)[2:],
            'prefix': f"{str(current_year)[2:]}年{current_month}月",
        },
        'last_month_full': (last_month_start, last_month_end),
        'last_month_to_now': (last_month_start, yesterday),
        'current_month_to_now': (current_month_start, yesterday),
        'is_month_start': today.day == 1,
    }


def identify_bookfairs(data_file, periods):
    """
    从数据底表识别上月和本月的书展供应商

    参数:
        data_file: 数据底表路径
        periods: get_bookfair_periods() 返回的字典

    返回:
        dict: {
            'last_month': [...],
            'current_month': [...]
        }
    """
    df = pd.read_excel(data_file, header=5)
    df["渠道组"] = df["渠道组"].ffill()
    df["供应商"] = df["供应商"].ffill()

    # 筛选线下HK渠道组
    xiaxian = df[df["渠道组"] == "线下HK"]

    # 上月书展
    last_prefix = periods['last_month']['prefix']
    last_month_fairs = xiaxian[
        xiaxian["供应商"].str.contains(last_prefix, na=False)
    ]["供应商"].unique().tolist()

    # 本月书展
    cur_prefix = periods['current_month']['prefix']
    current_month_fairs = xiaxian[
        xiaxian["供应商"].str.contains(cur_prefix, na=False)
    ]["供应商"].unique().tolist()

    return {
        'last_month': last_month_fairs,
        'current_month': current_month_fairs,
    }



def load_report_data(file_path):
    """加载报表数据，返回DataFrame"""
    df = pd.read_excel(file_path, header=5)
    # 向下填充渠道组和供应商
    df['渠道组'] = df['渠道组'].ffill()
    df['供应商'] = df['供应商'].ffill()
    return df


def find_supplier_rows(df, supplier_name):
    """
    找到指定供应商的所有行（包括渠道名称行和总计行）
    返回：{渠道名称: row_data} 字典
    """
    result = {}
    supplier_rows = df[df['供应商'] == supplier_name]

    for _, row in supplier_rows.iterrows():
        channel_name = row['渠道名称']
        if pd.notna(channel_name):
            result[channel_name] = row.to_dict()

    return result


def merge_bookfair_data():
    """整合书展数据（滚动展示逻辑）"""
    print("="*60)
    print("书展数据整合（滚动展示逻辑）")
    print("="*60)
    print()

    # 【Phase 3-Δ】自动下载最新数据
    print("[0] 自动下载最新数据...")
    from download_bookfair import download_bookfair_data, verify_bookfair_freshness, get_bookfair_periods
    periods = get_bookfair_periods()
    file_paths = download_bookfair_data(periods)
    verify_bookfair_freshness(file_paths, periods)
    print(f"  ✓ 数据已更新至 {periods['yesterday']}")
    print()

    # 从下载结果获取文件路径
    file_may_full = file_paths.get("may_full")
    file_may_to_now = file_paths.get("may_to_now")

    # 获取动态日期（已在上面获取）
    print("[1] 日期与书展识别")
    print(f"  上月: {periods['last_month']['prefix']} ({periods['last_month_to_now'][0]} ~ {periods['last_month_to_now'][1]})")
    print(f"  本月: {periods['current_month']['prefix']} ({periods['current_month_to_now'][0]} ~ {periods['current_month_to_now'][1]})")
    print()

    # 获取港澳流速数据底表（用于书展供应商识别和提取）
    FILE_JUNE = BASE_DIR.parent / "港澳流速" / "sample" / "海外港澳商务_各渠道主辅投数据 (2).xlsx"

    # 识别书展供应商
    bookfairs = identify_bookfairs(FILE_JUNE, periods)
    print("  上月书展: " + str(bookfairs['last_month']))
    print("  本月书展: " + str(bookfairs['current_month']))
    print()

    # 加载3个报表
    print("[2] 加载报表...")
    df_may_full = load_report_data(file_may_full)
    df_may_to_now = load_report_data(file_may_to_now)
    df_june = load_report_data(FILE_JUNE)
    print(f"  上月完整月: {len(df_may_full)} 行")
    print(f"  上月至今: {len(df_may_to_now)} 行")
    print(f"  本月至今: {len(df_june)} 行")

    print()

    # 处理上月书展（跨月整合）
    if bookfairs['last_month']:
        print(f"[3] 整合上月书展数据...")
        merge_last_month_fairs(df_may_full, df_may_to_now, df_june,
                                file_may_to_now,
                                bookfairs['last_month'], periods)
    else:
        print(f"[3] 上月无书展，跳过")

    print()

    # 处理本月书展（本月数据）
    if bookfairs['current_month']:
        print(f"[4] 提取本月书展数据...")
        merge_current_month_fairs(df_june, bookfairs['current_month'], periods)
    else:
        print(f"[4] 本月无书展，跳过")

    print()
    print(f"[完成]")


def merge_last_month_fairs(df_may_full, df_may_to_now, df_june,
                            base_file_path,
                            suppliers, periods):
    """
    整合上月书展数据（跨月累计）

    基底：上月至今表（5.1-昨天），其滚动 GMV/到课/例子数等已经是完整链路累计
    覆写字段：滚动消耗、滚动ROI2总成本、滚动ROI2

    参数:
        df_may_full:    上月完整月 DataFrame（5.1-5.31）
        df_may_to_now:  上月至今 DataFrame（5.1-昨天）—— 作为基底
        df_june:        本月至今 DataFrame（6.1-昨天）
        base_file_path: 上月至今 Excel 文件路径（作为输出基底）
        suppliers:      上月书展供应商列表
        periods:        日期信息
    """
    OUTPUT_DIR.mkdir(exist_ok=True)

    # 准备输出文件
    last_month_prefix = periods['last_month']['prefix'].replace('年', '').replace('月', '')
    cur_day = periods['yesterday'].strftime('%m.%d').lstrip('0').replace('.0', '.')
    output_file = OUTPUT_DIR / f"书展数据整合_上月_{last_month_prefix}-{cur_day}.xlsx"

    import shutil
    shutil.copy(base_file_path, output_file)

    # 用 pandas 读取数据（用于查找）
    df_output = pd.read_excel(output_file, header=5)
    df_output['渠道组'] = df_output['渠道组'].ffill()
    df_output['供应商'] = df_output['供应商'].ffill()

    # 用 openpyxl 修改 Excel 文件
    from openpyxl import load_workbook

    wb = load_workbook(output_file)
    ws = wb.active

    # 构建列名到列索引的映射（从 header 行）
    headers = {}
    for col_idx in range(1, ws.max_column + 1):
        h = ws.cell(6, col_idx).value
        if h:
            headers[str(h).strip()] = col_idx

    print(f"  [准备输出] {output_file.name}")

    for supplier in suppliers:
        print(f"  【{supplier}】")

        # 提取各报表的数据
        may_full_data   = find_supplier_rows(df_may_full, supplier)
        may_to_now_data = find_supplier_rows(df_may_to_now, supplier)
        june_data       = find_supplier_rows(df_june, supplier)

        # 遍历基底（上月至今）的所有渠道
        for channel_name in may_to_now_data.keys():
            may_row        = may_full_data.get(channel_name, {})
            may_to_now_row = may_to_now_data.get(channel_name, {})
            june_row       = june_data.get(channel_name, {})

            # 加和：滚动消耗 + 滚动ROI2总成本
            cost_may  = may_row.get('滚动消耗', 0) or 0
            cost_june = june_row.get('滚动消耗', 0) or 0
            cost_merged = cost_may + cost_june

            roi2_cost_may  = may_row.get('滚动ROI2总成本', 0) or 0
            roi2_cost_june = june_row.get('滚动ROI2总成本', 0) or 0
            roi2_cost_merged = roi2_cost_may + roi2_cost_june

            # GMV 用上月至今（5.1-昨天，已经是完整链路累计）
            gmv = may_to_now_row.get('滚动GMV', 0) or 0
            roi2 = gmv / roi2_cost_merged if roi2_cost_merged > 0 else 0

            # 在 DataFrame 中找到对应行并更新
            mask = (df_output['供应商'] == supplier) & (df_output['渠道名称'] == channel_name)
            if mask.any():
                idx = df_output[mask].index[0]
                df_output.loc[idx, '滚动消耗'] = cost_merged
                df_output.loc[idx, '滚动ROI2总成本'] = roi2_cost_merged
                df_output.loc[idx, '滚动ROI2'] = roi2

                if channel_name == '总计':
                    print(f"      {channel_name}: {cost_may:.0f} + {cost_june:.0f} = {cost_merged:.0f}")

    # 将修改后的 DataFrame 写回 Excel
    from openpyxl import load_workbook

    wb = load_workbook(output_file)
    ws = wb.active

    # 1. 先找到列名在 Excel 中的位置（第 6 行是 header）
    excel_headers = {}
    for col_idx in range(1, ws.max_column + 1):
        h = ws.cell(6, col_idx).value
        if h:
            excel_headers[str(h).strip()] = col_idx

    print(f"  [写入] 找到的列: 滚动消耗={excel_headers.get('滚动消耗')}, ROI2成本={excel_headers.get('滚动ROI2总成本')}, ROI2={excel_headers.get('滚动ROI2')}")

    # 2. 对每一行数据，找到它在 Excel 中对应的位置，然后更新
    for df_idx, row in df_output.iterrows():
        supplier = row.get('供应商')
        channel = row.get('渠道名称')

        # 在 Excel 中查找这一行
        for xlsx_row_idx in range(7, ws.max_row + 1):
            xlsx_supplier = ws.cell(xlsx_row_idx, 1).value or ws.cell(xlsx_row_idx, 3).value  # 第1列或第3列
            xlsx_channel = ws.cell(xlsx_row_idx, 2).value or ws.cell(xlsx_row_idx, 4).value   # 第2列或第4列

            if str(supplier).strip() == str(xlsx_supplier).strip() if xlsx_supplier else False:
                # 更新三个字段
                cost_col = excel_headers.get('滚动消耗')
                roi2_cost_col = excel_headers.get('滚动ROI2总成本')
                roi2_col = excel_headers.get('滚动ROI2')

                if cost_col:
                    ws.cell(xlsx_row_idx, cost_col).value = row.get('滚动消耗')
                if roi2_cost_col:
                    ws.cell(xlsx_row_idx, roi2_cost_col).value = row.get('滚动ROI2总成本')
                if roi2_col:
                    ws.cell(xlsx_row_idx, roi2_col).value = row.get('滚动ROI2')

    wb.save(output_file)
    wb.close()

    # 用 pandas 重新保存一遍，确保格式一致
    df_output.to_excel(output_file, sheet_name='Sheet', index=False, header=True, startrow=5)

    print(f"  ✅ 已保存: {output_file.name}")


def merge_current_month_fairs(df_june, suppliers, periods):
    """
    提取本月书展数据（不需要整合）

    参数:
        df_june: 本月至今DataFrame
        suppliers: 本月书展供应商列表
        periods: 日期信息
    """
    OUTPUT_DIR.mkdir(exist_ok=True)

    # 准备输出文件
    cur_month_prefix = periods['current_month']['prefix'].replace('年', '').replace('月', '')
    cur_day = periods['yesterday'].strftime('%m.%d').lstrip('0').replace('.0', '.')
    output_file = OUTPUT_DIR / f"书展数据整合_本月_{cur_month_prefix}-{cur_day}.xlsx"

    import shutil
    shutil.copy(FILE_JUNE, output_file)

    print(f"  [准备输出] {output_file.name}")

    for supplier in suppliers:
        print(f"  【{supplier}】")
        data = find_supplier_rows(df_june, supplier)
        print(f"    明细行数: {len(data)}")

    print(f"  ✅ 已保存: {output_file.name}")


if __name__ == "__main__":
    merge_bookfair_data()
