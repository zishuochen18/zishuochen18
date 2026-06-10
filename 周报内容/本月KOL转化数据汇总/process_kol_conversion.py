"""
处理 KOL 转化数据（v2 - 按列标题匹配 + 写入值）
1. 从数据底表提取钟嘉欣（图片+视频1+视频）和周家蔚的数据
2. 自己计算钟嘉欣三个供应商的汇总值（不依赖 Excel 公式）
3. 按列标题匹配填充各 sheet（写入"值"，不是公式）
"""
"""
本月 KOL 转化数据处理脚本（模块 2）

═══════════════════════════════════════════════════════════════
功能：从数据底表提取 KOLHK 转化数据，生成钟嘉欣/周家蔚的月度汇总
═══════════════════════════════════════════════════════════════

【输入】
- sample/海外港澳商务_各渠道主辅投数据-KOL汇总.xlsx（KOL 汇总模板）
- sample/本月KOL转化链路数据汇总.xlsx（链路数据模板）
- ../港澳流速/sample/海外港澳商务_各渠道主辅投数据.xlsx（共享数据底表）

【输出】
- output/海外港澳商务_各渠道主辅投数据-KOL汇总.xlsx
- output/本月KOL转化链路数据汇总.xlsx

【处理逻辑】

1. 钟嘉欣汇总（钟嘉欣 = 图片 + 视频1 + 视频）
   - sample 中【钟嘉欣--求和】sheet 写入 3 个供应商各自数据
   - 公式自动累加为汇总行
   - 把汇总值（粘贴值）写入【钟嘉欣】sheet 的本月行

2. 周家蔚（单一供应商）
   - 直接从数据底表取周家蔚的总计行
   - 写入【周家蔚】sheet 的本月行

3. 本月转化链路数据汇总
   - 钟嘉欣本月数据 + 钟嘉欣 1月至今汇总
   - 周家蔚本月数据 + 周家蔚 3月至今汇总
   - KOLHK 整体本月数据

4. 钟嘉欣图片&视频明细
   - 从【钟嘉欣--求和】sheet 取 3 个供应商各自数据

【列对应】（按列标题精确匹配，不依赖列索引）
- 滚动消耗、滚动消耗(不含赠课成本)、单例子成本
- 例子数、分发数、约课数、应到课数、滚动应到课数
- 到课数、滚动到课数、当月成交数、滚动成交数
- 当月GMV、滚动GMV、海外GMV占比、滚动ASP
- 各种比率、滚动ROI2总成本、滚动ROI2

【关键变量】
- CURRENT_MONTH_SERIAL：当前月份的 Excel 日期序列号
  例：6月 = 46174（2026-06-01）

【注意事项】
- 写入"值"而不是公式（避免数据混乱）
- 按列标题精确匹配，重复列名取第一次出现的列
- 钟嘉欣三个供应商（图片/视频1/视频）分别填入【钟嘉欣--求和】sheet
"""
import sys
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

sys.stdout.reconfigure(encoding="utf-8")

# 路径配置
BASE_DIR = Path(__file__).parent
SAMPLE_DIR = BASE_DIR / "sample"
DATA_FILE = Path(r"c:\Users\chenzishuo\Desktop\新建文件夹 (2)\周报内容\港澳流速\sample\海外港澳商务_各渠道主辅投数据 (2).xlsx")
KOL_SUMMARY_FILE = SAMPLE_DIR / "海外港澳商务_各渠道主辅投数据-KOL汇总.xlsx"
CONVERSION_FILE = SAMPLE_DIR / "本月KOL转化链路数据汇总.xlsx"
OUTPUT_DIR = BASE_DIR / "output"

# 当前月份（6月对应的 Excel 日期序列号）
CURRENT_MONTH_SERIAL = 46174


def load_data_table():
    """从数据底表提取 KOLHK 供应商数据"""
    df = pd.read_excel(DATA_FILE, header=5)
    df["渠道组"] = df["渠道组"].ffill()
    df["供应商"] = df["供应商"].ffill()

    df = df[df["渠道组"] == "KOLHK"]

    supplier_totals = df[df["渠道名称"].str.contains("总计", na=False)].copy()

    result = {}
    for _, row in supplier_totals.iterrows():
        name = row["供应商"]
        if name == "总计":
            result["KOLHK_总计"] = row.to_dict()
        else:
            result[name] = row.to_dict()

    return result


def load_zjx_video_channels():
    """提取【钟嘉欣视频】供应商的渠道明细（4个抖音渠道）"""
    df = pd.read_excel(DATA_FILE, header=5)
    df["渠道组"] = df["渠道组"].ffill()
    df["供应商"] = df["供应商"].ffill()

    # 筛选钟嘉欣视频供应商，排除总计行
    zjx_video = df[(df["供应商"] == "钟嘉欣视频") & (~df["渠道名称"].str.contains("总计", na=False))].copy()

    # 提取渠道简称（抖音-1/3/4/5）
    zjx_video["渠道简称"] = zjx_video["渠道名称"].str.extract(r'抖音-(\d+)')[0].apply(lambda x: f"抖音-{x}" if pd.notna(x) else "未知")

    result = []
    for _, row in zjx_video.iterrows():
        result.append({
            "渠道简称": row["渠道简称"],
            "渠道名称": row["渠道名称"],
            "例子数": row.get("例子数", 0),
            "约课数": row.get("约课数", 0),
            "滚动消耗": row.get("滚动消耗", 0),
            "滚动消耗(不含赠课成本)": row.get("滚动消耗(不含赠课成本)", 0),
            "分发数": row.get("分发数", 0),
            "应到课数": row.get("应到课数", 0),
            "滚动应到课数": row.get("滚动应到课数", 0),
            "到课数": row.get("到课数", 0),
            "滚动到课数": row.get("滚动到课数", 0),
            "当月成交数": row.get("当月成交数", 0),
            "滚动成交数": row.get("滚动成交数", 0),
            "当月GMV": row.get("当月GMV", 0),
            "滚动GMV": row.get("滚动GMV", 0),
        })

    # 计算各渠道的转化率指标
    for r in result:
        r["例子约课率"] = r["约课数"] / r["例子数"] if r["例子数"] > 0 else 0
        r["约课到课率"] = r["滚动到课数"] / r["约课数"] if r["约课数"] > 0 else 0
        r["到课转化率"] = r["当月成交数"] / r["到课数"] if r["到课数"] > 0 else 0
        r["滚动转化率"] = r["滚动成交数"] / r["例子数"] if r["例子数"] > 0 else 0
        r["滚动ROI2"] = 0  # ROI2需要ROI2总成本，底表渠道级别没有，暂时为0

    return result


def get(d, k):
    """安全获取数值"""
    if not d:
        return 0
    v = d.get(k)
    return v if isinstance(v, (int, float)) else 0


def calc_zjx_aggregate(suppliers):
    """按业务规则计算钟嘉欣三个供应商的汇总值"""
    pic = suppliers.get("钟嘉欣图片", {})
    v1 = suppliers.get("钟嘉欣视频1", {})
    v2 = suppliers.get("钟嘉欣视频", {})

    result = {}

    # 直接累加的字段
    sum_fields = [
        "滚动消耗", "滚动消耗(不含赠课成本)", "例子数", "分发数",
        "分发数\n(剔除毛例子)", "约课数", "应到课数", "滚动应到课数",
        "到课数", "滚动到课数", "当月成交数", "滚动成交数",
        "当月GMV", "滚动GMV",
        "滚动应到课人数", "滚动到课人数", "滚动到课人次",
        "运营成本", "销售固定成本", "销售变动成本", "TMK费用", "DEMO费用",
        "滚动ROI2总成本"
    ]
    for f in sum_fields:
        result[f] = get(pic, f) + get(v1, f) + get(v2, f)

    # 计算比率字段
    if result["例子数"] > 0:
        result["单例子成本"] = result["滚动消耗(不含赠课成本)"] / result["例子数"]
        result["例子分发率"] = result["分发数"] / result["例子数"]
        result["例子约课率"] = result["约课数"] / result["例子数"]
        result["应到课率"] = result["应到课数"] / result["例子数"]
    else:
        result["单例子成本"] = 0
        result["例子分发率"] = 0
        result["例子约课率"] = 0
        result["应到课率"] = 0

    result["分发约课率"] = result["约课数"] / result["分发数"] if result["分发数"] > 0 else 0
    result["约课到课率"] = result["滚动到课数"] / result["约课数"] if result["约课数"] > 0 else 0
    result["到课转化率"] = result["当月成交数"] / result["到课数"] if result["到课数"] > 0 else 0
    result["滚动应到课率"] = result["滚动到课数"] / result["滚动应到课数"] if result["滚动应到课数"] > 0 else 0
    # 滚动转化率 = 滚动成交数 / 例子数
    result["滚动转化率"] = result["滚动成交数"] / result["例子数"] if result["例子数"] > 0 else 0
    result["滚动ROI2"] = result["滚动GMV"] / result["滚动ROI2总成本"] if result["滚动ROI2总成本"] > 0 else 0
    result["滚动ASP"] = result["滚动GMV"] / result["滚动成交数"] if result["滚动成交数"] > 0 else 0

    # 海外GMV占比 ≈ 1（海外渠道全部归海外）
    result["海外GMV占比"] = 1.0
    # 注册转化率 = 当月成交数 / 例子数
    result["注册转化率"] = result["当月成交数"] / result["例子数"] if result["例子数"] > 0 else 0

    return result


def get_column_mapping(ws, header_row=2):
    """读取 sheet 的列标题，返回 {标题: 列索引} 映射（只取第一次出现的列）"""
    mapping = {}
    for col_idx in range(1, ws.max_column + 1):
        header = ws.cell(header_row, col_idx).value
        if header is not None:
            header_str = str(header).strip()
            if header_str not in mapping:  # 只取第一次出现，避免辅投/汇总重复列覆盖
                mapping[header_str] = col_idx
    return mapping


def fill_row_by_columns(ws, target_row, source_data, header_row=2):
    """按列标题匹配，将源数据值写入目标行"""
    col_mapping = get_column_mapping(ws, header_row)
    for col_name, col_idx in col_mapping.items():
        if col_name in source_data:
            val = source_data[col_name]
            if isinstance(val, (int, float)):
                ws.cell(target_row, col_idx).value = val


def find_or_insert_month_row(ws, month_serial):
    """找到月份对应的行，没有就在汇总行上方插入"""
    target_row = None
    for row_idx in range(3, ws.max_row + 1):
        cell_val = ws.cell(row_idx, 1).value
        if cell_val == month_serial:
            target_row = row_idx
            break

    if not target_row:
        for row_idx in range(ws.max_row, 2, -1):
            cell_val = ws.cell(row_idx, 1).value
            if cell_val and "汇总" in str(cell_val):
                target_row = row_idx
                ws.insert_rows(target_row)
                ws.cell(target_row, 1).value = month_serial
                break

    return target_row


def fill_kol_summary_file(data):
    """填充【海外港澳商务_各渠道主辅投数据-KOL汇总.xlsx】"""
    OUTPUT_DIR.mkdir(exist_ok=True)
    import shutil
    output_file = OUTPUT_DIR / KOL_SUMMARY_FILE.name
    shutil.copy(KOL_SUMMARY_FILE, output_file)

    wb = load_workbook(output_file)

    # --- 1. 填充【钟嘉欣--求和】sheet (图片/视频1/视频 三行) ---
    ws_sum = wb["钟嘉欣--求和"]
    # 标题行在行2（openpyxl 1-based）
    fill_row_by_columns(ws_sum, 3, data.get("钟嘉欣图片", {}), header_row=2)
    fill_row_by_columns(ws_sum, 4, data.get("钟嘉欣视频1", {}), header_row=2)
    fill_row_by_columns(ws_sum, 5, data.get("钟嘉欣视频", {}), header_row=2)
    print(f"  [钟嘉欣--求和] 已填充图片/视频1/视频供应商行")

    # --- 2. 计算钟嘉欣汇总值（用于填充【钟嘉欣】sheet 5月行）---
    zjx_aggregate = calc_zjx_aggregate(data)
    print(f"  钟嘉欣汇总: 例子数={zjx_aggregate['例子数']}, 约课数={zjx_aggregate['约课数']}, "
          f"滚动消耗={round(zjx_aggregate['滚动消耗'], 2)}")

    # --- 3. 填充【钟嘉欣】sheet 的5月行（写入计算后的值）---
    ws_zjx = wb["钟嘉欣"]
    target_row = find_or_insert_month_row(ws_zjx, CURRENT_MONTH_SERIAL)
    fill_row_by_columns(ws_zjx, target_row, zjx_aggregate, header_row=2)
    print(f"  [钟嘉欣] 已填充6月数据到行{target_row}（粘贴值）")

    # --- 4. 填充【周家蔚】sheet 的当月行 (6月暂不统计周家蔚，已跳过) ---
    # ws_zjw = wb["周家蔚"]
    # target_row = find_or_insert_month_row(ws_zjw, CURRENT_MONTH_SERIAL)
    # fill_row_by_columns(ws_zjw, target_row, data.get("周家蔚", {}), header_row=2)
    # print(f"  [周家蔚] 已填充6月数据到行{target_row}")
    print(f"  [周家蔚] 6月暂不统计，已跳过")

    wb.save(output_file)
    print(f"  已保存: {output_file}")
    return zjx_aggregate


def calc_kolhk_aggregate(data):
    """计算 KOLHK 渠道组当月汇总（基于所有 KOLHK 供应商）"""
    result = {}

    sum_fields = [
        "滚动消耗", "滚动消耗(不含赠课成本)", "例子数", "分发数",
        "分发数\n(剔除毛例子)", "约课数", "应到课数", "滚动应到课数",
        "到课数", "滚动到课数", "当月成交数", "滚动成交数",
        "当月GMV", "滚动GMV",
        "滚动应到课人数", "滚动到课人数", "滚动到课人次",
        "运营成本", "销售固定成本", "销售变动成本", "TMK费用", "DEMO费用",
        "滚动ROI2总成本"
    ]

    # KOLHK 总计已经在数据底表里
    if "KOLHK_总计" in data:
        kol_total = data["KOLHK_总计"]
        for f in sum_fields:
            result[f] = get(kol_total, f)
        # 比率字段直接读
        rate_fields = [
            "单例子成本", "例子分发率", "分发约课率", "例子约课率",
            "约课到课率", "应到课率", "滚动应到课率", "到课转化率",
            "注册转化率", "滚动转化率", "滚动ROI2", "滚动ASP", "海外GMV占比"
        ]
        for f in rate_fields:
            v = kol_total.get(f)
            result[f] = v if isinstance(v, (int, float)) else 0

    return result


def fill_conversion_file(data, zjx_aggregate):
    """填充【本月KOL转化链路数据汇总.xlsx】"""
    import shutil
    output_file = OUTPUT_DIR / CONVERSION_FILE.name
    shutil.copy(CONVERSION_FILE, output_file)

    wb = load_workbook(output_file)

    # --- 1. 填充【本月汇总数据】---
    ws_month = wb["本月汇总数据"]
    # 标题行在行2

    # 找到6月钟嘉欣行（A列含"6月钟嘉欣"）
    zjx_may_row = None
    for row_idx in range(3, ws_month.max_row + 1):
        cell_val = ws_month.cell(row_idx, 1).value
        if cell_val and "6月钟嘉欣" in str(cell_val):
            zjx_may_row = row_idx
            break

    if zjx_may_row:
        fill_row_by_columns(ws_month, zjx_may_row, zjx_aggregate, header_row=2)
        print(f"  [本月汇总数据] 钟嘉欣6月数据已填充到行{zjx_may_row}")

    # 找到6月周家蔚行
    zjw_may_row = None
    for row_idx in range(3, ws_month.max_row + 1):
        cell_val = ws_month.cell(row_idx, 1).value
        if cell_val and "6月周家蔚" in str(cell_val):
            zjw_may_row = row_idx
            break

    if zjw_may_row:
        fill_row_by_columns(ws_month, zjw_may_row, data.get("周家蔚", {}), header_row=2)
        print(f"  [本月汇总数据] 周家蔚6月数据已填充到行{zjw_may_row}")

    # 填充6月KOLHK汇总行
    kolhk_may_row = None
    for row_idx in range(3, ws_month.max_row + 1):
        cell_val = ws_month.cell(row_idx, 1).value
        if cell_val and "6月-KOLHK汇总" in str(cell_val):
            kolhk_may_row = row_idx
            break

    if kolhk_may_row:
        kolhk_data = calc_kolhk_aggregate(data)
        fill_row_by_columns(ws_month, kolhk_may_row, kolhk_data, header_row=2)
        print(f"  [本月汇总数据] KOLHK 6月汇总已填充到行{kolhk_may_row}")

    # --- 2. 填充【钟嘉欣图片&视频数据】（不分月份，直接按供应商名匹配 A 列）---
    ws_detail = wb["钟嘉欣图片&视频数据"]
    for row_idx in range(3, ws_detail.max_row + 1):
        a_val = ws_detail.cell(row_idx, 1).value
        if not a_val:
            continue
        a_str = str(a_val).strip()
        if a_str == "钟嘉欣图片":
            fill_row_by_columns(ws_detail, row_idx, data.get("钟嘉欣图片", {}), header_row=2)
            print(f"  [钟嘉欣图片&视频数据] 钟嘉欣图片已填充到行{row_idx}")
        elif a_str == "钟嘉欣视频1":
            fill_row_by_columns(ws_detail, row_idx, data.get("钟嘉欣视频1", {}), header_row=2)
            print(f"  [钟嘉欣图片&视频数据] 钟嘉欣视频1已填充到行{row_idx}")
        elif a_str == "钟嘉欣视频":
            v2_data = data.get("钟嘉欣视频", {})
            if v2_data:
                fill_row_by_columns(ws_detail, row_idx, v2_data, header_row=2)
                print(f"  [钟嘉欣图片&视频数据] 钟嘉欣视频已填充到行{row_idx}")

    # --- 3. 新增【钟嘉欣视频渠道明细】sheet ---
    video_channels = load_zjx_video_channels()
    if "钟嘉欣视频渠道明细" not in wb.sheetnames:
        ws_video_detail = wb.create_sheet("钟嘉欣视频渠道明细")
        # 写入标题行
        headers = ["渠道简称", "例子数", "约课数", "滚动消耗", "例子约课率", "约课到课率", "到课转化率", "滚动转化率", "滚动GMV", "滚动ROI2"]
        for col_idx, h in enumerate(headers, start=1):
            ws_video_detail.cell(1, col_idx).value = h

        # 写入数据行
        for row_idx, ch in enumerate(video_channels, start=2):
            ws_video_detail.cell(row_idx, 1).value = ch["渠道简称"]
            ws_video_detail.cell(row_idx, 2).value = ch["例子数"]
            ws_video_detail.cell(row_idx, 3).value = ch["约课数"]
            ws_video_detail.cell(row_idx, 4).value = ch["滚动消耗"]
            ws_video_detail.cell(row_idx, 5).value = ch["例子约课率"]
            ws_video_detail.cell(row_idx, 6).value = ch["约课到课率"]
            ws_video_detail.cell(row_idx, 7).value = ch["到课转化率"]
            ws_video_detail.cell(row_idx, 8).value = ch["滚动转化率"]
            ws_video_detail.cell(row_idx, 9).value = ch["滚动GMV"]
            ws_video_detail.cell(row_idx, 10).value = ch["滚动ROI2"]

        print(f"  [钟嘉欣视频渠道明细] 新建sheet，已写入{len(video_channels)}个渠道")
    else:
        print(f"  [钟嘉欣视频渠道明细] sheet已存在，跳过创建")

    wb.save(output_file)
    print(f"  已保存: {output_file}")


def main():
    print("=" * 60)
    print("KOL 转化数据处理 v2")
    print("=" * 60)

    print("\n[1] 加载数据底表...")
    data = load_data_table()
    print(f"  KOLHK 供应商: {[k for k in data.keys() if k != 'KOLHK_总计']}")

    print("\n[2] 处理【海外港澳商务_各渠道主辅投数据-KOL汇总.xlsx】...")
    zjx_aggregate = fill_kol_summary_file(data)

    print("\n[3] 处理【本月KOL转化链路数据汇总.xlsx】...")
    fill_conversion_file(data, zjx_aggregate)

    print("\n[完成]")


if __name__ == "__main__":
    main()
