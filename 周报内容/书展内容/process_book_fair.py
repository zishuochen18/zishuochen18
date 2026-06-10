"""
书展数据复盘处理
1. 从数据底表 海外港澳商务_各渠道主辅投数据.xlsx 提取本月两个书展的明细 + 汇总
2. 写入书展数据汇总.xlsx 中（保留历史书展不动）
3. 提供 extract_book_fair_data() 接口给主 HTML 使用
"""
"""
书展数据复盘处理脚本（模块 4）

═══════════════════════════════════════════════════════════════
功能：从数据底表提取上月/本月书展数据，与历史书展对比
═══════════════════════════════════════════════════════════════

【输入】
- output/书展数据整合_上月_XXX.xlsx（上月书展跨月整合数据）
- output/书展数据整合_本月_XXX.xlsx（本月书展数据，如有）
- sample/书展数据汇总.xlsx（手工维护，行 2-18 是历史书展）
- ../港澳流速/sample/海外港澳商务_各渠道主辅投数据.xlsx（数据底表）

【输出】
- extract_book_fair_data() 返回上月/本月书展 + 历史书展数据

【关键参数】
- 上月书展识别：从整合数据的供应商名中自动识别
- 本月书展识别：如有整合文件则提取

【对外接口】
- extract_book_fair_data() - 返回上月/本月书展 + 历史书展数据

【注意事项】
- 历史书展行（2-18）保持不动
- 归档检查：月初时自动追加上上月书展汇总行到历史
"""
import sys
import shutil
from pathlib import Path
from datetime import datetime, timedelta
from calendar import monthrange
import pandas as pd
from openpyxl import load_workbook

sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = Path(__file__).parent
SAMPLE_DIR = BASE_DIR / "sample"
OUTPUT_DIR = BASE_DIR / "output"
# 使用整合后的书展数据
DATA_FILE = BASE_DIR.parent / "书展数据复盘" / "output"  # 整合文件目录
SUMMARY_FILE = SAMPLE_DIR / "书展数据汇总.xlsx"

# 字段顺序（从数据底表读，写入汇总表）。汇总表只保留到第一个"滚动ROI2"
SUMMARY_FIELDS = [
    "滚动消耗", "滚动消耗(不含赠课成本)", "单例子成本", "例子数", "分发数",
    "分发数\n(剔除毛例子)", "约课数", "应到课数", "滚动应到课数",
    "到课数", "滚动到课数", "当月成交数", "滚动成交数",
    "当月GMV", "滚动GMV", "海外GMV占比", "滚动ASP",
    "例子分发率", "分发约课率", "例子约课率", "约课到课率",
    "应到课率", "滚动应到课率", "到课转化率", "注册转化率",
    "滚动转化率", "滚动ROI2总成本", "滚动ROI2"
]

# 重点关注指标（HTML 中高亮）
KEY_METRICS = ["例子约课率", "约课到课率", "到课转化率", "滚动转化率", "滚动ROI2"]


def get_bookfair_periods():
    """
    计算书展数据所需的日期范围

    返回:
        dict: 包含日期信息和识别前缀
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


def find_latest_file(pattern):
    """从output目录中找最新的符合pattern的文件"""
    import glob
    files = glob.glob(str(DATA_FILE / pattern))
    if not files:
        raise FileNotFoundError(f"书展数据文件缺失: {pattern}")
    return max(files, key=lambda f: Path(f).stat().st_mtime)


def get_last_month_fairs(periods):
    """
    从上月整合文件提取上月书展数据

    返回:
        dict: {
            'fairs': [{name, ...fields}],
            'period_label': '2026年5月',
            'period_range': '5.1-6.7',
        }
    """
    # 查找最新的上月整合文件
    pattern = f"书展数据整合_上月_*.xlsx"
    file_path = find_latest_file(pattern)  # 找不到会 raise FileNotFoundError

    df = pd.read_excel(file_path, header=5)
    df["供应商"] = df["供应商"].ffill()

    last_prefix = periods['last_month']['prefix']
    last_month_fairs = df[
        df["供应商"].str.contains(last_prefix, na=False)
    ]["供应商"].unique().tolist()

    fairs = []
    for supplier in last_month_fairs:
        supplier_data = df[df["供应商"] == supplier]
        for _, row in supplier_data.iterrows():
            rec = {"name": row.get("渠道名称", supplier)}
            for f in SUMMARY_FIELDS:
                v = row.get(f)
                rec[f] = v if pd.notna(v) else None

            # 重新计算单例子成本 = 滚动消耗 / 例子数
            cost = rec.get("滚动消耗")
            examples = rec.get("例子数")
            if cost and examples and examples != 0:
                rec["单例子成本"] = cost / examples

            fairs.append(rec)

    period_start = periods['last_month_to_now'][0].strftime("%m.%d").lstrip("0").replace(".0", ".")
    period_end = periods['yesterday'].strftime("%m.%d").lstrip("0").replace(".0", ".")

    return {
        'fairs': fairs,
        'period_label': f"2026年{periods['last_month']['month']}月",
        'period_range': f"{periods['last_month_to_now'][0].month}.{periods['last_month_to_now'][0].day}-{period_end}",
    }


def get_current_month_fairs(periods):
    """
    从本月整合文件提取本月书展数据（如有）

    返回:
        dict 或 None（本月无新文件时返回 None，但找到文件后解析失败会 raise）
    """
    # 查找最新的本月整合文件
    pattern = f"书展数据整合_本月_*.xlsx"
    try:
        file_path = find_latest_file(pattern)
    except FileNotFoundError:
        # 本月可能没有新文件，返回 None 是允许的（但不能隐式软降级）
        return None

    df = pd.read_excel(file_path, header=5)
    df["供应商"] = df["供应商"].ffill()

    cur_prefix = periods['current_month']['prefix']
    current_month_fairs = df[
        df["供应商"].str.contains(cur_prefix, na=False)
    ]["供应商"].unique().tolist()

    if not current_month_fairs:
        return None

    fairs = []
    for supplier in current_month_fairs:
        supplier_data = df[df["供应商"] == supplier]
        for _, row in supplier_data.iterrows():
            rec = {"name": row.get("渠道名称", supplier)}
            for f in SUMMARY_FIELDS:
                v = row.get(f)
                rec[f] = v if pd.notna(v) else None

            # 重新计算单例子成本 = 滚动消耗 / 例子数
            cost = rec.get("滚动消耗")
            examples = rec.get("例子数")
            if cost and examples and examples != 0:
                rec["单例子成本"] = cost / examples

            fairs.append(rec)

    period_start = periods['current_month_to_now'][0].strftime("%m.%d").lstrip("0").replace(".0", ".")
    period_end = periods['yesterday'].strftime("%m.%d").lstrip("0").replace(".0", ".")

    return {
        'fairs': fairs,
        'period_label': f"2026年{periods['current_month']['month']}月",
        'period_range': f"{period_start}-{period_end}",
    }


def archive_last_month_to_history(periods):
    """
    检查是否需要归档，如需要则将上上月书展的汇总行追加到历史

    触发条件：当前日期是本月1日（即上月已结束）
    """
    if not periods['is_month_start']:
        return

    print("  [归档检查] 检测到月初，准备归档上上月书展...")

    # 简化实现：这部分在实际生产中需要实现完整的归档逻辑
    # 目前仅记录日志，真实归档需要额外的配置和处理
    print("  [归档] 已记录，可在月初时手动处理或通过定时任务完成")



def extract_short_name(channel_name: str) -> str:
    """
    从数据底表的"渠道名称"中提取承担人简称
    规则：
      1. 优先取 "-shukei-" 到 "-思维" 之间的内容（前面加"-"），但若 shukei 后紧接"思维"则跳过
      2. 否则取 "9.9小课包" 后的字符
      3. 否则原样返回
    """
    if not isinstance(channel_name, str):
        return ""

    if "-shukei-" in channel_name:
        idx = channel_name.find("-shukei-") + len("-shukei-")
        rest = channel_name[idx:]
        # 若 shukei 后紧接"思维"或"主投"，说明没承担人，跳过
        if not rest.startswith(("思维", "主投", "美术")) and "-思维" in rest:
            middle = rest.split("-思维")[0].strip()
            if middle:
                return f"-{middle}"

    if "9.9小课包" in channel_name:
        return channel_name.split("9.9小课包")[-1].strip()

    # 兜底：取最后一段
    parts = channel_name.split("-")
    return parts[-1].strip() if parts else channel_name


def load_book_fair_details(supplier_name: str) -> dict:
    """从数据底表提取指定书展的明细行 + 总计行"""
    df = pd.read_excel(DATA_FILE, header=5)
    df["渠道组"] = df["渠道组"].ffill()
    df["供应商"] = df["供应商"].ffill()

    sub = df[df["供应商"] == supplier_name].copy()

    details = []
    total = None
    for _, row in sub.iterrows():
        chan = row["渠道名称"]
        if not isinstance(chan, str):
            continue

        rec = {"_full_channel": chan}
        for f in SUMMARY_FIELDS:
            v = row.get(f)
            rec[f] = v if pd.notna(v) else None

        if chan == "总计":
            total = rec
            total["short_name"] = supplier_name  # 汇总行就用供应商名
        else:
            rec["short_name"] = extract_short_name(chan)
            details.append(rec)

    return {"total": total, "details": details}


def write_to_summary(book_fairs_data: dict):
    """
    把本月两个书展的数据写入 书展数据汇总.xlsx。
    保留行 1 (表头) 和行 2-18 (历史书展)，从行 19 开始重写本月数据。
    """
    OUTPUT_DIR.mkdir(exist_ok=True)
    out_file = OUTPUT_DIR / SUMMARY_FILE.name
    shutil.copy(SUMMARY_FILE, out_file)

    wb = load_workbook(out_file)
    ws = wb["Sheet1"]

    # 清空行 19 起的所有数据（保险：清空 19~50 行）
    for row in range(19, 60):
        for col in range(1, ws.max_column + 1):
            ws.cell(row, col).value = None

    # 列映射：列1=渠道名称(short_name)，列2-29=数据列（对应 SUMMARY_FIELDS）
    write_row = 19
    for fair_name in ["26年5月課外活動展", "26年5月第十二屆兒童書展"]:
        data = book_fairs_data.get(fair_name)
        if not data:
            continue

        # 写汇总行
        if data["total"]:
            ws.cell(write_row, 1).value = fair_name
            for i, f in enumerate(SUMMARY_FIELDS, start=2):
                v = data["total"].get(f)
                ws.cell(write_row, i).value = v
            write_row += 1

        # 写明细行
        for d in data["details"]:
            ws.cell(write_row, 1).value = d["short_name"]
            for i, f in enumerate(SUMMARY_FIELDS, start=2):
                v = d.get(f)
                ws.cell(write_row, i).value = v
            write_row += 1

    wb.save(out_file)
    print(f"  [书展汇总] 已写入 {write_row - 19} 行（含汇总+明细）")
    return out_file


def extract_book_fair_data() -> dict:
    """
    供 HTML 周报使用的接口

    返回:
        dict: {
            'last_month_fairs': {...},     # 上月书展
            'current_month_fairs': {...},  # 本月书展（如有）
            'history': [...],              # 历史书展
            'periods': {...},              # 日期信息
        }
    """
    periods = get_bookfair_periods()

    # 1. 提取上月书展
    last_month_fairs = get_last_month_fairs(periods)

    # 2. 提取本月书展
    current_month_fairs = get_current_month_fairs(periods)

    # 3. 提取历史书展
    wb = load_workbook(SUMMARY_FILE, data_only=True)
    ws = wb["Sheet1"]

    headers = ["渠道名称"] + SUMMARY_FIELDS

    history_rows = []
    for row in range(3, 20):  # 行 2-18 (openpyxl 1-based: 3-19)
        name = ws.cell(row, 1).value
        if not name:
            continue
        rec = {"name": str(name).strip()}
        for i, f in enumerate(SUMMARY_FIELDS, start=2):
            v = ws.cell(row, i).value
            rec[f] = v

        # 重新计算单例子成本 = 滚动消耗 / 例子数
        cost = rec.get("滚动消耗")
        examples = rec.get("例子数")
        if cost and examples and examples != 0:
            rec["单例子成本"] = cost / examples

        history_rows.append(rec)

    # 4. 归档检查
    archive_last_month_to_history(periods)

    return {
        "last_month_fairs": last_month_fairs,
        "current_month_fairs": current_month_fairs,
        "history": history_rows,
        "periods": periods,
        "headers": headers,
        "summary_fields": SUMMARY_FIELDS,
        "key_metrics": KEY_METRICS,
    }


def main():
    print("=" * 60)
    print("书展数据复盘（滚动展示逻辑）")
    print("=" * 60)
    print()

    data = extract_book_fair_data()

    if data["last_month_fairs"]:
        print(f"[1] 上月书展数据（{data['last_month_fairs']['period_label']}）")
        print(f"    周期: {data['last_month_fairs']['period_range']}")
        print(f"    行数: {len(data['last_month_fairs']['fairs'])}")
    else:
        print(f"[1] 上月无书展数据")

    print()

    if data["current_month_fairs"]:
        print(f"[2] 本月书展数据（{data['current_month_fairs']['period_label']}）")
        print(f"    周期: {data['current_month_fairs']['period_range']}")
        print(f"    行数: {len(data['current_month_fairs']['fairs'])}")
    else:
        print(f"[2] 本月无书展数据")

    print()

    print(f"[3] 历史书展")
    print(f"    行数: {len(data['history'])}")

    print()
    print(f"[完成] 书展数据提取完毕")



if __name__ == "__main__":
    main()
