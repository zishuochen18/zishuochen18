"""
线下商超数据复盘处理 v2

完整逻辑：
1. 从数据底表提取线下HK商超数据（排除带"展"字眼的供应商）
2. 写入【当月数据整理汇总】sheet
3. 汇总当月合计数据到【分月数据汇总】sheet
4. 维护【分场次数据汇总】sheet（含承担人明细）
5. 识别重复场次做数据对比
6. 提供 extract_shangchao_data() 接口给 HTML 使用
"""
"""
线下商超复盘处理脚本（模块 5）

═══════════════════════════════════════════════════════════════
功能：从数据底表提取线下HK商超数据，统计本月分场次和分月趋势
═══════════════════════════════════════════════════════════════

【输入】
- sample/线下商超内容汇总.xlsx（手工维护，含分月汇总、分场次、当月明细 3 个 sheet）
- ../港澳流速/sample/海外港澳商务_各渠道主辅投数据.xlsx（数据底表）

【输出】
- output/线下商超内容汇总.xlsx

【关键参数】
- VENUE_DAYS_MAP：当月新场次的天数配置（每月维护）
  例：{"啟田商場3場": 4, "樂富商場市集": 4}
- EXCLUDE_KEYWORDS：从线下HK排除的展会关键词（"展", "STEM"）

【数据流】
1. 从数据底表提取线下HK渠道组数据
2. 排除带"展/STEM"的展会供应商（书展归到模块 4）
3. 按例子数 > 0 筛选有效场次
4. 按场次（供应商）汇总，含承担人明细
5. 识别重复场次（同一商场多次出摊，按基础名识别）
6. 计算分月趋势的当月行（所有非展会线下HK供应商总计）

【关键计算公式】
- CPS成本 = 滚动GMV × 0.35
- CPT成本 = 滚动消耗 - CPS成本
- 天均 = 例子数 / 出摊天数

【两个口径区分】
- 5.2 分月趋势：所有非展会线下HK供应商的总和（含例子=0但有滚动消耗的场次）
- 5.3 分场次明细：仅例子数 > 0 的有效场次（实际有客流的场次）

【对外接口】
- extract_shangchao_data() - 返回完整商超数据

【注意事项】
- VENUE_DAYS_MAP 配置脚本会自动检测新场次缺失天数并提醒
- 重复场次基础名识别规则：去除月份前缀、"X场/場"后缀、"商場/廣場"后缀
"""
import sys
import re
import shutil
from pathlib import Path
from copy import copy
from datetime import datetime, timedelta
from calendar import monthrange

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = Path(__file__).parent
SAMPLE_DIR = BASE_DIR / "sample"
OUTPUT_DIR = BASE_DIR / "output"
DATA_FILE = BASE_DIR.parent / "港澳流速" / "sample" / "海外港澳商务_各渠道主辅投数据 (2).xlsx"
SUMMARY_FILE = SAMPLE_DIR / "线下商超内容汇总.xlsx"

# 排除的展会关键词
EXCLUDE_KEYWORDS = ["展", "STEM", "商超宣傳單"]

# 重点关注指标
KEY_METRICS = ["例子约课率", "约课到课率", "到课转化率", "滚动转化率", "滚动ROI2"]

# 当月数据整理汇总的列映射（数据底表列名 → 写入顺序）
DATA_FIELDS = [
    "滚动消耗", "滚动消耗(不含赠课成本)", "单例子成本", "例子数",
    "分发数", "分发数\n(剔除毛例子)", "约课数", "应到课数",
    "滚动应到课数", "到课数", "滚动到课数", "当月成交数",
    "滚动成交数", "当月GMV", "滚动GMV", "海外GMV占比",
    "滚动ASP", "例子分发率", "分发约课率", "例子约课率",
    "约课到课率", "应到课率", "滚动应到课率", "到课转化率",
    "注册转化率", "滚动转化率", "滚动ROI2总成本", "滚动ROI2"
]

# 当月场次天数配置（每月维护，键=供应商名，值=出摊天数）
# 优先级：分场次sheet中已有的天数 > VENUE_DAYS_MAP 中的配置
# 当月新增场次（在分场次sheet里没有的），从这里取天数
VENUE_DAYS_MAP = {
    # 5月场次（用于跨月累计展示）
    "马鞍山WEGOMALL市集": 3,
    "秀茂坪2场": 3,
    "興華商場": 2,
    "栢麗商場3場": 3,
    "啟田商場3場": 4,
    "樂富商場市集": 4,
    # 6月场次（原有配置保持）
    "馬鞍山市集2場": 3,
    "東港城會所": 2,
    "石籬商場2場": 2,
}

# 书展模块的 sample 目录（复用跨月快照）
BOOKFAIR_SAMPLE_DIR = BASE_DIR.parent / "书展数据复盘" / "sample"


def _last_month_full_file(periods):
    """5.1-5.31 完整月快照路径"""
    return BOOKFAIR_SAMPLE_DIR / "海外港澳商务_各渠道主辅投数据_5.1-5.31.xlsx"


def _last_month_to_now_file(periods):
    """5.1-昨天 跨月快照路径（使用最新的匹配文件）"""
    import glob
    pattern = str(BOOKFAIR_SAMPLE_DIR / "海外港澳商务_各渠道主辅投数据_5.1-*.xlsx")
    files = glob.glob(pattern)
    if not files:
        return None
    # 返回最新的文件（排除 5.1-5.31 完整月）
    candidates = [f for f in files if "5.1-5.31" not in f]
    if candidates:
        return Path(max(candidates, key=lambda f: Path(f).stat().st_mtime))
    return None


# 分月数据汇总的列（从列0开始）
MONTHLY_FIELDS = [
    "月份", "天数", "天均", "CPS消耗", "CPT消耗", "滚动消耗",
    "滚动消耗(不含赠课成本)", "单例子成本", "例子数", "分发数",
    "剔除毛例子分发数", "约课数", "应到课数", "滚动应到课数",
    "到课数", "滚动到课数", "当月成交数", "滚动成交数",
    "当月GMV", "滚动GMV", "海外GMV占比", "滚动ASP",
    "例子分发率", "分发约课率", "例子约课率", "约课到课率",
    "应到课率", "滚动应到课率", "到课转化率", "注册转化率",
    "滚动转化率", "滚动ROI2总成本", "滚动ROI2"
]

def extract_shangchao_from_source():
    """
    步骤1：从数据底表提取线下HK商超数据（排除带"展"字眼的供应商）
    返回 DataFrame（含渠道明细 + 总计行）
    """
    df = pd.read_excel(DATA_FILE, header=5)
    df["渠道组"] = df["渠道组"].ffill()
    df["供应商"] = df["供应商"].ffill()

    # 只保留线下HK
    hk = df[df["渠道组"].astype(str).str.contains("线下HK", na=False)].copy()

    # 排除带"展"/"STEM"的供应商
    hk = hk[~hk["供应商"].astype(str).apply(
        lambda x: any(k in x for k in EXCLUDE_KEYWORDS)
    )]
    # 排除渠道组级别的"总计"行
    hk = hk[hk["供应商"] != "总计"]

    return hk


def extract_short_name(channel_name: str) -> str:
    """从渠道名称提取简称（取"小课包"后面的名字）"""
    if not isinstance(channel_name, str):
        return ""
    if channel_name == "总计":
        return "总计"

    # 取 "9.9小课包" 后面的字符
    if "小课包" in channel_name:
        parts = channel_name.split("小课包")
        tail = parts[-1].strip()
        if tail:
            return tail

    # 兜底：取 shukei 后到 思维 之间
    if "-shukei-" in channel_name:
        idx = channel_name.find("-shukei-") + len("-shukei-")
        rest = channel_name[idx:]
        if not rest.startswith(("思维", "主投", "美术")) and "-思维" in rest:
            middle = rest.split("-思维")[0].strip()
            if middle:
                return middle

    # 最后兜底
    parts = channel_name.split("-")
    return parts[-1].strip() if parts else channel_name


def get_shangchao_periods():
    """
    计算商超数据所需的日期范围（复用书展模块的 get_bookfair_periods 逻辑）

    返回:
        dict: {
            'today': date,
            'yesterday': date,
            'last_month': {'year': int, 'month': int, 'year_short': str, 'prefix': str},
            'current_month': {'year': int, 'month': int, 'year_short': str, 'prefix': str},
            'last_month_full': (start_date, end_date),
            'last_month_to_now': (start_date, yesterday),
            'current_month_to_now': (start_date, yesterday),
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


def get_venue_base_name(venue_name: str) -> str:
    """
    提取场次的基础名称（用于识别重复场次）
    例如：'秀茂坪商場' / '秀茂坪2场' → '秀茂坪'
         '啟田商場' / '啟田商場2場' / '啟田商場3場' → '啟田商場'
         '栢麗廣場2场' / '栢麗商場3場' → '栢麗'
    """
    name = venue_name.strip()
    # 去掉月份前缀（如 "5月"、"12月"）
    name = re.sub(r'^\d+月', '', name)
    # 去掉数字+场/場后缀
    name = re.sub(r'\d+[场場]$', '', name)
    # 去掉"商場"/"商场"/"廣場"/"广场" 后缀
    name = re.sub(r'(商場|商场|廣場|广场)$', '', name)
    return name.strip()


def identify_last_month_venues(periods):
    """
    从上月完整月底表中识别 5 月场次（线下HK ∧ 非展会 ∧ 例子数 > 0）
    返回供应商名列表
    """
    file_path = _last_month_full_file(periods)
    if not file_path or not file_path.exists():
        raise FileNotFoundError(f"上月完整月底表缺失: {file_path}")

    df = pd.read_excel(file_path, header=5)
    df["渠道组"] = df["渠道组"].ffill()
    df["供应商"] = df["供应商"].ffill()

    hk = df[df["渠道组"].astype(str).str.contains("线下HK", na=False)]
    hk = hk[~hk["供应商"].astype(str).apply(lambda x: any(k in x for k in EXCLUDE_KEYWORDS))]
    hk = hk[hk["供应商"] != "总计"]

    venues = []
    for supplier in hk["供应商"].unique():
        sub = hk[(hk["供应商"] == supplier) & (hk["渠道名称"] == "总计")]
        if not sub.empty:
            ex = sub.iloc[0].get("例子数")
            if isinstance(ex, (int, float)) and ex > 0:
                venues.append(supplier)

    return venues


def identify_current_month_venues(periods):
    """
    从本月底表中识别 6 月场次（线下HK ∧ 非展会 ∧ 例子数 > 0）
    返回供应商名列表
    """
    df = pd.read_excel(DATA_FILE, header=5)
    df["渠道组"] = df["渠道组"].ffill()
    df["供应商"] = df["供应商"].ffill()

    hk = df[df["渠道组"].astype(str).str.contains("线下HK", na=False)]
    hk = hk[~hk["供应商"].astype(str).apply(lambda x: any(k in x for k in EXCLUDE_KEYWORDS))]
    hk = hk[hk["供应商"] != "总计"]

    venues = []
    for supplier in hk["供应商"].unique():
        sub = hk[(hk["供应商"] == supplier) & (hk["渠道名称"] == "总计")]
        if not sub.empty:
            ex = sub.iloc[0].get("例子数")
            if isinstance(ex, (int, float)) and ex > 0:
                venues.append(supplier)

    return venues


def write_current_month_sheet(hk_df, out_wb):
    """
    步骤2：把当月商超数据写入【当月数据整理汇总】sheet
    只写每个供应商的总计行（不写明细渠道）
    """
    ws = out_wb["当月数据整理汇总"]

    # 清空数据区（保留行1-2表头，从行3开始写）
    max_clear = 100

    # 先找出"合计"行号
    total_row_num = None
    for r in range(3, max_clear):
        if ws.cell(r, 1).value and "合计" in str(ws.cell(r, 1).value):
            total_row_num = r
            break

    # 清空时跳过整个合计行
    for row in range(3, max_clear):
        if row == total_row_num:
            continue   # 整行不动（保留所有 SUM/比率公式）
        if row >= max_clear - 5:
            continue
        for col in range(1, 36):
            ws.cell(row, col).value = None

    # 按供应商分组写入：只写总计行
    write_row = 3
    suppliers = [s for s in hk_df["供应商"].unique() if s != "总计"]

    for supplier in suppliers:
        sub = hk_df[hk_df["供应商"] == supplier]
        # 找该供应商的总计行
        total_row = sub[sub["渠道名称"] == "总计"]
        if total_row.empty:
            continue
        row = total_row.iloc[0]

        # 写渠道组、供应商、渠道名称（这里渠道名称 = 总计）
        ws.cell(write_row, 2).value = "线下HK"
        ws.cell(write_row, 3).value = supplier
        ws.cell(write_row, 4).value = "总计"

        # 写数据列（从列8开始，对应 DATA_FIELDS）
        col_offset = 8
        for i, field in enumerate(DATA_FIELDS):
            v = row.get(field)
            if pd.notna(v):
                ws.cell(write_row, col_offset + i).value = v
        write_row += 1

    print(f"  [当月数据整理汇总] 写入 {write_row - 3} 行供应商总计")
    return write_row


def update_monthly_summary(hk_df, out_wb, current_month_name="6月"):
    """
    步骤3：把当月合计数据写入【分月数据汇总】对应月份行
    直接从 hk_df（线下HK非展会供应商总计行）聚合计算，不依赖合计行公式
    """
    import re

    ws_monthly = out_wb["分月数据汇总"]

    # 从 hk_df 直接聚合（取各供应商的总计行求和）
    all_totals = hk_df[hk_df["渠道名称"] == "总计"]

    sum_fields = [
        "滚动消耗", "滚动消耗(不含赠课成本)", "例子数", "分发数",
        "分发数\n(剔除毛例子)", "约课数", "应到课数", "滚动应到课数",
        "到课数", "滚动到课数", "当月成交数", "滚动成交数",
        "当月GMV", "滚动GMV", "滚动ROI2总成本"
    ]
    agg = {}
    for f in sum_fields:
        s = 0.0
        for _, row in all_totals.iterrows():
            v = row.get(f)
            if isinstance(v, (int, float)) and pd.notna(v):
                s += v
        agg[f] = s

    # 重算衍生指标
    ex = agg["例子数"]
    agg["单例子成本"] = round(agg["滚动消耗(不含赠课成本)"] / ex, 2) if ex > 0 else None
    agg["例子约课率"] = agg["约课数"] / ex if ex > 0 else None
    agg["例子分发率"] = agg["分发数"] / ex if ex > 0 else None
    agg["分发约课率"] = agg["约课数"] / agg["分发数"] if agg["分发数"] > 0 else None
    agg["约课到课率"] = agg["滚动到课数"] / agg["约课数"] if agg["约课数"] > 0 else None
    agg["到课转化率"] = agg["当月成交数"] / agg["到课数"] if agg["到课数"] > 0 else None
    agg["滚动转化率"] = agg["滚动成交数"] / agg["滚动到课数"] if agg["滚动到课数"] > 0 else None
    agg["滚动ASP"] = agg["滚动GMV"] / agg["滚动成交数"] if agg["滚动成交数"] > 0 else None
    agg["滚动ROI2"] = agg["滚动GMV"] / agg["滚动ROI2总成本"] if agg["滚动ROI2总成本"] > 0 else None

    # 计算天均（数据底表中无天数字段，此处不计算）
    # 天均仅在分场次数据汇总中维护
    agg["天数"] = None
    agg["天均"] = None

    # 计算注册转化率
    agg["注册转化率"] = agg["当月成交数"] / ex if ex > 0 else None

    # CPS/CPT
    cps, cpt = calc_cps_cpt(agg["滚动GMV"], agg["滚动消耗"])
    agg["CPS成本"] = cps
    agg["CPT成本"] = cpt

    # 找分月数据汇总中对应月份的行
    target_row = None
    summary_row = None
    last_month_row = 2

    # 删除重复的当月行（防御性修复）
    duplicate_rows = []
    for row in range(3, ws_monthly.max_row + 1):
        v = ws_monthly.cell(row, 1).value
        if not v:
            continue
        v_str = str(v).strip()
        if v_str == current_month_name:
            if target_row:  # 如果已经找到一个，后续的都是重复
                duplicate_rows.append(row)
            else:
                target_row = row
        elif "汇总" in v_str or "合计" in v_str:
            summary_row = row
        else:
            last_month_row = row

    # 从后往前删除重复行（避免行号变化影响）
    for row in sorted(duplicate_rows, reverse=True):
        ws_monthly.delete_rows(row, 1)
        print(f"  [分月数据汇总] 删除重复的 {current_month_name} 行（行 {row}）")

    # 没找到月份行，在汇总行之前插入新行
    if not target_row:
        if summary_row:
            ws_monthly.insert_rows(summary_row)
            target_row = summary_row
            new_summary_row = summary_row + 1

            # 扩展汇总行所有 SUM 公式范围 +1
            SUM_RANGE_RE = re.compile(r"SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)", re.IGNORECASE)
            for col in range(1, ws_monthly.max_column + 1):
                cell = ws_monthly.cell(new_summary_row, col)
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    def extend_range(m):
                        c1, r1, c2, r2 = m.group(1), m.group(2), m.group(3), m.group(4)
                        return f"SUM({c1}{r1}:{c2}{int(r2)+1})"
                    cell.value = SUM_RANGE_RE.sub(extend_range, cell.value)
        else:
            target_row = last_month_row + 1
        ws_monthly.cell(target_row, 1).value = current_month_name
        print(f"  [分月数据汇总] 未找到 {current_month_name} 行，自动新增到行 {target_row}")

    # 分月表列映射（列号 → 聚合字段名）
    monthly_col_map = {
        4: "CPS成本",
        5: "CPT成本",
        7: "滚动消耗(不含赠课成本)",
        8: "单例子成本",
        9: "例子数",
        10: "分发数",
        11: "分发数\n(剔除毛例子)",
        12: "约课数",
        13: "应到课数",
        14: "滚动应到课数",
        15: "到课数",
        16: "滚动到课数",
        17: "当月成交数",
        18: "滚动成交数",
        19: "当月GMV",
        20: "滚动GMV",
        22: "滚动ASP",
        23: "例子分发率",
        24: "分发约课率",
        25: "例子约课率",
        26: "约课到课率",
        29: "到课转化率",
        30: "注册转化率",
        31: "滚动转化率",
        32: "滚动ROI2总成本",
        33: "滚动ROI2",
    }

    for col, field in monthly_col_map.items():
        v = agg.get(field)
        if v is not None:
            ws_monthly.cell(target_row, col).value = v

    # 滚动消耗 = CPS + CPT（写到列6）
    ws_monthly.cell(target_row, 6).value = agg["滚动消耗"]

    # 为新增的26年6月行设置公式（包括天数/天均/CPS/CPT/注册转化率）
    if current_month_name.startswith("26年"):
        # 只加总本月有效场次（例子数>0）的天数
        total_days = 0
        valid_suppliers = [s for s in hk_df["供应商"].unique() if s != "总计"]
        for supplier_name in valid_suppliers:
            sub = hk_df[hk_df["供应商"] == supplier_name]
            total_row = sub[sub["渠道名称"] == "总计"]
            if not total_row.empty:
                ex = total_row.iloc[0].get("例子数")
                if isinstance(ex, (int, float)) and ex > 0:
                    # 该场次有效，累加其天数
                    total_days += VENUE_DAYS_MAP.get(supplier_name, 0)

        # 写入天数（数值）
        if total_days > 0:
            ws_monthly.cell(target_row, 2).value = total_days

        # 设置衍生字段公式
        ws_monthly.cell(target_row, 3).value = f"=I{target_row}/B{target_row}"   # 天均
        ws_monthly.cell(target_row, 4).value = f"=T{target_row}*0.35"             # CPS成本
        ws_monthly.cell(target_row, 5).value = f"=F{target_row}-D{target_row}"    # CPT成本
        ws_monthly.cell(target_row, 30).value = f"=Q{target_row}/I{target_row}"   # 注册转化率

    # 特殊处理：5月行天数从 VENUE_DAYS_MAP 的 5 月场次加和
    if "5月" in str(current_month_name):
        may_venues = ["马鞍山WEGOMALL市集", "秀茂坪2场", "興華商場", "栢麗商場3場", "啟田商場3場", "樂富商場市集"]
        total_days = sum(VENUE_DAYS_MAP.get(v, 0) for v in may_venues)
        if total_days > 0:
            ws_monthly.cell(target_row, 2).value = total_days
            print(f"  [分月数据汇总] {current_month_name} 天数从场次加和: {total_days} 天")
    summary_row = None
    for row in range(3, ws_monthly.max_row + 1):
        v = ws_monthly.cell(row, 1).value
        if v and ("汇总" in str(v) or "合计" in str(v)):
            summary_row = row
            break

    if summary_row:
        sr = summary_row
        end = sr - 1  # 数据范围最后一行（汇总行上一行）

        # 基础 SUM 字段（直接累加）
        sum_cols = {
            2:  "B",   # 天数
            6:  "F",   # 滚动消耗
            7:  "G",   # 滚动消耗(不含赠课成本)
            9:  "I",   # 例子数
            10: "J",   # 分发数
            11: "K",   # 剔除毛例子分发数
            12: "L",   # 约课数
            13: "M",   # 应到课数
            14: "N",   # 滚动应到课数
            15: "O",   # 到课数
            16: "P",   # 滚动到课数
            17: "Q",   # 当月成交数
            18: "R",   # 滚动成交数
            19: "S",   # 当月GMV
            20: "T",   # 滚动GMV
            32: "AF",  # 滚动ROI2总成本
        }
        for col_idx, col_letter in sum_cols.items():
            ws_monthly.cell(sr, col_idx).value = f"=SUM({col_letter}3:{col_letter}{end})"

        # 衍生字段（基于汇总行自身分子分母重算）
        derived_formulas = {
            3:  f"=I{sr}/B{sr}",        # 天均 = 例子/天数
            4:  f"=T{sr}*0.35",         # CPS = GMV*0.35
            5:  f"=F{sr}-D{sr}",        # CPT = 总消耗-CPS
            8:  f"=G{sr}/I{sr}",        # 单例子成本 = 滚动消耗(不含赠课)/例子数
            22: f"=T{sr}/R{sr}",        # 滚动ASP = 滚动GMV/滚动成交
            23: f"=J{sr}/I{sr}",        # 例子分发率 = 分发/例子
            24: f"=L{sr}/J{sr}",        # 分发约课率 = 约课/分发
            25: f"=L{sr}/I{sr}",        # 例子约课率 = 约课/例子
            26: f"=P{sr}/L{sr}",        # 约课到课率 = 滚动到课/约课
            29: f"=Q{sr}/O{sr}",        # 到课转化率 = 当月成交/到课
            30: f"=Q{sr}/I{sr}",        # 注册转化率 = 当月成交/例子
            31: f"=R{sr}/P{sr}",        # 滚动转化率 = 滚动成交/滚动到课
            33: f"=T{sr}/AF{sr}",       # 滚动ROI2 = 滚动GMV/滚动ROI2总成本
        }
        for col_idx, formula in derived_formulas.items():
            ws_monthly.cell(sr, col_idx).value = formula

        print(f"  [分月数据汇总] 已刷新汇总行（行 {sr}），数据范围 row 3 ~ row {end}")

    # 更新上月（5月）行的天数（从 VENUE_DAYS_MAP 的 5 月场次加和）
    may_row = None
    for row in range(3, ws_monthly.max_row + 1):
        v = ws_monthly.cell(row, 1).value
        if v and str(v).strip() in ["5月", "26年5月"]:
            may_row = row
            break

    if may_row:
        may_venues = ["马鞍山WEGOMALL市集", "秀茂坪2场", "興華商場", "栢麗商場3場", "啟田商場3場", "樂富商場市集"]
        total_days = sum(VENUE_DAYS_MAP.get(v, 0) for v in may_venues)
        if total_days > 0:
            ws_monthly.cell(may_row, 2).value = total_days
            print(f"  [分月数据汇总] 5月 行天数更新: {total_days} 天（从场次加和）")

    print(f"  [分月数据汇总] 已更新 {current_month_name} 行（例子={int(ex)}, "
          f"约课={int(agg['约课数'])}, CPS={cps:.0f}, CPT={cpt:.0f}）")


def build_venue_data(hk_df):
    """
    步骤4：构建分场次数据（按供应商分组 + 承担人明细）
    只保留例子数>0的场次（即本月有效场次）
    返回 [{total: {...}, details: [{...}]}]
    """
    venues = []
    suppliers = [s for s in hk_df["供应商"].unique() if s != "总计"]

    for supplier in suppliers:
        sub = hk_df[hk_df["供应商"] == supplier]
        total_row = sub[sub["渠道名称"] == "总计"]
        detail_rows = sub[sub["渠道名称"] != "总计"]

        if total_row.empty:
            continue

        t = total_row.iloc[0]
        # 只保留例子数>0的场次
        ex = t.get("例子数")
        if not isinstance(ex, (int, float)) or ex <= 0:
            continue

        venue = {
            "total": {
                "name": supplier,
                **{f: (t[f] if pd.notna(t.get(f)) else None) for f in DATA_FIELDS}
            },
            "details": []
        }

        for _, row in detail_rows.iterrows():
            chan = row.get("渠道名称")
            if not isinstance(chan, str):
                continue
            short = extract_short_name(chan)
            rec = {
                "name": short,
                **{f: (row[f] if pd.notna(row.get(f)) else None) for f in DATA_FIELDS}
            }
            venue["details"].append(rec)

        venues.append(venue)

    return venues


def _load_supplier_from_file(file_path, supplier):
    """从指定快照读取线下HK供应商数据（总计行 + 明细行）"""
    if not file_path.exists():
        raise FileNotFoundError(f"商超底表缺失: {file_path}")

    df = pd.read_excel(file_path, header=5)
    df["渠道组"] = df["渠道组"].ffill()
    df["供应商"] = df["供应商"].ffill()
    df = df[df["渠道组"].astype(str).str.contains("线下HK", na=False)]
    df = df[~df["供应商"].astype(str).apply(lambda x: any(k in x for k in EXCLUDE_KEYWORDS))]

    # 按供应商精确匹配
    sub = df[df["供应商"] == supplier]
    if sub.empty:
        return {"total": None, "details": []}

    total_row = sub[sub["渠道名称"] == "总计"]
    detail_rows = sub[sub["渠道名称"] != "总计"]

    return {
        "total": total_row.iloc[0].to_dict() if not total_row.empty else None,
        "details": [r.to_dict() for _, r in detail_rows.iterrows()],
    }


def merge_last_month_venue(supplier, periods):
    """
    跨月整合上月场次数据（5.1-5.31 + 6.1-至今）

    规则：
    - 滚动消耗 / 滚动ROI2总成本：5月完整 + 6月至今 加和
    - GMV / 到课 / 例子数等滚动指标：直接用 5.1-至今 快照
    - 滚动ROI2：重算 = 5.1-至今.GMV / 整合后.ROI2总成本
    - 单例子成本：重算 = 整合后.滚动消耗 / 5.1-至今.例子数

    返回：{'total': {...}, 'details': [...]}
    """
    may_full = _load_supplier_from_file(_last_month_full_file(periods), supplier)
    may_to_now = _load_supplier_from_file(_last_month_to_now_file(periods), supplier)
    cur_month = _load_supplier_from_file(DATA_FILE, supplier)

    if not may_to_now["total"]:
        return None

    # 以 5.1-至今 为基底
    base = dict(may_to_now["total"])

    # 累加消耗 + ROI2 成本
    cost_5full = (may_full["total"] or {}).get("滚动消耗", 0) or 0
    cost_6 = (cur_month["total"] or {}).get("滚动消耗", 0) or 0
    base["滚动消耗"] = cost_5full + cost_6

    roi2c_5full = (may_full["total"] or {}).get("滚动ROI2总成本", 0) or 0
    roi2c_6 = (cur_month["total"] or {}).get("滚动ROI2总成本", 0) or 0
    base["滚动ROI2总成本"] = roi2c_5full + roi2c_6

    # 重算 ROI2 和单例子成本
    gmv = base.get("滚动GMV", 0) or 0
    base["滚动ROI2"] = gmv / base["滚动ROI2总成本"] if base["滚动ROI2总成本"] > 0 else 0

    ex = base.get("例子数") or 0
    base["单例子成本"] = base["滚动消耗"] / ex if ex > 0 else None

    # 承担人明细同样整合
    detail_merged = []
    for d in may_to_now["details"]:
        chan = d.get("渠道名称")
        d_full = next((x for x in may_full["details"] if x.get("渠道名称") == chan), {})
        d_cur = next((x for x in cur_month["details"] if x.get("渠道名称") == chan), {})

        d_merged = dict(d)
        d_merged["滚动消耗"] = (d_full.get("滚动消耗", 0) or 0) + (d_cur.get("滚动消耗", 0) or 0)
        d_merged["滚动ROI2总成本"] = (d_full.get("滚动ROI2总成本", 0) or 0) + (d_cur.get("滚动ROI2总成本", 0) or 0)

        d_gmv = d_merged.get("滚动GMV", 0) or 0
        d_merged["滚动ROI2"] = d_gmv / d_merged["滚动ROI2总成本"] if d_merged["滚动ROI2总成本"] > 0 else 0

        d_ex = d_merged.get("例子数") or 0
        d_merged["单例子成本"] = d_merged["滚动消耗"] / d_ex if d_ex > 0 else None

        detail_merged.append(d_merged)

    # 计算 CPS/CPT
    gmv = base.get("滚动GMV") or 0
    cost = base.get("滚动消耗") or 0
    cps, cpt = calc_cps_cpt(gmv, cost)
    base["CPS成本"] = cps
    base["CPT成本"] = cpt

    # 计算注册转化率
    ex_cnt = base.get("例子数") or 0
    trans_cnt = base.get("当月成交数") or 0
    base["注册转化率"] = trans_cnt / ex_cnt if ex_cnt > 0 else None

    # 承担人明细同样计算
    for d in detail_merged:
        d_gmv = d.get("滚动GMV") or 0
        d_cost = d.get("滚动消耗") or 0
        d_cps, d_cpt = calc_cps_cpt(d_gmv, d_cost)
        d["CPS成本"] = d_cps
        d["CPT成本"] = d_cpt

        d_ex = d.get("例子数") or 0
        d_trans = d.get("当月成交数") or 0
        d["注册转化率"] = d_trans / d_ex if d_ex > 0 else None

    return {"total": base, "details": detail_merged}


def find_repeat_venues(all_venues):
    """
    步骤5：识别重复场次（基础名称相同的场次）
    返回 {base_name: [venue1, venue2, ...]}
    """
    from collections import defaultdict
    groups = defaultdict(list)

    for v in all_venues:
        name = v["total"]["name"]
        base = get_venue_base_name(name)
        if base:
            groups[base].append(v)

    # 只保留有2个以上的重复场次
    return {k: v for k, v in groups.items() if len(v) >= 2}


def find_new_venues(current_venues, all_venues, current_month_name="6月"):
    """
    识别当月新增场次（在数据底表里有但分场次sheet里没有的）
    返回 [{venue_dict, suggested_name}, ...]
    """
    existing_in_month = set()
    for v in all_venues:
        name = v["total"]["name"]
        if name.startswith(current_month_name):
            existing_in_month.add(name)

    new_venues = []
    for v in current_venues:
        supplier = v["total"]["name"]
        suggested_name = f"{current_month_name}{supplier}"
        base = get_venue_base_name(supplier)

        already_exists = False
        for existing in existing_in_month:
            existing_clean = re.sub(r'^\d+月', '', existing)
            existing_base = get_venue_base_name(existing_clean)
            if base == existing_base:
                already_exists = True
                break

        if not already_exists:
            ex = v["total"].get("例子数") or 0
            # 只识别例子数>0的为本月有效场次
            if ex > 0:
                cost = v["total"].get("滚动消耗") or 0
                new_venues.append({
                    "venue": v,
                    "suggested_name": suggested_name,
                    "examples": ex,
                    "cost": cost,
                })

    return new_venues


def update_venue_summary(out_wb, current_venues, venue_days_map=None,
                          current_month_name="6月"):
    """
    更新【分场次数据汇总】sheet：插入当月新场次（含承担人明细）
    venue_days_map: {供应商名: 天数}，由用户提供（无则跳过天均）
    """
    if venue_days_map is None:
        venue_days_map = {}

    ws = out_wb["分场次数据汇总"]

    # 找到表的末尾
    last_row = ws.max_row
    while last_row > 2 and ws.cell(last_row, 1).value is None:
        last_row -= 1

    write_row = last_row + 1

    # 列映射（基于 sample 中"分场次数据汇总" sheet 的实际列结构）
    col_map = {
        "天数": 2, "天均": 3,
        "滚动消耗": 6, "滚动消耗(不含赠课成本)": 7,
        "单例子成本": 8, "例子数": 9,
        "分发数": 10, "分发数\n(剔除毛例子)": 11,
        "约课数": 12, "到课数": 13,
        "应到课数": 14, "滚动到课数": 16,
        "当月成交数": 17, "滚动成交数": 18,
        "当月GMV": 19, "滚动GMV": 20,
        "海外GMV占比": 21, "滚动ASP": 22,
        "例子分发率": 23, "分发约课率": 24,
        "例子约课率": 25, "约课到课率": 26,
        "应到课率": 27, "滚动应到课率": 28,
        "到课转化率": 29, "注册转化率": 30,
        "滚动转化率": 31, "滚动ROI2总成本": 32,
        "滚动ROI2": 33,
    }

    new_count = 0
    for v in current_venues:
        supplier = v["total"]["name"]
        ex = v["total"].get("例子数") or 0

        # 只写入例子数>0的场次
        if ex <= 0:
            continue

        venue_label = f"{current_month_name}{supplier}"

        # 检查是否已存在
        already_exists = False
        for row in range(3, last_row + 1):
            v_existing = ws.cell(row, 1).value
            if v_existing and str(v_existing).strip() == venue_label:
                already_exists = True
                break
        if already_exists:
            continue

        # 写场次汇总行
        days = venue_days_map.get(supplier)
        ws.cell(write_row, 1).value = venue_label
        if days:
            ws.cell(write_row, 2).value = days
            if ex > 0:
                ws.cell(write_row, 3).value = round(ex / days, 2)

        for field, col in col_map.items():
            if field in ("天数", "天均"):
                continue
            value = v["total"].get(field)
            if value is not None:
                ws.cell(write_row, col).value = value
        write_row += 1

        # 写承担人明细
        for d in v["details"]:
            ws.cell(write_row, 1).value = d["name"]
            for field, col in col_map.items():
                if field in ("天数", "天均"):
                    continue
                value = d.get(field)
                if value is not None:
                    ws.cell(write_row, col).value = value
            write_row += 1

        new_count += 1

    print(f"  [分场次数据汇总] 新增 {new_count} 个场次，截至行 {write_row - 1}")


def read_monthly_summary(wb):
    """读取分月数据汇总（从已有 workbook）"""
    ws = wb["分月数据汇总"]
    rows = []
    for row_idx in range(3, ws.max_row + 1):
        name = ws.cell(row_idx, 1).value
        if not name:
            continue
        rec = {"name": str(name).strip()}
        for col_idx in range(2, min(ws.max_column + 1, 34)):
            field_idx = col_idx - 1
            if field_idx < len(MONTHLY_FIELDS):
                field = MONTHLY_FIELDS[field_idx]
            else:
                continue
            v = ws.cell(row_idx, col_idx).value
            rec[field] = v
        rows.append(rec)
    return rows


def read_venue_summary(wb):
    """读取分场次数据汇总（从已有 workbook）"""
    ws = wb["分场次数据汇总"]
    venues = []
    current_venue = None

    for row_idx in range(3, ws.max_row + 1):
        name = ws.cell(row_idx, 1).value
        if not name:
            continue
        name_str = str(name).strip()

        rec = {"name": name_str}
        for col_idx in range(2, min(ws.max_column + 1, 34)):
            field_idx = col_idx - 1
            if field_idx < len(MONTHLY_FIELDS):
                field = MONTHLY_FIELDS[field_idx]
            else:
                continue
            v = ws.cell(row_idx, col_idx).value
            rec[field] = v

        # 判断是场次汇总行还是承担人明细行
        # 规则：如果名称以月份开头（如"5月"、"26年5月"），则认为是场次汇总行
        is_venue_total = any(name_str.startswith(prefix) for prefix in [
            "1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月",
            "26年5月", "26年6月", "26年7月", "26年8月", "26年9月", "26年10月", "26年11月", "26年12月"
        ])

        if is_venue_total:
            current_venue = {"total": rec, "details": []}
            venues.append(current_venue)
        else:
            if current_venue:
                current_venue["details"].append(rec)

    return venues


def read_venue_days_from_sample(wb, current_month_name="5月"):
    """
    从【分场次数据汇总】读取当月各场次的天数
    返回 {供应商名（去掉月份前缀）: 天数}
    """
    ws = wb["分场次数据汇总"]
    days_map = {}

    for row in range(3, ws.max_row + 1):
        name = ws.cell(row, 1).value
        if not name:
            continue
        name_str = str(name).strip()
        if not name_str.startswith(current_month_name):
            continue
        # 去掉月份前缀
        supplier = name_str[len(current_month_name):]
        days = ws.cell(row, 2).value
        if isinstance(days, (int, float)) and days >= 0:
            days_map[supplier] = days

    return days_map


def attach_days_to_current_venues(current_venues, days_map):
    """把天数和天均补充到 current_venues 中"""
    for v in current_venues:
        supplier = v["total"]["name"]
        days = days_map.get(supplier)
        # fallback 到 VENUE_DAYS_MAP
        if days is None:
            days = VENUE_DAYS_MAP.get(supplier)

        v["total"]["天数"] = days
        ex = v["total"].get("例子数") or 0
        if days and days > 0 and ex > 0:
            v["total"]["天均"] = round(ex / days, 2)
        else:
            v["total"]["天均"] = None


def write_last_month_venues(out_wb, periods):
    """
    整合并写入上月场次数据（5.1-至今 快照为基底，滚动消耗/ROI2成本做累加）
    对已存在的行做原地 update（刷新滚动指标），新场次则插入新行。
    """
    last_month_venues = identify_last_month_venues(periods)
    if not last_month_venues:
        print(f"  [上月场次] 未识别到上月有效场次")
        return

    ws = out_wb["分场次数据汇总"]

    # 收集已有的场次标签（用于识别是否需要 update）
    existing_labels = {}
    for row in range(3, ws.max_row + 1):
        v = ws.cell(row, 1).value
        if v:
            label_str = str(v).strip()
            existing_labels[label_str] = row

    # 找末尾行（用于插入新场次）
    last_row = ws.max_row
    while last_row > 2 and ws.cell(last_row, 1).value is None:
        last_row -= 1
    write_row = last_row + 1

    col_map = {
        "天数": 2, "天均": 3,
        "CPS成本": 4, "CPT成本": 5,
        "滚动消耗": 6, "滚动消耗(不含赠课成本)": 7,
        "单例子成本": 8, "例子数": 9,
        "分发数": 10, "分发数\n(剔除毛例子)": 11,
        "约课数": 12, "到课数": 13,
        "应到课数": 14, "滚动到课数": 16,
        "当月成交数": 17, "滚动成交数": 18,
        "当月GMV": 19, "滚动GMV": 20,
        "例子约课率": 25, "约课到课率": 26,
        "到课转化率": 29, "注册转化率": 30,
        "滚动转化率": 31,
        "滚动ROI2总成本": 32, "滚动ROI2": 33,
    }

    last_month_prefix = periods["last_month"]["prefix"]
    updated = 0
    added = 0

    for supplier in last_month_venues:
        # 整合数据
        merged = merge_last_month_venue(supplier, periods)
        if not merged:
            continue

        venue_label = f"{last_month_prefix}{supplier}"

        # 检查是否已存在
        if venue_label in existing_labels:
            # 原地 update：找到行号后刷新数据列
            target_row = existing_labels[venue_label]
            t = merged["total"]

            # 如果天数为空，从 VENUE_DAYS_MAP 补充
            if ws.cell(target_row, 2).value is None:
                days = VENUE_DAYS_MAP.get(supplier)
                if days:
                    ws.cell(target_row, 2).value = days
                    ex = t.get("例子数") or 0
                    if ex > 0:
                        ws.cell(target_row, 3).value = round(ex / days, 2)

            for field, col in col_map.items():
                if field in ("天数", "天均"):
                    continue
                value = t.get(field)
                if value is not None:
                    ws.cell(target_row, col).value = value

            # 承担人明细行：按顺序找到下一行们，逐个 update
            detail_row_idx = target_row + 1
            for d in merged["details"]:
                short = d.get("name", "")
                for field, col in col_map.items():
                    if field in ("天数", "天均"):
                        continue
                    value = d.get(field)
                    if value is not None:
                        ws.cell(detail_row_idx, col).value = value
                detail_row_idx += 1

            updated += 1
            print(f"  [上月场次] ✅ 已更新 {venue_label}（滚动消耗={t.get('滚动消耗', 0):.0f}）")
        else:
            # 新增场次行
            t = merged["total"]
            ws.cell(write_row, 1).value = venue_label

            # 天数从 VENUE_DAYS_MAP 取
            days = VENUE_DAYS_MAP.get(supplier)
            if days:
                ws.cell(write_row, 2).value = days
                ex = t.get("例子数") or 0
                if ex > 0:
                    ws.cell(write_row, 3).value = round(ex / days, 2)

            # 写数据列
            for field, col in col_map.items():
                if field in ("天数", "天均"):
                    continue
                value = t.get(field)
                if value is not None:
                    ws.cell(write_row, col).value = value
            write_row += 1

            # 承担人明细
            for d in merged["details"]:
                short = d.get("name", "")
                ws.cell(write_row, 1).value = short
                for field, col in col_map.items():
                    if field in ("天数", "天均"):
                        continue
                    value = d.get(field)
                    if value is not None:
                        ws.cell(write_row, col).value = value
                write_row += 1

            added += 1
            print(f"  [上月场次] ✅ 已新增 {venue_label}（例子={t.get('例子数', 0):.0f}，{len(merged['details'])} 个承担人）")

    print(f"  [上月场次] 完成：更新 {updated} 个，新增 {added} 个场次")


def calc_cps_cpt(gmv, total_cost):
    """
    用公式计算 CPS/CPT 成本
    CPS成本 = 滚动GMV × 0.35
    CPT成本 = 滚动消耗 - CPS成本
    """
    if not isinstance(gmv, (int, float)):
        gmv = 0
    if not isinstance(total_cost, (int, float)):
        total_cost = 0
    cps = gmv * 0.35
    cpt = total_cost - cps
    return cps, cpt


def attach_cps_cpt_to_current_venues(current_venues, *args, **kwargs):
    """
    用数据底表的滚动GMV和滚动消耗计算 CPS/CPT
    CPS成本 = 滚动GMV × 0.35
    CPT成本 = 滚动消耗 - CPS成本
    同时给每个承担人明细行也算上 CPS/CPT
    """
    for v in current_venues:
        t = v["total"]
        gmv = t.get("滚动GMV") or 0
        total_cost = t.get("滚动消耗") or 0
        cps, cpt = calc_cps_cpt(gmv, total_cost)
        t["CPS成本"] = cps
        t["CPT成本"] = cpt

        # 承担人明细
        for d in v["details"]:
            d_gmv = d.get("滚动GMV") or 0
            d_cost = d.get("滚动消耗") or 0
            d_cps, d_cpt = calc_cps_cpt(d_gmv, d_cost)
            d["CPS成本"] = d_cps
            d["CPT成本"] = d_cpt


def extract_shangchao_data():
    """供 HTML 周报使用的接口：返回线下商超的完整数据"""
    # 读取输出文件（包含最新的上月/本月场次交叉数据）
    output_file = OUTPUT_DIR / SUMMARY_FILE.name
    if not output_file.exists():
        print("  [重新生成] 线下商超数据缓存不存在，调用process_shangchao...")
        main()

    wb = load_workbook(output_file, data_only=True)
    monthly = read_monthly_summary(wb)
    all_venues = read_venue_summary(wb)

    # 为所有月份行计算缺失的衍生字段（因为 Excel 中存的是公式，data_only=True 读不到缓存值）
    for r in monthly:
        if r["name"] == "汇总":
            continue

        # 计算衍生字段
        ex = r.get("例子数") or 0
        days = r.get("天数") or 0
        gmv = r.get("滚动GMV") or 0
        cost = r.get("滚动消耗") or 0
        distr = r.get("分发数") or 0
        visits = r.get("约课数") or 0
        oncourse = r.get("到课数") or 0
        oncourse_rolling = r.get("滚动到课数") or 0
        trans = r.get("滚动成交数") or 0
        cost_no_gift = r.get("滚动消耗(不含赠课成本)") or 0
        cost_total_roi2 = r.get("滚动ROI2总成本") or 0

        # 天均
        if days > 0 and ex > 0:
            r["天均"] = round(ex / days, 2)

        # CPS/CPT
        cps, cpt = calc_cps_cpt(gmv, cost)
        r["CPS消耗"] = cps
        r["CPT消耗"] = cpt

        # 单例子成本
        if ex > 0:
            r["单例子成本"] = round(cost_no_gift / ex, 2)
            r["例子约课率"] = visits / ex
            r["例子分发率"] = distr / ex

        # 分发约课率
        if distr > 0:
            r["分发约课率"] = visits / distr

        # 约课到课率
        if visits > 0:
            r["约课到课率"] = oncourse_rolling / visits

        # 到课转化率
        if oncourse > 0:
            r["到课转化率"] = r.get("当月成交数", 0) / oncourse

        # 滚动转化率（滚动成交数 / 例子数）
        if ex > 0:
            r["滚动转化率"] = trans / ex

        # 滚动ASP
        if trans > 0:
            r["滚动ASP"] = gmv / trans

        # 滚动ROI2
        if cost_total_roi2 > 0:
            r["滚动ROI2"] = gmv / cost_total_roi2

        # 注册转化率
        if ex > 0:
            r["注册转化率"] = (r.get("当月成交数") or 0) / ex

    # 当月行（最后一个非"汇总"行）
    current_month = None
    prev_month = None
    total_row = None
    for r in monthly:
        if r["name"] == "汇总":
            total_row = r
        else:
            prev_month = current_month
            current_month = r

    cur_month_name = current_month["name"] if current_month else "6月"

    # 从数据底表提取当月场次数据
    hk_df = extract_shangchao_from_source()
    current_venues = build_venue_data(hk_df)

    # 从 output 读已有的天数 + VENUE_DAYS_MAP 补充
    days_map = read_venue_days_from_sample(wb, cur_month_name)
    attach_days_to_current_venues(current_venues, days_map)

    # 从 output 读 CPS/CPT 数据
    attach_cps_cpt_to_current_venues(current_venues, wb, cur_month_name)

    # 从数据底表实时计算当月汇总数据
    if current_month:
        all_supplier_totals = hk_df[hk_df["渠道名称"] == "总计"]

        agg_fields = [
            "滚动消耗", "滚动消耗(不含赠课成本)", "例子数", "分发数",
            "分发数\n(剔除毛例子)", "约课数", "应到课数", "滚动应到课数",
            "到课数", "滚动到课数", "当月成交数", "滚动成交数",
            "当月GMV", "滚动GMV", "滚动ROI2总成本"
        ]
        cur_total = {}
        for f in agg_fields:
            s = 0
            for _, r in all_supplier_totals.iterrows():
                v = r.get(f)
                if isinstance(v, (int, float)) and pd.notna(v):
                    s += v
            cur_total[f] = s

        cur_total_days = 0
        for v in current_venues:
            d = v["total"].get("天数")
            if isinstance(d, (int, float)):
                cur_total_days += d

        cur_total_cps, cur_total_cpt = calc_cps_cpt(
            cur_total["滚动GMV"], cur_total["滚动消耗"]
        )

        current_month["天数"] = cur_total_days
        if cur_total_days > 0 and cur_total["例子数"] > 0:
            current_month["天均"] = round(cur_total["例子数"] / cur_total_days, 2)
        current_month["CPS消耗"] = cur_total_cps
        current_month["CPT消耗"] = cur_total_cpt
        current_month["滚动消耗"] = cur_total["滚动消耗"]
        current_month["滚动消耗(不含赠课成本)"] = cur_total["滚动消耗(不含赠课成本)"]
        current_month["例子数"] = cur_total["例子数"]
        current_month["分发数"] = cur_total["分发数"]
        current_month["剔除毛例子分发数"] = cur_total["分发数\n(剔除毛例子)"]
        current_month["约课数"] = cur_total["约课数"]
        current_month["应到课数"] = cur_total["应到课数"]
        current_month["滚动应到课数"] = cur_total["滚动应到课数"]
        current_month["到课数"] = cur_total["到课数"]
        current_month["滚动到课数"] = cur_total["滚动到课数"]
        current_month["当月成交数"] = cur_total["当月成交数"]
        current_month["滚动成交数"] = cur_total["滚动成交数"]
        current_month["当月GMV"] = cur_total["当月GMV"]
        current_month["滚动GMV"] = cur_total["滚动GMV"]
        current_month["滚动ROI2总成本"] = cur_total["滚动ROI2总成本"]
        if cur_total["例子数"] > 0:
            current_month["单例子成本"] = round(cur_total["滚动消耗(不含赠课成本)"] / cur_total["例子数"], 2)
            current_month["例子约课率"] = cur_total["约课数"] / cur_total["例子数"]
            current_month["例子分发率"] = cur_total["分发数"] / cur_total["例子数"]
        if cur_total["分发数"] > 0:
            current_month["分发约课率"] = cur_total["约课数"] / cur_total["分发数"]
        if cur_total["约课数"] > 0:
            current_month["约课到课率"] = cur_total["滚动到课数"] / cur_total["约课数"]
        if cur_total["到课数"] > 0:
            current_month["到课转化率"] = cur_total["当月成交数"] / cur_total["到课数"]
        if cur_total["滚动到课数"] > 0:
            current_month["滚动转化率"] = cur_total["滚动成交数"] / cur_total["滚动到课数"]
        if cur_total["滚动成交数"] > 0:
            current_month["滚动ASP"] = cur_total["滚动GMV"] / cur_total["滚动成交数"]
        if cur_total["滚动ROI2总成本"] > 0:
            current_month["滚动ROI2"] = cur_total["滚动GMV"] / cur_total["滚动ROI2总成本"]

    # 为汇总行也计算所有衍生字段（从 monthly 中所有非汇总行聚合）
    if total_row:
        summary_fields = [
            "天数", "滚动消耗", "滚动消耗(不含赠课成本)", "例子数", "分发数",
            "分发数\n(剔除毛例子)", "约课数", "应到课数", "滚动应到课数",
            "到课数", "滚动到课数", "当月成交数", "滚动成交数",
            "当月GMV", "滚动GMV", "滚动ROI2总成本"
        ]
        summary_agg = {}
        for f in summary_fields:
            s = 0
            for r in monthly:
                if r["name"] != "汇总":
                    v = r.get(f)
                    if isinstance(v, (int, float)):
                        s += v
            summary_agg[f] = s

        # 计算汇总行的衍生字段
        summary_cps, summary_cpt = calc_cps_cpt(
            summary_agg.get("滚动GMV", 0), summary_agg.get("滚动消耗", 0)
        )

        total_row["天数"] = summary_agg.get("天数", 0)
        if summary_agg.get("天数", 0) > 0 and summary_agg.get("例子数", 0) > 0:
            total_row["天均"] = round(summary_agg["例子数"] / summary_agg["天数"], 2)
        total_row["CPS消耗"] = summary_cps
        total_row["CPT消耗"] = summary_cpt
        total_row["滚动消耗"] = summary_agg.get("滚动消耗", 0)
        total_row["滚动消耗(不含赠课成本)"] = summary_agg.get("滚动消耗(不含赠课成本)", 0)
        total_row["例子数"] = summary_agg.get("例子数", 0)
        total_row["分发数"] = summary_agg.get("分发数", 0)
        total_row["剔除毛例子分发数"] = summary_agg.get("分发数\n(剔除毛例子)", 0)
        total_row["约课数"] = summary_agg.get("约课数", 0)
        total_row["应到课数"] = summary_agg.get("应到课数", 0)
        total_row["滚动应到课数"] = summary_agg.get("滚动应到课数", 0)
        total_row["到课数"] = summary_agg.get("到课数", 0)
        total_row["滚动到课数"] = summary_agg.get("滚动到课数", 0)
        total_row["当月成交数"] = summary_agg.get("当月成交数", 0)
        total_row["滚动成交数"] = summary_agg.get("滚动成交数", 0)
        total_row["当月GMV"] = summary_agg.get("当月GMV", 0)
        total_row["滚动GMV"] = summary_agg.get("滚动GMV", 0)
        total_row["滚动ROI2总成本"] = summary_agg.get("滚动ROI2总成本", 0)

        ex = summary_agg.get("例子数", 0)
        if ex > 0:
            total_row["单例子成本"] = round(summary_agg.get("滚动消耗(不含赠课成本)", 0) / ex, 2)
            total_row["例子约课率"] = summary_agg.get("约课数", 0) / ex
            total_row["例子分发率"] = summary_agg.get("分发数", 0) / ex

        distr = summary_agg.get("分发数", 0)
        if distr > 0:
            total_row["分发约课率"] = summary_agg.get("约课数", 0) / distr

        visits = summary_agg.get("约课数", 0)
        if visits > 0:
            total_row["约课到课率"] = summary_agg.get("滚动到课数", 0) / visits

        oncourse = summary_agg.get("到课数", 0)
        if oncourse > 0:
            total_row["到课转化率"] = summary_agg.get("当月成交数", 0) / oncourse

        oncourse_rolling = summary_agg.get("滚动到课数", 0)
        trans = summary_agg.get("滚动成交数", 0)
        ex_summary = summary_agg.get("例子数", 0)

        # 滚动转化率（滚动成交数 / 例子数）
        if ex_summary > 0:
            total_row["滚动转化率"] = trans / ex_summary

        # 滚动ASP
        if trans > 0:
            total_row["滚动ASP"] = summary_agg.get("滚动GMV", 0) / trans

        cost = summary_agg.get("滚动ROI2总成本", 0)
        if cost > 0:
            total_row["滚动ROI2"] = summary_agg.get("滚动GMV", 0) / cost
    periods = get_shangchao_periods()
    last_month_prefix = periods["last_month"]["prefix"]  # "26年5月"
    cur_month_prefix = periods["current_month"]["prefix"]  # "26年6月"

    last_month_venues = [v for v in all_venues
                         if v.get("total", {}).get("name", "").startswith(last_month_prefix)]

    # 为上月场次也计算缺失的衍生字段
    for venue in last_month_venues:
        t = venue["total"]

        # 计算 total 的衍生字段
        ex = t.get("例子数") or 0
        days = t.get("天数") or 0
        gmv = t.get("滚动GMV") or 0
        cost = t.get("滚动消耗") or 0
        distr = t.get("分发数") or 0
        visits = t.get("约课数") or 0
        oncourse = t.get("到课数") or 0
        oncourse_rolling = t.get("滚动到课数") or 0
        trans = t.get("滚动成交数") or 0
        cost_no_gift = t.get("滚动消耗(不含赠课成本)") or 0
        cost_total_roi2 = t.get("滚动ROI2总成本") or 0

        # 天均
        if days > 0 and ex > 0:
            t["天均"] = round(ex / days, 2)

        # CPS/CPT
        cps, cpt = calc_cps_cpt(gmv, cost)
        t["CPS成本"] = cps
        t["CPT成本"] = cpt

        # 单例子成本
        if ex > 0:
            t["单例子成本"] = round(cost_no_gift / ex, 2)
            t["例子约课率"] = visits / ex
            t["例子分发率"] = distr / ex

        # 分发约课率
        if distr > 0:
            t["分发约课率"] = visits / distr

        # 约课到课率
        if visits > 0:
            t["约课到课率"] = oncourse_rolling / visits

        # 到课转化率
        if oncourse > 0:
            t["到课转化率"] = (t.get("当月成交数") or 0) / oncourse

        # 滚动转化率（滚动成交数 / 例子数）
        if ex > 0:
            t["滚动转化率"] = trans / ex

        # 滚动ASP
        if trans > 0:
            t["滚动ASP"] = gmv / trans

        # 滚动ROI2
        if cost_total_roi2 > 0:
            t["滚动ROI2"] = gmv / cost_total_roi2

        # 注册转化率
        if ex > 0:
            t["注册转化率"] = (t.get("当月成交数") or 0) / ex

        # 同样为 details 中的每个承担人计算衍生字段
        for d in venue.get("details", []):
            d_ex = d.get("例子数") or 0
            d_days = d.get("天数") or 0
            d_gmv = d.get("滚动GMV") or 0
            d_cost = d.get("滚动消耗") or 0
            d_distr = d.get("分发数") or 0
            d_visits = d.get("约课数") or 0
            d_oncourse = d.get("到课数") or 0
            d_oncourse_rolling = d.get("滚动到课数") or 0
            d_trans = d.get("滚动成交数") or 0
            d_cost_no_gift = d.get("滚动消耗(不含赠课成本)") or 0
            d_cost_total_roi2 = d.get("滚动ROI2总成本") or 0

            if d_days > 0 and d_ex > 0:
                d["天均"] = round(d_ex / d_days, 2)

            d_cps, d_cpt = calc_cps_cpt(d_gmv, d_cost)
            d["CPS成本"] = d_cps
            d["CPT成本"] = d_cpt

            if d_ex > 0:
                d["单例子成本"] = round(d_cost_no_gift / d_ex, 2)
                d["例子约课率"] = d_visits / d_ex
                d["例子分发率"] = d_distr / d_ex

            if d_distr > 0:
                d["分发约课率"] = d_visits / d_distr

            if d_visits > 0:
                d["约课到课率"] = d_oncourse_rolling / d_visits

            if d_oncourse > 0:
                d["到课转化率"] = (d.get("当月成交数") or 0) / d_oncourse

            # 滚动转化率（滚动成交数 / 例子数）
            if d_ex > 0:
                d["滚动转化率"] = d_trans / d_ex

            if d_trans > 0:
                d["滚动ASP"] = d_gmv / d_trans

            if d_cost_total_roi2 > 0:
                d["滚动ROI2"] = d_gmv / d_cost_total_roi2

            if d_ex > 0:
                d["注册转化率"] = (d.get("当月成交数") or 0) / d_ex

    # 提取本月场次：直接使用current_venues（从数据底表实时构建，例子数>0）
    current_month_venues = current_venues

    # 重复场次：合并历史场次和当月场次（带月份前缀）
    combined_venues = list(all_venues)
    for v in current_venues:
        ex = v["total"].get("例子数") or 0
        if ex <= 0:
            continue
        venue_label = f"{cur_month_name}{v['total']['name']}"
        if not any(x["total"]["name"] == venue_label for x in combined_venues):
            cur_v = {
                "total": {**v["total"], "name": venue_label},
                "details": v["details"]
            }
            combined_venues.append(cur_v)

    repeat_venues = find_repeat_venues(combined_venues)

    return {
        "monthly": monthly,
        "all_venues": all_venues,
        "current_venues": current_venues,
        "last_month_venues": last_month_venues,
        "current_month_venues": current_month_venues,
        "current_month": current_month,
        "prev_month": prev_month,
        "total": total_row,
        "repeat_venues": repeat_venues,
        "periods": {
            "last_month": {"prefix": last_month_prefix},
            "current_month": {"prefix": cur_month_prefix},
            "yesterday": periods["yesterday"],
        },
        "key_metrics": KEY_METRICS,
        "monthly_fields": MONTHLY_FIELDS,
    }


def main():
    print("=" * 60)
    print("线下商超数据复盘 v3（跨月累计）")
    print("=" * 60)

    # 获取日期信息
    periods = get_shangchao_periods()
    print(f"\n[0] 日期信息")
    print(f"  上月: {periods['last_month']['prefix']} ({periods['last_month_to_now'][0]} ~ {periods['last_month_to_now'][1]})")
    print(f"  本月: {periods['current_month']['prefix']} ({periods['current_month_to_now'][0]} ~ {periods['current_month_to_now'][1]})")
    print()

    print("[1] 从数据底表提取线下HK商超数据...")
    hk_df = extract_shangchao_from_source()
    suppliers = [s for s in hk_df["供应商"].unique() if s != "总计"]
    print(f"  本月商超供应商: {len(suppliers)} 个")
    active_count = 0
    for s in suppliers:
        sub = hk_df[(hk_df["供应商"] == s) & (hk_df["渠道名称"] == "总计")]
        if len(sub) > 0:
            ex = sub.iloc[0].get("例子数", 0) or 0
            cost = sub.iloc[0].get("滚动消耗", 0) or 0
            if ex > 0:
                active_count += 1
                print(f"    ✓ {s}: 例子={ex}, 滚动消耗={cost:.0f}")
    print(f"  本月有数据场次（例子数>0）: {active_count} 个")

    print("\n[2] 写入当月数据整理汇总...")
    OUTPUT_DIR.mkdir(exist_ok=True)
    out_file = OUTPUT_DIR / SUMMARY_FILE.name
    if not out_file.exists():
        shutil.copy(SUMMARY_FILE, out_file)
    out_wb = load_workbook(out_file)
    write_current_month_sheet(hk_df, out_wb)

    print("\n[3] 更新分月数据汇总...")
    update_monthly_summary(hk_df, out_wb, periods["current_month"]["prefix"])

    print("\n[4] 构建本月分场次数据...")
    current_venues = build_venue_data(hk_df)
    print(f"  本月场次: {len(current_venues)} 个（含承担人明细）")

    print("\n[5] 检查新增场次...")
    wb_read = load_workbook(out_file, data_only=True)
    all_venues = read_venue_summary(wb_read)
    new_venues = find_new_venues(current_venues, all_venues, periods["current_month"]["prefix"])
    if new_venues:
        print(f"  ⚠️ 发现 {len(new_venues)} 个新增场次（需要补充天数信息）：")
        for nv in new_venues:
            print(f"    · {nv['suggested_name']}: 例子={nv['examples']}, 滚动消耗={nv['cost']:.0f}")
    else:
        print("  无新增场次")

    print("\n[6] 更新分场次数据汇总（本月新场次）...")
    venue_days_map = VENUE_DAYS_MAP
    update_venue_summary(out_wb, current_venues, venue_days_map, periods["current_month"]["prefix"])

    print(f"\n[6b] 整合上月场次数据（跨月累计）...")
    write_last_month_venues(out_wb, periods)

    out_wb.save(out_file)
    print(f"\n  输出: {out_file}")

    print("\n[7] 识别重复场次...")
    wb_final = load_workbook(out_file, data_only=True)
    final_venues = read_venue_summary(wb_final)
    repeats = find_repeat_venues(final_venues)
    print(f"  重复场次: {len(repeats)} 组")
    for base, group in repeats.items():
        names = [v["total"]["name"] for v in group]
        print(f"    {base}: {names}")

    print("\n[完成]")


if __name__ == "__main__":
    main()
