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

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = Path(__file__).parent
SAMPLE_DIR = BASE_DIR / "sample"
OUTPUT_DIR = BASE_DIR / "output"
DATA_FILE = BASE_DIR.parent / "港澳流速" / "sample" / "海外港澳商务_各渠道主辅投数据.xlsx"
SUMMARY_FILE = SAMPLE_DIR / "线下商超内容汇总.xlsx"

# 排除的展会关键词
EXCLUDE_KEYWORDS = ["展", "STEM"]

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
    "啟田商場3場": 4,
    "樂富商場市集": 4,
}

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


def write_current_month_sheet(hk_df, out_wb):
    """
    步骤2：把当月商超数据写入【当月数据整理汇总】sheet
    保留合计行的公式不动，只写明细行
    """
    ws = out_wb["当月数据整理汇总"]

    # 清空数据区（保留行1-2表头，从行3开始写）
    # 先找合计行位置（通常在最后）
    max_clear = 100
    for row in range(3, max_clear):
        for col in range(1, 36):
            cell = ws.cell(row, col)
            # 跳过合计行（保留公式）
            if col == 1 and cell.value and "合计" in str(cell.value):
                continue
            if row < max_clear - 5:  # 留最后几行给合计
                cell.value = None

    # 按供应商分组写入
    write_row = 3
    suppliers = [s for s in hk_df["供应商"].unique() if s != "总计"]

    for supplier in suppliers:
        sub = hk_df[hk_df["供应商"] == supplier]
        for _, row in sub.iterrows():
            chan = row.get("渠道名称")
            if not isinstance(chan, str) or chan == "总计":
                continue

            # 写渠道组、供应商、渠道名称
            ws.cell(write_row, 2).value = "线下HK"
            ws.cell(write_row, 3).value = supplier
            ws.cell(write_row, 4).value = chan

            # 写数据列（从列8开始，对应 DATA_FIELDS）
            col_offset = 8
            for i, field in enumerate(DATA_FIELDS):
                v = row.get(field)
                if pd.notna(v):
                    ws.cell(write_row, col_offset + i).value = v
            write_row += 1

    print(f"  [当月数据整理汇总] 写入 {write_row - 3} 行明细")
    return write_row


def update_monthly_summary(out_wb, current_month_name="5月"):
    """
    步骤3：把当月合计数据写入【分月数据汇总】对应月份行
    从【当月数据整理汇总】的合计行读取值，写入分月表
    """
    ws_cur = out_wb["当月数据整理汇总"]
    ws_monthly = out_wb["分月数据汇总"]

    # 找当月数据整理汇总的合计行
    total_row_idx = None
    for row in range(3, ws_cur.max_row + 1):
        v = ws_cur.cell(row, 1).value
        if v and "合计" in str(v):
            total_row_idx = row
            break

    if not total_row_idx:
        print("  ⚠️ 未找到当月数据整理汇总的合计行")
        return

    # 找分月数据汇总中对应月份的行
    target_row = None
    for row in range(3, ws_monthly.max_row + 1):
        v = ws_monthly.cell(row, 1).value
        if v and str(v).strip() == current_month_name:
            target_row = row
            break

    if not target_row:
        print(f"  ⚠️ 分月数据汇总中未找到 {current_month_name} 行")
        return

    # 从当月合计行读取数据，写入分月表对应列
    # 注意：分月表的"CPS消耗(列4)"和"CPT消耗(列5)"是手工维护的，不覆盖
    field_mapping = {
        # 6: 8,   # 滚动消耗 - 改为公式 =CPS消耗+CPT消耗（保留分月表已有的公式或值）
        7: 9,   # 滚动消耗(不含赠课成本)
        8: 10,  # 单例子成本
        9: 11,  # 例子数
        10: 12, # 分发数
        11: 13, # 剔除毛例子分发数
        12: 14, # 约课数
        13: 15, # 应到课数
        14: 16, # 滚动应到课数
        15: 17, # 到课数
        16: 18, # 滚动到课数
        17: 19, # 当月成交数
        18: 20, # 滚动成交数
        19: 21, # 当月GMV
        20: 22, # 滚动GMV
        21: 23, # 海外GMV占比
        22: 24, # 滚动ASP
        23: 25, # 例子分发率
        24: 26, # 分发约课率
        25: 27, # 例子约课率
        26: 28, # 约课到课率
        27: 29, # 应到课率
        28: 30, # 滚动应到课率
        29: 31, # 到课转化率
        30: 32, # 注册转化率
        31: 33, # 滚动转化率
        32: 34, # 滚动ROI2总成本
        33: 35, # 滚动ROI2
    }

    for monthly_col, cur_col in field_mapping.items():
        v = ws_cur.cell(total_row_idx, cur_col).value
        if v is not None:
            ws_monthly.cell(target_row, monthly_col).value = v

    print(f"  [分月数据汇总] 已更新 {current_month_name} 行（CPS/CPT 列保留不动）")


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


def find_new_venues(current_venues, all_venues, current_month_name="5月"):
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
                          current_month_name="5月"):
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
        has_days = rec.get("天数")
        if has_days and isinstance(has_days, (int, float)) and has_days > 0:
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
    if not SUMMARY_FILE.exists():
        return None

    # 读取已有的汇总表
    wb = load_workbook(SUMMARY_FILE, data_only=True)
    monthly = read_monthly_summary(wb)
    all_venues = read_venue_summary(wb)

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

    cur_month_name = current_month["name"] if current_month else "5月"

    # 从数据底表提取当月场次数据
    hk_df = extract_shangchao_from_source()
    current_venues = build_venue_data(hk_df)

    # 从 sample 读已有的天数 + VENUE_DAYS_MAP 补充
    days_map = read_venue_days_from_sample(wb, cur_month_name)
    attach_days_to_current_venues(current_venues, days_map)

    # 从 sample 读 CPS/CPT 数据
    attach_cps_cpt_to_current_venues(current_venues, wb, cur_month_name)

    # 从数据底表实时计算当月汇总数据
    # 重要：分月趋势的5月行 = 所有非展会线下HK供应商的总计行汇总
    # （不是只看例子>0的有效场次，因为像啟田商場2場、太和商場等虽然例子=0
    # 但有滚动消耗，也属于商超费用支出）
    if current_month:
        # 取所有非展会线下HK供应商的总计行
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

        # 天数：累加 current_venues 中的天数（仅有效场次有天数）+ VENUE_DAYS_MAP
        cur_total_days = 0
        for v in current_venues:
            d = v["total"].get("天数")
            if isinstance(d, (int, float)):
                cur_total_days += d

        # 用公式计算 CPS/CPT
        cur_total_cps, cur_total_cpt = calc_cps_cpt(
            cur_total["滚动GMV"], cur_total["滚动消耗"]
        )

        # 写回 current_month
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

    # 重复场次：合并历史场次和当月场次（带月份前缀）
    combined_venues = list(all_venues)
    for v in current_venues:
        ex = v["total"].get("例子数") or 0
        if ex <= 0:
            continue
        # 检查是否已存在
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
        "current_month": current_month,
        "prev_month": prev_month,
        "total": total_row,
        "repeat_venues": repeat_venues,
        "key_metrics": KEY_METRICS,
        "monthly_fields": MONTHLY_FIELDS,
    }


def main():
    print("=" * 60)
    print("线下商超数据复盘 v2")
    print("=" * 60)

    print("\n[1] 从数据底表提取线下HK商超数据...")
    hk_df = extract_shangchao_from_source()
    suppliers = [s for s in hk_df["供应商"].unique() if s != "总计"]
    print(f"  商超供应商: {len(suppliers)} 个")
    active_count = 0
    for s in suppliers:
        sub = hk_df[(hk_df["供应商"] == s) & (hk_df["渠道名称"] == "总计")]
        if len(sub) > 0:
            ex = sub.iloc[0].get("例子数", 0) or 0
            cost = sub.iloc[0].get("滚动消耗", 0) or 0
            if ex > 0:
                active_count += 1
                print(f"    ✓ {s}: 例子={ex}, 滚动消耗={cost:.0f}")
    print(f"  当月有数据场次（例子数>0）: {active_count} 个")

    print("\n[2] 写入当月数据整理汇总...")
    OUTPUT_DIR.mkdir(exist_ok=True)
    out_file = OUTPUT_DIR / SUMMARY_FILE.name
    shutil.copy(SUMMARY_FILE, out_file)
    out_wb = load_workbook(out_file)
    write_current_month_sheet(hk_df, out_wb)

    print("\n[3] 更新分月数据汇总...")
    update_monthly_summary(out_wb, "5月")

    print("\n[4] 构建当月分场次数据...")
    current_venues = build_venue_data(hk_df)
    print(f"  当月场次: {len(current_venues)} 个（含承担人明细）")

    print("\n[5] 检查新增场次...")
    wb_read = load_workbook(SUMMARY_FILE, data_only=True)
    all_venues = read_venue_summary(wb_read)
    new_venues = find_new_venues(current_venues, all_venues, "5月")
    if new_venues:
        print(f"  ⚠️ 发现 {len(new_venues)} 个新增场次（需要补充天数信息）：")
        for nv in new_venues:
            print(f"    · {nv['suggested_name']}: 例子={nv['examples']}, 滚动消耗={nv['cost']:.0f}")
        print("  💡 请在 main() 中传入 venue_days_map 参数，或后续手动维护天数")
    else:
        print("  无新增场次")

    print("\n[6] 更新分场次数据汇总...")
    # venue_days_map: {供应商名: 天数}，可由用户在调用时提供
    venue_days_map = VENUE_DAYS_MAP
    update_venue_summary(out_wb, current_venues, venue_days_map, "5月")

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
