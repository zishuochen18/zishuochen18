"""
转介绍打卡 - 后端非手推达成情况
"""
"""
转介绍打卡处理脚本（模块 6）

═══════════════════════════════════════════════════════════════
功能：从销售明细 + 带量明细 + 流速表，生成两个子模块数据：
  6.1 后端非手推达成情况（手推 vs 非手推）
  6.2 打卡链路数据（可打卡学员 → 打卡 → 例子 → 约课 → 到课 → 转化）
═══════════════════════════════════════════════════════════════

【输入】
- sample/益智海外用户销售明细_末次渠道.xlsx（BI 导出）
- sample/后端转介绍流速.xlsx（手工维护，按月分 sheet）
- sample/{月}.{日}-{月}.{日}海外正式课学员带量明细_末次渠道_新.xlsx（本期+上期，2 个文件）

【统计周期】（动态计算）
- 本期：当月 1 日 ~ 昨天
- 上期：上月 1 日 ~ 上月对应日期
- 例：今天 6.3，本期 6.1-6.2，上期 5.1-5.2

【销售明细关键列索引】
- COL_DATE = 7：末次渠道更新日期
- COL_CHANNEL = 24：末次渠道名称
- COL_BOOK_TIME = 75：首次体验课约课时间
- COL_ATTEND_TIME = 93：首次体验课出席时间
- COL_SIGN_TIME = 113：首签时间
- COL_SIGN_AMOUNT = 114：首签金额
- COL_ATTRIB = 208：推荐人业绩归属（首消）

【6.1 后端非手推达成 - 筛选规则】
- 共同前提：推荐人业绩归属（首消）= 班主任
- 手推：末次渠道名称 = "手推链接"
- 非手推：末次渠道名称 ≠ "手推链接"

【6.1 各指标计算】
- 例子数：末次渠道更新日期在统计周期内的记录数
- 目标例子数：从【后端转介绍流速】对应月份 sheet 累加 1日~统计日的目标
- 约课/到课/成单：在例子集合中，分别按约课时间/出席时间/首签时间在统计周期内筛选
- 各种比率：约课率、约课到课率、到课转化率、转化率、ASP、GMV占比（公式）

【6.2 打卡链路 - 数据来源】
- 可打卡学员：带量明细中【是否可打卡学员】= 是 的人数
- 打卡人数：带量明细中【打卡次数】> 0 的人数
- 打卡次数：带量明细中【打卡次数】> 0 的求和
- 例子数：销售明细中渠道名称为以下之一：
  · 【海外益智】平台小程序连报（海报）
  · 【海外】抽奖活动
  · 【海外益智】平台小程序联报（FB/WhatsApp）
- 约课/到课/成单：在例子集合中，按对应时间筛选（同 6.1 逻辑）

【流速表 sheet 命名约定】
- "{年}年{月}月后端转介绍流速"，如 "26年5月后端转介绍流速"
- 列：[后台日期, 星期, 手推目标, 非手推目标]

【带量明细文件命名约定】
- "{月}.{日}-{月}.{日}海外正式课学员带量明细_末次渠道_新.xlsx"
- 按文件名排序识别本期/上期（月份大的为本期）
- 表头在第 11 行（数据从第 12 行起）

【对外接口】
- extract_referral_data() - 6.1 后端非手推达成数据
- extract_punch_card_data() - 6.2 打卡链路数据

【注意事项】
- 销售明细列索引基于当前 BI 表结构（210 列），如表结构变化需调整
- 带量明细文件 35MB+，每周报需重新导出 2 份
- 后端转介绍流速.xlsx 每月新增 sheet 时按命名规范创建
"""
import sys
from pathlib import Path
from datetime import datetime, timedelta
from calendar import monthrange

import pandas as pd
from openpyxl import load_workbook

sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = Path(__file__).parent
SAMPLE_DIR = BASE_DIR / "sample"
OUTPUT_DIR = BASE_DIR / "output"

SALES_FILE = SAMPLE_DIR / "益智海外用户销售明细_末次渠道.xlsx"
SPEED_FILE = SAMPLE_DIR / "后端转介绍流速.xlsx"

# 带量明细文件（本期和上期）
CUR_PUNCH_FILE = None   # 动态查找
PREV_PUNCH_FILE = None  # 动态查找

# 打卡链路中"例子数"对应的渠道名称
PUNCH_CHANNELS = [
    "【海外益智】平台小程序连报（海报）",
    "【海外】抽奖活动",
    "【海外益智】平台小程序联报（FB/WhatsApp）",
]

# 销售明细列索引（1-based, openpyxl 风格）
COL_DATE = 7              # 末次渠道更新日期
COL_CHANNEL = 24          # 末次渠道名称
COL_BOOK_TIME = 75        # 首次体验课约课时间
COL_ATTEND_TIME = 93      # 首次体验课出席时间
COL_SIGN_TIME = 113       # 首签时间
COL_SIGN_AMOUNT = 114     # 首签金额
COL_ATTRIB = 208          # 推荐人业绩归属（首消）

SALES_HEADER_ROW = 8
SALES_DATA_START_ROW = 9


def get_periods(today=None):
    """
    返回 (本期, 上期) 的日期范围
    本期：当月1号 ~ 昨天
    上期：上月1号 ~ 上月对应日期
    """
    if today is None:
        today = datetime.now().date()

    yesterday = today - timedelta(days=1)
    cur_start = today.replace(day=1)
    cur_end = yesterday

    # 上月同期
    if cur_start.month == 1:
        prev_year = cur_start.year - 1
        prev_month = 12
    else:
        prev_year = cur_start.year
        prev_month = cur_start.month - 1

    prev_start = cur_start.replace(year=prev_year, month=prev_month)
    # 上月对应日期 = 当前日 - 1（同样取昨天那天对应的上月日期）
    prev_end_day = min(yesterday.day, monthrange(prev_year, prev_month)[1])
    prev_end = datetime(prev_year, prev_month, prev_end_day).date()

    return (cur_start, cur_end), (prev_start, prev_end)


def fmt_period(start, end):
    """格式化日期范围为 5.1-5.2"""
    if start.month == end.month:
        return f"{start.month}.{start.day}-{end.month}.{end.day}"
    return f"{start.month}.{start.day}-{end.month}.{end.day}"


def to_date(v):
    """将各种格式的值转为 date 对象"""
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if hasattr(v, 'date'):
        return v.date()
    if isinstance(v, str):
        try:
            return datetime.strptime(v[:10], "%Y-%m-%d").date()
        except ValueError:
            return None
    return None


def load_sales_data():
    """读取销售明细，返回筛选过的记录（业绩归属=班主任）"""
    wb = load_workbook(SALES_FILE, data_only=True)
    ws = wb.active

    records = []
    for row in ws.iter_rows(min_row=SALES_DATA_START_ROW, values_only=True):
        if not row or row[0] is None:
            continue
        attrib = row[COL_ATTRIB - 1]
        # 先筛业绩归属为班主任
        if attrib != "班主任":
            continue

        records.append({
            "date": to_date(row[COL_DATE - 1]),
            "channel": row[COL_CHANNEL - 1],
            "book_time": to_date(row[COL_BOOK_TIME - 1]),
            "attend_time": to_date(row[COL_ATTEND_TIME - 1]),
            "sign_time": to_date(row[COL_SIGN_TIME - 1]),
            "sign_amount": row[COL_SIGN_AMOUNT - 1] or 0,
        })

    return records


def load_all_sales_data():
    """读取销售明细全部记录（不筛业绩归属，用于打卡链路）"""
    wb = load_workbook(SALES_FILE, data_only=True)
    ws = wb.active

    records = []
    for row in ws.iter_rows(min_row=SALES_DATA_START_ROW, values_only=True):
        if not row or row[0] is None:
            continue

        records.append({
            "date": to_date(row[COL_DATE - 1]),
            "channel": row[COL_CHANNEL - 1],
            "book_time": to_date(row[COL_BOOK_TIME - 1]),
            "attend_time": to_date(row[COL_ATTEND_TIME - 1]),
            "sign_time": to_date(row[COL_SIGN_TIME - 1]),
            "sign_amount": row[COL_SIGN_AMOUNT - 1] or 0,
        })

    return records


def is_in_range(d, start, end):
    """检查 date d 是否在 [start, end] 区间内"""
    if d is None:
        return False
    return start <= d <= end


def calc_metrics(records, period_start, period_end, channel_filter):
    """
    计算指标
    channel_filter: "shoutui" / "non_shoutui"
    返回 dict 含 例子数, 约课数, 到课数, 成单数, GMV
    """
    examples = 0
    bookings = 0
    attendances = 0
    signups = 0
    gmv = 0

    for r in records:
        # 渠道筛选
        ch = r["channel"]
        if channel_filter == "shoutui":
            if ch != "手推链接":
                continue
        elif channel_filter == "non_shoutui":
            if ch == "手推链接":
                continue

        # 末次渠道更新日期在统计周期内 = 例子
        if is_in_range(r["date"], period_start, period_end):
            examples += 1

            # 在例子集合中再筛
            if is_in_range(r["book_time"], period_start, period_end):
                bookings += 1
            if is_in_range(r["attend_time"], period_start, period_end):
                attendances += 1
            if is_in_range(r["sign_time"], period_start, period_end):
                signups += 1
                gmv += r["sign_amount"] or 0

    return {
        "examples": examples,
        "bookings": bookings,
        "attendances": attendances,
        "signups": signups,
        "gmv": gmv,
    }


def load_target(period_start, period_end):
    """
    从【后端转介绍流速】读目标
    返回 (手推目标, 非手推目标)
    手推目标 = 流速表第3列 1日~end的求和
    非手推目标 = 流速表第4列 1日~end的求和
    """
    wb = load_workbook(SPEED_FILE, data_only=True)

    year = period_start.year
    month = period_start.month
    yy = year - 2000  # 26

    # sheet 名格式: "26年5月后端转介绍流速"
    sheet_name = None
    for s in wb.sheetnames:
        if f"{yy}年{month}月" in s or f"{month}月" in s:
            sheet_name = s
            break

    if sheet_name is None:
        print(f"  ⚠️ 未找到 {yy}年{month}月 流速sheet")
        return 0, 0

    ws = wb[sheet_name]
    shoutui = 0
    non_shoutui = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        d = row[0]
        if d is None:
            continue
        if isinstance(d, datetime):
            d_date = d.date()
        elif hasattr(d, 'date'):
            d_date = d.date()
        else:
            continue

        if period_start <= d_date <= period_end:
            shoutui += row[2] or 0       # 第3列：手推
            non_shoutui += row[3] or 0   # 第4列：非手推

    return shoutui, non_shoutui


def build_period_data(records, period_start, period_end):
    """
    构建一个时期的三行数据：手推、非手推、合计
    """
    shoutui = calc_metrics(records, period_start, period_end, "shoutui")
    non_shoutui = calc_metrics(records, period_start, period_end, "non_shoutui")

    # 目标
    target_shoutui, target_non_shoutui = load_target(period_start, period_end)

    rows = []

    # 注：样例中 "非手推链接" 在 "手推链接" 之前显示
    # 行1：非手推链接
    rows.append({
        "name": "非手推链接",
        "examples": non_shoutui["examples"],
        "target": target_non_shoutui,
        "bookings": non_shoutui["bookings"],
        "attendances": non_shoutui["attendances"],
        "signups": non_shoutui["signups"],
        "gmv": non_shoutui["gmv"],
    })
    # 行2：手推链接
    rows.append({
        "name": "手推链接",
        "examples": shoutui["examples"],
        "target": target_shoutui,
        "bookings": shoutui["bookings"],
        "attendances": shoutui["attendances"],
        "signups": shoutui["signups"],
        "gmv": shoutui["gmv"],
    })
    # 行3：合计
    total_ex = shoutui["examples"] + non_shoutui["examples"]
    total_bk = shoutui["bookings"] + non_shoutui["bookings"]
    total_at = shoutui["attendances"] + non_shoutui["attendances"]
    total_sn = shoutui["signups"] + non_shoutui["signups"]
    total_gmv = shoutui["gmv"] + non_shoutui["gmv"]
    total_target = target_shoutui + target_non_shoutui
    rows.append({
        "name": "合计",
        "examples": total_ex,
        "target": total_target,
        "bookings": total_bk,
        "attendances": total_at,
        "signups": total_sn,
        "gmv": total_gmv,
    })

    # 计算公式字段
    for r in rows:
        ex = r["examples"]
        tg = r["target"]
        bk = r["bookings"]
        at = r["attendances"]
        sn = r["signups"]
        gv = r["gmv"]

        r["mtd_rate"] = ex / tg if tg else 0
        r["example_pct"] = ex / total_ex if total_ex else 0
        r["book_rate"] = bk / ex if ex else 0
        r["attend_rate"] = at / bk if bk else 0
        r["conv_rate"] = sn / at if at else 0
        r["total_conv_rate"] = sn / ex if ex else 0
        r["asp"] = gv / sn if sn else 0
        r["gmv_pct"] = gv / total_gmv if total_gmv else 0

    return rows


def extract_referral_data(today=None):
    """对外接口：返回完整的后端非手推达成情况数据"""
    (cur_start, cur_end), (prev_start, prev_end) = get_periods(today)

    print(f"  本期: {cur_start} ~ {cur_end}")
    print(f"  上期: {prev_start} ~ {prev_end}")

    records = load_sales_data()
    print(f"  销售明细记录数（业绩归属=班主任）: {len(records)}")

    cur_rows = build_period_data(records, cur_start, cur_end)
    prev_rows = build_period_data(records, prev_start, prev_end)

    return {
        "cur_period": fmt_period(cur_start, cur_end),
        "prev_period": fmt_period(prev_start, prev_end),
        "cur_rows": cur_rows,
        "prev_rows": prev_rows,
    }


def find_punch_files():
    """动态查找本期和上期的带量明细文件"""
    files = list(SAMPLE_DIR.glob("*海外正式课学员带量明细*.xlsx"))
    files = [f for f in files if not f.name.startswith("~$")]
    if not files:
        raise FileNotFoundError(f"sample/ 下未找到带量明细文件")
    # 按名称中的月份排序
    # 文件名格式：6.1-6.2海外正式课学员带量明细_末次渠道_新.xlsx
    files_sorted = sorted(files, key=lambda x: x.name)
    if len(files_sorted) >= 2:
        return files_sorted[1], files_sorted[0]  # 本期（月份大的）在后
    elif len(files_sorted) == 1:
        return files_sorted[0], None
    raise FileNotFoundError(f"sample/ 下未找到带量明细文件")


def load_punch_data(filepath):
    """读取带量明细文件，返回 DataFrame"""
    if filepath is None or not filepath.exists():
        raise FileNotFoundError(f"带量明细文件缺失: {filepath}")
    df = pd.read_excel(filepath, header=10)
    return df


def calc_punch_card_metrics(punch_df, sales_records, period_start, period_end):
    """
    计算打卡链路数据指标
    - punch_df: 带量明细 DataFrame
    - sales_records: 已加载的销售明细记录
    - period_start, period_end: 统计周期
    """
    if punch_df is None or len(punch_df) == 0:
        return {
            "can_punch": 0, "punched": 0, "punch_count": 0,
            "punch_rate": 0, "avg_punch": 0,
            "examples": 0, "split_rate": 0,
            "bookings": 0, "book_rate": 0,
            "attendances": 0, "attend_rate": 0,
            "signups": 0, "conv_rate": 0,
            "total_conv_rate": 0, "gmv": 0, "asp": 0,
        }

    # 可打卡学员
    col_can = "是否可打卡学员(截止至结束日期)"
    can_punch = (punch_df[col_can] == "是").sum()

    # 打卡人数和打卡次数
    col_punch = "打卡次数"
    punch_series = pd.to_numeric(punch_df[col_punch], errors="coerce").fillna(0)
    punched = (punch_series > 0).sum()
    punch_count = int(punch_series[punch_series > 0].sum())

    # 打卡率和人均
    punch_rate = punched / can_punch if can_punch > 0 else 0
    avg_punch = punch_count / punched if punched > 0 else 0

    # 例子数：从销售明细中筛选特定渠道
    examples = 0
    bookings = 0
    attendances = 0
    signups = 0
    gmv = 0

    for r in sales_records:
        ch = r["channel"]
        if ch not in PUNCH_CHANNELS:
            continue
        if is_in_range(r["date"], period_start, period_end):
            examples += 1
            if is_in_range(r["book_time"], period_start, period_end):
                bookings += 1
            if is_in_range(r["attend_time"], period_start, period_end):
                attendances += 1
            if is_in_range(r["sign_time"], period_start, period_end):
                signups += 1
                gmv += r["sign_amount"] or 0

    # 打卡裂变率
    split_rate = examples / punch_count if punch_count > 0 else 0

    # 比率
    book_rate = bookings / examples if examples > 0 else 0
    attend_rate = attendances / bookings if bookings > 0 else 0
    conv_rate = signups / attendances if attendances > 0 else 0
    total_conv_rate = signups / examples if examples > 0 else 0
    asp = gmv / signups if signups > 0 else 0

    return {
        "can_punch": can_punch,
        "punched": punched,
        "punch_count": punch_count,
        "punch_rate": punch_rate,
        "avg_punch": avg_punch,
        "examples": examples,
        "split_rate": split_rate,
        "bookings": bookings,
        "book_rate": book_rate,
        "attendances": attendances,
        "attend_rate": attend_rate,
        "signups": signups,
        "conv_rate": conv_rate,
        "total_conv_rate": total_conv_rate,
        "gmv": gmv,
        "asp": asp,
    }


def extract_punch_card_data(today=None):
    """对外接口：返回打卡链路数据"""
    (cur_start, cur_end), (prev_start, prev_end) = get_periods(today)

    print(f"  本期: {cur_start} ~ {cur_end}")
    print(f"  上期: {prev_start} ~ {prev_end}")

    # 查找带量明细文件
    cur_file, prev_file = find_punch_files()
    print(f"  本期文件: {cur_file.name if cur_file else 'None'}")
    print(f"  上期文件: {prev_file.name if prev_file else 'None'}")

    # 加载数据（本期必须存在）
    cur_df = load_punch_data(cur_file)

    # 上期如果不存在则设为 None（不抛异常，上期可能没有数据）
    prev_df = None
    if prev_file:
        try:
            prev_df = load_punch_data(prev_file)
        except FileNotFoundError:
            prev_df = None

    print(f"  本期行数: {len(cur_df) if cur_df is not None else 0}")
    print(f"  上期行数: {len(prev_df) if prev_df is not None else 0}")

    # 加载销售明细（不筛业绩归属，因为打卡链路只看渠道名）
    all_records = load_all_sales_data()
    print(f"  销售明细总记录数: {len(all_records)}")

    cur_metrics = calc_punch_card_metrics(cur_df, all_records, cur_start, cur_end)
    prev_metrics = calc_punch_card_metrics(prev_df, all_records, prev_start, prev_end)

    return {
        "cur_period": fmt_period(cur_start, cur_end),
        "prev_period": fmt_period(prev_start, prev_end),
        "cur": cur_metrics,
        "prev": prev_metrics,
    }


def main():
    print("=" * 60)
    print("转介绍打卡 - 后端非手推达成情况")
    print("=" * 60)

    data = extract_referral_data()

    print(f"\n=== 本期 {data['cur_period']} ===")
    print(f"  {'类型':<12} {'例子数':>6} {'目标':>6} {'达成率':>8} "
          f"{'约课':>5} {'到课':>5} {'成单':>5} {'GMV':>10}")
    for r in data["cur_rows"]:
        print(f"  {r['name']:<12} {r['examples']:>6} {r['target']:>6} "
              f"{r['mtd_rate']*100:>7.2f}% "
              f"{r['bookings']:>5} {r['attendances']:>5} "
              f"{r['signups']:>5} {r['gmv']:>10.1f}")

    print(f"\n=== 上期 {data['prev_period']} ===")
    print(f"  {'类型':<12} {'例子数':>6} {'目标':>6} {'达成率':>8} "
          f"{'约课':>5} {'到课':>5} {'成单':>5} {'GMV':>10}")
    for r in data["prev_rows"]:
        print(f"  {r['name']:<12} {r['examples']:>6} {r['target']:>6} "
              f"{r['mtd_rate']*100:>7.2f}% "
              f"{r['bookings']:>5} {r['attendances']:>5} "
              f"{r['signups']:>5} {r['gmv']:>10.1f}")

    print("\n" + "=" * 60)
    print("转介绍打卡 - 打卡链路数据")
    print("=" * 60)

    punch_data = extract_punch_card_data()

    print(f"\n=== 本期 {punch_data['cur_period']} ===")
    c = punch_data["cur"]
    print(f"  可打卡学员: {c['can_punch']}")
    print(f"  打卡人数: {c['punched']}")
    print(f"  打卡次数: {c['punch_count']}")
    print(f"  打卡率: {c['punch_rate']*100:.2f}%")
    print(f"  人均打卡次数: {c['avg_punch']:.2f}")
    print(f"  例子数: {c['examples']}")
    print(f"  打卡裂变率: {c['split_rate']:.4f}")
    print(f"  约课数: {c['bookings']}, 约课率: {c['book_rate']*100:.2f}%")
    print(f"  到课数: {c['attendances']}, 约课到课率: {c['attend_rate']*100:.2f}%")
    print(f"  转化例子数: {c['signups']}, 到课转化率: {c['conv_rate']*100:.2f}%")
    print(f"  注册转化率: {c['total_conv_rate']*100:.2f}%")
    print(f"  GMV: {c['gmv']:.1f}, ASP: {c['asp']:.2f}")

    print(f"\n=== 上期 {punch_data['prev_period']} ===")
    p = punch_data["prev"]
    print(f"  可打卡学员: {p['can_punch']}")
    print(f"  打卡人数: {p['punched']}")
    print(f"  打卡次数: {p['punch_count']}")
    print(f"  打卡率: {p['punch_rate']*100:.2f}%")
    print(f"  人均打卡次数: {p['avg_punch']:.2f}")
    print(f"  例子数: {p['examples']}")
    print(f"  打卡裂变率: {p['split_rate']:.4f}")
    print(f"  约课数: {p['bookings']}, 约课率: {p['book_rate']*100:.2f}%")
    print(f"  到课数: {p['attendances']}, 约课到课率: {p['attend_rate']*100:.2f}%")
    print(f"  转化例子数: {p['signups']}, 到课转化率: {p['conv_rate']*100:.2f}%")
    print(f"  注册转化率: {p['total_conv_rate']*100:.2f}%")
    print(f"  GMV: {p['gmv']:.1f}, ASP: {p['asp']:.2f}")

    print("\n[完成]")


if __name__ == "__main__":
    main()
