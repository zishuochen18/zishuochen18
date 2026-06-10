"""
转介绍打卡 - SmartBI 报表下载 + 手工文件校验

下载流程：
1. sales_detail_referral - 自动下载（浏览器）
2. student_带量_current - 手工提供（下载前放入 sample/ 目录）
3. student_带量_last - 手工提供（下载前放入 sample/ 目录）

校验流程：
1. student_带量 文件存在性 + mtime 新鲜度（24 小时窗口）
2. 后端转介绍流速.xlsx - 必须存在当月 sheet
3. 转介绍打卡内容.xlsx - 仅校验存在
"""
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = Path(__file__).parent
CONFIG_DIR = BASE_DIR.parent / "configs"
SAMPLE_DIR = BASE_DIR / "sample"

from utils.smartbi_helper import (
    get_periods,
    render_runtime_config,
    run_smartbi_cli_task,
    run_smartbi_browser_task,
    rename_downloaded_file,
    assert_smartbi_credentials,
)


def validate_manual_files(periods):
    """按 glob 前缀匹配 student_带量 文件，缺失/过期 → 打印指引并 raise FileNotFoundError。

    glob 前缀匹配（不强制 _新.xlsx 精确文件名）— 因为浏览器下载会自动加 (31)/(1) 等后缀。
    与 process_referral.find_punch_files() 行为对齐。
    """
    from datetime import datetime
    from calendar import monthrange
    import glob as glob_module

    cur_range = f"{periods['current_month_to_now'][0].month}.{periods['current_month_to_now'][0].day}-{periods['yesterday'].month}.{periods['yesterday'].day}"
    # 上期日期范围：last_month_start ~ yesterday在上月内对应的日期
    last_month_start = periods['last_month_to_now'][0]
    yesterday = periods['yesterday']
    # 计算上月对应日期（yesterday的日期数，限制在上月天数范围内）
    days_in_last_month = monthrange(last_month_start.year, last_month_start.month)[1]
    last_month_corresponding_day_num = min(yesterday.day, days_in_last_month)
    last_month_corresponding = last_month_start.replace(day=last_month_corresponding_day_num)
    last_range = f"{last_month_start.month}.{last_month_start.day}-{last_month_corresponding.month}.{last_month_corresponding.day}"
    expected = {
        "本期": (cur_range, periods['current_month_to_now'][0], periods['yesterday']),
        "上期": (last_range, last_month_start, last_month_corresponding),
    }

    # 24小时新鲜度窗口
    fresh_cutoff = datetime.now().timestamp() - 24 * 3600
    missing, stale, found = [], [], {}

    for label, (date_range, start, end) in expected.items():
        # glob: 比如 "6.1-6.9*海外正式课学员带量明细*.xlsx"
        candidates = list(SAMPLE_DIR.glob(f"{date_range}*海外正式课学员带量明细*.xlsx"))
        candidates = [c for c in candidates if not c.name.startswith("~$")]
        if not candidates:
            missing.append((label, date_range, start, end))
            continue
        # 取最新的
        f = max(candidates, key=lambda x: x.stat().st_mtime)
        if f.stat().st_mtime < fresh_cutoff:
            stale.append((label, f, start, end))
        else:
            found[label] = f

    if missing or stale:
        print("\n" + "=" * 60)
        print("❌ 需要手工下载 student_带量 报表")
        print("=" * 60)
        print("  报表：海外正式课学员带量明细_末次渠道_新")
        print("  报表 ID：I2c928087019a727e727e35b1019a776750975c74")
        print("  路径：分析报表/海外直播业务线/海外后端/思维-后端/转介绍/转介绍明细/海外正式课学员带量明细_末次渠道_新")
        print(f"  目标目录：{SAMPLE_DIR}")
        print()
        for label, date_range, start, end in missing:
            print(f"  [缺失] {label}：文件名前缀 {date_range}*海外正式课学员带量明细*.xlsx")
            print(f"         筛选条件 → 开始日期={start}, 结束日期={end}")
        for label, f, start, end in stale:
            mtime_str = datetime.fromtimestamp(f.stat().st_mtime).strftime('%Y-%m-%d %H:%M')
            print(f"  [过期] {label}：{f.name}（mtime={mtime_str}，需24小时内导出）")
            print(f"         筛选条件 → 开始日期={start}, 结束日期={end}")
        print("=" * 60 + "\n")
        raise FileNotFoundError("student_带量 文件缺失或过期，请手工导出后重新运行")

    return {
        "带量_current": found["本期"],
        "带量_last":    found["上期"],
    }


def run_download(periods=None):
    """
    下载转介绍数据：销售明细自动下载 + student_带量 手工校验 + 手工文件校验

    参数:
        periods: 日期窗口字典（可选）

    返回:
        {
            "sales_detail": Path,
            "带量_current": Path,
            "带量_last": Path,
            "后端流速": Path,
            "打卡内容": Path,
        }

    异常:
        FileNotFoundError / RuntimeError
    """
    if periods is None:
        periods = get_periods()

    print("[下载步骤] 转介绍模块（1 自动 + 2 手工 + 2 手工文件校验）")

    assert_smartbi_credentials()

    # 渲染配置
    config_path = render_runtime_config(
        CONFIG_DIR / "referral_reports.json",
        {
            "current_month_start": periods['current_month_to_now'][0].strftime('%Y-%m-%d'),
            "yesterday": periods['yesterday'].strftime('%Y-%m-%d'),
            "last_month_start": periods['last_month_to_now'][0].strftime('%Y-%m-%d'),
            "last_month_corresponding_day": periods['last_month_to_now'][1].strftime('%Y-%m-%d'),
        }
    )

    # 读配置文件
    import json
    with open(CONFIG_DIR / "referral_reports.json", 'r', encoding='utf-8') as f:
        config = json.load(f)

    # 只清销售明细旧文件（不清手工的 student_带量 文件）
    import glob
    old_patterns = [
        SAMPLE_DIR / "益智海外用户销售明细*.xlsx",
    ]
    for pattern in old_patterns:
        for f in glob.glob(str(pattern)):
            try:
                Path(f).unlink()
                print(f"  [清理] {Path(f).name}")
            except:
                pass

    # 下载报表
    result = {}

    # 读渲染后的配置文件（获取最新的 filters）
    import json
    with open(config_path, 'r', encoding='utf-8') as f:
        runtime_config = json.load(f)

    # Task 1: 销售明细（走浏览器导出）
    print(f"  [下载] sales_detail_referral...")
    task_cfg = runtime_config["tasks"]["sales_detail_referral"]
    filters = [
        (override["key"], override["value"], override["displayValue"])
        for override in task_cfg["filters"].get("overrides", [])
    ]
    output_path = SAMPLE_DIR / "益智海外用户销售明细_末次渠道.xlsx"
    result["sales_detail"] = run_smartbi_browser_task(
        report_id=task_cfg["report"]["id"],
        filters=filters,
        output_path=output_path,
    )

    # Task 2 & 3: student_带量 手工校验
    print(f"  [校验] student_带量 手工文件...")
    manual_files = validate_manual_files(periods)
    result["带量_current"] = manual_files["带量_current"]
    result["带量_last"] = manual_files["带量_last"]

    # 手工文件校验
    print(f"  [校验] 其他手工文件...")

    # 1. 后端转介绍流速.xlsx - 必须存在当月 sheet
    flow_file = SAMPLE_DIR / "后端转介绍流速.xlsx"
    if not flow_file.exists():
        print(f"  ⚠️ 后端转介绍流速.xlsx 缺失，将使用现有数据")
    else:
        try:
            from openpyxl import load_workbook
            wb = load_workbook(flow_file)
            expected_sheet = f"{periods['current_month']['year_short']}年{periods['current_month']['month']}月后端转介绍流速"
            if expected_sheet not in wb.sheetnames:
                print(f"  ⚠️ 流速表缺少 sheet '{expected_sheet}'，将使用现有数据")
            else:
                print(f"  ✓ 后端转介绍流速.xlsx（含当月 sheet）")
            wb.close()
        except Exception as e:
            print(f"  ⚠️ 后端流速校验异常: {e}")

    result["后端流速"] = flow_file

    # 2. 转介绍打卡内容.xlsx - 仅校验存在
    punch_file = SAMPLE_DIR / "转介绍打卡内容.xlsx"
    if not punch_file.exists():
        raise FileNotFoundError(f"  ❌ 转介绍打卡内容.xlsx 缺失: {punch_file}")
    print(f"  ✓ 转介绍打卡内容.xlsx")
    result["打卡内容"] = punch_file

    print(f"  ✓ 转介绍模块下载校验完成")
    return result


def main():
    """单测"""
    print("=" * 60)
    print("转介绍打卡 - 下载校验测试")
    print("=" * 60)

    try:
        periods = get_periods()
        result = run_download(periods)
        print(f"\n✓ 成功: {result}")
    except Exception as e:
        print(f"\n❌ 错误: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
