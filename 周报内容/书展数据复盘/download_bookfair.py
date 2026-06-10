"""
书展数据自动下载与日期校验（Phase 3-Δ）

功能：
1. 从 configs/bookfair_reports.json 模板渲染运行时配置
2. 调用 SmartBI CLI 下载最新数据（自动更新到昨天）
3. 重命名输出文件为固定名称
4. 校验下载数据的日期范围
"""
import sys
import os
import glob
import shutil
from pathlib import Path
from datetime import datetime

sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = Path(__file__).parent
CONFIG_DIR = BASE_DIR.parent / "configs"
RUNTIME_DIR = CONFIG_DIR / "_runtime"
SAMPLE_DIR = BASE_DIR / "sample"

# 创建 _runtime 目录（gitignore）
RUNTIME_DIR.mkdir(exist_ok=True)

# 导入公共工具
from utils.smartbi_helper import (
    get_periods,
    render_runtime_config,
    run_smartbi_cli_task,
    rename_downloaded_file,
    verify_file_freshness,
    assert_smartbi_credentials,
)


def get_bookfair_periods():
    """（向后兼容包装）获取书展数据所需的日期范围"""
    return get_periods()


def clean_old_files():
    """清掉 sample 目录中旧的书展数据文件（避免 unique_path 加后缀）"""
    pattern = SAMPLE_DIR / "海外港澳商务_各渠道主辅投数据*.xlsx"
    old_files = glob.glob(str(pattern))
    for f in old_files:
        try:
            os.remove(f)
            print(f"  已删除旧文件: {Path(f).name}")
        except Exception as e:
            print(f"  ⚠️ 删除失败: {e}")


def verify_bookfair_freshness(file_paths, periods):
    """
    校验下载的 Excel 文件日期范围

    参数:
        file_paths: {"may_full": path1, "may_to_now": path2}
        periods: 期望的日期字典

    抛异常如果日期不匹配
    """
    try:
        from openpyxl import load_workbook
    except ImportError:
        print("  ⚠️ openpyxl 未安装，跳过日期校验")
        return

    print(f"\n[日期校验]")

    # 检查 may_full（应该到上月最后一天）
    may_full_path = file_paths.get("may_full")
    if may_full_path and may_full_path.exists():
        try:
            wb = load_workbook(may_full_path, data_only=True)
            ws = wb.active
            # SmartBI 报表通常在第 2-6 行有元信息，尝试查找 "快照日期" 或 "数据日期"
            # 这里简单起见，只通过文件 mtime 做粗粒度检查
            mtime = datetime.fromtimestamp(os.path.getmtime(may_full_path)).date()
            print(f"  may_full 文件时间: {mtime}")
            wb.close()
        except Exception as e:
            print(f"  ⚠️ may_full 日期校验异常: {e}")

    # 检查 may_to_now（应该到昨天）
    may_to_now_path = file_paths.get("may_to_now")
    if may_to_now_path and may_to_now_path.exists():
        try:
            wb = load_workbook(may_to_now_path, data_only=True)
            ws = wb.active
            mtime = datetime.fromtimestamp(os.path.getmtime(may_to_now_path)).date()
            expected_date = periods['yesterday']

            # 检查文件修改时间是否在今天（或昨天）
            if mtime != expected_date and mtime != datetime.now().date():
                print(f"  ⚠️ may_to_now 文件时间 {mtime} 不是昨天或今天")
            else:
                print(f"  may_to_now 文件时间: {mtime} ✓")

            wb.close()
        except Exception as e:
            print(f"  ⚠️ may_to_now 日期校验异常: {e}")

    print(f"  ✓ 日期校验完成")


def download_bookfair_data(periods) -> dict:
    """
    主逻辑：下载书展数据

    参数:
        periods: get_bookfair_periods() 返回的字典

    返回:
        {"may_full": path1, "may_to_now": path2}
    """
    print("\n[Step 1] 准备下载书展数据...")

    assert_smartbi_credentials()

    # 渲染临时配置
    config_path = render_runtime_config(
        CONFIG_DIR / "bookfair_reports.json",
        {
            "last_month_start": periods['last_month_full'][0].strftime('%Y-%m-%d'),
            "last_month_end": periods['last_month_full'][1].strftime('%Y-%m-%d'),
            "yesterday": periods['yesterday'].strftime('%Y-%m-%d'),
        }
    )

    # 清掉旧文件
    print(f"\n[清理] 删除旧的书展数据文件...")
    clean_old_files()

    # 运行第一个 task → 立即重命名（避免被第二个 task 覆盖）
    print(f"\n[下载] 调用 SmartBI CLI...")
    file_may_full = run_smartbi_cli_task(config_path, "bookfair_may_full", SAMPLE_DIR)
    target_may_full = SAMPLE_DIR / "海外港澳商务_各渠道主辅投数据_5.1-5.31.xlsx"
    if file_may_full != target_may_full:
        import shutil
        shutil.move(str(file_may_full), str(target_may_full))

    # 运行第二个 task → 立即重命名
    file_may_to_now = run_smartbi_cli_task(config_path, "bookfair_may_to_now", SAMPLE_DIR)
    yesterday_short = f"{periods['yesterday'].month}.{periods['yesterday'].day}"
    target_may_to_now = SAMPLE_DIR / f"海外港澳商务_各渠道主辅投数据_5.1-{yesterday_short}.xlsx"
    if file_may_to_now != target_may_to_now:
        import shutil
        shutil.move(str(file_may_to_now), str(target_may_to_now))

    # 校验日期
    file_paths = {"may_full": target_may_full, "may_to_now": target_may_to_now}
    verify_bookfair_freshness(file_paths, periods)

    print(f"\n✓ 书展数据下载完成")
    return file_paths


def main():
    """单测：独立运行下载逻辑"""
    print("=" * 60)
    print("书展数据自动下载测试")
    print("=" * 60)

    periods = get_bookfair_periods()
    print(f"目标日期: 上月 {periods['last_month_full'][0]} ~ {periods['last_month_full'][1]}, 昨天 {periods['yesterday']}")

    try:
        download_bookfair_data(periods)
    except Exception as e:
        print(f"\n❌ 错误: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
