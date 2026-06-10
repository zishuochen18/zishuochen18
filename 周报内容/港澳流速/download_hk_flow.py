"""
港澳流速 - 共享底表自动下载

下载 KOL/商超/书展 三个模块共用的数据底表：
海外港澳商务_各渠道主辅投数据 (2).xlsx

落盘到: 港澳流速/sample/海外港澳商务_各渠道主辅投数据 (2).xlsx
"""
import sys
from pathlib import Path
from datetime import datetime

sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = Path(__file__).parent
CONFIG_DIR = BASE_DIR.parent / "configs"
SAMPLE_DIR = BASE_DIR / "sample"

from utils.smartbi_helper import (
    get_periods,
    render_runtime_config,
    run_smartbi_cli_task,
    rename_downloaded_file,
    verify_file_freshness,
    assert_smartbi_credentials,
)


def run_download(periods=None):
    """
    下载港澳流速底表

    参数:
        periods: 日期窗口字典（可选，默认调用 get_periods()）

    返回:
        {"hk_flow_current": Path} 下载文件映射

    异常:
        RuntimeError / FileNotFoundError
    """
    if periods is None:
        periods = get_periods()

    print("[下载步骤] 港澳流速底表（KOL/商超/书展共用）")

    assert_smartbi_credentials()
    config_path = render_runtime_config(CONFIG_DIR / "hk_flow_reports.json", {
        "last_month_start": periods['current_month_to_now'][0].strftime('%Y-%m-%d'),
        "last_month_end": periods['yesterday'].strftime('%Y-%m-%d'),
        "yesterday": periods['yesterday'].strftime('%Y-%m-%d'),
    })

    # 清掉旧文件
    old_pattern = str(SAMPLE_DIR / "海外港澳商务_各渠道主辅投数据*.xlsx")
    import glob
    for f in glob.glob(old_pattern):
        try:
            Path(f).unlink()
            print(f"  [清理] {Path(f).name}")
        except:
            pass

    # 下载
    downloaded_file = run_smartbi_cli_task(config_path, "hk_flow_current", SAMPLE_DIR)

    # 直接覆盖标准文件名（不改名），供所有下游模块使用
    target = SAMPLE_DIR / "海外港澳商务_各渠道主辅投数据 (2).xlsx"
    if downloaded_file != target:
        import shutil
        shutil.move(str(downloaded_file), str(target))

    # 校验新鲜度
    verify_file_freshness(target, periods['yesterday'])

    print(f"  ✓ 港澳流速底表已更新至 {target.name}")
    return {"hk_flow_current": target}


def main():
    """单测：独立运行"""
    print("=" * 60)
    print("港澳流速 - 底表下载测试")
    print("=" * 60)

    try:
        periods = get_periods()
        print(f"目标日期: 当月 {periods['current_month_to_now'][0]} ~ {periods['yesterday']}")
        result = run_download(periods)
        print(f"\n✓ 下载成功: {result}")
    except Exception as e:
        print(f"\n❌ 错误: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
