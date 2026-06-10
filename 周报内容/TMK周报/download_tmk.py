"""
TMK 周报 - 自动下载 3 个 SmartBI 报表

下载配置已存在于 configs/tmk_reports.json，包含 3 个 task：
- tmk_work_monitor
- tmk_unscheduled_monitor
- tmk_dnd_summary
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
    rename_downloaded_file,
    assert_smartbi_credentials,
)


def run_download(periods=None):
    """
    下载 TMK 3 个报表

    参数:
        periods: 日期窗口字典（可选）

    返回:
        {
            "tmk_work_monitor": Path,
            "tmk_unscheduled_monitor": Path,
            "tmk_dnd_summary": Path,
        }

    异常:
        RuntimeError / FileNotFoundError
    """
    if periods is None:
        periods = get_periods()

    print("[下载步骤] TMK 周报（3 个报表）")

    assert_smartbi_credentials()

    # TMK 配置不含日期占位符（固定的渠道筛选），但仍走统一渲染以保持接口一致
    config_path = render_runtime_config(CONFIG_DIR / "tmk_reports.json")

    # 清掉旧文件
    import glob
    old_patterns = [
        SAMPLE_DIR / "海外TMK*.xlsx",
    ]
    for pattern in old_patterns:
        for f in glob.glob(str(pattern)):
            try:
                Path(f).unlink()
                print(f"  [清理] {Path(f).name}")
            except:
                pass

    # 下载三个 task
    tasks = [
        ("tmk_work_monitor", "海外TMK做工监控.xlsx"),
        ("tmk_unscheduled_monitor", "海外TMK未邀约做工监控播报.xlsx"),
        ("tmk_dnd_summary", "海外TMK做工勿扰情况汇总.xlsx"),
    ]

    result = {}
    for task_key, target_name in tasks:
        print(f"  [下载] {task_key}...")
        downloaded_file = run_smartbi_cli_task(config_path, task_key, SAMPLE_DIR)
        target = rename_downloaded_file(downloaded_file, target_name)
        result[task_key] = target

    print(f"  ✓ TMK 3 个报表下载完成")
    return result


def main():
    """单测"""
    print("=" * 60)
    print("TMK 周报 - 下载测试")
    print("=" * 60)

    try:
        periods = get_periods()
        result = run_download(periods)
        print(f"\n✓ 下载成功: {result}")
    except Exception as e:
        print(f"\n❌ 错误: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
