"""
线下商超 - 依赖文件校验

商超模块依赖的文件：
1. 共享底表（由 download_hk_flow 负责下载）
2. 跨月快照（5 月完整 + 5 月至今，由书展 download_bookfair 触发的下载落到 书展数据复盘/sample/）

本脚本仅校验这些文件的存在性。
"""
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = Path(__file__).parent
BOOKFAIR_SAMPLE_DIR = BASE_DIR.parent / "书展数据复盘" / "sample"
HK_FLOW_DATA = BASE_DIR.parent / "港澳流速" / "sample" / "海外港澳商务_各渠道主辅投数据 (2).xlsx"

from utils.smartbi_helper import verify_file_freshness, get_periods


def run_download(periods=None):
    """
    商超依赖文件校验（不实际下载）

    参数:
        periods: 日期窗口字典（可选）

    返回:
        {"hk_flow": Path, "bookfair_snapshot": {full, to_now}}

    异常:
        FileNotFoundError: 任一文件缺失
    """
    if periods is None:
        periods = get_periods()

    print("[校验步骤] 线下商超模块依赖文件")

    # 共享底表
    if not HK_FLOW_DATA.exists():
        raise FileNotFoundError(f"  ❌ 港澳流速共享底表缺失: {HK_FLOW_DATA}")
    print(f"  ✓ 港澳流速共享底表存在")

    # 校验新鲜度
    try:
        verify_file_freshness(HK_FLOW_DATA, periods['yesterday'])
    except Exception as e:
        raise RuntimeError(f"  ❌ 港澳流速底表校验失败: {e}")

    # 跨月快照（由书展下载）
    yesterday_short = f"{periods['yesterday'].month}.{periods['yesterday'].day}"
    bookfair_full = BOOKFAIR_SAMPLE_DIR / "海外港澳商务_各渠道主辅投数据_5.1-5.31.xlsx"
    bookfair_to_now = BOOKFAIR_SAMPLE_DIR / f"海外港澳商务_各渠道主辅投数据_5.1-{yesterday_short}.xlsx"

    for file_path, desc in [
        (bookfair_full, "上月完整月快照"),
        (bookfair_to_now, "上月至今快照"),
    ]:
        if not file_path.exists():
            raise FileNotFoundError(f"  ❌ {desc} 缺失: {file_path}")
        print(f"  ✓ {desc} 存在")

    print(f"  ✓ 商超模块依赖文件校验完成")
    return {
        "hk_flow": HK_FLOW_DATA,
        "bookfair_snapshot": {
            "full": bookfair_full,
            "to_now": bookfair_to_now,
        }
    }


def main():
    """单测"""
    print("=" * 60)
    print("线下商超 - 文件校验测试")
    print("=" * 60)

    try:
        periods = get_periods()
        result = run_download(periods)
        print(f"\n✓ 校验成功: {result}")
    except Exception as e:
        print(f"\n❌ 错误: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
