"""
KOL 模块 - 文件存在性校验

KOL 模块依赖的三个文件：
1. 海外港澳商务_各渠道主辅投数据-KOL汇总.xlsx（模板，人工维护）
2. 本月KOL转化链路数据汇总.xlsx（模板，人工维护）
3. ../港澳流速/sample/海外港澳商务_各渠道主辅投数据 (2).xlsx（由 download_hk_flow 负责）

本脚本仅校验这三个文件的存在性和新鲜度。无需从 SmartBI 下载。
"""
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = Path(__file__).parent
SAMPLE_DIR = BASE_DIR / "sample"
HK_FLOW_DATA = BASE_DIR.parent / "港澳流速" / "sample" / "海外港澳商务_各渠道主辅投数据 (2).xlsx"

from utils.smartbi_helper import verify_file_freshness, get_periods


def run_download(periods=None):
    """
    KOL 文件校验（不实际下载）

    参数:
        periods: 日期窗口字典（可选）

    返回:
        {"kol_template": Path, "conversion_template": Path, "hk_flow": Path}

    异常:
        FileNotFoundError: 任一文件缺失
    """
    if periods is None:
        periods = get_periods()

    print("[校验步骤] KOL 模块依赖文件")

    # 三个必需文件
    kol_template = SAMPLE_DIR / "海外港澳商务_各渠道主辅投数据-KOL汇总.xlsx"
    conversion_template = SAMPLE_DIR / "本月KOL转化链路数据汇总.xlsx"

    files_to_check = [
        (kol_template, "KOL 汇总模板"),
        (conversion_template, "KOL 转化链路模板"),
        (HK_FLOW_DATA, "港澳流速共享底表"),
    ]

    for file_path, desc in files_to_check:
        if not file_path.exists():
            raise FileNotFoundError(f"  ❌ {desc} 缺失: {file_path}")
        print(f"  ✓ {desc} 存在")

    # 校验共享底表新鲜度
    try:
        verify_file_freshness(HK_FLOW_DATA, periods['yesterday'])
    except Exception as e:
        raise RuntimeError(f"  ❌ 港澳流速底表校验失败: {e}")

    print(f"  ✓ KOL 模块文件校验完成")
    return {
        "kol_template": kol_template,
        "conversion_template": conversion_template,
        "hk_flow": HK_FLOW_DATA,
    }


def main():
    """单测"""
    print("=" * 60)
    print("KOL 模块 - 文件校验测试")
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
