"""
数据完整性检查脚本

用法：
    python check_data.py

功能：
    在生成周报前，检查所有 sample 目录下的数据底表是否齐全。
    若缺失文件，会列出缺失项并提示从何处获取。
"""
import sys
from datetime import datetime, timedelta
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = Path(__file__).parent

# 必备数据文件清单
REQUIRED_FILES = [
    {
        "module": "1. 港澳商务流速",
        "files": [
            {
                "path": "港澳流速/sample/海外港澳商务_各渠道主辅投数据.xlsx",
                "required": True,
                "source": "BI报表 → 海外直播业务线 → 海外商务 → 港澳商务",
                "freq": "每月",
            },
            {
                "path_glob": "港澳流速/sample/*港澳市场流速*.xlsx",
                "required": True,
                "source": "流速表（手工维护，含每日维度数据）",
                "freq": "每月",
            },
        ],
    },
    {
        "module": "2. 本月KOL转化数据",
        "files": [
            {
                "path": "本月KOL转化数据汇总/sample/海外港澳商务_各渠道主辅投数据-KOL汇总.xlsx",
                "required": True,
                "source": "KOL 汇总模板（手工维护）",
                "freq": "每月",
            },
            {
                "path": "本月KOL转化数据汇总/sample/本月KOL转化链路数据汇总.xlsx",
                "required": True,
                "source": "链路数据模板（手工维护）",
                "freq": "每月",
            },
        ],
    },
    {
        "module": "3. TMK做工周报",
        "files": [
            {
                "path": "TMK周报/sample/海外TMK做工监控.xlsx",
                "required": True,
                "source": "BI报表 → 海外直播业务线 → 海外前端 → TMK",
                "freq": "每周",
            },
            {
                "path": "TMK周报/sample/海外TMK未邀约做工监控播报.xlsx",
                "required": True,
                "source": "BI报表 → 海外直播业务线 → 海外前端 → TMK",
                "freq": "每周",
            },
            {
                "path": "TMK周报/sample/海外TMK做工勿扰情况汇总.xlsx",
                "required": True,
                "source": "BI报表 → 海外直播业务线 → 海外前端 → TMK",
                "freq": "每周",
            },
        ],
    },
    {
        "module": "4. 书展数据复盘",
        "files": [
            {
                "path": "书展内容/sample/书展数据汇总.xlsx",
                "required": True,
                "source": "书展汇总表（手工维护）",
                "freq": "每月",
            },
        ],
    },
    {
        "module": "5. 线下商超复盘",
        "files": [
            {
                "path": "线下商超内容/sample/线下商超内容汇总.xlsx",
                "required": True,
                "source": "商超汇总表（手工维护）",
                "freq": "每月",
            },
        ],
    },
    {
        "module": "6. 转介绍打卡",
        "files": [
            {
                "path": "转介绍打卡内容/sample/益智海外用户销售明细_末次渠道.xlsx",
                "required": True,
                "source": "BI报表 → 海外直播业务线 → 海外前端 → 益智_海外前端 → 转化漏斗",
                "freq": "每周",
            },
            {
                "path": "转介绍打卡内容/sample/后端转介绍流速.xlsx",
                "required": True,
                "source": "后端转介绍流速表（手工维护，按月分sheet）",
                "freq": "每月",
            },
            {
                "path_glob": "转介绍打卡内容/sample/*海外正式课学员带量明细*.xlsx",
                "required": True,
                "source": "BI报表 → 海外直播业务线 → 海外后端 → 思维-后端 → 转介绍 → 转介绍明细",
                "freq": "每周（本期+上期 2 份）",
                "min_count": 2,
            },
        ],
    },
]


def check_file(item):
    """检查单个文件项，返回 (是否通过, 状态描述, 实际文件列表)"""
    if "path" in item:
        full_path = BASE_DIR / item["path"]
        if full_path.exists():
            # 检查是否被占用
            try:
                with open(full_path, "rb") as f:
                    f.read(1)
                size_kb = full_path.stat().st_size / 1024
                return True, f"✓ ({size_kb:.0f} KB)", [full_path.name]
            except PermissionError:
                return False, "⚠️ 文件被占用（请关闭 Excel 后重试）", [full_path.name]
        else:
            return False, "✗ 文件缺失", []

    elif "path_glob" in item:
        # 通配符匹配
        glob_path = item["path_glob"]
        parent_glob = "/".join(glob_path.split("/")[:-1])
        pattern = glob_path.split("/")[-1]
        parent_dir = BASE_DIR / parent_glob

        if not parent_dir.exists():
            return False, f"✗ 目录不存在 ({parent_glob})", []

        files = sorted(
            [f for f in parent_dir.glob(pattern) if not f.name.startswith("~$")]
        )
        min_count = item.get("min_count", 1)
        if len(files) < min_count:
            return False, f"✗ 文件不足（需 {min_count} 个，找到 {len(files)} 个）", [f.name for f in files]

        return True, f"✓ 找到 {len(files)} 个", [f.name for f in files]

    return False, "✗ 配置错误", []


def main():
    print("=" * 70)
    print("港澳商务周报 - 数据完整性检查")
    print(f"检查时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 70)

    today = datetime.now().date()
    yesterday = today - timedelta(days=1)
    print(f"\n📅 周报统计周期：")
    print(f"   本期：{today.replace(day=1)} ~ {yesterday}")

    if today.month == 1:
        prev_month_year, prev_month = today.year - 1, 12
    else:
        prev_month_year, prev_month = today.year, today.month - 1
    prev_start = today.replace(year=prev_month_year, month=prev_month, day=1)
    print(f"   上期：{prev_start.strftime('%Y-%m')}-01 ~ {prev_start.strftime('%Y-%m')}-{yesterday.day:02d}")

    print()
    all_pass = True
    missing_count = 0
    locked_count = 0

    for module_block in REQUIRED_FILES:
        print(f"\n📂 {module_block['module']}")
        print("-" * 70)

        for item in module_block["files"]:
            path_show = item.get("path") or item.get("path_glob")
            ok, status, files = check_file(item)

            if not ok:
                all_pass = False
                if "占用" in status:
                    locked_count += 1
                else:
                    missing_count += 1

            print(f"  {status:<32} {path_show}")
            if not ok:
                print(f"    ↳ 来源：{item['source']}")
                print(f"    ↳ 频率：{item['freq']}")
            elif files and len(files) > 1:
                for f in files:
                    print(f"      · {f}")

    print("\n" + "=" * 70)
    if all_pass:
        print("✅ 全部通过！可以运行 python generate_weekly_report.py")
    else:
        print(f"❌ 检查未通过 - 缺失 {missing_count} 项, 被占用 {locked_count} 项")
        if locked_count > 0:
            print("\n💡 提示：关闭 Excel 后再运行")
    print("=" * 70)

    return 0 if all_pass else 1


if __name__ == "__main__":
    sys.exit(main())
