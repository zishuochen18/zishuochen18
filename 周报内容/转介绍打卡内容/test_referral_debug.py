"""
转介绍打卡内容 - 独立调试脚本
测试 extract_referral_data() 和 extract_punch_card_data() 函数
"""
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = Path(__file__).parent
sys.path.insert(0, str(BASE_DIR.parent))

from 转介绍打卡内容 import process_referral

print("=" * 60)
print("转介绍打卡内容 - 独立调试")
print("=" * 60)

# 测试下载是否完成
sample_dir = BASE_DIR / "sample"
required_files = {
    "销售明细": sample_dir / "益智海外用户销售明细_末次渠道.xlsx",
    "后端流速": sample_dir / "后端转介绍流速.xlsx",
    "打卡内容": sample_dir / "转介绍打卡内容.xlsx",
}

print("\n[检查文件]")
for label, fpath in required_files.items():
    exists = fpath.exists()
    status = "✓" if exists else "✗"
    print(f"  {status} {label}: {fpath.name}")

# 查找 student_带量 文件
print("\n[检查 student_带量 手工文件]")
punch_files = list(sample_dir.glob("*海外正式课学员带量明细*.xlsx"))
punch_files = [f for f in punch_files if not f.name.startswith("~$")]
punch_files = sorted(punch_files, key=lambda x: x.name)
if punch_files:
    for f in punch_files:
        print(f"  ✓ {f.name}")
else:
    print("  ✗ 未找到 student_带量 文件")

# 测试数据提取
print("\n" + "=" * 60)
print("[6.1] 后端非手推达成情况")
print("=" * 60)

try:
    referral_data = process_referral.extract_referral_data()

    print(f"\n本期: {referral_data['cur_period']}")
    print(f"{'类型':<12} {'例子数':>6} {'目标':>6} {'达成率':>8} {'约课':>5} {'到课':>5} {'成单':>5} {'GMV':>10}")
    for r in referral_data["cur_rows"]:
        print(f"{r['name']:<12} {r['examples']:>6} {r['target']:>6} {r['mtd_rate']*100:>7.2f}% {r['bookings']:>5} {r['attendances']:>5} {r['signups']:>5} {r['gmv']:>10.1f}")

    print(f"\n上期: {referral_data['prev_period']}")
    print(f"{'类型':<12} {'例子数':>6} {'目标':>6} {'达成率':>8} {'约课':>5} {'到课':>5} {'成单':>5} {'GMV':>10}")
    for r in referral_data["prev_rows"]:
        print(f"{r['name']:<12} {r['examples']:>6} {r['target']:>6} {r['mtd_rate']*100:>7.2f}% {r['bookings']:>5} {r['attendances']:>5} {r['signups']:>5} {r['gmv']:>10.1f}")

    print(f"\n✓ 后端非手推达成数据提取成功")
except Exception as e:
    print(f"\n✗ 后端非手推达成数据提取失败: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 60)
print("[6.2] 打卡链路数据")
print("=" * 60)

try:
    punch_data = process_referral.extract_punch_card_data()

    print(f"\n本期: {punch_data['cur_period']}")
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
    print(f"  转化数: {c['signups']}, 到课转化率: {c['conv_rate']*100:.2f}%")
    print(f"  注册转化率: {c['total_conv_rate']*100:.2f}%")
    print(f"  GMV: {c['gmv']:.1f}, ASP: {c['asp']:.2f}")

    print(f"\n上期: {punch_data['prev_period']}")
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
    print(f"  转化数: {p['signups']}, 到课转化率: {p['conv_rate']*100:.2f}%")
    print(f"  注册转化率: {p['total_conv_rate']*100:.2f}%")
    print(f"  GMV: {p['gmv']:.1f}, ASP: {p['asp']:.2f}")

    print(f"\n✓ 打卡链路数据提取成功")
except Exception as e:
    print(f"\n✗ 打卡链路数据提取失败: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 60)
print("[完成] 转介绍打卡内容调试")
print("=" * 60)
