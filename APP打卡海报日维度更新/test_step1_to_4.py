"""测试 Step 1-4 流程：Sensors 登录 → 进入书签 → 选日期导出 → 计算排序"""

import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent
sys.path.insert(0, str(SCRIPT_DIR))

from playwright.sync_api import sync_playwright
from poster_update import (
    step1_login_sensors,
    step2_enter_bookmark,
    step3_export_data,
    step4_calculate_conversion_rate,
    SAMPLE_DIR
)


def test_steps_1_to_4():
    SAMPLE_DIR.mkdir(exist_ok=True)

    print("\n" + "="*70)
    print("APP打卡海报日维度更新 - Step 1-4 测试")
    print("="*70)
    print("\n流程：")
    print("  Step 1: 登录 Sensors（自动填写账号密码）")
    print("  Step 2: 进入书签数据表")
    print("  Step 3: 选择日期【过去 14 天】+ 等待 5 秒缓存 + 导出数据")
    print("  Step 4: 计算转换率并排序")
    print("="*70 + "\n")

    import time
    time.sleep(1)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        try:
            # Step 1
            print("[Step 1] 登录 Sensors...")
            result1 = step1_login_sensors(page)
            print(f"  {'✓' if result1 else '✗'} Step 1: {'成功' if result1 else '失败'}")
            if not result1:
                return False

            # Step 2
            print("\n[Step 2] 进入书签数据表...")
            result2 = step2_enter_bookmark(page)
            print(f"  {'✓' if result2 else '✗'} Step 2: {'成功' if result2 else '失败'}")
            if not result2:
                return False

            # Step 3
            print("\n[Step 3] 选择日期【过去 14 天】并导出数据...")
            result3 = step3_export_data(page, context)
            print(f"  {'✓' if result3 else '✗'} Step 3: {'成功' if result3 else '失败'}")
            if not result3:
                return False

            # Step 4
            print("\n[Step 4] 计算转换率并排序...")
            excel_files = list(SAMPLE_DIR.glob("*SensorsAnalytics.xlsx"))
            if excel_files:
                latest_excel = max(excel_files, key=lambda p: p.stat().st_mtime)
                result4 = step4_calculate_conversion_rate(latest_excel)
                print(f"  {'✓' if result4 else '✗'} Step 4: {'成功' if result4 else '失败'}")
            else:
                print("  ✗ 未找到导出的 Excel 文件")
                result4 = False

            # 总结
            print("\n" + "="*70)
            if result1 and result2 and result3 and result4:
                print("✓ Step 1-4 全部成功！")
                print("="*70)
                return True
            else:
                print("✗ 流程中存在失败步骤")
                print("="*70)
                return False

        except KeyboardInterrupt:
            print("\n\n用户中断流程")
            return False
        except Exception as e:
            print(f"\n✗ 发生异常: {str(e)[:200]}")
            import traceback
            traceback.print_exc()
            return False
        finally:
            browser.close()


if __name__ == "__main__":
    success = test_steps_1_to_4()
    sys.exit(0 if success else 1)
