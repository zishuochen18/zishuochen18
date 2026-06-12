"""
测试 Step 5-9：平台中心操作
直接使用 output 中的排序文件，不需要重新运行 Step 1-4
"""
import sys
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding="utf-8")

# 配置
BIZCENTER_URL = "https://bizcenter-h5-cms.61info.cn/#/groupPoster"
POSTER_GROUP_CODE = "siweidaka"
BUSINESS_GROUP_NAME = "海外益智海报-非台湾"
OUTPUT_DIR = Path("output")

# 导入步骤函数
from poster_update import (
    step5_login_bizcenter,
    step6_search_poster_group,
    step7_remove_business_group,
    step8_add_sorted_posters,
    step9_save
)

def main():
    print("=" * 60)
    print("APP打卡海报日维度更新 - Step 5-9 测试")
    print("=" * 60)
    print()

    # 查找排序文件
    sorted_files = list(OUTPUT_DIR.glob("*裂变率排序*.xlsx"))
    if not sorted_files:
        print("❌ 未找到排序文件，请先运行 Step 1-4")
        return

    sorted_file = sorted_files[0]
    print(f"[准备] 使用排序文件: {sorted_file.name}\n")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        try:
            # Step 5: 登录平台中心
            print("[步骤 5] 登录平台中心...")
            if step5_login_bizcenter(page):
                print("  ✓ 平台中心已登录\n")
            else:
                print("  ❌ 登录失败\n")
                return

            input("按 Enter 继续到第六步...\n")

            # Step 6: 查询海报组
            print("[步骤 6] 查询海报组...")
            if step6_search_poster_group(page):
                print("  ✓ 海报组查询完成\n")
            else:
                print("  ❌ 查询失败\n")

            input("按 Enter 继续到第七步...\n")

            # Step 7: 移除旧业务分组
            print("[步骤 7] 移除旧业务分组...")
            if step7_remove_business_group(page):
                print("  ✓ 业务分组已移除\n")
            else:
                print("  ❌ 移除失败\n")

            input("按 Enter 继续到第八步...\n")

            # Step 8: 输入新海报ID
            print("[步骤 8] 输入新海报ID...")
            if step8_add_sorted_posters(page, sorted_file):
                print("  ✓ 海报ID已输入\n")
            else:
                print("  ❌ 输入失败\n")

            input("按 Enter 继续到第九步...\n")

            # Step 9: 保存更新
            print("[步骤 9] 保存更新...")
            if step9_save(page):
                print("  ✓ 更新已保存\n")
            else:
                print("  ❌ 保存失败\n")

            print("✓ Step 5-9 测试完成！")

        except Exception as e:
            print(f"\n❌ 错误: {e}")
            import traceback
            traceback.print_exc()
        finally:
            context.close()
            browser.close()

if __name__ == "__main__":
    main()
