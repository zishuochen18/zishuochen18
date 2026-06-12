"""
测试 Step 5-6：登录平台中心 + 查询海报组
"""
import sys
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding="utf-8")

# 配置
BIZCENTER_URL = "https://bizcenter-h5-cms.61info.cn/#/groupPoster"
POSTER_GROUP_CODE = "siweidaka"
OUTPUT_DIR = Path("output")

# 导入步骤函数
from poster_update import (
    step5_login_bizcenter,
    step6_search_poster_group
)

def main():
    print("=" * 60)
    print("APP打卡海报日维度更新 - Step 5-6 测试")
    print("=" * 60)
    print()

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

            # Step 6: 查询海报组
            print("[步骤 6] 查询海报组...")
            if step6_search_poster_group(page):
                print("  ✓ 海报组查询完成\n")
            else:
                print("  ❌ 查询失败\n")

            print("=" * 60)
            print("✓ Step 5-6 测试完成！")
            print("=" * 60)

            input("\n按 Enter 关闭浏览器...")

        except Exception as e:
            print(f"\n❌ 错误: {e}")
            import traceback
            traceback.print_exc()
        finally:
            context.close()
            browser.close()

if __name__ == "__main__":
    main()

