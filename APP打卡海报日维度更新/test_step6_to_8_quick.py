"""
快速验证：只测试 Step 7-8（删除旧海报 + 添加新海报）
跳过 Step 5-6 登录，直接进入已登录状态
"""
import sys
from pathlib import Path
from playwright.sync_api import sync_playwright
import pandas as pd

sys.stdout.reconfigure(encoding="utf-8")

from poster_update import (
    step6_search_poster_group,
    step7_remove_business_group,
    step8_add_sorted_posters,
    BIZCENTER_URL,
    POSTER_GROUP_CODE,
)

SCRIPT_DIR = Path(__file__).parent
OUTPUT_DIR = SCRIPT_DIR / "output"

def main():
    print("=" * 100)
    print("快速验证：Step 6-8（查询 → 删除旧海报 → 添加新海报）")
    print("=" * 100)
    print()

    sorted_file = OUTPUT_DIR / "海报裂变率排序_20260611.xlsx"

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        try:
            # 假设已登录，直接导航到海报组编辑页面
            print("[准备] 导航到海报组编辑页面...")
            page.goto(BIZCENTER_URL)
            page.wait_for_load_state("load", timeout=10000)
            page.wait_for_timeout(2000)

            # Step 6
            print("[步骤 6] 查询海报组...")
            if not step6_search_poster_group(page):
                print("  ❌ Step 6 失败\n")
                return
            print("  ✓ Step 6 完成\n")

            # Step 7
            print("[步骤 7] 删除旧海报...")
            if not step7_remove_business_group(page):
                print("  ❌ Step 7 失败\n")
                return
            print("  ✓ Step 7 完成\n")

            # Step 8
            print("[步骤 8] 添加新海报...")
            if not step8_add_sorted_posters(page, sorted_file):
                print("  ❌ Step 8 失败\n")
                return
            print("  ✓ Step 8 完成\n")

            print("=" * 100)
            print("✓ Step 6-8 验证完成！")
            print("=" * 100)

        except Exception as e:
            print(f"\n❌ 错误: {e}")
            import traceback
            traceback.print_exc()
        finally:
            input("\n按 Enter 关闭浏览器...")
            browser.close()

if __name__ == "__main__":
    main()
