"""
调试脚本：通过点击 <label> 来勾选checkbox

重点：确认点击 label 能成功勾选
"""
import sys
from pathlib import Path
from playwright.sync_api import sync_playwright
import pandas as pd

sys.stdout.reconfigure(encoding="utf-8")

from poster_update import (
    step5_login_bizcenter,
    step6_search_poster_group,
    GROUP_WRAPPER_SELECTOR,
    SELECT_POSTER_BUTTON_TEXT,
)

SCRIPT_DIR = Path(__file__).parent
OUTPUT_DIR = SCRIPT_DIR / "output"

def main():
    sorted_file = OUTPUT_DIR / "海报裂变率排序_20260611.xlsx"
    df = pd.read_excel(sorted_file, sheet_name='裂变率排序')
    poster_ids = [int(x) for x in df['海报ID-汇总处理'].tolist()]
    test_id = poster_ids[0]

    print(f"测试点击 <label> 来勾选海报 {test_id}")
    print()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        try:
            if not step5_login_bizcenter(page) or not step6_search_poster_group(page):
                input("\n按 Enter 关闭浏览器...")
                return

            wrapper = page.locator(GROUP_WRAPPER_SELECTOR).first
            if not wrapper.is_visible():
                input("\n按 Enter 关闭浏览器...")
                return

            # 打开弹窗
            select_btn = wrapper.locator(f'button:has-text("{SELECT_POSTER_BUTTON_TEXT}")').first
            select_btn.click()
            page.wait_for_timeout(1500)

            dialog = page.locator(".el-dialog:visible").first

            # 翻到第 5 页
            for i in range(4):
                pagination = dialog.locator(".el-pagination").first
                pag_buttons = pagination.locator("button").all()
                pag_buttons[1].click()
                page.wait_for_timeout(600)

            # 找海报1758
            rows = dialog.locator('tr.el-table__row').all()
            target_row = None
            for row in rows:
                if "1758" in row.text_content():
                    target_row = row
                    break

            if not target_row:
                print("❌ 未找到海报1758")
                input("\n按 Enter 关闭浏览器...")
                return

            print("✓ 找到海报1758\n")

            # 获取checkbox的前置状态
            checkbox = target_row.locator('input[type="checkbox"]').first
            old_checked = checkbox.evaluate("el => el.checked")
            print(f"勾选前状态: {old_checked}")

            # 尝试点击 <label> 元素
            print("\n[方案 1] 点击 <label class='el-checkbox'> 元素...")
            label = target_row.locator('label.el-checkbox').first
            if label.is_visible():
                print("  label 可见")
                label.click()
            else:
                print("  label 不可见，尝试 force 点击...")
                label.click(force=True)

            page.wait_for_timeout(500)

            new_checked = checkbox.evaluate("el => el.checked")
            print(f"勾选后状态: {new_checked}")

            if new_checked != old_checked:
                print(f"\n✓ 成功！状态已从 {old_checked} 变为 {new_checked}")
            else:
                print(f"\n❌ 状态未改变")

            # 再点一下试试看是否能取消
            print("\n[测试取消] 再点击一次...")
            label.click(force=True)
            page.wait_for_timeout(500)
            final_checked = checkbox.evaluate("el => el.checked")
            print(f"第二次点击后状态: {final_checked}")

        finally:
            input("\n按 Enter 关闭浏览器...")
            browser.close()

if __name__ == "__main__":
    main()
