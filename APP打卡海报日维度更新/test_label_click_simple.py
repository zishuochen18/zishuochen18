"""
简化版调试：只测试label点击
"""
import sys
from pathlib import Path
from playwright.sync_api import sync_playwright
import pandas as pd
import time

sys.stdout.reconfigure(encoding="utf-8")

from poster_update import (
    step5_login_bizcenter,
    step6_search_poster_group,
    GROUP_WRAPPER_SELECTOR,
    SELECT_POSTER_BUTTON_TEXT,
)

SCRIPT_DIR = Path(__file__).parent
OUTPUT_DIR = SCRIPT_DIR / "output"

sorted_file = OUTPUT_DIR / "海报裂变率排序_20260611.xlsx"
df = pd.read_excel(sorted_file, sheet_name='裂变率排序')
poster_ids = [int(x) for x in df['海报ID-汇总处理'].tolist()]
test_id = poster_ids[0]

print(f"测试点击 <label> 来勾选海报 {test_id}\n")

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    context = browser.new_context(accept_downloads=True)
    page = context.new_page()

    try:
        # Step 5
        if step5_login_bizcenter(page):
            print("✓ Step 5 完成\n")
        else:
            print("❌ Step 5 失败\n")
            raise Exception("Login failed")

        # Step 6
        if step6_search_poster_group(page):
            print("✓ Step 6 完成\n")
        else:
            print("❌ Step 6 失败\n")
            raise Exception("Search failed")

        # 打开弹窗
        wrapper = page.locator(GROUP_WRAPPER_SELECTOR).first
        select_btn = wrapper.locator(f'button:has-text("{SELECT_POSTER_BUTTON_TEXT}")').first
        select_btn.click()
        page.wait_for_timeout(1500)

        dialog = page.locator(".el-dialog:visible").first

        # 翻到第 5 页
        print("[翻页] 前往第 5 页...")
        for i in range(4):
            pagination = dialog.locator(".el-pagination").first
            pag_buttons = pagination.locator("button").all()
            pag_buttons[1].click()
            page.wait_for_timeout(600)
        print("✓ 已到第 5 页\n")

        # 找海报1758
        print("[查找] 查找海报 1758...")
        rows = dialog.locator('tr.el-table__row').all()
        target_row = None
        for row in rows:
            if "1758" in row.text_content():
                target_row = row
                break

        if not target_row:
            print("❌ 未找到")
            raise Exception("Poster not found")
        print("✓ 找到\n")

        # 获取checkbox前置状态
        checkbox = target_row.locator('input[type="checkbox"]').first
        old_checked = checkbox.evaluate("el => el.checked")
        print(f"勾选前: {old_checked}")

        # 点击label
        print("[点击] 点击 <label>...")
        label = target_row.locator('label.el-checkbox').first
        label.click(force=True)
        page.wait_for_timeout(500)

        new_checked = checkbox.evaluate("el => el.checked")
        print(f"勾选后: {new_checked}\n")

        if new_checked != old_checked:
            print(f"✓ 成功！状态改变：{old_checked} → {new_checked}")
        else:
            print(f"❌ 失败！状态未改变")

    finally:
        time.sleep(2)
        browser.close()
