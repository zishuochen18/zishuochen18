"""
调试脚本：查找缺失海报的位置和原因

缺失的海报ID：1716, 1668, 1649, 1696, 1726
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
    print("=" * 80)
    print("调试脚本：查找缺失海报的位置")
    print("=" * 80)
    print()

    sorted_file = OUTPUT_DIR / "海报裂变率排序_20260611.xlsx"
    df = pd.read_excel(sorted_file, sheet_name='裂变率排序')
    poster_ids = [int(x) for x in df['海报ID-汇总处理'].tolist()]

    # 缺失的海报
    missing_ids = [1716, 1668, 1649, 1696, 1726]
    print(f"[准备] 要查找的缺失海报ID：{missing_ids}\n")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        try:
            if not step5_login_bizcenter(page):
                print("❌ Step 5 失败\n")
                return

            if not step6_search_poster_group(page):
                print("❌ Step 6 失败\n")
                return

            wrapper = page.locator(GROUP_WRAPPER_SELECTOR).first
            if not wrapper.is_visible():
                print("❌ 未找到分组 wrapper\n")
                input("\n按 Enter 关闭浏览器...")
                return

            print("✓ 登录和查询完成\n")

            # 打开弹窗
            print("[步骤] 打开【选择海报】弹窗...\n")
            select_btn = wrapper.locator(f'button:has-text("{SELECT_POSTER_BUTTON_TEXT}")').first
            select_btn.click()
            page.wait_for_timeout(1500)

            dialog = page.locator(".el-dialog:visible").first

            # 对每个缺失的海报进行扫描
            for missing_id in missing_ids:
                print(f"[查找] 查找海报 {missing_id}...")

                # 每次查找前都从第一页开始
                page.evaluate("window.scrollTo(0, 0)")
                page.wait_for_timeout(300)

                # 回到第一页
                pagination = dialog.locator(".el-pagination").first
                if pagination.is_visible():
                    # 找到第一页按钮
                    pag_buttons = pagination.locator("button").all()
                    # 通常第一个按钮是"上一页"，需要找到页码按钮或直接滑到第一页
                    page.evaluate("document.querySelector('.el-dialog').scrollTop = 0")
                    page.wait_for_timeout(300)

                found = False
                for page_num in range(1, 201):
                    # 获取当前页的所有行
                    rows = dialog.locator('tr.el-table__row').all()

                    # 在当前页查找
                    for idx, row in enumerate(rows):
                        text = row.text_content()
                        if str(missing_id) in text:
                            print(f"  ✓ 在第 {page_num} 页找到！位置在行 [{idx}]")
                            print(f"    完整文本: {text[:150]}")
                            found = True
                            break

                    if found:
                        break

                    # 翻到下一页
                    try:
                        pag_buttons = pagination.locator("button").all()
                        if len(pag_buttons) > 1:
                            next_btn = pag_buttons[1]
                            if next_btn.is_visible() and not next_btn.is_disabled():
                                next_btn.click()
                                page.wait_for_timeout(600)
                            else:
                                break
                        else:
                            break
                    except:
                        break

                if not found:
                    print(f"  ❌ 未找到海报 {missing_id}（扫描了全部页面）")

                print()

            print("=" * 80)
            print("调试完成")
            print("=" * 80)

        finally:
            input("\n按 Enter 关闭浏览器...")
            browser.close()

if __name__ == "__main__":
    main()
