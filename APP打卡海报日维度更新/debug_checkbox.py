"""
调试脚本：精确找出为什么 :has-text() 不工作，并测试勾选checkbox

重点：找到正确的选择器和checkbox勾选方式
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
    print("调试脚本：测试selector和checkbox")
    print("=" * 80)
    print()

    sorted_file = OUTPUT_DIR / "海报裂变率排序_20260611.xlsx"
    df = pd.read_excel(sorted_file, sheet_name='裂变率排序')
    poster_ids = [int(x) for x in df['海报ID-汇总处理'].tolist()]
    test_id = poster_ids[0]  # 1758
    print(f"[准备] 要查找的海报ID：{test_id}\n")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        try:
            if not step5_login_bizcenter(page):
                return
            if not step6_search_poster_group(page):
                return

            wrapper = page.locator(GROUP_WRAPPER_SELECTOR).first
            if not wrapper.is_visible():
                input("\n按 Enter 关闭浏览器...")
                return

            print("[步骤 1] 打开弹窗...")
            select_btn = wrapper.locator(f'button:has-text("{SELECT_POSTER_BUTTON_TEXT}")').first
            select_btn.click()
            page.wait_for_timeout(1500)

            dialog = page.locator(".el-dialog:visible").first
            if not dialog.is_visible():
                input("\n按 Enter 关闭浏览器...")
                return

            print("✓ 弹窗已出现\n")

            # 翻到第 5 页找海报1758
            print("[步骤 2] 翻页到第 5 页...")
            for i in range(4):  # 需要翻 4 次才能到第 5 页
                pagination = dialog.locator(".el-pagination").first
                pag_buttons = pagination.locator("button").all()
                next_btn = pag_buttons[1]
                next_btn.click()
                page.wait_for_timeout(600)
            print("✓ 已到第 5 页\n")

            # 现在查找海报1758
            print("[步骤 3] 查找海报 1758 的行...")
            rows = dialog.locator('tr.el-table__row').all()
            print(f"当前页行数: {len(rows)}")
            print()

            target_row = None
            target_idx = -1
            for idx, row in enumerate(rows):
                text = row.text_content()
                if "1758" in text:
                    target_row = row
                    target_idx = idx
                    print(f"✓ 在行 [{idx}] 找到！")
                    print(f"  完整文本: {text[:150]}\n")
                    break

            if not target_row:
                print("❌ 未找到\n")
                input("\n按 Enter 关闭浏览器...")
                return

            # 检查该行的HTML结构
            print("[步骤 4] 分析该行的 HTML 结构...")
            html = target_row.evaluate("el => el.outerHTML")
            print(f"HTML 长度: {len(html)} 字符")
            print(f"HTML 前 300 字:\n{html[:300]}\n")

            # 检查该行内的 checkbox
            print("[步骤 5] 查找该行内的 checkbox...")
            checkboxes = target_row.locator('input[type="checkbox"]').all()
            print(f"找到 {len(checkboxes)} 个 checkbox")

            if checkboxes:
                checkbox = checkboxes[0]
                checked = checkbox.evaluate("el => el.checked")
                print(f"  checkbox状态 (checked): {checked}")
                print(f"  checkbox是否可见: {checkbox.is_visible()}")
                print(f"  checkbox是否禁用: {checkbox.is_disabled()}")

                # 尝试点击checkbox
                print("\n[步骤 6] 点击checkbox...")
                try:
                    checkbox.click()
                    page.wait_for_timeout(300)
                    new_checked = checkbox.evaluate("el => el.checked")
                    print(f"✓ 点击成功！新状态: {new_checked}")
                except Exception as e:
                    print(f"❌ 点击失败: {e}")

                # 检查附近是否有其他可点击元素（比如单元格本身）
                print("\n[步骤 7] 检查行内的其他可点击元素...")
                cells = target_row.locator('td').all()
                print(f"找到 {len(cells)} 个 <td>")
                if cells:
                    first_cell = cells[0]
                    cell_html = first_cell.evaluate("el => el.outerHTML")
                    print(f"第一个<td> HTML 前 200 字:\n{cell_html[:200]}\n")

                    # 检查第一个cell是否包含checkbox
                    first_cell_checkbox = first_cell.locator('input[type="checkbox"]').first
                    if first_cell_checkbox.is_visible():
                        print("✓ 第一个<td>内包含checkbox")
                        print(f"  是否被checked: {first_cell_checkbox.evaluate('el => el.checked')}")

            else:
                print("❌ 未找到 checkbox！检查是否使用了其他的勾选方式...\n")

                # 尝试找其他可能的勾选元素
                print("[步骤 6b] 查找其他勾选元素...")
                icons = target_row.locator('i.el-icon').all()
                print(f"  找到 {len(icons)} 个 <i class='el-icon'> 元素")

                spans = target_row.locator('span').all()
                print(f"  找到 {len(spans)} 个 <span>")

                buttons = target_row.locator('button').all()
                print(f"  找到 {len(buttons)} 个 <button>")

        finally:
            input("\n按 Enter 关闭浏览器...")
            browser.close()

if __name__ == "__main__":
    main()
