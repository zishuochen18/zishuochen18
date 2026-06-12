"""
调试脚本：精确测试如何在弹窗中用 Playwright 选择器找到指定海报ID

重点：验证 selector 的有效性，找出最稳定的查找方式
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
    POSTER_CARD_SELECTOR,
    SELECT_POSTER_BUTTON_TEXT,
)

SCRIPT_DIR = Path(__file__).parent
OUTPUT_DIR = SCRIPT_DIR / "output"

def main():
    print("=" * 80)
    print("调试脚本：测试【选择海报】弹窗中的海报ID选择器")
    print("=" * 80)
    print()

    # 读 Excel 获取第一个海报ID
    sorted_file = OUTPUT_DIR / "海报裂变率排序_20260611.xlsx"
    df = pd.read_excel(sorted_file, sheet_name='裂变率排序')
    poster_ids = [int(x) for x in df['海报ID-汇总处理'].tolist()]
    test_id = poster_ids[0]  # 第一个海报: 1758
    print(f"[准备] 要查找的海报ID：{test_id}\n")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        try:
            # Step 5: 登录
            if not step5_login_bizcenter(page):
                print("❌ 登录失败\n")
                return

            # Step 6: 查询
            if not step6_search_poster_group(page):
                print("❌ 查询失败\n")
                return

            # 定位 wrapper
            wrapper = page.locator(GROUP_WRAPPER_SELECTOR).first
            if not wrapper.is_visible():
                print("❌ 未找到分组 wrapper\n")
                input("\n按 Enter 关闭浏览器...")
                return

            # 点击【选择海报】
            print("[步骤 1] 打开【选择海报】弹窗...")
            select_btn = wrapper.locator(f'button:has-text("{SELECT_POSTER_BUTTON_TEXT}")').first
            select_btn.click()
            page.wait_for_timeout(1500)

            dialog = page.locator(".el-dialog:visible").first
            if not dialog.is_visible():
                print("❌ 弹窗未出现\n")
                input("\n按 Enter 关闭浏览器...")
                return

            print("✓ 弹窗已出现\n")

            # ===== 开始测试不同的选择器 =====
            print("[步骤 2] 测试不同的选择器查找海报...")
            print()

            # 方案 1: 直接 :has-text()
            print(f"[方案 1] 使用 :has-text() 选择器")
            print(f"  selector: tr.el-table__row:has-text(\"{test_id}\")")
            try:
                card = dialog.locator(f'tr.el-table__row:has-text("{test_id}")').first
                if card.is_visible(timeout=2000):
                    text = card.text_content().strip()[:100]
                    print(f"  ✓ 找到！text: {text}")
                else:
                    print(f"  ❌ 定位器未返回可见元素")
            except Exception as e:
                print(f"  ❌ 异常: {e}")

            print()

            # 方案 2: 遍历所有行，用 text_content() 检查
            print(f"[方案 2] 遍历所有行，用 text_content() 检查")
            all_rows = dialog.locator('tr.el-table__row').all()
            print(f"  当前页行数: {len(all_rows)}")
            found_idx = -1
            for idx, row in enumerate(all_rows):
                text = row.text_content()
                if str(test_id) in text:
                    print(f"  ✓ 在行 [{idx}] 找到！text[:100]: {text[:100]}")
                    found_idx = idx
                    break
            if found_idx == -1:
                print(f"  ❌ 未在当前页找到")

            print()

            # 方案 3: 用 filter() 方法
            print(f"[方案 3] 使用 filter() 方法")
            print(f"  locator: tr.el-table__row filtered by has(:has-text())")
            try:
                filtered = dialog.locator(f'tr.el-table__row').filter(has_text=str(test_id)).first
                if filtered.is_visible(timeout=2000):
                    text = filtered.text_content().strip()[:100]
                    print(f"  ✓ 找到！text: {text}")
                else:
                    print(f"  ❌ 定位器未返回可见元素")
            except Exception as e:
                print(f"  ❌ 异常: {e}")

            print()

            # 方案 4: 检查翻页器并尝试翻页
            print(f"[方案 4] 检查当前页数，尝试翻页")
            pagination = dialog.locator(".el-pagination").first
            if pagination.is_visible():
                # 从翻页器中获取当前页数
                total_text = pagination.locator(".el-pagination__total").first
                if total_text:
                    print(f"  翻页器信息: {total_text.text_content()}")

                # 尝试找到下一页按钮并点击
                pag_buttons = pagination.locator("button").all()
                print(f"  翻页器按钮数: {len(pag_buttons)}")
                if len(pag_buttons) > 1:
                    next_btn = pag_buttons[1]
                    if not next_btn.is_disabled():
                        print(f"  下一页按钮可用，点击...")
                        next_btn.click()
                        page.wait_for_timeout(800)

                        # 再次尝试查找
                        print(f"  翻页后，再试一次...")
                        all_rows_page2 = dialog.locator('tr.el-table__row').all()
                        print(f"  第 2 页行数: {len(all_rows_page2)}")
                        for idx, row in enumerate(all_rows_page2):
                            text = row.text_content()
                            if str(test_id) in text:
                                print(f"  ✓ 在第 2 页行 [{idx}] 找到！text[:100]: {text[:100]}")
                                break
                    else:
                        print(f"  下一页按钮已禁用（可能只有 1 页）")
            else:
                print(f"  ❌ 翻页器不可见")

            print()
            print("=" * 80)
            print("测试完成")
            print("=" * 80)

        finally:
            input("\n按 Enter 关闭浏览器...")
            browser.close()

if __name__ == "__main__":
    main()
