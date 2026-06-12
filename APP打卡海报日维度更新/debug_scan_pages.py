"""
调试脚本：扫描翻页找到海报ID的正确位置

重点：找出海报1758实际在第几页
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
    print("调试脚本：扫描弹窗，找到海报ID 1758 的位置")
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
                print("❌ 登录失败\n")
                return

            if not step6_search_poster_group(page):
                print("❌ 查询失败\n")
                return

            wrapper = page.locator(GROUP_WRAPPER_SELECTOR).first
            if not wrapper.is_visible():
                print("❌ 未找到分组 wrapper\n")
                input("\n按 Enter 关闭浏览器...")
                return

            print("[步骤] 打开【选择海报】弹窗...")
            select_btn = wrapper.locator(f'button:has-text("{SELECT_POSTER_BUTTON_TEXT}")').first
            select_btn.click()
            page.wait_for_timeout(1500)

            dialog = page.locator(".el-dialog:visible").first
            if not dialog.is_visible():
                print("❌ 弹窗未出现\n")
                input("\n按 Enter 关闭浏览器...")
                return

            print("✓ 弹窗已出现\n")
            print(f"[扫描] 翻页查找海报ID {test_id}...\n")

            # 扫描前 20 页
            for page_num in range(1, 21):
                rows = dialog.locator('tr.el-table__row').all()
                print(f"[第 {page_num} 页] {len(rows)} 行 ", end="")

                # 打印这一页的所有海报ID
                found_ids = []
                for row in rows:
                    text = row.text_content()
                    # 尝试从文本中提取海报ID（通常在最前面）
                    parts = text.split()
                    if parts and parts[0].isdigit():
                        found_ids.append(parts[0])

                print(f"包含ID: {', '.join(found_ids[:5])}", end="")

                # 检查是否有目标ID
                for row in rows:
                    text = row.text_content()
                    if str(test_id) in text:
                        print(f" ✓ 找到 {test_id}！")
                        print(f"\n[成功] 海报 {test_id} 在第 {page_num} 页")
                        print(f"完整文本: {text[:150]}")
                        input("\n按 Enter 关闭浏览器...")
                        return

                print()

                # 翻到下一页
                pagination = dialog.locator(".el-pagination").first
                if pagination.is_visible():
                    pag_buttons = pagination.locator("button").all()
                    if len(pag_buttons) > 1:
                        next_btn = pag_buttons[1]
                        if next_btn.is_visible() and not next_btn.is_disabled():
                            next_btn.click()
                            page.wait_for_timeout(600)
                        else:
                            print(f"\n[结束] 已到最后一页，海报 {test_id} 未找到")
                            input("\n按 Enter 关闭浏览器...")
                            return
                    else:
                        print(f"\n[结束] 无翻页按钮")
                        input("\n按 Enter 关闭浏览器...")
                        return
                else:
                    print(f"\n[结束] 翻页器消失")
                    input("\n按 Enter 关闭浏览器...")
                    return

            print(f"\n[结束] 扫描了 20 页，海报 {test_id} 仍未找到")

        finally:
            input("\n按 Enter 关闭浏览器...")
            browser.close()

if __name__ == "__main__":
    main()
