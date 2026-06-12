"""
测试：单次打开对话框，选择多个海报，一次性确认

这个测试与 step 8 current 不同：
- step 8 current: 对每个海报打开一次对话框 → 选择 → 确认 → 关闭 (16 次循环)
- 这个测试: 打开一次对话框 → 在同一对话框内找 16 个海报并全部勾选 → 一次确认 → 关闭
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
    step7_remove_business_group,
    GROUP_WRAPPER_SELECTOR,
    SELECT_POSTER_BUTTON_TEXT,
    DIALOG_SELECTOR,
    POSTER_CARD_SELECTOR,
    DIALOG_CONFIRM_BUTTON_TEXT,
)

SCRIPT_DIR = Path(__file__).parent
OUTPUT_DIR = SCRIPT_DIR / "output"

def main():
    print("=" * 100)
    print("测试：单次打开对话框，在同一对话框内选择全部 16 个海报")
    print("=" * 100)
    print()

    sorted_file = OUTPUT_DIR / "海报裂变率排序_20260611.xlsx"
    df = pd.read_excel(sorted_file, sheet_name='裂变率排序')
    poster_ids = [int(x) for x in df['海报ID-汇总处理'].tolist()]

    print(f"[准备] 要处理的 16 个海报ID：\n{poster_ids}\n")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        try:
            # Step 5-7
            if not step5_login_bizcenter(page):
                print("❌ Step 5 失败\n")
                return

            if not step6_search_poster_group(page):
                print("❌ Step 6 失败\n")
                return

            if not step7_remove_business_group(page):
                print("❌ Step 7 失败\n")
                return

            print("✓ Step 5-7 完成\n")

            wrapper = page.locator(GROUP_WRAPPER_SELECTOR).first

            # ============ 单次打开对话框 ============
            print("[步骤] 打开【选择海报】对话框...")
            try:
                select_btn = wrapper.locator(f'button:has-text("{SELECT_POSTER_BUTTON_TEXT}")').first
                select_btn.click()
                page.wait_for_timeout(1500)
            except Exception as e:
                print(f"❌ 打开弹窗失败: {e}")
                return

            dialog = page.locator(DIALOG_SELECTOR).first
            if not dialog.is_visible(timeout=3000):
                print(f"❌ 弹窗未出现")
                return

            print(f"✓ 弹窗已打开\n")

            # ============ 在同一对话框内逐个查找和勾选所有海报 ============
            success_count = 0
            fail_list = []

            for idx, pid in enumerate(poster_ids):
                print(f"[{idx+1}/16] 查找海报 {pid}...", end=" ", flush=True)

                found = False
                for page_num in range(1, 251):
                    rows = dialog.locator(POSTER_CARD_SELECTOR).all()

                    # 在当前页查找
                    target_row = None
                    for row in rows:
                        text = row.text_content()
                        if str(pid) in text:
                            target_row = row
                            found_page = page_num
                            break

                    if target_row:
                        # 勾选
                        try:
                            label = target_row.locator('label.el-checkbox').first
                            label.click(force=True)
                            page.wait_for_timeout(300)
                            success_count += 1
                            print(f"✓ 已勾选")
                            found = True
                        except Exception as e:
                            print(f"❌ 勾选失败")
                            fail_list.append((pid, "勾选失败"))
                        break

                    if page_num > 1 and len(rows) == 0:
                        print(f"❌ 翻页后无数据")
                        break

                    # 翻页
                    try:
                        pagination = dialog.locator(".el-pagination").first
                        if pagination.is_visible(timeout=1000):
                            pag_buttons = pagination.locator("button").all()
                            if len(pag_buttons) > 1:
                                next_btn = pag_buttons[1]
                                if next_btn.is_visible() and not next_btn.is_disabled():
                                    next_btn.click()
                                    page.wait_for_timeout(700)
                                else:
                                    break
                            else:
                                break
                        else:
                            break
                    except:
                        break

                if not found:
                    print(f"❌ 未找到")
                    fail_list.append((pid, "未找到"))

            print()

            # ============ 全部勾选完后，一次确认 ============
            print(f"\n[步骤] 点击【确认】按钮...")
            try:
                confirm_btn = dialog.locator(f'button:has-text("{DIALOG_CONFIRM_BUTTON_TEXT}")').first
                confirm_btn.click()
                page.wait_for_timeout(1500)
                print(f"✓ 已确认\n")
            except Exception as e:
                print(f"❌ 确认失败: {e}\n")
                return

            # ============ 验证结果 ============
            page.wait_for_timeout(2000)
            rows = wrapper.locator('tr.el-table__row').all()
            actual_count = len(rows)

            print("=" * 100)
            print(f"【结果统计】")
            print("=" * 100)
            print(f"✓ 尝试选择：{len(poster_ids)} 个海报")
            print(f"✓ 成功勾选：{success_count} 个海报")
            print(f"✓ 实际添加到分组：{actual_count} 个海报")

            if fail_list:
                print(f"\n❌ 失败 {len(fail_list)} 个：")
                for pid, reason in fail_list:
                    print(f"   - {pid}: {reason}")

            if actual_count == len(poster_ids):
                print(f"\n✓✓✓ 成功！所有 {len(poster_ids)} 个海报都已添加！")
            else:
                print(f"\n⚠️  预期 {len(poster_ids)} 个，实际只有 {actual_count} 个")
                print(f"   → 损失了 {len(poster_ids) - actual_count} 个海报")

        finally:
            input("\n按 Enter 关闭浏览器...")
            browser.close()

if __name__ == "__main__":
    main()
