"""
深度调试脚本：逐个海报详细追踪

目标：找出为什么只有 11 个海报被添加，其他 5 个失败的原因
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
    print("深度调试脚本：逐个海报详细追踪")
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

            # 逐个处理每个海报
            success_list = []
            fail_list = []

            for idx, pid in enumerate(poster_ids):
                print("=" * 100)
                print(f"[{idx+1}/16] 处理海报 {pid}")
                print("=" * 100)

                # 打开弹窗
                print(f"  [1] 打开【选择海报】弹窗...")
                try:
                    select_btn = wrapper.locator(f'button:has-text("{SELECT_POSTER_BUTTON_TEXT}")').first
                    select_btn.click()
                    page.wait_for_timeout(1500)
                    print(f"      ✓ 弹窗打开命令已执行")
                except Exception as e:
                    print(f"      ❌ 打开弹窗失败: {str(e)[:80]}")
                    fail_list.append((pid, "打开弹窗失败"))
                    continue

                # 等待弹窗出现
                print(f"  [2] 等待弹窗出现...")
                try:
                    dialog = page.locator(DIALOG_SELECTOR).first
                    if not dialog.is_visible(timeout=3000):
                        print(f"      ❌ 弹窗未在 3s 内出现")
                        fail_list.append((pid, "弹窗加载超时"))
                        continue
                    print(f"      ✓ 弹窗已出现")
                except Exception as e:
                    print(f"      ❌ 弹窗检查失败: {str(e)[:80]}")
                    fail_list.append((pid, "弹窗检查异常"))
                    continue

                # 翻页查找
                print(f"  [3] 翻页查找海报 {pid}...")
                found = False
                found_page = 0

                for page_num in range(1, 251):
                    try:
                        rows = dialog.locator(POSTER_CARD_SELECTOR).all()
                        print(f"      [第 {page_num} 页] 行数: {len(rows)}", end=" ")

                        if len(rows) == 0 and page_num > 1:
                            print(f"❌ 翻页失败（0行）")
                            break

                        # 查找海报
                        target_row = None
                        for row in rows:
                            text = row.text_content()
                            if str(pid) in text:
                                target_row = row
                                found_page = page_num
                                break

                        if target_row:
                            print(f"✓ 找到！")

                            # 勾选
                            print(f"      [勾选] 点击 checkbox...", end=" ")
                            try:
                                label = target_row.locator('label.el-checkbox').first
                                label.click(force=True)
                                page.wait_for_timeout(400)
                                print(f"✓")
                            except Exception as e:
                                print(f"❌ {str(e)[:40]}")
                                fail_list.append((pid, f"勾选失败"))
                                break

                            # 确认
                            print(f"      [确认] 点击【确认】按钮...", end=" ")
                            try:
                                confirm_btn = dialog.locator(f'button:has-text("{DIALOG_CONFIRM_BUTTON_TEXT}")').first
                                confirm_btn.click()
                                page.wait_for_timeout(1200)
                                print(f"✓")
                                found = True
                                success_list.append(pid)
                            except Exception as e:
                                print(f"❌ {str(e)[:40]}")
                                fail_list.append((pid, f"确认失败"))
                            break

                        else:
                            print(f"(未找到)", end=" ")

                        # 翻页
                        try:
                            pagination = dialog.locator(".el-pagination").first
                            if pagination.is_visible(timeout=1000):
                                pag_buttons = pagination.locator("button").all()
                                if len(pag_buttons) > 1:
                                    next_btn = pag_buttons[1]
                                    if next_btn.is_visible() and not next_btn.is_disabled():
                                        next_btn.click()
                                        page.wait_for_timeout(800)
                                        page.wait_for_timeout(300)
                                    else:
                                        print(f"(已到最后一页)")
                                        break
                                else:
                                    break
                            else:
                                break
                        except Exception as e:
                            print(f"(翻页异常)")
                            break

                    except Exception as e:
                        print(f"(扫描异常: {str(e)[:30]})")
                        break

                if not found:
                    print(f"      ❌ 扫描完全部页面后未找到")
                    fail_list.append((pid, f"整页扫描未找到"))

                print()
                time.sleep(0.5)

            # 输出统计
            print("=" * 100)
            print("【最终统计】")
            print("=" * 100)
            print(f"\n✓ 成功添加 {len(success_list)} 个海报：{success_list}\n")

            if fail_list:
                print(f"❌ 失败 {len(fail_list)} 个海报：")
                for pid, reason in fail_list:
                    print(f"   - {pid}: {reason}")
            else:
                print("✓ 全部成功！")

        finally:
            input("\n按 Enter 关闭浏览器...")
            browser.close()

if __name__ == "__main__":
    main()
