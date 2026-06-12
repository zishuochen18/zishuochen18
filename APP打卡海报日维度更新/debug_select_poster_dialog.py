"""
调试脚本：查看【选择海报】弹窗的 HTML 结构、卡片格式、翻页器
"""
import sys
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding="utf-8")

from poster_update import (
    step5_login_bizcenter,
    step6_search_poster_group,
    GROUP_WRAPPER_SELECTOR,
    BUSINESS_GROUP_NAME
)

def main():
    print("=" * 80)
    print("调试脚本：【选择海报】弹窗结构")
    print("=" * 80)
    print()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        try:
            # Step 5: 登录平台中心（增加更长的超时以应对网络波动）
            print("[步骤 5] 登录平台中心（超时时间已增加至 90s）...")
            try:
                if step5_login_bizcenter(page):
                    print("  ✓ 平台中心已登录\n")
                else:
                    print("  ❌ 登录失败\n")
                    return
            except Exception as e:
                print(f"  ⚠️ 登录过程出错: {e}")
                print(f"  请手动确认已登录，按 Enter 继续...")
                input()

            # Step 6: 查询海报组
            print("[步骤 6] 查询海报组...")
            if step6_search_poster_group(page):
                print("  ✓ 海报组查询完成，已进入编辑界面\n")
            else:
                print("  ❌ 查询失败\n")
                return

            # 主要调试逻辑：定位目标分组，点击【选择海报】，分析弹窗结构
            print("=" * 80)
            print(f"[调试] 定位【{BUSINESS_GROUP_NAME}】的【选择海报】弹窗")
            print("=" * 80)
            print()

            # 定位 wrapper
            wrapper = page.locator(GROUP_WRAPPER_SELECTOR).first
            if not wrapper.is_visible():
                print(f"❌ 未找到分组 wrapper")
                input("\n按 Enter 关闭浏览器...")
                return

            print(f"✓ 找到分组 wrapper\n")

            # 打印 wrapper 内所有按钮
            print("[1] Wrapper 内所有按钮：")
            all_buttons = wrapper.locator('button').all()
            for idx, btn in enumerate(all_buttons):
                try:
                    text = btn.text_content().strip()[:30]
                    visible = btn.is_visible()
                    disabled = btn.evaluate("el => el.disabled")
                    print(f"  [{idx}] 文案='{text}' visible={visible} disabled={disabled}")
                except:
                    print(f"  [{idx}] (获取信息失败)")

            # 找【选择海报】按钮
            print(f"\n[2] 查找【选择海报】按钮...")
            select_btn = None
            for candidate_text in ["选择海报", "新增海报", "添加海报"]:
                try:
                    btn = wrapper.locator(f'button:has-text("{candidate_text}")').first
                    if btn.is_visible():
                        select_btn = btn
                        print(f"  ✓ 找到：'{candidate_text}'")
                        break
                except:
                    pass

            if not select_btn:
                print(f"  ⚠️ 未找到【选择海报】类按钮，终止")
                input("\n按 Enter 关闭浏览器...")
                return

            # 点击【选择海报】
            print(f"\n[3] 点击【选择海报】按钮...")
            select_btn.click()
            page.wait_for_timeout(1500)

            # 等待弹窗出现
            print(f"\n[4] 等待弹窗出现...")
            dialog = None
            for selector in [".el-dialog__wrapper:visible", "[role='dialog']", ".el-dialog:visible"]:
                try:
                    d = page.locator(selector).first
                    if d.is_visible(timeout=2000):
                        dialog = d
                        print(f"  ✓ 找到弹窗：{selector}")
                        break
                except:
                    pass

            if not dialog:
                print(f"  ⚠️ 弹窗未出现，终止")
                input("\n按 Enter 关闭浏览器...")
                return

            page.wait_for_timeout(500)

            # 打印弹窗结构
            print(f"\n[5] 弹窗顶层 outerHTML（前 500 字）：")
            dialog_html = dialog.evaluate("el => el.outerHTML")
            print(f"  {dialog_html[:500]}")
            print()

            # 打印弹窗内所有按钮
            print(f"[6] 弹窗内所有按钮：")
            dialog_buttons = dialog.locator('button').all()
            for idx, btn in enumerate(dialog_buttons[:15]):  # 最多打印 15 个
                try:
                    text = btn.text_content().strip()[:30]
                    print(f"  [{idx}] '{text}'")
                except:
                    print(f"  [{idx}] (获取失败)")

            # 打印弹窗内所有 input
            print(f"\n[7] 弹窗内所有 input：")
            dialog_inputs = dialog.locator('input').all()
            for idx, inp in enumerate(dialog_inputs[:10]):
                try:
                    placeholder = inp.get_attribute("placeholder") or "(无)"
                    print(f"  [{idx}] placeholder='{placeholder}'")
                except:
                    print(f"  [{idx}] (获取失败)")

            # 打印海报卡片结构
            print(f"\n[8] 海报卡片（前 5 个）：")
            # 候选选择器
            for selector in [".el-card", "tr.el-table__row", "li.poster-item", ".poster-item", "[data-id]"]:
                cards = dialog.locator(selector).all()
                if cards:
                    print(f"  找到 {len(cards)} 个卡片（selector: {selector}）")
                    for i, card in enumerate(cards[:5]):
                        try:
                            html = card.evaluate("el => el.outerHTML")
                            text = card.text_content().strip()[:100]
                            data_id = card.get_attribute("data-id") or "(无)"
                            print(f"    [{i}] data-id='{data_id}' text='{text}' HTML[:100]='{html[:100]}'")
                        except:
                            print(f"    [{i}] (获取失败)")
                    break

            # 打印翻页器结构
            print(f"\n[9] 翻页器结构：")
            pagination = dialog.locator(".el-pagination").first
            if pagination.is_visible(timeout=2000):
                print(f"  ✓ 找到翻页器")
                pag_html = pagination.evaluate("el => el.outerHTML")
                print(f"    outerHTML[:200]: {pag_html[:200]}")
                # 打印翻页器内的按钮
                pag_buttons = pagination.locator("button").all()
                for idx, btn in enumerate(pag_buttons):
                    try:
                        text = btn.text_content().strip()
                        disabled = btn.is_disabled()
                        print(f"    按钮 [{idx}] text='{text}' disabled={disabled}")
                    except:
                        pass
            else:
                print(f"  ⚠️ 未找到翻页器")

            # 测试已知海报ID（1758）穿透翻页
            print(f"\n[10] 测试查找已知海报ID（1758）：")
            found = False
            for page_num in range(1, 6):  # 最多翻 5 页
                print(f"  [页 {page_num}] 查找...")
                # 候选选择器：文本 / data-id
                card = dialog.locator("*:has-text('1758')").first
                try:
                    if card.is_visible(timeout=1000):
                        print(f"    ✓ 在本页找到海报 1758")
                        html = card.evaluate("el => el.outerHTML")
                        print(f"      所在元素 HTML[:150]: {html[:150]}")
                        found = True
                        break
                except:
                    pass

                # 翻下一页
                next_btns = dialog.locator("button:has-text('下一页')").all()
                if next_btns:
                    for nb in next_btns:
                        try:
                            if nb.is_visible() and not nb.is_disabled():
                                nb.click()
                                page.wait_for_timeout(600)
                                break
                        except:
                            pass

            if found:
                print(f"  ✓ 穿透翻页成功命中海报 1758")
            else:
                print(f"  ⚠️ 未在前 5 页找到海报 1758（可能已有更多页）")

            # 输出建议的 selector 集合
            print(f"\n" + "=" * 80)
            print("建议的 selector 集合（复制到 poster_update.py 顶部）：")
            print("=" * 80)
            print(f"""
SELECT_POSTER_BUTTON_TEXT = "选择海报"           # 由上面【按钮文案】确认
DIALOG_SELECTOR = ".el-dialog__wrapper:visible"  # 根据【弹窗结构】调整
POSTER_CARD_SELECTOR = ".el-card"                # 根据【海报卡片】调整
PAGINATION_NEXT_SELECTOR = "button:has-text('下一页')"  # 根据【翻页器】调整
DIALOG_CONFIRM_BUTTON_TEXT = "确定"              # 根据【弹窗按钮】调整
            """)

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
