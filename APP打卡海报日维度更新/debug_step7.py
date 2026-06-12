"""
调试脚本：查看【海外益智海报-非台湾】分组和海报的 HTML 结构
"""
import sys
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding="utf-8")

# 导入步骤函数
from poster_update import (
    step5_login_bizcenter,
    step6_search_poster_group
)

BUSINESS_GROUP_NAME = "海外益智海报-非台湾"

def main():
    print("=" * 60)
    print("调试脚本：查看分组结构")
    print("=" * 60)
    print()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        try:
            # Step 5: 登录平台中心
            print("[步骤 5] 登录平台中心...")
            if step5_login_bizcenter(page):
                print("  ✓ 平台中心已登录\n")
            else:
                print("  ❌ 登录失败\n")
                return

            # Step 6: 查询海报组
            print("[步骤 6] 查询海报组...")
            if step6_search_poster_group(page):
                print("  ✓ 海报组查询完成，已进入编辑界面\n")
            else:
                print("  ❌ 查询失败\n")
                return

            # 调试：找到【海外益智海报-非台湾】的位置
            print(f"\n[查找分组位置]")
            group_text_element = page.locator(f'text="{BUSINESS_GROUP_NAME}"').first
            if group_text_element.is_visible():
                # 获取该元素的信息
                group_html = group_text_element.evaluate("el => el.outerHTML")
                parent_html = group_text_element.evaluate("el => el.parentElement.outerHTML")

                print(f"找到文本元素")
                print(f"  该元素的HTML：{group_html[:300]}")
                print(f"  父元素的HTML：{parent_html[:300]}")

                # 尝试找到父级容器（可能是 div、td、th 等）
                parent_locator = group_text_element.locator("xpath=ancestor::div | ancestor::td | ancestor::th | ancestor::tr").first
                if parent_locator:
                    parent_info = parent_locator.evaluate("el => ({tag: el.tagName, text: el.textContent.substring(0,100), classes: el.className})")
                    print(f"  父容器：{parent_info}")

            print(f"\n[页面上所有包含【海外】的元素]")
            all_elements = page.locator("*:has-text('海外')").all()
            print(f"找到 {len(all_elements)} 个包含【海外】的元素\n")

            for idx, elem in enumerate(all_elements[:10]):
                try:
                    text = elem.text_content().strip()[:50]
                    tag = elem.evaluate("el => el.tagName")
                    print(f"  [{idx}] <{tag}> {text}")
                except:
                    pass


            # 打印所有行，找出分组标题
            print(f"\n[打印所有行找出分组名称]")
            all_rows = page.locator("table tbody tr").all()
            print(f"总共 {len(all_rows)} 行\n")

            # 查找所有可能的分组标题
            all_divs_with_title = page.locator("div.title").all()
            print(f"找到 {len(all_divs_with_title)} 个分组标题")
            for idx, div in enumerate(all_divs_with_title):
                title_text = div.text_content().strip()[:100]
                print(f"  [{idx}] {title_text}")

            print()



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
