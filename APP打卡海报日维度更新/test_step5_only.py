"""
独立测试 Step 5：登录平台中心
用于验证扫码登录方式是否正常工作
"""
import sys
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding="utf-8")

# 导入 Step 5 函数
from poster_update import step5_login_bizcenter

def main():
    print("=" * 60)
    print("独立测试 Step 5：登录平台中心")
    print("=" * 60)
    print()
    print("说明：")
    print("1. 浏览器将打开平台中心登录页面")
    print("2. 请扫码登录")
    print("3. 脚本会每 5 秒检查一次登录状态（最多 12 次，总 60 秒）")
    print("4. 检测到登录标记'欢乐童年'后，自动导航到海报组管理页面")
    print()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        try:
            # 运行 Step 5
            print("开始 Step 5...")
            result = step5_login_bizcenter(page)

            if result:
                print("\n" + "=" * 60)
                print("✓ Step 5 测试成功！")
                print("=" * 60)

                # 检查当前页面
                current_url = page.url
                print(f"\n当前 URL: {current_url}")

                # 列出页面上的主要元素
                print("\n页面内容检查：")

                # 检查是否有"欢乐童年"标记
                logo = page.locator("text=欢乐童年").first
                if logo.is_visible():
                    print("  ✓ 找到登录标记：欢乐童年")

                # 检查是否有菜单或主要内容
                menu = page.locator("[class*='menu'], [class*='nav'], [class*='sidebar']").first
                if menu.is_visible():
                    print("  ✓ 找到菜单/导航元素")

                # 检查是否有搜索框
                search = page.locator("input[placeholder*='搜索'], input[placeholder*='查询']").first
                if search.is_visible():
                    print("  ✓ 找到搜索框")

                print("\n请检查浏览器中的页面是否正确显示海报组管理界面")
                input("\n确认无误后，按 Enter 退出...")

            else:
                print("\n❌ Step 5 测试失败")

        except Exception as e:
            print(f"\n❌ 错误: {e}")
            import traceback
            traceback.print_exc()
        finally:
            print("\n正在关闭浏览器...")
            context.close()
            browser.close()
            print("✓ 浏览器已关闭")

if __name__ == "__main__":
    main()
