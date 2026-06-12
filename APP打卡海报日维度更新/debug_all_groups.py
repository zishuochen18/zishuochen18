"""
调试脚本：查找所有分组及其对应的行数范围
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

def main():
    print("=" * 60)
    print("调试脚本：查找所有分组和行数范围")
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

            # 调试：找出所有分组及其对应的表格
            print("=" * 60)
            print("所有分组及其对应的表格")
            print("=" * 60)

            # 找出所有分组标题
            all_group_titles = page.locator("div.title").all()
            print(f"找到 {len(all_group_titles)} 个分组标题\n")

            for idx, title_div in enumerate(all_group_titles):
                try:
                    # 获取分组名称（用多种方式尝试）
                    group_name = None

                    # 方式1：尝试获取 p 标签的文本
                    try:
                        title_elem = title_div.locator("p").first
                        if title_elem:
                            group_name = title_elem.text_content(timeout=5000).strip()
                    except:
                        pass

                    # 方式2：如果方式1失败，直接从 div 获取文本
                    if not group_name:
                        group_name = title_div.text_content(timeout=5000).strip()[:50]

                    # 方式3：如果还是失败，直接获取 HTML
                    if not group_name:
                        group_html = title_div.evaluate("el => el.innerHTML")
                        group_name = f"(HTML: {group_html[:50]})"

                    print(f"[{idx}] 分组名称：{group_name}")

                    # 获取该分组标题的 HTML
                    try:
                        title_html = title_div.evaluate("el => el.outerHTML")
                        print(f"    标题 HTML（前200字）：{title_html[:200]}")
                    except:
                        print(f"    无法获取标题 HTML")

                    # 找到紧跟其后的表格
                    try:
                        next_table = title_div.locator("xpath=following::table[1]").first
                        if next_table:
                            rows = next_table.locator("tbody tr").all()
                            print(f"    对应表格行数：{len(rows)}")
                        else:
                            print(f"    未找到对应的表格")
                    except Exception as e:
                        print(f"    查找表格时出错：{e}")

                    print()
                except Exception as e:
                    print(f"[{idx}] 处理分组时出错：{e}")
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
