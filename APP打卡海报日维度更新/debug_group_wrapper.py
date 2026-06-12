"""
调试脚本：定位【海外益智海报-非台湾】分组的 wrapper 容器
沿着祖先链向上走，找到第一个既包含目标 title 又包含【移除】按钮的容器
"""
import sys
from pathlib import Path
from playwright.sync_api import sync_playwright
import json

sys.stdout.reconfigure(encoding="utf-8")

from poster_update import (
    step5_login_bizcenter,
    step6_search_poster_group,
    BUSINESS_GROUP_NAME
)

def main():
    print("=" * 80)
    print("调试脚本：定位分组 wrapper 容器")
    print("=" * 80)
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

            # 主要调试逻辑：沿着祖先链向上走
            print("=" * 80)
            print(f"[调试] 定位【{BUSINESS_GROUP_NAME}】的 wrapper 容器")
            print("=" * 80)
            print()

            # 尝试找到目标 title 的 <p> 元素
            title_p = page.locator(f'div.title:has(p:has-text("{BUSINESS_GROUP_NAME}")) p').first

            if not title_p.is_visible():
                print(f"❌ 未找到【{BUSINESS_GROUP_NAME}】的标题元素")
                input("\n按 Enter 关闭浏览器...")
                return

            print(f"✓ 找到【{BUSINESS_GROUP_NAME}】的 <p> 元素\n")

            # 用 JavaScript 沿着祖先链向上走
            ancestors_info = title_p.evaluate("""
                el => {
                    const out = [];
                    let cur = el;
                    for (let i = 0; i < 10 && cur; i++) {
                        cur = cur.parentElement;
                        if (!cur) break;

                        // 计数：该容器内有多少个【移除】按钮
                        const removeButtons = Array.from(cur.querySelectorAll('button'))
                            .filter(b => b.textContent.trim() === '移除');
                        const removeCount = removeButtons.length;

                        // 计数：该容器内有多少行（tr.el-table__row）
                        const rows = cur.querySelectorAll('tr.el-table__row');
                        const rowCount = rows.length;

                        // 计数：该容器内有多少个 div.title
                        const titles = cur.querySelectorAll('div.title');
                        const titleCount = titles.length;

                        // 获取容器的信息
                        const info = {
                            level: i + 1,
                            tag: cur.tagName,
                            className: cur.className,
                            id: cur.id || '(无)',
                            hasVueScope: cur.hasAttribute('data-v-fbffa806'),
                            vueScopeValue: cur.getAttribute('data-v-fbffa806') || '(无)',
                            removeBtnCount: removeCount,
                            rowCount: rowCount,
                            titleCount: titleCount,
                            outerHTMLPreview: cur.outerHTML.substring(0, 200),
                            // 是否是理想的 wrapper（只包含 1 个 title 且有【移除】按钮）
                            isIdealWrapper: titleCount === 1 && removeCount > 0
                        };
                        out.push(info);
                    }
                    return out;
                }
            """)

            # 打印结果
            print("祖先链分析（从 <p> 元素向上）：\n")
            for info in ancestors_info:
                level = info['level']
                tag = info['tag']
                className = info['className']
                hasVue = info['hasVueScope']
                removeCnt = info['removeBtnCount']
                rowCnt = info['rowCount']
                titleCnt = info['titleCount']
                isIdeal = info['isIdealWrapper']

                marker = "👈 【理想 wrapper】" if isIdeal else ""
                vue_marker = f" [Vue-{info['vueScopeValue']}]" if hasVue else ""

                print(f"[Level {level}] <{tag}> class='{className}'{vue_marker}")
                print(f"  📊 统计: title={titleCnt}, row={rowCnt}, 【移除】={removeCnt} {marker}")
                print(f"  🔍 HTML preview: {info['outerHTMLPreview'][:100]}...")
                print()

            # 找出第一个理想 wrapper
            ideal_wrapper = next((info for info in ancestors_info if info['isIdealWrapper']), None)
            if ideal_wrapper:
                print("=" * 80)
                print(f"✅ 找到理想的 wrapper 容器！")
                print("=" * 80)
                print(f"\n第一个【理想 wrapper】在 Level {ideal_wrapper['level']}：")
                print(f"  Tag: <{ideal_wrapper['tag']}>")
                print(f"  ClassName: {ideal_wrapper['className']}")
                print(f"  HasVueScope: {ideal_wrapper['hasVueScope']}")
                print(f"  ContentInfo:")
                print(f"    - 包含的 title 数: {ideal_wrapper['titleCount']}")
                print(f"    - 包含的行数: {ideal_wrapper['rowCount']}")
                print(f"    - 包含的【移除】按钮数: {ideal_wrapper['removeBtnCount']}")
                print()

                # 给出建议的 selector
                className = ideal_wrapper['className'].strip()
                if className:
                    suggested_selector = f"div.{className.split()[0]}:has(p:has-text(\"{BUSINESS_GROUP_NAME}\"))"
                else:
                    suggested_selector = f"{ideal_wrapper['tag'].lower()}:has(p:has-text(\"{BUSINESS_GROUP_NAME}\"))"

                print(f"建议的 Playwright selector（用于 phase B）：")
                print(f"  GROUP_WRAPPER_SELECTOR = \"{suggested_selector}\"")
                print()
            else:
                print("=" * 80)
                print("⚠️ 未找到理想的 wrapper 容器")
                print("=" * 80)
                print("\n所有祖先都是平的，或者 title 数 > 1。")
                print("这种情况下需要用 Phase B 的备选方案（compareDocumentPosition）。")
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
