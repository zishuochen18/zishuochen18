"""
APP打卡海报日维度更新 - 完整自动化工作流

步骤：
1. 登录神策数据网页
2. 进入指定书签
3. 导出数据表格
4. 计算曝光裂变率并排序
5. 登录平台中心
6. 编辑海报组（移除旧海报）
7. 按排序结果输入新海报ID
8. 保存更新
"""
import sys
import os
from pathlib import Path
from datetime import datetime
import time

sys.stdout.reconfigure(encoding="utf-8")

SCRIPT_DIR = Path(__file__).parent
SAMPLE_DIR = SCRIPT_DIR / "sample"
OUTPUT_DIR = SCRIPT_DIR / "output"
LOGS_DIR = SCRIPT_DIR / "logs"

# 神策数据配置
SENSORS_URL = "https://sensors.61info.cn/bookmark/?project=platform_production&product=sensors_analysis"
SENSORS_USERNAME = "chenzishuo@hltn.com"
SENSORS_PASSWORD = "pYM49395$"
BOOKMARK_NAME = "思维海外海报效率分析-日维度更新"

# 平台中心配置
BIZCENTER_URL = "https://bizcenter-h5-cms.61info.cn/#/groupPoster"
POSTER_GROUP_CODE = "siweidaka"
BUSINESS_GROUP_NAME = "海外益智海报-非台湾"
# wrapper 容器选择器（通过 debug_group_wrapper.py 确认）
GROUP_WRAPPER_SELECTOR = 'div.oneGroup:has(p:has-text("海外益智海报-非台湾"))'

# 【选择海报】弹窗相关选择器（通过 debug_select_poster_dialog.py 确认）
SELECT_POSTER_BUTTON_TEXT = "选择海报"
DIALOG_SELECTOR = ".el-dialog.select-product-dialog"
POSTER_CARD_SELECTOR = "tr.el-table__row"
DIALOG_SEARCH_INPUT_SELECTOR = "input[placeholder='请输入']"
DIALOG_CONFIRM_BUTTON_TEXT = "确认"


def step1_login_sensors(page):
    """第一步：登录神策数据网页（在已创建的 page 对象上操作）"""
    print("[步骤 1] 登录神策数据网页...")
    print(f"  URL: {SENSORS_URL}")
    print(f"  用户名: {SENSORS_USERNAME}")

    # 访问网页
    page.goto(SENSORS_URL, timeout=60000)
    page.wait_for_load_state("networkidle", timeout=30000)
    print("  ✓ 网页已加载")

    # 等待表单加载
    page.wait_for_timeout(2000)

    # 查找用户名和密码输入框
    print("  查找登录表单...")

    # 尝试多种选择器方式找输入框
    username_selectors = [
        'input[placeholder*="用户"]',
        'input[placeholder*="邮箱"]',
        'input[placeholder*="账"]',
        'input[name*="user"]',
        'input[name*="email"]',
        'input[type="text"]'
    ]

    username_field = None
    for selector in username_selectors:
        try:
            elem = page.locator(selector).first
            if elem.is_visible():
                username_field = elem
                print(f"  ✓ 找到用户名输入框: {selector}")
                break
        except:
            pass

    if not username_field:
        print("  ⚠️ 未找到用户名输入框，请打开浏览器手动检查页面结构")
        input("  按 Enter 继续...")
        return False

    # 查找密码输入框
    password_field = page.locator('input[type="password"]').first
    if not password_field.is_visible():
        print("  ⚠️ 未找到密码输入框")
        input("  按 Enter 继续...")
        return False

    # 输入用户名
    username_field.fill(SENSORS_USERNAME)
    print(f"  ✓ 已输入用户名")

    # 输入密码
    password_field.fill(SENSORS_PASSWORD)
    print(f"  ✓ 已输入密码")

    # 等待登录按钮加载
    page.wait_for_timeout(1000)

    # 尝试多种方式找登录按钮
    print("  查找登录按钮...")
    login_button = None

    button_selectors = [
        'button:has-text("登 录")',
        'button:has-text("登录")',
        'button:has-text("Login")',
        'button[type="submit"]',
        'input[type="submit"]',
        'div[role="button"]:has-text("登录")',
        'a:has-text("登录")'
    ]

    for selector in button_selectors:
        try:
            elem = page.locator(selector).first
            if elem.is_visible():
                login_button = elem
                print(f"  ✓ 找到登录按钮: {selector}")
                break
        except:
            pass

    if login_button:
        # 点击登录按钮
        login_button.click()
        print("  ✓ 已点击登录按钮")

        # 等待登录完成
        page.wait_for_load_state("networkidle", timeout=30000)
        print("  ✓ 登录成功")
    else:
        print("  ⚠️ 未找到登录按钮，请打开浏览器手动点击登录")
        input("  登录完成后，按 Enter 继续...")

    return True


def step2_enter_bookmark(page):
    """第二步：找到书签并进入（在已创建的 page 对象上操作）"""
    print(f"[步骤 2] 进入书签: {BOOKMARK_NAME}...")

    # 增加等待时间，确保动态内容加载
    print("  等待页面内容加载...")
    page.wait_for_timeout(5000)

    # 检查当前 URL，用于调试
    current_url = page.url
    print(f"  当前 URL: {current_url[:80]}")

    # 方法：直接搜索包含书签名称的链接
    print(f"  查找书签: 【{BOOKMARK_NAME}】...")

    # 先调试：列出所有页面上的链接
    all_links = page.locator("a").all()
    print(f"  [调试] 页面上共有 {len(all_links)} 个链接")

    # 查找包含书签名称的链接
    found_bookmarks = []
    for idx, link in enumerate(all_links):
        try:
            text = link.text_content().strip()
            is_visible = link.is_visible()
            if idx < 10:  # 显示前 10 个
                print(f"    [{idx}] 文本: {text[:40] if text else '(空)'} | 可见: {is_visible}")

            if BOOKMARK_NAME in text or "思维海外海报" in text:
                found_bookmarks.append((text, link))
                print(f"  ✓ 找到匹配的链接: {text[:50]}")
        except:
            pass

    # 查找书签表格中的链接（多种选择器尝试）
    selectors_to_try = [
        f'a:has-text("{BOOKMARK_NAME}")',
        f'a:has-text("思维海外海报")',
        f'tr:has-text("{BOOKMARK_NAME}") a',
        f'*:has-text("{BOOKMARK_NAME}")'
    ]

    bookmark_link = None
    for selector in selectors_to_try:
        try:
            elem = page.locator(selector).first
            if elem.is_visible():
                print(f"  ✓ 用选择器找到链接: {selector}")
                bookmark_link = elem
                break
        except:
            pass

    if bookmark_link:
        print(f"  ✓ 找到书签链接")
        print(f"  准备点击进入...")

        # 滚动到元素可见
        try:
            bookmark_link.scroll_into_view_if_needed(timeout=3000)
            print(f"  ✓ 已滚动到视图内")
        except:
            pass

        # 获取元素信息用于调试
        try:
            href = bookmark_link.get_attribute("href")
            print(f"  链接地址: {href[:60] if href else '(空)'}")
        except:
            pass

        # 尝试多种点击方式
        print(f"  执行点击操作...")
        try:
            # 先尝试强制点击
            bookmark_link.click(force=True)
            print(f"  ✓ 已执行点击")
        except Exception as e:
            print(f"  ⚠️ 点击异常: {e}")
            # 如果点击失败，尝试用 JavaScript 点击
            try:
                bookmark_link.evaluate("el => el.click()")
                print(f"  ✓ 已用 JS 执行点击")
            except Exception as e2:
                print(f"  ❌ JS 点击也失败: {e2}")

        # 等待页面加载
        print(f"  等待页面加载...")
        try:
            page.wait_for_load_state("networkidle", timeout=30000)
            print(f"  ✓ 书签页面已加载完成")
        except Exception as e:
            print(f"  ⚠️ 等待超时或异常: {e}")
            page.wait_for_timeout(3000)
            print(f"  已等待 3 秒，继续...")

        return True
    else:
        # 如果没找到书签，显示更详细的页面信息
        print(f"  ⚠️ 未找到可点击的书签链接")

        # 检查是否仍在登录页
        login_field = page.locator('input[type="password"]').first
        if login_field.is_visible():
            print(f"  [提示] 似乎仍在登录页面，请检查登录状态")
            return False
        else:
            print(f"  [提示] 页面已加载，但未找到书签")
            print(f"  请在浏览器中手动点击书签【{BOOKMARK_NAME}】进入")
            input("  完成后，按 Enter 继续...")
            return True


def step3_export_data(page, context):
    """第三步：导出数据表格

    需要先选择日期范围【过去14天】，等待5秒数据缓存，再下载
    """
    import os
    import glob
    from pathlib import Path

    print("[步骤 3] 导出数据表格...")

    # ============ 第一步：选择日期 ============
    print("  [子步骤] 选择日期范围【过去 14 天】...")

    try:
        # 1. 查找并点击日期输入框
        #    Sensors 使用 ant-input，value 包含"过去"
        date_input = page.locator('input.ant-input[value*="过去"]').first
        if not date_input.is_visible(timeout=3000):
            # 备选：通过 date-range-picker 容器找
            date_input = page.locator('.date-range-picker-wrapper-v3 input').first

        if date_input and date_input.is_visible():
            date_input.click(force=True)
            print(f"    ✓ 已点击日期选择框")
            page.wait_for_timeout(1500)
        else:
            print(f"    ⚠️ 未找到日期选择框")

        # 2. 在弹出的面板中，点击【过去 14 天】快捷选项
        #    Sensors 面板使用 div.date-range-picker-ranges-link-v3，文本带空格
        print("    查找【过去 14 天】选项...")
        past_14_btn = page.locator('div.date-range-picker-ranges-link-v3:has-text("过去 14 天")').first

        if past_14_btn and past_14_btn.is_visible(timeout=2000):
            past_14_btn.click(force=True)
            print(f"    ✓ 已点击【过去 14 天】")
            page.wait_for_timeout(1000)

            # 3. 点击【确定】按钮确认日期范围
            print("    查找并点击【确定】按钮...")
            confirm_btn = page.locator('button:has-text("确定")').first
            if confirm_btn and confirm_btn.is_visible(timeout=2000):
                confirm_btn.click(force=True)
                print(f"    ✓ 已点击【确定】")
            else:
                # 备选：可能面板自动关闭，无需确认
                print(f"    ⚠️ 未找到确定按钮，可能已自动确认")

            # 4. 等待 5 秒数据缓存
            print("    等待 5 秒数据缓存...")
            page.wait_for_timeout(5000)
            print(f"    ✓ 数据已缓存")
        else:
            print(f"    ⚠️ 未找到【过去 14 天】按钮，继续使用当前日期")

    except Exception as e:
        print(f"    ⚠️ 选择日期出错: {str(e)[:100]}")
        print(f"    继续使用当前日期进行导出")

    # ============ 第二步：导出数据 ============
    # 创建输出目录
    SAMPLE_DIR.mkdir(exist_ok=True)

    print("  等待 5s 网页加载...")
    page.wait_for_timeout(5000)

    # 查找导出按钮（多种选择器尝试）
    print("  查找导出按钮...")
    export_button = None

    export_selectors = [
        'button:has-text("导出")',
        'button:has-text("下载")',
        'button:has-text("Export")',
        'a:has-text("导出")',
        'a:has-text("下载")',
        '[class*="export"]',
        '[class*="download"]'
    ]

    for selector in export_selectors:
        try:
            elem = page.locator(selector).first
            if elem.is_visible():
                print(f"  ✓ 找到导出按钮: {selector}")
                export_button = elem
                break
        except:
            pass

    if not export_button:
        # 列出页面上的所有按钮用于调试
        all_buttons = page.locator("button").all()
        print(f"  [调试] 页面上共有 {len(all_buttons)} 个按钮")
        for idx, btn in enumerate(all_buttons[:15]):
            try:
                text = btn.text_content().strip()
                print(f"    [{idx}] {text[:40] if text else '(空)'}")
            except:
                pass

        print("  ⚠️ 未找到导出按钮，尝试等待文件自动下载...")

        # 不再使用 input()，改为等待下载事件
        try:
            page.wait_for_download(timeout=30000)
            return True
        except:
            print("  ⚠️ 没有检测到下载，返回 False")
            return False

    # 准备下载
    print("  准备处理导出...")

    # 点击导出按钮
    export_button.click(force=True)
    print("  ✓ 已点击导出按钮")

    # 等待可能的弹窗出现
    page.wait_for_timeout(2000)

    # 检查是否有对话框或确认按钮出现
    print("  检查是否有下载确认对话框...")

    # 尝试找确认/下载按钮
    try:
        with page.expect_download(timeout=30000) as download_info:
            # 尝试多种确认按钮
            confirm_selectors = [
                'button:has-text("下载文件")',
                'button:has-text("下载")',
                'button:has-text("确认下载")',
                'button:has-text("确定")',
            ]

            clicked = False
            for selector in confirm_selectors:
                try:
                    btn = page.locator(selector).first
                    if btn.is_visible(timeout=1000):
                        print(f"  ✓ 找到确认按钮: {selector}")
                        btn.click(force=True)
                        clicked = True
                        break
                except:
                    pass

            if not clicked:
                print("  ⚠️ 未找到确认按钮，等待下载...")

            page.wait_for_timeout(2000)

        # 获取下载的文件
        download = download_info.value
        download_path = SAMPLE_DIR / download.suggested_filename
        download.save_as(str(download_path))
        print(f"  ✓ 数据已导出: {download_path.name}")
        return download_path

    except Exception as e:
        print(f"  ⚠️ 下载异常: {str(e)[:100]}")
        return None


def step4_calculate_conversion_rate(excel_file: Path) -> Path:
    """第四步：计算曝光裂变率并排序"""
    import pandas as pd

    print("[步骤 4] 计算曝光裂变率...")

    # 读取数据
    df = pd.read_excel(excel_file, sheet_name=0)
    print(f"  ✓ 读取 {len(df)} 行数据")

    # 查找列名（处理中文编码问题）
    # 需要找到: 海报ID, [A]海报曝光次数, [F]小程序允许授权UV

    poster_id_col = None
    exposure_col = None  # [A]
    auth_col = None      # [F]

    for col in df.columns:
        if "海报ID" in col:
            poster_id_col = col
        if "[A]" in col and "曝光" in col:
            exposure_col = col
        if "[F]" in col and "授权" in col:
            auth_col = col

    if not all([poster_id_col, exposure_col, auth_col]):
        print(f"  ❌ 找不到必要的列")
        print(f"    海报ID: {poster_id_col}")
        print(f"    [A]曝光: {exposure_col}")
        print(f"    [F]授权: {auth_col}")
        print(f"  可用列: {list(df.columns)}")
        return None

    print(f"  ✓ 找到必要的列:")
    print(f"    海报ID: {poster_id_col}")
    print(f"    [A]曝光: {exposure_col}")
    print(f"    [F]授权: {auth_col}")

    # 计算裂变率
    # 公式：裂变率 = [F] 小程序允许授权UV / [A] 海报曝光次数
    df_calc = df.copy()
    df_calc['裂变率'] = df_calc[auth_col] / df_calc[exposure_col]

    # 替换 inf 和 nan
    df_calc['裂变率'] = df_calc['裂变率'].replace([float('inf'), float('-inf')], 0)
    df_calc['裂变率'] = df_calc['裂变率'].fillna(0)

    # 筛选裂变率 > 0 的数据
    df_filtered = df_calc[df_calc['裂变率'] > 0].copy()
    print(f"  筛选结果: {len(df_filtered)} / {len(df_calc)} (裂变率 > 0)")

    # 排序（从高到低）
    df_sorted = df_filtered.sort_values('裂变率', ascending=False).reset_index(drop=True)

    # 保存结果
    output_name = f"海报裂变率排序_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx"
    output_path = OUTPUT_DIR / output_name

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_sorted.to_excel(writer, sheet_name='裂变率排序', index=False)

    print(f"  ✓ 结果已保存: {output_path.name}")
    print(f"  排序 TOP 10:")
    for idx, row in df_sorted.head(10).iterrows():
        poster_id = row[poster_id_col]
        exposure = row[exposure_col]
        auth = row[auth_col]
        rate = row['裂变率']
        print(f"    [{idx+1}] 海报ID: {poster_id}, 曝光: {exposure}, 授权: {auth}, 裂变率: {rate:.4f}")

    return output_path


def step5_login_bizcenter(page):
    """第五步：登录平台中心

    步骤：
    1. 进入登录网页 https://login.61info.cn/
    2. 确认登录状态
    3. 导航到 https://bizcenter-h5-cms.61info.cn/#/groupPoster
    4. 确认进入海报组界面
    """
    import time

    print("[步骤 5] 登录平台中心...")

    # 步骤 1：进入登录页面
    print("  [子步骤 1] 进入登录网页...")
    page.goto("https://login.61info.cn/", timeout=60000)
    try:
        page.wait_for_load_state("domcontentloaded", timeout=15000)
    except:
        print("  ⚠️ 页面加载较慢，继续执行...")
    page.wait_for_timeout(2000)
    print("  ✓ 登录页面已加载")

    # 步骤 2：确认登录状态
    print("  [子步骤 2] 确认登录状态...")
    max_attempts = 12
    check_interval = 5
    logged_in = False

    for attempt in range(1, max_attempts + 1):
        print(f"    [检查 {attempt}/{max_attempts}]...")
        page.wait_for_timeout(1000)

        # 检查是否有"欢乐童年"登录标记
        try:
            logo_element = page.locator("text=欢乐童年").first
            if logo_element.is_visible():
                print(f"    ✓ 检测到登录标记：欢乐童年")
                logged_in = True
                break
        except:
            pass

        # 检查是否已离开登录页面（URL 不再是 login）
        current_url = page.url
        if "login.61info.cn" not in current_url:
            print(f"    ✓ 已离开登录页面")
            logged_in = True
            break

        if attempt < max_attempts:
            print(f"    {check_interval} 秒后重试（请在浏览器中完成扫码登录）...")
            time.sleep(check_interval)
            page.reload(timeout=30000)

    if not logged_in:
        print(f"  ⚠️ 检查超时，请确认是否已完成扫码登录")
        input("  确认后按 Enter 继续...")

    print(f"  ✓ 登录状态已确认")

    # 步骤 3：导航到海报组管理页面
    print("  [子步骤 3] 导航到海报组管理页面...")

    # 第一次导航
    page.goto("https://bizcenter-h5-cms.61info.cn/#/groupPoster", timeout=60000)
    page.wait_for_load_state("networkidle", timeout=30000)
    page.wait_for_timeout(2000)
    print("  ✓ 第一次导航完成")

    # 检查是否需要再次导航
    current_url = page.url
    print(f"    当前 URL: {current_url}")

    if "groupPoster" not in current_url:
        print("  ✓ 进行第二次导航...")
        page.goto("https://bizcenter-h5-cms.61info.cn/#/groupPoster", timeout=60000)
        page.wait_for_load_state("networkidle", timeout=30000)
        page.wait_for_timeout(2000)
        print("  ✓ 第二次导航完成")

    # 检查是否仍需导航
    current_url = page.url
    if "groupPoster" not in current_url:
        print("  ✓ 进行第三次导航...")
        page.goto("https://bizcenter-h5-cms.61info.cn/#/groupPoster", timeout=60000)
        page.wait_for_load_state("networkidle", timeout=30000)
        page.wait_for_timeout(2000)
        print("  ✓ 第三次导航完成")

    print("  ✓ 页面已加载")

    # 步骤 4：确认进入海报组界面
    print("  [子步骤 4] 确认进入海报组界面...")
    final_url = page.url
    print(f"    当前 URL: {final_url}")

    if "bizcenter-h5-cms" in final_url and "groupPoster" in final_url:
        print(f"  ✓ 已成功进入海报组管理界面")
        return True
    else:
        print(f"  ⚠️ 未进入预期页面，当前 URL: {final_url}")
        input("  按 Enter 继续...")
        return True


def step6_search_poster_group(page, poster_group_code=None):
    """第六步：查询海报组并进入编辑界面

    Args:
        page: Playwright page 对象
        poster_group_code: 海报组代码（默认使用全局 POSTER_GROUP_CODE）
    """
    if poster_group_code is None:
        poster_group_code = POSTER_GROUP_CODE

    print(f"[步骤 6] 查询海报组: {poster_group_code}...")

    # 等待页面加载
    page.wait_for_timeout(2000)

    # 首先列出所有输入框（用于调试）
    print("  [调试] 页面上的所有输入框:")
    all_inputs = page.locator('input').all()
    for idx, inp in enumerate(all_inputs):
        try:
            placeholder = inp.get_attribute('placeholder') or "(无)"
            name = inp.get_attribute('name') or "(无)"
            input_type = inp.get_attribute('type') or "text"
            is_readonly = inp.get_attribute('readonly') is not None
            is_visible = inp.is_visible()
            print(f"    [{idx}] type={input_type}, placeholder={placeholder}, name={name}, readonly={is_readonly}, visible={is_visible}")
        except:
            pass

    # 查找搜索框（跳过只读的）
    search_selectors = [
        'input[placeholder*="组code"]',
        'input[placeholder*="Code"]',
        'input[placeholder*="代码"]',
        'input[placeholder*="海报"]',
        'input[placeholder*="查询"]',
        'input[placeholder*="搜索"]:not([readonly])',
        'input[name*="code"]',
        'input[name*="groupCode"]',
        'input[type="text"]:not([readonly])'
    ]

    search_input = None
    for selector in search_selectors:
        try:
            elem = page.locator(selector).first
            if elem.is_visible():
                # 检查是否为只读
                is_readonly = elem.get_attribute('readonly') is not None
                if not is_readonly:
                    placeholder = elem.get_attribute('placeholder') or "(无)"
                    print(f"  ✓ 找到搜索框 (placeholder={placeholder}): {selector}")
                    search_input = elem
                    break
        except:
            pass

    if search_input:
        # 输入海报组 code
        try:
            search_input.fill(poster_group_code)
            print(f"  ✓ 已输入: {poster_group_code}")
        except Exception as e:
            print(f"  ⚠️ 输入失败: {e}")
            print(f"  请手动输入: {poster_group_code}")
            input("  输入完成后按 Enter 继续...")

        # 查找搜索/查询按钮
        search_buttons = [
            'button:has-text("搜索")',
            'button:has-text("查询")',
            'button:has-text("Search")',
            'button[type="submit"]'
        ]

        for selector in search_buttons:
            try:
                btn = page.locator(selector).first
                if btn.is_visible():
                    print(f"  ✓ 找到查询按钮: {selector}")
                    btn.click()
                    page.wait_for_timeout(2000)
                    break
            except:
                pass

        print(f"  ✓ 查询完成")

        # 查找并点击【编辑】按钮
        print(f"  [子步骤] 查找【编辑】按钮...")

        # 先列出所有可见的按钮（用于调试）
        print(f"  [调试] 页面上的所有按钮:")
        all_buttons = page.locator('button').all()
        for idx, btn in enumerate(all_buttons):
            try:
                text = btn.text_content().strip() if btn.text_content() else "(空)"
                is_visible = btn.is_visible()
                if is_visible or "编辑" in text:
                    print(f"    [{idx}] {text[:50]}, visible={is_visible}")
            except:
                pass

        # 查找所有【编辑】按钮，优先选可见的
        edit_buttons = page.locator('button:has-text("编辑")').all()
        print(f"  找到 {len(edit_buttons)} 个【编辑】按钮")

        edit_button = None
        # 优先选择可见的编辑按钮
        for btn in edit_buttons:
            try:
                if btn.is_visible():
                    edit_button = btn
                    print(f"  ✓ 找到可见的【编辑】按钮")
                    break
            except:
                pass

        if edit_button:
            edit_button.click()
            print(f"  ✓ 已点击【编辑】按钮，进入编辑界面")
            page.wait_for_load_state("networkidle", timeout=15000)
            page.wait_for_timeout(2000)
            print(f"  ✓ 编辑界面已加载")
        else:
            print(f"  ⚠️ 未找到可见的【编辑】按钮，请手动点击")
            input("  点击【编辑】后，按 Enter 继续...")

        return True
    else:
        print(f"  ⚠️ 未找到搜索框，请手动搜索: {poster_group_code}")
        input("  搜索完成后，按 Enter 继续...")
        return True


def step7_remove_business_group(page):
    """第七步：删除【海外益智海报-非台湾】分组下的所有海报

    使用 wrapper-scoped 查询确保只删除目标分组内的海报，不影响其他分组。
    """
    print(f"[步骤 7] 删除【{BUSINESS_GROUP_NAME}】下的海报...")
    page.wait_for_timeout(2000)

    # 用 wrapper selector 定位到目标分组的容器
    wrapper = page.locator(GROUP_WRAPPER_SELECTOR).first

    if not wrapper.is_visible():
        print(f"  ⚠️ 未找到分组容器（selector: {GROUP_WRAPPER_SELECTOR}）")
        print(f"  请手动移除该分组下的海报")
        input("  完成后，按 Enter 继续...")
        return True

    print(f"  ✓ 找到分组容器: {BUSINESS_GROUP_NAME}")

    # 安全检查：确保 wrapper 内有数据表格（tr 行），说明这是分组内容而不是分组本身
    table_rows = wrapper.locator('tr').all()
    if not table_rows:
        print(f"  ⚠️ 分组容器内未找到表格行，可能结构不对")
        print(f"  ✓ 分组可能已为空，跳过删除步骤")
        return True

    # while 循环：每次重新查询 wrapper 内的【移除】按钮，逐个点击直到全部删除
    deleted_count = 0
    max_iterations = 50  # 防止无限循环

    for iteration in range(max_iterations):
        # 在 wrapper 范围内查找【移除】按钮（一定在分组内部的表格中）
        # 关键：只查找 wrapper 内的按钮，不会误查全局的【删除分组】按钮
        remove_buttons = wrapper.locator('tr button:has-text("移除")').all()

        if not remove_buttons:
            print(f"  ✓ 该分组下海报已全部移除（共 {deleted_count} 个）")
            return True

        btn = remove_buttons[0]

        if not btn.is_visible():
            print(f"  ⚠️ 第一个【移除】按钮不可见，停止删除")
            print(f"  ✓ 已删除 {deleted_count} 个海报")
            break

        print(f"  [{deleted_count+1}] 点击【移除】按钮...")
        try:
            btn.click()
            page.wait_for_timeout(500)
        except Exception as e:
            print(f"    ⚠️ 点击按钮失败: {e}")
            continue

        # 处理可能的确认弹窗（在整个页面范围查找，因为弹窗是全局的）
        confirmed = False
        for confirm_attempt in range(5):  # 最多尝试 5 次找确认按钮
            try:
                confirm_buttons = page.locator('button:has-text("确定"), button:has-text("确认")').all()
                if confirm_buttons:
                    for confirm_btn in confirm_buttons:
                        try:
                            if confirm_btn.is_visible():
                                confirm_btn.click()
                                page.wait_for_timeout(800)
                                confirmed = True
                                break
                        except:
                            pass
                if confirmed:
                    break
            except:
                pass
            page.wait_for_timeout(200)

        if not confirmed:
            print(f"    ⚠️ 未找到/点击确认按钮，继续")

        deleted_count += 1
        page.wait_for_timeout(500)  # 额外等待 DOM 更新

    print(f"  ⚠️ 删除循环超过 {max_iterations} 次，可能发生了错误")
    print(f"  ✓ 已删除 {deleted_count} 个海报")

    # 重新查询确认是否还有【移除】按钮
    remaining = wrapper.locator('button:has-text("移除")').all()
    if not remaining:
        print(f"  ✓ 确认：该分组下已全部清空")
        return True
    else:
        print(f"  ⚠️ 该分组下仍有 {len(remaining)} 个【移除】按钮未删除")
        return False


def step8_add_sorted_posters(page, sorted_posters_file: Path):
    """第八步：按裂变率从高到低，逐个选择海报

    正确流程：对每个海报 → 打开对话框 → 翻页查找 → 勾选 → 确认关闭 → 等待关闭 → 下一个
    """
    import pandas as pd

    print("[步骤 8] 按裂变率从高到低选择海报...")
    page.wait_for_timeout(2000)

    # 1. 读 Excel
    df = pd.read_excel(sorted_posters_file, sheet_name='裂变率排序')
    poster_ids = [int(x) for x in df['海报ID-汇总处理'].tolist()]
    print(f"  ✓ 读取 {len(poster_ids)} 个海报ID（动态数量）")

    # 2. 定位目标分组 wrapper
    wrapper = page.locator(GROUP_WRAPPER_SELECTOR).first
    if not wrapper.is_visible():
        print(f"  ⚠️ 未找到分组容器")
        return False

    # 3. 对每个海报ID 逐个处理（每次都打开/关闭对话框）
    selected = 0
    not_found = []

    for idx, pid in enumerate(poster_ids):
        print(f"  [{idx+1}/{len(poster_ids)}] 海报 {pid}...", end=" ", flush=True)

        # ===== 步骤A：打开对话框 =====
        try:
            select_btn = wrapper.locator(f'button:has-text("{SELECT_POSTER_BUTTON_TEXT}")').first
            select_btn.click()
            page.wait_for_timeout(1500)
        except Exception as e:
            print(f"❌ 打开弹窗失败")
            continue

        # ===== 步骤B：等待对话框出现 =====
        dialog = page.locator(DIALOG_SELECTOR).first
        try:
            if not dialog.is_visible(timeout=3000):
                print(f"❌ 弹窗未出现")
                continue
        except:
            print(f"❌ 弹窗加载失败")
            continue

        # ===== 步骤C：翻页查找该海报 =====
        found = False
        for page_num in range(1, 251):
            rows = dialog.locator(POSTER_CARD_SELECTOR).all()
            if len(rows) == 0:
                page.wait_for_timeout(500)
                rows = dialog.locator(POSTER_CARD_SELECTOR).all()
                if len(rows) == 0:
                    break

            target_row = None
            for row in rows:
                if str(pid) in row.text_content():
                    target_row = row
                    break

            if target_row:
                # 找到！勾选
                label = target_row.locator('label.el-checkbox').first
                label.click(force=True)
                page.wait_for_timeout(300)
                found = True
                print(f"✓ 已勾选", end=" ", flush=True)
                break

            # 翻到下一页
            pagination = dialog.locator(".el-pagination").first
            if not pagination.is_visible(timeout=1000):
                break
            pag_buttons = pagination.locator("button").all()
            if len(pag_buttons) <= 1:
                break
            next_btn = pag_buttons[1]
            if not next_btn.is_visible() or next_btn.is_disabled():
                break
            next_btn.click()
            page.wait_for_timeout(800)

        if not found:
            print(f"❌ 未找到")
            not_found.append(pid)
            # 关闭对话框（点取消或关闭按钮）
            try:
                cancel_btn = dialog.locator('button:has-text("取消")').first
                if cancel_btn.is_visible():
                    cancel_btn.click()
                    page.wait_for_timeout(800)
            except:
                pass
            continue

        # ===== 步骤D：点击【确认】关闭对话框 =====
        try:
            confirm_btn = dialog.locator(f'button:has-text("{DIALOG_CONFIRM_BUTTON_TEXT}")').first
            confirm_btn.click()
            page.wait_for_timeout(1200)
        except:
            print(f"❌ 确认失败")
            continue

        # ===== 步骤E：等待对话框完全消失 =====
        for _ in range(15):
            try:
                if not dialog.is_visible(timeout=200):
                    break
            except:
                break
            page.wait_for_timeout(200)

        selected += 1
        print(f"✓ 已确认", flush=True)

    # 4. 验证
    page.wait_for_timeout(2000)
    rows = wrapper.locator('tr.el-table__row').all()
    print(f"\n  ✓ 成功添加 {selected}/{len(poster_ids)} 个海报")
    print(f"  ✓ 分组内当前海报数: {len(rows)}")

    if not_found:
        print(f"  ⚠️ 未找到的海报：{not_found}")

    return selected == len(poster_ids)


def step8_5_verify_posters(page, sorted_posters_file: Path):
    """Step 8.5：验证所有海报ID是否都正确添加到分组中

    在保存之前，逐个检查每个海报ID是否都在【海外益智海报-非台湾】分组中
    """
    import pandas as pd

    print("[步骤 8.5] 验证所有海报是否都已正确添加到分组...")
    page.wait_for_timeout(2000)

    # 1. 读 Excel 获取应该添加的海报 ID 列表
    df = pd.read_excel(sorted_posters_file, sheet_name='裂变率排序')
    expected_poster_ids = [int(x) for x in df['海报ID-汇总处理'].tolist()]
    print(f"  ✓ 需要验证的海报ID数: {len(expected_poster_ids)}")
    print(f"    期望的海报ID: {expected_poster_ids}\n")

    # 2. 定位目标分组 wrapper
    wrapper = page.locator(GROUP_WRAPPER_SELECTOR).first
    if not wrapper.is_visible():
        print(f"  ❌ 未找到分组容器")
        return False

    # 3. 获取分组中所有的海报行
    print("  [子步骤] 读取分组中的所有海报...")
    rows = wrapper.locator('tr.el-table__row').all()
    print(f"  ✓ 分组内共有 {len(rows)} 个海报\n")

    # 4. 提取每行中的海报ID
    print("  [子步骤] 逐行检查海报ID...")
    found_poster_ids = []
    not_found_ids = []

    for idx, row in enumerate(rows):
        row_text = row.text_content()

        # 尝试从行文本中提取海报ID
        found_in_row = False
        for expected_id in expected_poster_ids:
            if str(expected_id) in row_text:
                found_poster_ids.append(expected_id)
                print(f"    [{idx+1}/{len(rows)}] ✓ 海报 {expected_id} 已找到")
                found_in_row = True
                break

        if not found_in_row:
            print(f"    [{idx+1}/{len(rows)}] ⚠️ 该行海报ID: {row_text[:50]}...")

    # 5. 验证所有期望的ID都被找到
    print(f"\n  [子步骤] 验证结果...")
    for expected_id in expected_poster_ids:
        if expected_id not in found_poster_ids:
            not_found_ids.append(expected_id)

    print(f"  ✓ 找到 {len(found_poster_ids)}/{len(expected_poster_ids)} 个海报")

    if not_found_ids:
        print(f"  ❌ 以下海报未找到: {not_found_ids}")
        print(f"\n  ⚠️ 验证失败！请检查是否所有海报都已添加")
        return False
    else:
        print(f"  ✓ 所有 {len(expected_poster_ids)} 个海报都已正确添加到分组\n")
        return True


def step9_save(page):
    """第九步：保存更新

    在点击保存前，需要验证其他业务分组未被改动，然后滑动到最底部点击保存
    """
    print("[步骤 9] 保存更新...")
    page.wait_for_timeout(2000)

    # 定义要检查的其他分组名称（排除【海外益智海报-非台湾】）
    other_groups = [
        "海外非港澳台益智",
        "台湾益智",
        "国内益智",
        "海外益智-wonderclass"
    ]

    print("\n  [子步骤] 验证其他分组未被改动...")

    for group_name in other_groups:
        try:
            # 查找该分组的 wrapper
            group_wrapper = page.locator(f'div.oneGroup:has(p:has-text("{group_name}"))').first
            if group_wrapper.is_visible(timeout=2000):
                # 检查该分组内的海报行数
                poster_rows = group_wrapper.locator('tr.el-table__row').all()

                print(f"    ✓ 【{group_name}】: {len(poster_rows)} 个海报（未改动）")
            else:
                print(f"    ⚠️ 【{group_name}】: 分组未找到或已隐藏")
        except Exception as e:
            print(f"    ⚠️ 【{group_name}】: 检查失败 - {str(e)[:50]}")

    print(f"\n  ✓ 所有其他分组验证完成\n")

    # ============ 滑动到页面最底部 ============
    print("  [子步骤] 滑动到页面最底部...")

    # 多次滑动以确保到达最底部（可能是嵌套的滚动容器）
    for i in range(10):
        page.evaluate("window.scrollBy(0, window.innerHeight)")
        page.wait_for_timeout(200)

    # 直接滑动到最底部
    page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
    page.wait_for_timeout(800)

    # 也尝试滚动任何可能的容器
    page.evaluate("""
        const scrollable = document.querySelector('.el-scrollbar__wrap') ||
                          document.querySelector('[class*="scroll"]') ||
                          document.documentElement;
        if (scrollable) {
            scrollable.scrollTop = scrollable.scrollHeight;
        }
    """)
    page.wait_for_timeout(800)

    # 打印页面滚动信息（调试用）
    scroll_info = page.evaluate("""
        ({
            bodyHeight: document.body.scrollHeight,
            windowHeight: window.innerHeight,
            currentScroll: window.scrollY,
            documentElement: document.documentElement.scrollHeight
        })
    """)
    print(f"    页面滚动信息: 总高度={scroll_info['bodyHeight']}px, 当前位置={scroll_info['currentScroll']}px")

    print(f"  ✓ 已滑动到页面底部\n")

    # ============ 查找并点击保存按钮 ============
    print("  [子步骤] 查找并点击【保存】按钮...\n")

    # 先寻找【保存】按钮
    save_button = page.locator('button:has-text("保存")').first

    if not save_button.is_visible(timeout=1000):
        # 如果没找到【保存】，再找其他可能的按钮
        save_button = page.locator('button:has-text("提交")').first
        if not save_button.is_visible(timeout=1000):
            save_button = page.locator('button:has-text("确定")').first

    try:
        if save_button.is_visible(timeout=2000):
            btn_text = save_button.text_content().strip()
            print(f"    ✓ 找到【{btn_text}】按钮")

            # 尝试使用 evaluate 将按钮滚动到视图中
            page.evaluate("""
                const btn = document.evaluate(
                    "//button[contains(text(), '保存')]",
                    document,
                    null,
                    XPathResult.FIRST_ORDERED_NODE_TYPE,
                    null
                ).singleNodeValue;
                if (btn) {
                    btn.scrollIntoView({ behavior: 'smooth', block: 'center' });
                }
            """)
            page.wait_for_timeout(500)

            # 点击保存按钮
            print(f"    ⏳ 点击【{btn_text}】按钮...")
            save_button.click()
            print(f"    ✓ 已点击【{btn_text}】按钮")
            page.wait_for_timeout(1500)

            # 等待保存完成
            try:
                page.wait_for_load_state("networkidle", timeout=15000)
                print(f"    ✓ 保存成功\n")
            except:
                print(f"    ✓ 保存请求已发送\n")

            return True
        else:
            print(f"    ⚠️ 【保存】按钮未在视图范围内\n")
    except Exception as e:
        print(f"    ⚠️ 操作异常: {str(e)[:80]}\n")

    print(f"  ⚠️ 保存按钮查找/点击失败，请确认【保存】按钮是否可见")
    input("  确认保存完成后，按 Enter 继续...")
    return True


def main():
    """主流程 - 在共享浏览器会话中运行所有步骤"""
    from playwright.sync_api import sync_playwright
    import os

    print("=" * 60)
    print("APP打卡海报日维度更新 - 自动化工作流")
    print("=" * 60)
    print()

    with sync_playwright() as p:
        # 创建下载目录
        SAMPLE_DIR.mkdir(exist_ok=True)
        print(f"[设置] 下载目录: {SAMPLE_DIR}")

        # 启动浏览器（一次）
        browser = p.chromium.launch(headless=False)

        # 创建上下文，启用下载功能并指定下载目录
        context = browser.new_context(
            accept_downloads=True,
            ignore_https_errors=True
        )

        # 设置下载目录（这样文件会自动保存到指定目录）
        context.tracing.start(screenshots=False, snapshots=False)

        page = context.new_page()

        try:
            # 第一步：登录
            if not step1_login_sensors(page):
                print("❌ 登录失败")
                context.close()
                browser.close()
                return

            input("  按 Enter 继续到第二步...")

            # 第二步：进入书签
            if not step2_enter_bookmark(page):
                print("⚠️ 第二步未完成，请检查")
                input("  按 Enter 继续到第三步...")

            input("  按 Enter 继续到第三步...")

            # 第三步：导出数据
            export_file = step3_export_data(page, context)
            if export_file:
                print(f"  ✓ 数据导出成功: {export_file}")
                # 检查文件是否存在
                if export_file.exists():
                    file_size = export_file.stat().st_size
                    print(f"  文件大小: {file_size / 1024:.2f} KB")
                else:
                    print(f"  ⚠️ 文件不存在: {export_file}")
            else:
                print("  ⚠️ 数据导出可能未成功")

            input("  按 Enter 继续到第四步...")

            # 第四步：计算裂变率
            calc_file = step4_calculate_conversion_rate(export_file)
            if calc_file:
                print(f"  ✓ 裂变率计算成功")
            else:
                print("  ⚠️ 裂变率计算失败")

            input("  按 Enter 继续到第五步...")

            # 第五步：登录平台中心
            if not step5_login_bizcenter(page):
                print("❌ 平台中心登录失败")

            input("  按 Enter 继续到第六步...")

            # 第六步：查询海报组
            if not step6_search_poster_group(page):
                print("⚠️ 海报组查询失败")

            input("  按 Enter 继续到第七步...")

            # 第七步：移除旧业务分组
            if not step7_remove_business_group(page):
                print("⚠️ 业务分组移除失败")

            input("  按 Enter 继续到第八步...")

            # 第八步：输入新海报ID
            if calc_file and calc_file.exists():
                if not step8_add_sorted_posters(page, calc_file):
                    print("⚠️ 海报ID输入失败")
            else:
                print("⚠️ 未找到排序文件，跳过海报输入")

            input("  按 Enter 继续到第九步...")

            # 第九步：保存更新
            if not step9_save(page):
                print("⚠️ 保存失败")

            input("  按 Enter 完成...")

            print("\n✓ 所有步骤已完成！")

        except Exception as e:
            print(f"\n❌ 错误: {e}")
            import traceback
            traceback.print_exc()
        finally:
            context.close()
            browser.close()

        # 最后检查 sample 目录中的文件
        print("\n[检查] sample 目录中的文件:")
        if SAMPLE_DIR.exists():
            files = list(SAMPLE_DIR.glob("*.*"))
            if files:
                for f in sorted(files, key=lambda x: x.stat().st_mtime, reverse=True)[:5]:
                    size_kb = f.stat().st_size / 1024
                    mtime = f.stat().st_mtime
                    print(f"  - {f.name} ({size_kb:.2f} KB)")
            else:
                print("  (目录为空)")


if __name__ == "__main__":
    main()
