"""
CRM 自动创建渠道 + 推广 + 二维码 工具

用法：
1. 安装依赖：
   pip install playwright requests
   playwright install chromium

2. 复制配置模板：
   cp config.example.json config.local.json
   # 然后编辑 config.local.json，填入你公司 CRM 的真实业务参数

3. 运行（交互式）：
   python crm_create_channel.py

4. 运行（飞书多维表格轮询）：
   python feishu_bitable_worker.py

脚本将自动完成三段操作：
- 创建渠道
- 创建推广（推广名称复用渠道名称）
- 创建推广二维码（在工作台生成短链并下载二维码）

最终输出推广链接和二维码文件。

注意：
- 真实业务参数请仅放在 config.local.json，该文件已被 .gitignore 忽略，不会进入版本库
- 若 CRM 改版导致按钮文字 / placeholder / DOM 结构变更，可在 config.local.json 中调整对应字段
"""

import json
import os
import time
from pathlib import Path
from playwright.sync_api import sync_playwright


# ============ 配置加载 ============

CONFIG_FILE = Path(__file__).parent / "config.local.json"
EXAMPLE_FILE = Path(__file__).parent / "config.example.json"
DOWNLOAD_DIR = os.path.dirname(os.path.abspath(__file__))
AUTH_STATE_FILE = os.path.join(DOWNLOAD_DIR, "auth_state.json")


def load_config() -> dict:
    if not CONFIG_FILE.exists():
        raise FileNotFoundError(
            f"未找到配置文件 {CONFIG_FILE.name}。\n"
            f"请将 {EXAMPLE_FILE.name} 复制为 {CONFIG_FILE.name}，并填入真实参数。"
        )
    with open(CONFIG_FILE, encoding="utf-8") as f:
        return json.load(f)


# ============ 通用工具函数（与具体业务无关） ============

def fill_filterable_dropdown(page, label_text: str, value: str, timeout: int = 10000):
    """
    Element UI 可搜索下拉框：
    1. 点击 label 旁的下拉框触发器，让它获得焦点并展开浮层
    2. 用键盘直接输入文字（Element UI 会用焦点上的输入触发筛选）
    3. 点击浮层中匹配的选项

    注意：触发器本身可能是 readonly 的，所以不能 fill，必须用键盘 type。
    """
    form_item = page.locator(f".el-form-item:has-text('{label_text}')").first
    if form_item.count() == 0:
        form_item = page.locator(f"text={label_text}").first.locator("..")
    trigger = form_item.locator(".el-select, .el-input, input").first
    trigger.click(timeout=timeout)
    time.sleep(0.8)

    page.keyboard.type(value, delay=50)
    time.sleep(2.0)

    panel_selectors = [
        ".el-select-dropdown:visible",
        ".el-popper:visible",
        ".el-cascader-panel:visible",
    ]
    for panel_sel in panel_selectors:
        try:
            panel = page.locator(panel_sel).last
            if panel.count() == 0:
                continue
            exact_option = panel.locator(f"li:has-text('{value}')").filter(visible=True).first
            if exact_option.count() > 0:
                exact_option.click(timeout=timeout)
                return
            first_option = panel.locator("li.el-select-dropdown__item, li").filter(visible=True).first
            if first_option.count() > 0:
                first_option.click(timeout=timeout)
                return
        except Exception:
            continue

    print(f"❌ 没找到 '{label_text}' 下拉的 '{value}' 选项")
    print("当前可见的下拉选项：")
    try:
        items = page.locator(".el-select-dropdown:visible li, .el-popper:visible li").all()
        for el in items[:30]:
            try:
                txt = el.inner_text(timeout=300).strip()
                if txt:
                    print(f"  - {txt!r}")
            except Exception:
                continue
    except Exception:
        pass
    raise RuntimeError(f"找不到下拉选项：{label_text} -> {value}")


def fill_input_by_label(page, label_text: str, value: str, timeout: int = 10000):
    """
    通过 label 文字定位旁边的 input 并填入。
    适用于 placeholder 是"请输入"等通用文案、无法靠 placeholder 区分的字段。
    依次尝试四种父节点层级。
    """
    candidates = [
        page.locator(f".el-form-item:has-text('{label_text}')").first,
        page.locator(f"text={label_text}").first.locator(".."),
        page.locator(f"text={label_text}").first.locator("..").locator(".."),
        page.locator(f"text={label_text}").first.locator("../../.."),
    ]
    for container in candidates:
        try:
            if container.count() == 0:
                continue
            inp = container.locator("input:visible, textarea:visible").first
            if inp.count() > 0:
                inp.fill(value, timeout=timeout)
                return
        except Exception:
            continue
    raise RuntimeError(f"找不到 label '{label_text}' 旁的输入框")


def fill_input_by_placeholder(page, placeholder_keyword: str, value: str, timeout: int = 10000):
    """通过 placeholder 关键词模糊匹配输入框，失败时列出所有可见输入框"""
    try:
        loc = page.get_by_placeholder(placeholder_keyword).filter(visible=True).first
        if loc.count() > 0:
            loc.fill(value, timeout=timeout)
            return
    except Exception:
        pass

    selectors = [
        f"input[placeholder*='{placeholder_keyword}']",
        f"textarea[placeholder*='{placeholder_keyword}']",
    ]
    for sel in selectors:
        try:
            loc = page.locator(sel).filter(visible=True).first
            if loc.count() > 0:
                loc.fill(value, timeout=timeout)
                return
        except Exception:
            continue

    print(f"❌ 没找到 placeholder 含 '{placeholder_keyword}' 的输入框。当前可见输入框：")
    try:
        all_inputs = page.locator("input:visible, textarea:visible").all()
        for el in all_inputs[:30]:
            try:
                ph = el.get_attribute("placeholder") or ""
                name = el.get_attribute("name") or ""
                print(f"  - placeholder={ph!r}, name={name!r}")
            except Exception:
                continue
    except Exception:
        pass
    raise RuntimeError(f"找不到输入框：placeholder 含 {placeholder_keyword!r}")


def select_dropdown_option(page, label_text: str, option_text: str, timeout: int = 10000):
    """普通下拉选择：点击触发器 → 在浮层里点选项"""
    clicked = False
    try:
        loc = page.locator(f"text={label_text}").first.locator("..").locator(
            ".el-select, .el-input, select, .ant-select"
        ).first
        if loc.count() > 0:
            loc.click(timeout=timeout)
            clicked = True
    except Exception:
        pass
    if not clicked:
        page.locator(f".el-form-item:has-text('{label_text}')").first.locator(".el-select").first.click(timeout=timeout)

    time.sleep(0.5)

    dropdown_panel_selectors = [
        ".el-select-dropdown:visible",
        ".el-cascader-panel:visible",
        ".ant-select-dropdown:visible",
    ]
    for panel_sel in dropdown_panel_selectors:
        try:
            panel = page.locator(panel_sel).last
            if panel.count() > 0:
                option = panel.locator(
                    f"li:has-text('{option_text}'), .el-select-dropdown__item:has-text('{option_text}')"
                ).filter(visible=True).first
                if option.count() > 0:
                    option.click(timeout=timeout)
                    return
        except Exception:
            continue

    print(f"❌ 下拉框 '{label_text}' 中没找到选项 '{option_text}'。当前可见的下拉选项：")
    try:
        items = page.locator(".el-select-dropdown:visible li, .ant-select-dropdown:visible li").all()
        for el in items[:30]:
            try:
                txt = el.inner_text(timeout=300).strip()
                if txt:
                    print(f"  - {txt!r}")
            except Exception:
                continue
    except Exception:
        pass
    raise RuntimeError(f"找不到下拉选项：{label_text} -> {option_text}")


def click_button_with_text(page, text: str, timeout: int = 10000):
    """点击包含指定文字的按钮，兼容 Element UI 的"确 定"风格（中文按钮自动加空格）"""
    text_variants = [text]
    if len(text) == 2:
        text_variants.append(" ".join(text))

    for txt in text_variants:
        selectors = [
            f"button:has-text('{txt}')",
            f"[role='button']:has-text('{txt}')",
            f".el-button:has-text('{txt}')",
            f".ant-btn:has-text('{txt}')",
        ]
        for sel in selectors:
            try:
                loc = page.locator(sel).filter(visible=True).last
                if loc.count() > 0:
                    loc.click(timeout=timeout)
                    return
            except Exception:
                continue

    print(f"❌ 没找到文字含 '{text}' 的按钮。当前页面可见按钮文字：")
    try:
        all_btns = page.locator("button:visible, [role='button']:visible").all()
        for b in all_btns[:30]:
            try:
                txt2 = b.inner_text(timeout=300).strip()
                if txt2:
                    print(f"  - {txt2!r}")
            except Exception:
                continue
    except Exception:
        pass
    raise RuntimeError(f"找不到按钮：{text}")


# ============ 业务流程：创建渠道 ============

def create_channel(page, channel_name: str, cfg: dict):
    """
    第一部分：【创建渠道】
    要求 page 已经进入 CRM 主页。
    """
    print("\n" + "=" * 50)
    print("  开始执行【创建渠道】")
    print("=" * 50)

    ch = cfg["channel"]

    # 1. 推广管理 → 渠道管理 → 创建渠道
    page.get_by_text("推广管理").first.click()
    time.sleep(1)
    page.get_by_text("渠道管理").first.click()
    time.sleep(1)
    page.get_by_text("创建渠道").first.click()
    time.sleep(1)

    # 2. 渠道名称
    fill_input_by_placeholder(page, "渠道名称", channel_name)
    time.sleep(0.5)

    # 3. 选择渠道（按渠道编号搜索）
    page.get_by_text("选择渠道").first.click()
    time.sleep(1)
    fill_input_by_placeholder(page, "渠道编号", ch["channel_no"])
    click_button_with_text(page, "查询")
    time.sleep(1)
    page.get_by_text(ch["channel_no"]).first.click()
    time.sleep(0.5)
    click_button_with_text(page, "确定")
    time.sleep(1)

    # 4. 渠道等级
    select_dropdown_option(page, "渠道等级", ch["channel_level"])
    time.sleep(0.5)

    # 5. 选择业务线
    page.get_by_text(ch["business_line"]).first.click()
    time.sleep(1)

    # 6. 点击"创建 X 渠道"（X 是业务线名）
    page.get_by_text(f"创建{ch['business_line']}渠道").first.click()
    time.sleep(1)

    # 7. 子渠道信息（复用名称 + 渠道 ID）
    page.get_by_placeholder("请输入渠道名称").last.fill(channel_name)
    time.sleep(0.5)
    page.get_by_text("选择渠道").last.click()
    time.sleep(1)
    fill_input_by_placeholder(page, "渠道名称或ID", ch["sub_channel_id"])
    click_button_with_text(page, "查询")
    time.sleep(1)
    page.get_by_text(ch["sub_channel_id"]).first.click()
    time.sleep(0.5)
    click_button_with_text(page, "确定")
    time.sleep(1)

    # 8. 渠道评级、所属市场、默认入库
    select_dropdown_option(page, "渠道评级", ch["channel_rating"])
    time.sleep(0.5)
    fill_filterable_dropdown(page, "所属市场", ch["operator_name"])
    time.sleep(0.5)
    select_dropdown_option(page, "默认入库", ch["default_storage"])
    time.sleep(0.5)
    click_button_with_text(page, "确认")
    time.sleep(1)
    click_button_with_text(page, "确定")
    time.sleep(1)

    # 9. 保存
    click_button_with_text(page, "保存")
    time.sleep(2)

    print(">>> 【创建渠道】完成！")


# ============ 业务流程：创建推广 ============

def create_promotion(page, channel_name: str, cfg: dict) -> str:
    """
    第二部分：【创建推广】
    返回最终生成的推广链接（如能从页面解析出）。
    """
    print("\n" + "=" * 50)
    print("  开始执行【创建推广】")
    print("=" * 50)

    pr = cfg["promotion"]
    home_url = cfg["crm"]["home_url"]

    # 1. 回到 CRM 主页
    page.goto(home_url)
    page.wait_for_load_state("networkidle")
    time.sleep(2)

    # 2. 推广管理 → 推广列表
    promo_list = page.get_by_text("推广列表").filter(visible=True)
    if promo_list.count() == 0:
        page.get_by_text("推广管理").filter(visible=True).first.click()
        time.sleep(1)
    page.get_by_text("推广列表").filter(visible=True).first.click()
    time.sleep(1)

    # 3. 点击【创建推广】，在"请选择推广页面类型"弹窗里选 H5 + 选择码良页面，确 定
    page.get_by_text("创建推广").first.click()
    time.sleep(1.5)

    dialog = page.locator("[role='dialog'][aria-label='请选择推广页面类型']").filter(visible=True).last

    # 选页面类型
    page_type_clicked = False
    for sel in [
        f"tr:has-text('{pr['page_type']}') input[type='radio']",
        f"tr:has-text('{pr['page_type']}') .el-radio",
        f"tr:has-text('{pr['page_type']}') td",
        f".el-radio:has-text('{pr['page_type']}')",
        f"label:has-text('{pr['page_type']}')",
    ]:
        try:
            loc = dialog.locator(sel).filter(visible=True).first
            if loc.count() > 0:
                loc.click(timeout=5000)
                page_type_clicked = True
                break
        except Exception:
            continue
    if not page_type_clicked:
        raise RuntimeError(f"找不到页面类型行：{pr['page_type']}")
    time.sleep(0.8)

    # 选页面配置
    page_config_clicked = False
    for sel in [
        f"tr:has-text('{pr['page_config']}') input[type='radio']",
        f"tr:has-text('{pr['page_config']}') .el-radio",
        f"tr:has-text('{pr['page_config']}') td",
        f".el-radio:has-text('{pr['page_config']}')",
        f"label:has-text('{pr['page_config']}')",
    ]:
        try:
            loc = dialog.locator(sel).filter(visible=True).first
            if loc.count() > 0:
                loc.click(timeout=5000)
                page_config_clicked = True
                break
        except Exception:
            continue
    if not page_config_clicked:
        raise RuntimeError(f"找不到页面配置行：{pr['page_config']}")
    time.sleep(0.8)

    # 弹窗内的【确 定】会打开新标签页，捕获之
    context = page.context
    dialog_ok = dialog.locator("button:has-text('确'), .el-button:has-text('确')").filter(visible=True).last
    try:
        with context.expect_page(timeout=10000) as new_page_info:
            if dialog_ok.count() > 0:
                dialog_ok.click(timeout=10000)
            else:
                click_button_with_text(page, "确定")
        new_page = new_page_info.value
        new_page.wait_for_load_state("networkidle")
        new_page.bring_to_front()
        page = new_page
    except Exception:
        time.sleep(2)
        for pg in context.pages:
            if "popularize/newEdit" in pg.url or "promotionPageType" in pg.url:
                pg.bring_to_front()
                page = pg
                break
    time.sleep(2)

    # 4. 推广名称
    fill_input_by_label(page, "推广名称", channel_name)
    time.sleep(0.5)

    # 5. 链接域名 → 自定义 → 域名
    page.get_by_text("自定义", exact=True).first.click()
    time.sleep(0.8)
    select_dropdown_option(page, "链接域名", pr["link_domain"])
    time.sleep(0.5)

    # 6. 关联商品
    page.get_by_text("请选择绑定一个单品或组合商品").first.click()
    time.sleep(1)
    fill_input_by_label(page, "商品ID", pr["product_id"])
    click_button_with_text(page, "查询")
    time.sleep(1)
    page.get_by_text(pr["product_id"]).first.click()
    time.sleep(0.5)
    click_button_with_text(page, "确定")
    time.sleep(1)

    # 7. 选择平台渠道
    page.get_by_text("请选择关联一个平台渠道").first.click()
    time.sleep(1)
    fill_input_by_label(page, "渠道名称", channel_name)
    click_button_with_text(page, "查询")
    time.sleep(1)
    page.get_by_text(channel_name).last.click()
    time.sleep(0.5)
    click_button_with_text(page, "确定")
    time.sleep(1)

    # 8. 选择落地页（远程搜索需要额外等待）
    time.sleep(2)
    fill_filterable_dropdown(page, "选择落地页", pr["landing_page_id"])
    time.sleep(0.5)

    # 9. 静默授权公众号 + 验证码短信签名
    select_dropdown_option(page, "静默授权公众号", pr["wechat_account"])
    time.sleep(0.5)
    select_dropdown_option(page, "验证码短信签名", pr["sms_signature"])
    time.sleep(0.5)

    # 10. 发布推广
    click_button_with_text(page, "发布推广")
    time.sleep(2)

    # 11. 等待"推广创建成功"，点击复制链接，读出推广链接
    try:
        page.get_by_text("推广创建成功").first.wait_for(state="visible", timeout=15000)
    except Exception:
        print(">>> ⚠️ 未检测到【推广创建成功】文字，仍尝试点击【复制链接】")
    time.sleep(1)

    click_button_with_text(page, "复制链接")
    time.sleep(1)

    promo_link = ""
    for sel in [
        "input[readonly][value^='http']",
        "input[value^='http']",
        ".el-dialog:visible input[value^='http']",
        "textarea:has-text('http')",
    ]:
        try:
            loc = page.locator(sel).filter(visible=True).first
            if loc.count() > 0:
                val = loc.input_value(timeout=2000)
                if val and val.startswith("http"):
                    promo_link = val
                    break
        except Exception:
            continue
    if not promo_link:
        try:
            text = page.locator(".el-dialog:visible, .el-message-box:visible, body").first.inner_text(timeout=2000)
            import re
            m = re.search(r"https?://\S+", text)
            if m:
                promo_link = m.group(0)
        except Exception:
            pass

    print(">>> 【创建推广】完成！")
    return promo_link


# ============ 业务流程：创建推广二维码 ============

def create_qrcode(page, channel_name: str, promo_link: str, cfg: dict) -> str:
    """
    第三部分：【创建推广二维码】
    在工作台创建短链并下载二维码。返回二维码文件路径。
    """
    print("\n" + "=" * 50)
    print("  开始执行【创建推广二维码】")
    print("=" * 50)

    qr = cfg["qrcode"]

    # 1. 打开工作台
    print("\n>>> 第一步：打开工作台")
    page.bring_to_front()
    page.goto(qr["workbench_url"])
    page.wait_for_load_state("networkidle")
    time.sleep(3)

    # 2. 技术设置 → 短链管理
    print(">>> 第二步：点击【技术设置】→【短链管理】")
    tech_menu = page.get_by_text("技术设置").filter(visible=True)
    if tech_menu.count() == 0:
        raise RuntimeError("找不到【技术设置】菜单")
    tech_menu.first.click()
    time.sleep(1)
    page.get_by_text("短链管理").filter(visible=True).first.click()
    time.sleep(3)
    page.wait_for_load_state("networkidle")

    # 3. 新建短链
    print(">>> 第三步：点击【新建短链】")
    click_button_with_text(page, "新建短链")
    time.sleep(1.5)

    # 4. 填写弹窗
    print(">>> 第四步：填写短链信息")
    time.sleep(1)
    dialog = page.locator("[role='dialog']").filter(visible=True).last
    if dialog.count() == 0:
        dialog = page.locator(".el-dialog:visible").last

    # 源网址
    try:
        inp = dialog.locator("input:visible, textarea:visible").first
        if inp.count() > 0:
            inp.fill(promo_link)
    except Exception:
        fill_input_by_label(page, "源网址", promo_link)
    time.sleep(0.5)

    fill_input_by_label(page, "短链标题", channel_name)
    time.sleep(0.5)
    select_dropdown_option(page, "短链分组", qr["short_link_group"])
    time.sleep(0.5)
    select_dropdown_option(page, "选择域名", qr["short_link_domain"])
    time.sleep(0.5)

    # 确 定
    dialog_ok = dialog.locator("button:has-text('确'), .el-button:has-text('确')").filter(visible=True).last
    if dialog_ok.count() > 0:
        dialog_ok.click(timeout=10000)
    else:
        click_button_with_text(page, "确定")
    time.sleep(2)

    # 处理成功弹窗
    try:
        success_dialog = page.locator(".el-message-box:visible, .el-dialog:visible").filter(
            has_text="成功"
        ).last
        if success_dialog.count() > 0:
            ok_btn = success_dialog.locator(
                "button:has-text('确'), .el-button:has-text('确')"
            ).filter(visible=True).last
            if ok_btn.count() > 0:
                ok_btn.click(timeout=5000)
        else:
            try:
                click_button_with_text(page, "确定")
            except Exception:
                pass
    except Exception:
        pass
    time.sleep(2)

    # 5. 切换到分组 → 找到行 → 下载二维码
    print(">>> 第五步：点击二维码图标 → 下载二维码")

    # 点击左侧分组
    try:
        group_node = page.get_by_text(qr["short_link_group"], exact=True).filter(visible=True).first
        if group_node.count() == 0:
            group_node = page.locator(
                f".el-tree :text('{qr['short_link_group']}'), .el-menu :text('{qr['short_link_group']}')"
            ).filter(visible=True).first
        group_node.click(timeout=5000)
    except Exception:
        pass
    time.sleep(2)
    page.wait_for_load_state("networkidle")

    # 找到行
    row = page.locator(f"tr:has-text('{channel_name}')").filter(visible=True).first
    if row.count() == 0:
        page.reload()
        page.wait_for_load_state("networkidle")
        time.sleep(2)
        row = page.locator(f"tr:has-text('{channel_name}')").filter(visible=True).first

    # 行内找二维码图标
    qr_icon_clicked = False
    for sel in [
        "[aria-label*='二维码']", "[title*='二维码']",
        ".qr-code, .qrcode, .icon-qrcode",
        "svg[class*='qr']", "img[src*='qr']", "i[class*='qr']",
    ]:
        try:
            icon = row.locator(sel).filter(visible=True).first
            if icon.count() > 0:
                icon.click(timeout=5000)
                qr_icon_clicked = True
                break
        except Exception:
            continue
    if not qr_icon_clicked:
        raise RuntimeError("找不到二维码图标")
    time.sleep(1)

    # 下载二维码
    qr_file = ""
    try:
        with page.expect_download(timeout=15000) as download_info:
            click_button_with_text(page, "下载二维码")
        download = download_info.value
        save_name = channel_name.replace("/", "_").replace("\\", "_") + ".jpg"
        save_path = os.path.join(DOWNLOAD_DIR, save_name)
        download.save_as(save_path)
        qr_file = save_path
        print(f">>>   二维码已保存：{save_path}")
    except Exception as e:
        print(f">>>   ⚠️ 下载捕获失败（{e}）")

    print("\n>>> 【创建推广二维码】完成！")
    return qr_file


# ============ 登录状态管理 ============

def ensure_login(browser, cfg: dict = None):
    """
    确保登录状态有效，返回 (context, page)。
    如果 auth_state.json 存在且有效，直接复用；否则等待扫码登录。
    """
    if cfg is None:
        cfg = load_config()
    home_url = cfg["crm"]["home_url"]

    if os.path.exists(AUTH_STATE_FILE):
        print(">>> 检测到已保存的登录状态，尝试复用...")
        context = browser.new_context(
            accept_downloads=True, storage_state=AUTH_STATE_FILE
        )
    else:
        context = browser.new_context(accept_downloads=True)

    page = context.new_page()
    page.goto(home_url)
    page.wait_for_load_state("networkidle")
    time.sleep(3)

    def _on_crm_page(url: str) -> bool:
        return url.startswith(home_url.split("#")[0]) and "/login" not in url

    need_login = not _on_crm_page(page.url)
    if need_login:
        print(">>> 登录状态已过期或不存在，需要扫码登录")
        input(">>> 请完成扫码登录，登录成功后按回车继续...")
        page.goto(home_url)
        page.wait_for_load_state("networkidle")
        time.sleep(3)
        if not _on_crm_page(page.url):
            raise RuntimeError(f"跳转失败（URL={page.url}），请确认已登录")
        context.storage_state(path=AUTH_STATE_FILE)
        print(f">>> 登录状态已保存到 {AUTH_STATE_FILE}")
    else:
        print(f">>> 登录状态有效，已进入 CRM 页面")

    return context, page


# ============ 主流程 ============

def run_single(channel_name: str, browser=None, context=None, page=None, cfg: dict = None):
    """
    纯执行逻辑（无 input 阻塞），供 worker 或外部调用。
    返回 (promo_link, qr_file_path)。
    """
    if cfg is None:
        cfg = load_config()

    own_browser = browser is None
    pw_context = None

    try:
        if own_browser:
            pw_context = sync_playwright().start()
            browser = pw_context.chromium.launch(headless=False)
            context, page = ensure_login(browser, cfg)

        # 每次执行前先回到 CRM 主页（上一次任务可能停留在工作台或其他页面）
        page.goto(cfg["crm"]["home_url"])
        page.wait_for_load_state("networkidle")
        time.sleep(2)

        create_channel(page, channel_name, cfg)
        promo_link = create_promotion(page, channel_name, cfg)

        qr_file = ""
        if promo_link:
            qr_file = create_qrcode(page, channel_name, promo_link, cfg)
        else:
            print("\n⚠️ 未获取到推广链接，跳过二维码创建")

        return promo_link, qr_file
    finally:
        if own_browser:
            browser.close()
            if pw_context:
                pw_context.stop()


def run_interactive(channel_name: str, cfg: dict):
    """交互式入口（手动运行时用）"""
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context, page = ensure_login(browser, cfg)

        promo_link, qr_file = run_single(
            channel_name, browser=browser, context=context, page=page, cfg=cfg
        )

        print("\n" + "=" * 50)
        print(">>> 所有操作完成！")
        print(f">>> 渠道名称：{channel_name}")
        if promo_link:
            print(f">>> 推广链接：{promo_link}")
        if qr_file:
            print(f">>> 二维码文件：{qr_file}")
        print("=" * 50)
        input(">>> 按回车关闭浏览器...")
        browser.close()


if __name__ == "__main__":
    print("=" * 50)
    print("  CRM 自动创建渠道 + 推广 + 二维码 工具")
    print("=" * 50)
    cfg = load_config()
    name = input("请输入【渠道名称】: ").strip()
    if not name:
        print("渠道名称不能为空！")
    else:
        run_interactive(name, cfg)
