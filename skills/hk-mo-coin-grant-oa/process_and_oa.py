"""
港澳渠道成交用户赠送豌豆币 OA 申请自动化
Step 1: 处理销售明细 → 生成豌豆币发放表格
Step 2: 登录 OA 系统 → 填写豌豆币申请表单 → 暂存草稿
"""
import os
import sys
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

sys.stdout.reconfigure(encoding="utf-8")

SCRIPT_DIR = Path(__file__).parent
SAMPLE_DIR = SCRIPT_DIR / "sample"
OUTPUT_DIR = SCRIPT_DIR / "output"
AUTH_STATE_FILE = SCRIPT_DIR / "oa_auth_state.json"

SALES_FILE = SAMPLE_DIR / "海外用户销售明细_末次渠道.xlsx"
SALES_HEADER_ROW = 7
COL_STUDENT_ID = "学员ID"
COL_FIRST_ORDER = "首签订单号"

VIRTUAL_AMOUNT = 32000
TEMPLATE_COLUMNS = [
    "豌豆/魔力用户id*\n（禁止用豌豆用户id发送魔力币\n或魔力用户id发送豌豆币）",
    "学科品类（\n豌豆：VIP_WanDou \n魔力：VIP_MoLi）*",
    "发放虚拟币数量*",
    "发放原因*",
    "发放备注*",
    "部门归属:第一级id*",
    "部门归属:最后一级id*",
    "部门归属:发放类型*",
    "关联的订单号",
    "是否合同内",
]
FIXED_VALUES = {
    "学科品类": "VIP_WanDou",
    "发放原因": "新签套餐奖励",
    "发放备注": "港澳商务-新签成交用户奖励",
    "第一级id": 5672,
    "最后一级id": 6204,
    "发放类型": 16,
    "是否合同内": "是",
}

OA_LOGIN_URL = "<your-sso-login-url>"
OA_PORTAL_URL = "<your-oa-portal-url>"


def get_last_week_range():
    """计算上周一到上周日，返回 (周一日期, 周日日期, 'M.D-M.D' 格式字符串)"""
    today = datetime.now()
    days_since_monday = today.weekday()
    this_monday = today - timedelta(days=days_since_monday)
    last_monday = this_monday - timedelta(days=7)
    last_sunday = last_monday + timedelta(days=6)
    range_str = f"{last_monday.month}.{last_monday.day}-{last_sunday.month}.{last_sunday.day}"
    return last_monday, last_sunday, range_str


def make_output_name():
    _, _, range_str = get_last_week_range()
    year = datetime.now().strftime("%y")
    return f"{year}年{range_str}港澳商务渠道成交用户赠送4节课豌豆币"


def step1_build_table() -> Path:
    OUTPUT_DIR.mkdir(exist_ok=True)

    df = pd.read_excel(SALES_FILE, header=SALES_HEADER_ROW)
    print(f"[原始数据] {len(df)} 行 / {len(df.columns)} 列")

    student_ids = df[COL_STUDENT_ID].dropna()
    student_ids = student_ids[student_ids.astype(str).str.strip() != ""]
    print(f"[学员ID] 有效记录 {len(student_ids)} 条")

    valid_idx = student_ids.index
    order_nos = df.loc[valid_idx, COL_FIRST_ORDER].fillna("").astype(str)

    out = pd.DataFrame({
        TEMPLATE_COLUMNS[0]: student_ids.astype(str).values,
        TEMPLATE_COLUMNS[1]: FIXED_VALUES["学科品类"],
        TEMPLATE_COLUMNS[2]: VIRTUAL_AMOUNT,
        TEMPLATE_COLUMNS[3]: FIXED_VALUES["发放原因"],
        TEMPLATE_COLUMNS[4]: FIXED_VALUES["发放备注"],
        TEMPLATE_COLUMNS[5]: FIXED_VALUES["第一级id"],
        TEMPLATE_COLUMNS[6]: FIXED_VALUES["最后一级id"],
        TEMPLATE_COLUMNS[7]: FIXED_VALUES["发放类型"],
        TEMPLATE_COLUMNS[8]: order_nos.values,
        TEMPLATE_COLUMNS[9]: FIXED_VALUES["是否合同内"],
    })

    out_name = make_output_name()
    out_path = OUTPUT_DIR / f"{out_name}.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        out.to_excel(writer, sheet_name="Sheet1", index=False)
        ws = writer.sheets["Sheet1"]
        from openpyxl.styles import numbers
        # A 列（用户ID）和 I 列（订单号）强制文本格式，避免科学计数法
        for row in ws.iter_rows(min_row=2, min_col=1, max_col=1):
            for cell in row:
                cell.number_format = numbers.FORMAT_TEXT
        for row in ws.iter_rows(min_row=2, min_col=9, max_col=9):
            for cell in row:
                cell.number_format = numbers.FORMAT_TEXT
    print(f"[完成] 输出: {out_path}")
    print(f"  总条数: {len(out)}")
    print(f"  豌豆币总数: {out[TEMPLATE_COLUMNS[2]].sum()}")
    return out_path


# ─── STEP 2: OA 表单提交 ───

def step2_submit_oa(excel_path: Path, total_coins: int, total_users: int):
    """登录 OA 系统，填写豌豆币申请表单，暂存草稿"""
    from playwright.sync_api import sync_playwright
    import time

    doc_name = make_output_name()

    with sync_playwright() as p:
        if AUTH_STATE_FILE.exists():
            context = p.chromium.launch(headless=False).new_context(
                storage_state=str(AUTH_STATE_FILE)
            )
            print("[登录] 使用已保存的 cookie")
        else:
            context = p.chromium.launch(headless=False).new_context()
            print("[登录] 首次登录，请手动完成 SSO 认证")

        page = context.new_page()
        page.goto(OA_LOGIN_URL, timeout=60000)
        time.sleep(3)

        # 检查登录状态
        needs_login = page.locator('input[type="password"]').count() > 0
        if needs_login:
            print("[提示] cookie 已过期或首次登录，请在浏览器中完成 SSO 登录...")
            page.wait_for_url("**code=**", timeout=120000)
            time.sleep(2)
            context.storage_state(path=str(AUTH_STATE_FILE))
            print("[登录] cookie 已保存")
        else:
            print("[登录] 登录状态有效")

        # 点击 OA系统 入口（新标签打开）
        time.sleep(2)
        oa_link = page.locator('text=OA系统').first
        if oa_link.count() > 0:
            print("[OA] 点击 OA系统 入口...")
            with page.expect_popup() as popup_info:
                oa_link.click()
            page = popup_info.value
            page.wait_for_load_state("networkidle", timeout=30000)
            time.sleep(15)  # OA 门户的 iframe 需要更长时间渲染
        else:
            page.goto(OA_PORTAL_URL, timeout=30000)
            time.sleep(15)

        print("[OA] 已进入 OA 门户")

        # 在 iframe 中找到包含【豌豆币添加申请】的 frame（带重试）
        target_frame = None
        for attempt in range(8):
            for f in page.frames:
                try:
                    has_link = f.evaluate("""() => {
                        let links = document.querySelectorAll('a');
                        for (let a of links) {
                            if ((a.innerText || '').trim().includes('豌豆币添加申请')) return true;
                        }
                        return false;
                    }""")
                    if has_link:
                        target_frame = f
                        break
                except Exception:
                    continue
            if target_frame:
                break
            print(f"[OA] 等待 iframe 加载... (尝试 {attempt+1}/8)")
            time.sleep(5)

        if not target_frame:
            print("[错误] 找不到豌豆币添加申请入口")
            context.close()
            return

        # 用 JS 在 iframe 内直接点击该链接（避开 Playwright text 匹配的转义问题）
        with context.expect_page() as new_page_info:
            target_frame.evaluate("""() => {
                let links = document.querySelectorAll('a');
                for (let a of links) {
                    if ((a.innerText || '').trim().includes('豌豆币添加申请')) {
                        a.click();
                        return;
                    }
                }
            }""")
        fp = new_page_info.value
        fp.wait_for_load_state("networkidle", timeout=30000)
        time.sleep(5)
        print("[OA] 进入豌豆币申请表单")

        # ─── 填写表单字段 ───

        def check_radio(field_id, value):
            """选中指定 radio（用 force 处理可能隐藏的情况）"""
            selector = f'input[name*="{field_id}"][value="{value}"]'
            el = fp.locator(selector).first
            el.click(force=True)
            el.evaluate("e => { if(e.onclick) e.onclick(); }")
            time.sleep(1)

        def fill_text(field_id, value):
            """填写文本框（处理可能隐藏的情况，必要时用 JS 设值）"""
            selector = f'input[name*="{field_id}"]:not([type=hidden]), textarea[name*="{field_id}"]'
            el = fp.locator(selector).first
            try:
                el.scroll_into_view_if_needed(timeout=3000)
                el.click(timeout=3000)
                el.fill(str(value))
            except Exception:
                fp.evaluate(
                    """(args) => {
                        let el = document.querySelector(args.sel);
                        if (el) { el.value = args.val; el.dispatchEvent(new Event('change', {bubbles:true})); }
                    }""",
                    {"sel": selector, "val": str(value)}
                )
            time.sleep(0.3)

        # 申请类型：益智豌豆币活动申请（relation_radio，用标签点击）
        fp.locator('tr:has(td.td_normal_title:has-text("申请类型"))').first.locator(
            'label:has-text("益智豌豆币活动申请")').first.click()
        time.sleep(2)
        print("[表单] 申请类型: 益智豌豆币活动申请")

        # 申请原因
        fill_text("fd_3a3df66de93d0e", doc_name)
        print(f"[表单] 申请原因: {doc_name}")

        # 科目：益智 = value 1
        check_radio("fd_3c1a3c86602ae8", "1")
        print("[表单] 科目: 益智")

        # 是否合同赠送：是 = value 1
        check_radio("fd_3c1a3ca61c3ad4", "1")
        print("[表单] 是否合同赠送: 是")

        # 关联订单号填写
        fill_text("fd_3c1a3ca3b400da", "详见附件")
        print("[表单] 关联订单号: 详见附件")

        # 虚拟币类型：豌豆币 = value 1
        check_radio("fd_3bb1d847ec6a14", "1")
        print("[表单] 虚拟币类型: 豌豆币")

        # 积分成本归属部门（manifest 地址选择器：点击容器 → 输入 → 键盘选第一个）
        try:
            container = fp.locator('xformflag[property*="fd_3a43a4a8d4e396"] div.inputselectsgl').first
            container.click()
            time.sleep(1)
            manifest_input = container.locator('input[type="text"]').last
            manifest_input.click(force=True)
            manifest_input.type("港澳市场组", delay=100)
            time.sleep(4)
            manifest_input.press("ArrowDown")
            time.sleep(0.5)
            manifest_input.press("Enter")
            time.sleep(1)
            print("[表单] 积分成本归属部门: 港澳市场组")
        except Exception as e:
            print(f"[警告] 积分成本归属部门: {e}")

        # 用户ID
        fill_text("fd_3a43a502778b84", "详见附件")
        print("[表单] 用户ID: 详见附件")

        # 需要添加豌豆币/魔力币数量总和
        fill_text("fd_3a43a5037ba0", str(total_coins))
        print(f"[表单] 豌豆币数量总和: {total_coins}")

        # 活动时间（上周日期范围）
        _, _, range_str = get_last_week_range()
        year = datetime.now().strftime("%Y")
        activity_time = f"{year}.{range_str}"
        fill_text("fd_3a43a6f08955d4", activity_time)
        print(f"[表单] 活动时间: {activity_time}")

        # 活动规则
        fill_text("fd_3a43a66b8984e2", doc_name)
        print(f"[表单] 活动规则: {doc_name}")

        # 预计参与人数
        fill_text("fd_3a43a673df46", str(total_users))
        print(f"[表单] 预计参与人数: {total_users}")

        # 是否需要人工发放豌豆币：是
        check_radio("fd_3a43a6803ccf", "是")
        print("[表单] 是否需要人工发放: 是")

        # 上传附件
        attach_file_input = fp.locator('xformflag[property*="fd_3a3df64e955efe"] input[type="file"]').first
        attach_file_input.set_input_files(str(excel_path))
        time.sleep(5)
        print(f"[附件] 已上传: {excel_path.name}")

        # 关闭可能存在的弹窗遮罩
        fp.evaluate("() => { document.querySelectorAll('.lui_dialog_mask').forEach(e => e.style.display='none'); }")
        time.sleep(0.5)

        # 点击右上角【暂存】
        fp.locator('div.lui_toolbar_btn:has-text("暂存")').click(force=True)
        time.sleep(3)
        print("[完成] OA 表单已暂存")

        try:
            context.storage_state(path=str(AUTH_STATE_FILE))
        except Exception:
            pass
        context.close()


def main():
    out_path = step1_build_table()

    df = pd.read_excel(out_path)
    total_coins = int(df[TEMPLATE_COLUMNS[2]].sum())
    total_users = len(df)

    step2_submit_oa(out_path, total_coins, total_users)


if __name__ == "__main__":
    main()
