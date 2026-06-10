"""
飞书文档发布脚本 v2

用法：
    python publish_to_feishu.py            # 推送当天周报到飞书

功能：
    1. 用 generate_weekly_report.py 中的数据生成飞书文档
    2. 文档存放到"港澳地区周报归档"文件夹
    3. 当天若已存在同名文档，覆盖更新内容
    4. 返回飞书文档链接

特点：
    - 表格使用代码块渲染（Markdown风格，无列数限制）
    - 包含完整的分析文字和综合判断
    - 自动添加重点关注、环比对比、综合判断等内容
"""
import sys
import json
import time
from pathlib import Path
from datetime import datetime, timedelta

import requests

sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = Path(__file__).parent
CONFIG_FILE = BASE_DIR / "feishu_config.local.json"

# 把各模块脚本路径加入 sys.path
sys.path.insert(0, str(BASE_DIR / "TMK周报"))
sys.path.insert(0, str(BASE_DIR / "书展内容"))
sys.path.insert(0, str(BASE_DIR / "线下商超内容"))
sys.path.insert(0, str(BASE_DIR / "转介绍打卡内容"))

# 飞书 API 端点
FEISHU_BASE = "https://open.feishu.cn/open-apis"


# ============ 飞书 API 调用 ============

def load_config():
    if not CONFIG_FILE.exists():
        print(f"❌ 配置文件不存在: {CONFIG_FILE}")
        sys.exit(1)
    return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))


def get_tenant_access_token(app_id: str, app_secret: str) -> str:
    """获取 tenant_access_token（2小时有效）"""
    url = f"{FEISHU_BASE}/auth/v3/tenant_access_token/internal"
    resp = requests.post(url, json={"app_id": app_id, "app_secret": app_secret})
    resp.raise_for_status()
    data = resp.json()
    if data.get("code") != 0:
        raise Exception(f"获取 token 失败: {data}")
    return data["tenant_access_token"]


def find_or_create_folder(token: str, folder_name: str) -> str:
    """在我的空间根目录查找或创建文件夹"""
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(
        f"{FEISHU_BASE}/drive/v1/files",
        headers=headers,
        params={"folder_token": "", "page_size": 200}
    )
    resp.raise_for_status()
    data = resp.json()
    if data.get("code") == 0:
        for f in data.get("data", {}).get("files", []):
            if f.get("type") == "folder" and f.get("name") == folder_name:
                print(f"  找到现有文件夹: {folder_name}")
                return f["token"]

    print(f"  创建新文件夹: {folder_name}")
    resp = requests.post(
        f"{FEISHU_BASE}/drive/v1/files/create_folder",
        headers=headers,
        json={"name": folder_name, "folder_token": ""}
    )
    resp.raise_for_status()
    data = resp.json()
    if data.get("code") != 0:
        raise Exception(f"创建文件夹失败: {data}")
    return data["data"]["token"]


def find_doc_by_title(token: str, folder_token: str, title: str) -> str:
    """在指定文件夹下按标题查找文档"""
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(
        f"{FEISHU_BASE}/drive/v1/files",
        headers=headers,
        params={"folder_token": folder_token, "page_size": 200}
    )
    resp.raise_for_status()
    data = resp.json()
    if data.get("code") == 0:
        for f in data.get("data", {}).get("files", []):
            if f.get("type") == "docx" and f.get("name") == title:
                return f["token"]
    return None


def create_doc(token: str, folder_token: str, title: str) -> str:
    """创建新文档"""
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.post(
        f"{FEISHU_BASE}/docx/v1/documents",
        headers=headers,
        json={"folder_token": folder_token, "title": title}
    )
    resp.raise_for_status()
    data = resp.json()
    if data.get("code") != 0:
        raise Exception(f"创建文档失败: {data}")
    return data["data"]["document"]["document_id"]


def get_doc_blocks(token: str, document_id: str) -> list:
    """获取文档的所有块"""
    headers = {"Authorization": f"Bearer {token}"}
    all_items = []
    page_token = None
    while True:
        params = {"page_size": 500}
        if page_token:
            params["page_token"] = page_token
        resp = requests.get(
            f"{FEISHU_BASE}/docx/v1/documents/{document_id}/blocks",
            headers=headers,
            params=params
        )
        resp.raise_for_status()
        data = resp.json()
        if data.get("code") != 0:
            raise Exception(f"获取块失败: {data}")
        all_items.extend(data["data"].get("items", []))
        if not data["data"].get("has_more"):
            break
        page_token = data["data"].get("page_token")
    return all_items


def delete_blocks(token: str, document_id: str, parent_block_id: str, start_index: int, end_index: int):
    """删除指定范围的子块（左闭右开）"""
    if start_index >= end_index:
        return
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.delete(
        f"{FEISHU_BASE}/docx/v1/documents/{document_id}/blocks/{parent_block_id}/children/batch_delete",
        headers=headers,
        json={"start_index": start_index, "end_index": end_index}
    )
    resp.raise_for_status()
    data = resp.json()
    if data.get("code") != 0:
        raise Exception(f"删除块失败: {data}")


def append_blocks(token: str, document_id: str, parent_block_id: str, blocks: list, index: int = -1):
    """批量追加子块，返回创建的子块列表（含 block_id）"""
    if not blocks:
        return []
    headers = {"Authorization": f"Bearer {token}"}
    payload = {"children": blocks}
    if index >= 0:
        payload["index"] = index
    resp = requests.post(
        f"{FEISHU_BASE}/docx/v1/documents/{document_id}/blocks/{parent_block_id}/children",
        headers=headers,
        json=payload
    )
    resp.raise_for_status()
    data = resp.json()
    if data.get("code") != 0:
        raise Exception(f"追加块失败: {data}")
    return data["data"].get("children", [])


def batch_update_blocks(token: str, document_id: str, requests_list: list):
    """批量更新块"""
    if not requests_list:
        return
    headers = {"Authorization": f"Bearer {token}"}
    # 飞书 API 最多 50 个/次
    chunk_size = 50
    for i in range(0, len(requests_list), chunk_size):
        chunk = requests_list[i:i+chunk_size]
        resp = requests.patch(
            f"{FEISHU_BASE}/docx/v1/documents/{document_id}/blocks/batch_update",
            headers=headers,
            json={"requests": chunk}
        )
        resp.raise_for_status()
        data = resp.json()
        if data.get("code") != 0:
            print(f"  ⚠️ 批量更新失败: {data}")
        time.sleep(0.3)


# ============ 飞书文档 Block 构造 ============

def text_run(text: str, bold=False, color=None):
    """构造文本片段
    color: 0=默认, 1=粉红, 2=橙红, 3=黄, 4=绿, 5=蓝, 6=紫, 7=灰
    """
    style = {}
    if bold:
        style["bold"] = True
    if color is not None:
        style["text_color"] = color
    return {
        "text_run": {
            "content": text,
            "text_element_style": style
        }
    }


def heading_block(text: str, level: int = 1):
    """标题块"""
    block_type = 3 + (level - 1)
    key = f"heading{level}"
    return {
        "block_type": block_type,
        key: {
            "elements": [text_run(text, bold=True)],
            "style": {}
        }
    }


def text_block(elements):
    """文本段落"""
    if isinstance(elements, str):
        elements = [text_run(elements)]
    return {
        "block_type": 2,
        "text": {
            "elements": elements,
            "style": {}
        }
    }


def parse_analysis_text(text: str):
    """
    把分析文字解析成 elements 列表，支持以下标记：
    **xxx** - 加粗
    [+]xxx[+] - 绿色
    [-]xxx[-] - 红色
    """
    import re
    elements = []
    # 用正则匹配各种标记
    pattern = re.compile(r'\*\*(.+?)\*\*|\[\+\](.+?)\[\+\]|\[-\](.+?)\[-\]')
    last_end = 0

    for match in pattern.finditer(text):
        # 匹配前的普通文本
        if match.start() > last_end:
            elements.append(text_run(text[last_end:match.start()]))

        bold_text, green_text, red_text = match.groups()
        if bold_text:
            elements.append(text_run(bold_text, bold=True))
        elif green_text:
            elements.append(text_run(green_text, bold=True, color=4))  # 绿
        elif red_text:
            elements.append(text_run(red_text, bold=True, color=1))  # 红

        last_end = match.end()

    # 末尾剩余文本
    if last_end < len(text):
        elements.append(text_run(text[last_end:]))

    if not elements:
        elements.append(text_run(text))

    return elements


# ============ 数据收集 ============

def collect_data():
    """从各模块收集数据"""
    import process_tmk
    import process_book_fair
    import process_shangchao
    import process_referral

    sys.path.insert(0, str(BASE_DIR))
    import generate_weekly_report as gwr

    print("[1] 提取流速数据...")
    flow_data = gwr.extract_flow_data()

    print("[2] 提取KOL转化数据...")
    kol_data = gwr.extract_kol_data()

    print("[3] 提取TMK做工数据...")
    tmk_data = process_tmk.extract_tmk_data()

    print("[4] 提取书展数据...")
    book_fair_data = process_book_fair.extract_book_fair_data()

    print("[5] 提取线下商超数据...")
    sc_data = process_shangchao.extract_shangchao_data()

    print("[6] 提取转介绍打卡数据...")
    referral_data = process_referral.extract_referral_data()
    punch_data = process_referral.extract_punch_card_data()

    return {
        "flow": flow_data,
        "kol": kol_data,
        "tmk": tmk_data,
        "book_fair": book_fair_data,
        "shangchao": sc_data,
        "referral": referral_data,
        "punch": punch_data,
    }


# ============ 格式化辅助函数 ============

def fmt_pct(v, force_decimal=False):
    if v is None or not isinstance(v, (int, float)):
        return "-"
    if abs(v) <= 1.5:
        return f"{v * 100:.1f}%"
    return f"{v:.1f}%"


def fmt_num(v, decimal=0):
    if v is None or not isinstance(v, (int, float)):
        return "-"
    if decimal == 0:
        return f"{v:,.0f}"
    return f"{v:,.{decimal}f}"


def fmt_roi(v):
    if v is None or not isinstance(v, (int, float)):
        return "-"
    return f"{v:.2f}"


def disp_metric(v, metric_name):
    """根据指标类型格式化显示"""
    if not isinstance(v, (int, float)):
        return "-"
    if metric_name == "滚动ROI2":
        return f"{v:.2f}"
    if v <= 1.5:
        return f"{v * 100:.1f}%"
    return f"{v:.2f}"


def calc_env(cur, prev):
    """计算环比，返回 (变化率, 趋势字符)"""
    if not isinstance(cur, (int, float)) or not isinstance(prev, (int, float)):
        return None, ""
    if prev == 0:
        return None, ""
    change = (cur - prev) / abs(prev) * 100
    trend = "↑" if change > 0 else "↓"
    return change, trend


# ============ 内容构建（与 HTML 对齐） ============

def build_sections(data: dict, date_str: str):
    """构建文档章节列表，每个 section 是一个 dict"""
    sections = []

    # 标题
    sections.append({"type": "h1", "text": f"港澳商务周报"})
    sections.append({"type": "text", "text": f"📅 数据截至 {date_str}（MTD累计）"})

    # ============ 1. 港澳商务流速 ============
    if data["flow"]:
        sections.append({"type": "h2", "text": "1. 港澳商务流速"})

        kol_names = ["钟嘉欣图片", "周家蔚", "钟嘉欣视频1", "其他汇总"]
        sc_names = ["代理商场", "书展", "代理人汇总"]
        kol_rows = [r for r in data["flow"] if r["name"] in kol_names]
        sc_rows = [r for r in data["flow"] if r["name"] in sc_names]
        kol_summary = next((r for r in data["flow"] if r["name"] == "KOL-汇总"), None)
        sc_summary = next((r for r in data["flow"] if r["name"] == "商超&社群合计"), None)
        total_row = next((r for r in data["flow"] if r["name"] == "合计"), None)

        # 1.1 重点关注
        sections.append({"type": "h3", "text": "1.1 重点关注"})
        behind = [r for r in kol_rows + sc_rows if r["gap"] < 0]
        ahead = [r for r in kol_rows + sc_rows if r["gap"] > 0]

        if behind:
            text = "**未达进度（例子）**：\n"
            for r in sorted(behind, key=lambda x: x["gap"]):
                rate = r["actual"] / r["mtd_target"] * 100 if r["mtd_target"] else 0
                text += f"  · {r['name']} [-]gap {int(r['gap'])}[-] 达成率 {rate:.1f}%\n"
            sections.append({"type": "analysis", "text": text.strip()})

        if ahead:
            text = "**超预期达成**：\n"
            for r in sorted(ahead, key=lambda x: -x["gap"]):
                rate = r["actual"] / r["mtd_target"] * 100 if r["mtd_target"] else 0
                text += f"  · {r['name']} [+]+{int(r['gap'])}[+] 达成率 {rate:.1f}%\n"
            sections.append({"type": "analysis", "text": text.strip()})

        # 1.2 KOL 模块
        sections.append({"type": "h3", "text": "1.2 KOL 模块"})
        headers = ["供应商", "MTD目标", "实际例子", "例子gap", "例子达成率",
                   "约课目标", "实际约课", "约课gap", "约课达成率", "约课成本"]
        rows = []
        for r in kol_rows + ([kol_summary] if kol_summary else []):
            rate = r["actual"] / r["mtd_target"] * 100 if r["mtd_target"] else 0
            lesson_rate = r["lesson_actual"] / r["lesson_target"] * 100 if r["lesson_target"] else 0
            rows.append([
                r["name"],
                fmt_num(r["mtd_target"]),
                fmt_num(r["actual"]),
                fmt_num(r["gap"]),
                f"{rate:.1f}%",
                fmt_num(r["lesson_target"]) if r["lesson_target"] else "-",
                fmt_num(r["lesson_actual"]),
                fmt_num(r["lesson_gap"]) if r["lesson_gap"] else "-",
                f"{lesson_rate:.1f}%",
                fmt_num(r["cost"], decimal=1) if r["cost"] else "-",
            ])
        sections.append({"type": "table", "headers": headers, "rows": rows})

        # 1.3 商超&社群
        sections.append({"type": "h3", "text": "1.3 商超&社群 模块"})
        rows = []
        for r in sc_rows + ([sc_summary] if sc_summary else []):
            rate = r["actual"] / r["mtd_target"] * 100 if r["mtd_target"] else 0
            lesson_rate = r["lesson_actual"] / r["lesson_target"] * 100 if r["lesson_target"] else 0
            rows.append([
                r["name"],
                fmt_num(r["mtd_target"]),
                fmt_num(r["actual"]),
                fmt_num(r["gap"]),
                f"{rate:.1f}%",
                fmt_num(r["lesson_target"]) if r["lesson_target"] else "-",
                fmt_num(r["lesson_actual"]),
                fmt_num(r["lesson_gap"]) if r["lesson_gap"] else "-",
                f"{lesson_rate:.1f}%",
                fmt_num(r["cost"], decimal=1) if r["cost"] else "-",
            ])
        sections.append({"type": "table", "headers": headers, "rows": rows})

        # 1.4 总计
        if kol_summary and sc_summary and total_row:
            sections.append({"type": "h3", "text": "1.4 总计"})
            rows = []
            for r in [kol_summary, sc_summary, total_row]:
                rate = r["actual"] / r["mtd_target"] * 100 if r["mtd_target"] else 0
                lesson_rate = r["lesson_actual"] / r["lesson_target"] * 100 if r["lesson_target"] else 0
                rows.append([
                    r["name"],
                    fmt_num(r["mtd_target"]),
                    fmt_num(r["actual"]),
                    fmt_num(r["gap"]),
                    f"{rate:.1f}%",
                    fmt_num(r["lesson_target"]) if r["lesson_target"] else "-",
                    fmt_num(r["lesson_actual"]),
                    fmt_num(r["lesson_gap"]) if r["lesson_gap"] else "-",
                    f"{lesson_rate:.1f}%",
                    fmt_num(r["cost"], decimal=1) if r["cost"] else "-",
                ])
            sections.append({"type": "table", "headers": headers, "rows": rows})

    # ============ 2. KOL 转化 ============
    if data["kol"]:
        sections.append({"type": "h2", "text": "2. 本月KOL转化数据"})
        kol_all = data["kol"].get("all_rows", [])

        today = datetime.now()
        cur_month = today.month
        prev_month = cur_month - 1 if cur_month > 1 else 12

        # 2.1 KOLHK 整体趋势
        kolhk_rows = [r for r in kol_all if "KOL汇总" in r.get("name", "") or "KOLHK汇总" in r.get("name", "")]
        kolhk_cur = next((r for r in kol_all if f"{cur_month}月-KOLHK汇总" in r.get("name", "")), None)
        kolhk_prev = next((r for r in kol_all if f"{prev_month}月-KOLHK汇总" in r.get("name", "")), None)

        sections.append({"type": "h3", "text": "2.1 KOLHK 整体月度趋势"})

        # 环比分析文字
        if kolhk_cur and kolhk_prev:
            text = f"**KOLHK 整体（{cur_month}月 vs {prev_month}月）**：\n"
            for metric in ["例子约课率", "约课到课率", "到课转化率", "滚动转化率", "滚动ROI2"]:
                v_cur = kolhk_cur.get(metric)
                v_prev = kolhk_prev.get(metric)
                change, trend = calc_env(v_cur, v_prev)
                if change is not None:
                    color = "[+]" if change > 0 else "[-]"
                    text += f"  · {metric}：本月 **{disp_metric(v_cur, metric)}**，上月 {disp_metric(v_prev, metric)}（环比 {color}{trend} {abs(change):.1f}%{color[:1]}{color[2]}）\n"
            sections.append({"type": "analysis", "text": text.strip()})

        # 表格
        kol_table_headers = ["项目", "滚动消耗", "例子数", "约课数", "到课数",
                              "成交数", "GMV", "例子约课率", "约课到课率",
                              "到课转化率", "滚动转化率", "滚动ROI2"]
        if kolhk_rows:
            rows = []
            for r in kolhk_rows:
                rows.append([
                    r["name"],
                    fmt_num(r.get("滚动消耗")),
                    fmt_num(r.get("例子数")),
                    fmt_num(r.get("约课数")),
                    fmt_num(r.get("到课数")),
                    fmt_num(r.get("当月成交数")),
                    fmt_num(r.get("当月GMV")),
                    fmt_pct(r.get("例子约课率")),
                    fmt_pct(r.get("约课到课率")),
                    fmt_pct(r.get("到课转化率")),
                    fmt_pct(r.get("滚动转化率")),
                    fmt_roi(r.get("滚动ROI2")),
                ])
            sections.append({"type": "table", "headers": kol_table_headers, "rows": rows})

        # 2.2 钟嘉欣
        sections.append({"type": "h3", "text": "2.2 钟嘉欣转化数据（1月至今）"})
        zjx_rows = [r for r in kol_all if "钟嘉欣" in r.get("name", "")
                    and "图片" not in r.get("name", "") and "视频" not in r.get("name", "")]
        zjx_cur = next((r for r in kol_all if f"{cur_month}月钟嘉欣" in r.get("name", "")), None)
        zjx_prev = next((r for r in kol_all if f"{prev_month}月钟嘉欣" in r.get("name", "")), None)
        zjx_sum = next((r for r in kol_all if r["name"] == "钟嘉欣-汇总"), None)

        if zjx_cur and zjx_prev:
            text = f"**钟嘉欣（{cur_month}月 vs {prev_month}月）**：\n"
            for metric in ["例子约课率", "约课到课率", "到课转化率", "滚动转化率", "滚动ROI2"]:
                v_cur = zjx_cur.get(metric)
                v_prev = zjx_prev.get(metric)
                v_sum = zjx_sum.get(metric) if zjx_sum else None
                change, trend = calc_env(v_cur, v_prev)
                line = f"  · {metric}：本月 **{disp_metric(v_cur, metric)}**，上月 {disp_metric(v_prev, metric)}"
                if change is not None:
                    color = "[+]" if change > 0 else "[-]"
                    line += f"（环比 {color}{trend} {abs(change):.1f}%{color[:1]}{color[2]}）"
                if v_sum is not None:
                    line += f"，1-{cur_month}月汇总 {disp_metric(v_sum, metric)}"
                text += line + "\n"
            sections.append({"type": "analysis", "text": text.strip()})

        if zjx_rows:
            rows = []
            for r in zjx_rows:
                rows.append([
                    r["name"],
                    fmt_num(r.get("滚动消耗")),
                    fmt_num(r.get("例子数")),
                    fmt_num(r.get("约课数")),
                    fmt_num(r.get("到课数")),
                    fmt_num(r.get("当月成交数")),
                    fmt_num(r.get("当月GMV")),
                    fmt_pct(r.get("例子约课率")),
                    fmt_pct(r.get("约课到课率")),
                    fmt_pct(r.get("到课转化率")),
                    fmt_pct(r.get("滚动转化率")),
                    fmt_roi(r.get("滚动ROI2")),
                ])
            sections.append({"type": "table", "headers": kol_table_headers, "rows": rows})

        # 2.3 钟嘉欣图片&视频对比
        detail_rows = data["kol"].get("detail_rows", [])
        if detail_rows:
            sections.append({"type": "h3", "text": "2.3 钟嘉欣图片&视频明细数据"})
            by_name = {r["name"]: r for r in detail_rows}
            pic = by_name.get("钟嘉欣图片")
            v1 = by_name.get("钟嘉欣视频1")

            if pic and v1:
                text = "**钟嘉欣 图片 vs 视频对比**：\n"
                for metric in ["例子约课率", "约课到课率", "到课转化率", "滚动转化率", "滚动ROI2"]:
                    v_pic = pic.get(metric)
                    v_v1 = v1.get(metric)
                    if isinstance(v_pic, (int, float)) and isinstance(v_v1, (int, float)):
                        better = "图片" if v_pic > v_v1 else "视频1"
                        text += f"  · {metric}：图片 **{disp_metric(v_pic, metric)}** vs 视频1 **{disp_metric(v_v1, metric)}** → [+]{better}更优[+]\n"
                sections.append({"type": "analysis", "text": text.strip()})

            rows = []
            for r in detail_rows:
                rows.append([
                    r["name"],
                    fmt_num(r.get("滚动消耗")),
                    fmt_num(r.get("例子数")),
                    fmt_num(r.get("约课数")),
                    fmt_num(r.get("到课数")),
                    fmt_num(r.get("当月成交数")),
                    fmt_num(r.get("当月GMV")),
                    fmt_pct(r.get("例子约课率")),
                    fmt_pct(r.get("约课到课率")),
                    fmt_pct(r.get("到课转化率")),
                    fmt_pct(r.get("滚动转化率")),
                    fmt_roi(r.get("滚动ROI2")),
                ])
            sections.append({"type": "table", "headers": kol_table_headers, "rows": rows})

        # 2.4 周家蔚
        zjw_rows = [r for r in kol_all if "周家蔚" in r.get("name", "")]
        zjw_cur = next((r for r in kol_all if f"{cur_month}月周家蔚" in r.get("name", "")), None)
        zjw_prev = next((r for r in kol_all if f"{prev_month}月周家蔚" in r.get("name", "")), None)
        zjw_sum = next((r for r in kol_all if r["name"] == "周家蔚-汇总"), None)

        if zjw_rows:
            sections.append({"type": "h3", "text": "2.4 周家蔚转化数据（3月至今）"})

            if zjw_cur and zjw_prev:
                text = f"**周家蔚（{cur_month}月 vs {prev_month}月）**：\n"
                for metric in ["例子约课率", "约课到课率", "到课转化率", "滚动转化率", "滚动ROI2"]:
                    v_cur = zjw_cur.get(metric)
                    v_prev = zjw_prev.get(metric)
                    v_sum = zjw_sum.get(metric) if zjw_sum else None
                    change, trend = calc_env(v_cur, v_prev)
                    line = f"  · {metric}：本月 **{disp_metric(v_cur, metric)}**，上月 {disp_metric(v_prev, metric)}"
                    if change is not None:
                        color = "[+]" if change > 0 else "[-]"
                        line += f"（环比 {color}{trend} {abs(change):.1f}%{color[:1]}{color[2]}）"
                    if v_sum is not None:
                        line += f"，3-{cur_month}月汇总 {disp_metric(v_sum, metric)}"
                    text += line + "\n"
                sections.append({"type": "analysis", "text": text.strip()})

            rows = []
            for r in zjw_rows:
                rows.append([
                    r["name"],
                    fmt_num(r.get("滚动消耗")),
                    fmt_num(r.get("例子数")),
                    fmt_num(r.get("约课数")),
                    fmt_num(r.get("到课数")),
                    fmt_num(r.get("当月成交数")),
                    fmt_num(r.get("当月GMV")),
                    fmt_pct(r.get("例子约课率")),
                    fmt_pct(r.get("约课到课率")),
                    fmt_pct(r.get("到课转化率")),
                    fmt_pct(r.get("滚动转化率")),
                    fmt_roi(r.get("滚动ROI2")),
                ])
            sections.append({"type": "table", "headers": kol_table_headers, "rows": rows})

        # 2.5 综合判断
        if zjx_cur and kolhk_cur:
            zjx_roi = zjx_cur.get("滚动ROI2")
            kol_roi = kolhk_cur.get("滚动ROI2")
            if isinstance(zjx_roi, (int, float)) and isinstance(kol_roi, (int, float)):
                sections.append({"type": "h3", "text": "2.5 综合判断"})
                if zjx_roi > kol_roi:
                    text = f"✅ 钟嘉欣 {cur_month}月滚动ROI2 ({zjx_roi:.2f}) 高于 KOLHK 整体 ({kol_roi:.2f})，[+]表现优异[+]"
                else:
                    text = f"⚠️ 钟嘉欣 {cur_month}月滚动ROI2 ({zjx_roi:.2f}) 低于 KOLHK 整体 ({kol_roi:.2f})，[-]需要关注[-]"
                sections.append({"type": "analysis", "text": text})

    # ============ 3. TMK 做工 ============
    if data["tmk"]:
        sections.append({"type": "h2", "text": "3. TMK 做工周报"})
        params = data["tmk"].get("params", {})
        overall = data["tmk"].get("overall", {})

        # 3.1 整体数据概览
        sections.append({"type": "h3", "text": "3.1 TMK 做工 - 整体数据概览"})
        period = f"{params.get('start_date', '?')} ~ {params.get('end_date', '?')}"
        text = f"📅 **周期**：{period}\n"
        if "日均通次" in overall:
            text += f"  · 团队日均通次 **{int(overall['日均通次'])}** 次\n"
        if "日均通次环比" in overall:
            v = overall["日均通次环比"]
            color = "[+]" if v > 0 else "[-]"
            direction = "提升" if v > 0 else "下降"
            text += f"  · 环比{direction} {color}{abs(v) * 100:.1f}%{color[:1]}{color[2]}\n"
        if "日均通时（分）" in overall:
            text += f"  · 团队日均通时 **{overall['日均通时（分）']:.1f}** 分钟\n"
        if "生均跟进时效达成率" in overall:
            rate = overall['生均跟进时效达成率'] * 100
            color = "[+]" if rate >= 50 else ("[-]" if rate < 30 else "")
            if color:
                text += f"  · 生均跟进时效达成率 {color}{rate:.1f}%{color[:1]}{color[2]}\n"
            else:
                text += f"  · 生均跟进时效达成率 **{rate:.1f}%**\n"
        sections.append({"type": "analysis", "text": text.strip()})

        # 3.2 异常情况
        alerts = data["tmk"].get("alerts", [])
        if alerts:
            sections.append({"type": "h3", "text": f"3.2 TMK 做工 - 个人异常情况"})
            from collections import defaultdict
            by_name = defaultdict(list)
            for a in alerts:
                by_name[a["name"]].append(a["desc"])
            text = ""
            for name, descs in by_name.items():
                text += f"⚠️ **{name}**：{' ；'.join(descs)}\n"
            sections.append({"type": "analysis", "text": text.strip()})

    # ============ 4. 书展 ============
    if data["book_fair"]:
        sections.append({"type": "h2", "text": "4. 书展数据复盘"})

        # 4.1 复盘分析
        sections.append({"type": "h3", "text": "4.1 复盘分析"})

        current = data["book_fair"]["current"]
        history = data["book_fair"]["history"]
        fair_names = data["book_fair"]["current_names"]

        # 本月双展对比
        if len(fair_names) >= 2:
            fair_a = current.get(fair_names[0], {}).get("total")
            fair_b = current.get(fair_names[1], {}).get("total")
            if fair_a and fair_b:
                text = f"**本月双展对比 — {fair_names[0]} vs {fair_names[1]}**：\n"
                for metric in ["例子约课率", "约课到课率", "到课转化率", "滚动转化率", "滚动ROI2"]:
                    va = fair_a.get(metric)
                    vb = fair_b.get(metric)
                    if isinstance(va, (int, float)) and isinstance(vb, (int, float)):
                        better = fair_names[0] if va > vb else fair_names[1]
                        short_better = better.replace("26年5月", "").replace("26年6月", "")
                        text += f"  · {metric}：{disp_metric(va, metric)} vs {disp_metric(vb, metric)} → [+]{short_better}更优[+]\n"
                sections.append({"type": "analysis", "text": text.strip()})

        # 与历史均值对比
        if history:
            hist_means = {}
            for metric in ["例子约课率", "约课到课率", "到课转化率", "滚动转化率", "滚动ROI2"]:
                vals = [h.get(metric) for h in history if isinstance(h.get(metric), (int, float))]
                if vals:
                    hist_means[metric] = sum(vals) / len(vals)

            for fair_name in fair_names:
                cur = current.get(fair_name, {}).get("total")
                if not cur:
                    continue
                text = f"**{fair_name} vs 历史{len(history)}场均值**：\n"
                for metric in ["例子约课率", "约课到课率", "到课转化率", "滚动转化率", "滚动ROI2"]:
                    v = cur.get(metric)
                    h = hist_means.get(metric)
                    change, trend = calc_env(v, h)
                    if change is not None:
                        color = "[+]" if change > 0 else "[-]"
                        text += f"  · {metric}：本场 **{disp_metric(v, metric)}**，历史均值 {disp_metric(h, metric)}（{color}{trend} {abs(change):.1f}%{color[:1]}{color[2]}）\n"
                sections.append({"type": "analysis", "text": text.strip()})

        # 4.2 本月书展数据明细
        sections.append({"type": "h3", "text": "4.2 本月书展数据明细"})
        bf_headers = ["渠道名称", "滚动消耗", "例子数", "约课数", "到课数", "成交数",
                      "GMV", "例子约课率", "约课到课率", "到课转化率", "滚动转化率", "滚动ROI2"]
        rows = []
        for fair_name in fair_names:
            fair = current.get(fair_name)
            if not fair:
                continue
            if fair.get("total"):
                t = fair["total"]
                rows.append([
                    t["name"],
                    fmt_num(t.get("滚动消耗")),
                    fmt_num(t.get("例子数")),
                    fmt_num(t.get("约课数")),
                    fmt_num(t.get("到课数")),
                    fmt_num(t.get("当月成交数")),
                    fmt_num(t.get("当月GMV")),
                    fmt_pct(t.get("例子约课率")),
                    fmt_pct(t.get("约课到课率")),
                    fmt_pct(t.get("到课转化率")),
                    fmt_pct(t.get("滚动转化率")),
                    fmt_roi(t.get("滚动ROI2")),
                ])
            for d in fair.get("details", []):
                rows.append([
                    d["name"],
                    fmt_num(d.get("滚动消耗")),
                    fmt_num(d.get("例子数")),
                    fmt_num(d.get("约课数")),
                    fmt_num(d.get("到课数")),
                    fmt_num(d.get("当月成交数")),
                    fmt_num(d.get("当月GMV")),
                    fmt_pct(d.get("例子约课率")),
                    fmt_pct(d.get("约课到课率")),
                    fmt_pct(d.get("到课转化率")),
                    fmt_pct(d.get("滚动转化率")),
                    fmt_roi(d.get("滚动ROI2")),
                ])
        if rows:
            sections.append({"type": "table", "headers": bf_headers, "rows": rows})

    # ============ 5. 线下商超 ============
    if data["shangchao"]:
        sections.append({"type": "h2", "text": "5. 线下商超复盘"})

        cur = data["shangchao"]["current_month"]
        prev = data["shangchao"]["prev_month"]
        total = data["shangchao"]["total"]

        # 5.1 当月分析
        if cur:
            sections.append({"type": "h3", "text": f"5.1 线下商超 - 当月分析"})

            # 数据概览
            text = f"**{cur['name']}商超数据概览**：\n"
            if cur.get("天数"):
                text += f"  · 出摊天数 **{int(cur['天数'])}** 天，天均例子 **{cur.get('天均', 0):.1f}**\n"
            text += f"  · 例子数 **{int(cur.get('例子数', 0))}**，约课数 **{int(cur.get('约课数', 0))}**\n"
            sections.append({"type": "analysis", "text": text.strip()})

            # 环比上月
            if prev:
                text = f"**{cur['name']} vs {prev['name']}（环比）**：\n"
                for metric in ["例子约课率", "约课到课率", "到课转化率", "滚动转化率", "滚动ROI2"]:
                    vc = cur.get(metric)
                    vp = prev.get(metric)
                    change, trend = calc_env(vc, vp)
                    if change is not None:
                        color = "[+]" if change > 0 else "[-]"
                        text += f"  · {metric}：本月 **{disp_metric(vc, metric)}**，上月 {disp_metric(vp, metric)}（{color}{trend} {abs(change):.1f}%{color[:1]}{color[2]}）\n"
                sections.append({"type": "analysis", "text": text.strip()})

            # vs 历史均值
            if total:
                text = f"**{cur['name']} vs 历史整体均值**：\n"
                for metric in ["例子约课率", "约课到课率", "到课转化率", "滚动转化率", "滚动ROI2"]:
                    vc = cur.get(metric)
                    vt = total.get(metric)
                    change, trend = calc_env(vc, vt)
                    if change is not None:
                        color = "[+]" if change > 0 else "[-]"
                        text += f"  · {metric}：本月 **{disp_metric(vc, metric)}**，历史均值 {disp_metric(vt, metric)}（{color}{trend} {abs(change):.1f}%{color[:1]}{color[2]}）\n"
                sections.append({"type": "analysis", "text": text.strip()})

        # 5.2 分月趋势
        sections.append({"type": "h3", "text": "5.2 线下商超 - 分月趋势"})
        sc_monthly_headers = ["月份", "天数", "天均", "CPS消耗", "CPT消耗", "滚动消耗",
                               "例子数", "约课数", "到课数", "成交数", "GMV",
                               "例子约课率", "约课到课率", "到课转化率", "滚动转化率", "滚动ROI2"]
        rows = []
        for r in data["shangchao"]["monthly"]:
            rows.append([
                r["name"],
                fmt_num(r.get("天数")),
                f"{r.get('天均', 0):.1f}" if r.get("天均") else "-",
                fmt_num(r.get("CPS消耗")),
                fmt_num(r.get("CPT消耗")),
                fmt_num(r.get("滚动消耗")),
                fmt_num(r.get("例子数")),
                fmt_num(r.get("约课数")),
                fmt_num(r.get("到课数")),
                fmt_num(r.get("当月成交数") or r.get("滚动成交数")),
                fmt_num(r.get("当月GMV") or r.get("滚动GMV")),
                fmt_pct(r.get("例子约课率")),
                fmt_pct(r.get("约课到课率")),
                fmt_pct(r.get("到课转化率")),
                fmt_pct(r.get("滚动转化率")),
                fmt_roi(r.get("滚动ROI2")),
            ])
        if rows:
            sections.append({"type": "table", "headers": sc_monthly_headers, "rows": rows})

        # 5.3 当月分场次明细
        valid_venues = [v for v in data["shangchao"]["current_venues"]
                        if (v["total"].get("例子数") or 0) > 0]
        if valid_venues:
            sections.append({"type": "h3", "text": f"5.3 线下商超 - {cur['name']}分场次明细"})
            sc_venue_headers = ["场次", "天数", "天均", "滚动消耗", "例子数", "约课数",
                                 "到课数", "成交数", "GMV",
                                 "例子约课率", "约课到课率", "到课转化率", "滚动转化率", "滚动ROI2"]
            rows = []
            for v in valid_venues:
                t = v["total"]
                rows.append([
                    t["name"],
                    fmt_num(t.get("天数")),
                    f"{t.get('天均', 0):.1f}" if t.get("天均") else "-",
                    fmt_num(t.get("滚动消耗")),
                    fmt_num(t.get("例子数")),
                    fmt_num(t.get("约课数")),
                    fmt_num(t.get("到课数")),
                    fmt_num(t.get("当月成交数")),
                    fmt_num(t.get("当月GMV")),
                    fmt_pct(t.get("例子约课率")),
                    fmt_pct(t.get("约课到课率")),
                    fmt_pct(t.get("到课转化率")),
                    fmt_pct(t.get("滚动转化率")),
                    fmt_roi(t.get("滚动ROI2")),
                ])
            sections.append({"type": "table", "headers": sc_venue_headers, "rows": rows})

        # 5.4 重复场次对比
        repeats = data["shangchao"].get("repeat_venues", {})
        if repeats:
            sections.append({"type": "h3", "text": "5.4 线下商超 - 重复场次对比"})
            for base, group in sorted(repeats.items()):
                sections.append({"type": "h3", "text": f"📍 {base}（{len(group)} 次场次）"})
                rows = []
                for v in group:
                    t = v["total"]
                    rows.append([
                        t["name"],
                        fmt_num(t.get("天数")),
                        f"{t.get('天均', 0):.1f}" if t.get("天均") else "-",
                        fmt_num(t.get("滚动消耗")),
                        fmt_num(t.get("例子数")),
                        fmt_num(t.get("约课数")),
                        fmt_pct(t.get("例子约课率")),
                        fmt_pct(t.get("约课到课率")),
                        fmt_pct(t.get("到课转化率")),
                        fmt_pct(t.get("滚动转化率")),
                        fmt_roi(t.get("滚动ROI2")),
                    ])
                sections.append({"type": "table",
                    "headers": ["场次", "天数", "天均", "滚动消耗", "例子数", "约课数",
                                "例子约课率", "约课到课率", "到课转化率", "滚动转化率", "滚动ROI2"],
                    "rows": rows})

                # 对比分析
                analysis = []
                for metric in ["例子约课率", "约课到课率", "到课转化率", "滚动转化率", "滚动ROI2"]:
                    vals = [(v["total"]["name"], v["total"].get(metric)) for v in group
                            if isinstance(v["total"].get(metric), (int, float))]
                    if len(vals) >= 2:
                        best = max(vals, key=lambda x: x[1])
                        worst = min(vals, key=lambda x: x[1])
                        if best[0] != worst[0]:
                            analysis.append(
                                f"  · {metric}：[+]{best[0]} {disp_metric(best[1], metric)} 最优[+]，"
                                f"{worst[0]} {disp_metric(worst[1], metric)} 最差"
                            )
                if analysis:
                    sections.append({"type": "analysis", "text": "\n".join(analysis)})

    # ============ 6. 转介绍打卡 ============
    if data["referral"]:
        sections.append({"type": "h2", "text": "6. 转介绍打卡"})

        # 6.1 后端非手推达成
        sections.append({"type": "h3", "text": "6.1 转介绍 - 后端非手推达成情况"})

        # 本期表格
        sections.append({"type": "h3", "text": f"{data['referral']['cur_period']}（本期）"})
        ref_headers = ["类型", "例子数", "目标例子数", "MTD达成率", "例子占比",
                       "约课人数", "到课人数", "成单数", "GMV",
                       "约课率", "约课到课率", "到课转化率", "转化率", "ASP", "GMV占比"]
        rows = []
        for r in data["referral"]["cur_rows"]:
            rows.append([
                r["name"],
                str(r["examples"]),
                str(r["target"]),
                fmt_pct(r["mtd_rate"]),
                fmt_pct(r["example_pct"]),
                str(r["bookings"]),
                str(r["attendances"]),
                str(r["signups"]),
                fmt_num(r["gmv"]),
                fmt_pct(r["book_rate"]),
                fmt_pct(r["attend_rate"]),
                fmt_pct(r["conv_rate"]),
                fmt_pct(r["total_conv_rate"]),
                fmt_num(r["asp"]),
                fmt_pct(r["gmv_pct"]),
            ])
        sections.append({"type": "table", "headers": ref_headers, "rows": rows})

        # 上期表格
        sections.append({"type": "h3", "text": f"{data['referral']['prev_period']}（上期）"})
        rows = []
        for r in data["referral"]["prev_rows"]:
            rows.append([
                r["name"],
                str(r["examples"]),
                str(r["target"]),
                fmt_pct(r["mtd_rate"]),
                fmt_pct(r["example_pct"]),
                str(r["bookings"]),
                str(r["attendances"]),
                str(r["signups"]),
                fmt_num(r["gmv"]),
                fmt_pct(r["book_rate"]),
                fmt_pct(r["attend_rate"]),
                fmt_pct(r["conv_rate"]),
                fmt_pct(r["total_conv_rate"]),
                fmt_num(r["asp"]),
                fmt_pct(r["gmv_pct"]),
            ])
        sections.append({"type": "table", "headers": ref_headers, "rows": rows})

        # 6.2 打卡链路数据
        if data["punch"]:
            sections.append({"type": "h3", "text": "6.2 转介绍 - 打卡链路数据"})

            # 环比分析
            cur = data["punch"]["cur"]
            prev = data["punch"]["prev"]
            text = f"**环比上期（{data['punch']['cur_period']} vs {data['punch']['prev_period']}）**\n"
            metrics_env = [
                ("可打卡学员", cur["can_punch"], prev["can_punch"]),
                ("打卡人数", cur["punched"], prev["punched"]),
                ("打卡率", cur["punch_rate"], prev["punch_rate"]),
                ("例子数", cur["examples"], prev["examples"]),
                ("打卡裂变率", cur["split_rate"], prev["split_rate"]),
                ("约课率", cur["book_rate"], prev["book_rate"]),
                ("约课到课率", cur["attend_rate"], prev["attend_rate"]),
            ]
            for name, cv, pv in metrics_env:
                change, trend = calc_env(cv, pv)
                if change is not None:
                    color = "[+]" if change > 0 else "[-]"
                    text += f"  · {name}：{color}{trend} {abs(change):.1f}%{color[:1]}{color[2]}\n"
            sections.append({"type": "analysis", "text": text.strip()})

            # 打卡链路表格
            punch_headers = ["日期", "可打卡学员", "打卡人数", "打卡次数", "打卡率",
                             "人均打卡次数", "例子数", "打卡裂变率", "约课数", "约课率",
                             "到课数", "约课到课率", "转化例子数", "到课转化率",
                             "注册转化率", "GMV", "ASP"]
            rows = []
            for label, d in [(data['punch']['cur_period'], cur), (data['punch']['prev_period'], prev)]:
                rows.append([
                    label,
                    fmt_num(d["can_punch"]),
                    fmt_num(d["punched"]),
                    fmt_num(d["punch_count"]),
                    fmt_pct(d["punch_rate"]),
                    f"{d['avg_punch']:.2f}",
                    str(d["examples"]),
                    fmt_pct(d["split_rate"]),
                    str(d["bookings"]),
                    fmt_pct(d["book_rate"]),
                    str(d["attendances"]),
                    fmt_pct(d["attend_rate"]),
                    str(d["signups"]),
                    fmt_pct(d["conv_rate"]),
                    fmt_pct(d["total_conv_rate"]),
                    fmt_num(d["gmv"]),
                    fmt_num(d["asp"]),
                ])
            sections.append({"type": "table", "headers": punch_headers, "rows": rows})

    return sections


# ============ 写入飞书文档 ============

def render_table_as_text(headers: list, rows: list) -> str:
    """把表格渲染为对齐的文本（Markdown 风格）"""
    # 计算每列最大宽度
    col_widths = [len(str(h)) for h in headers]
    for row in rows:
        for i, v in enumerate(row):
            if i < len(col_widths):
                col_widths[i] = max(col_widths[i], len(str(v)))

    # 渲染表头
    lines = []
    header_line = " | ".join(f"{str(h):<{col_widths[i]}}" for i, h in enumerate(headers))
    lines.append(header_line)

    # 分隔线
    sep_line = "-+-".join("-" * w for w in col_widths)
    lines.append(sep_line)

    # 渲染数据行
    for row in rows:
        row_line = " | ".join(f"{str(row[i] if i < len(row) else ''):<{col_widths[i]}}" for i in range(len(headers)))
        lines.append(row_line)

    return "\n".join(lines)


def section_to_block(section):
    """把单个 section 转换为飞书 block（使用代码块渲染表格）"""
    t = section["type"]
    if t == "h1":
        return heading_block(section["text"], level=1)
    elif t == "h2":
        return heading_block(section["text"], level=2)
    elif t == "h3":
        return heading_block(section["text"], level=3)
    elif t == "text":
        return text_block(section["text"])
    elif t == "analysis":
        # 分析块：解析特殊标记，按行拆分
        lines = section["text"].split("\n")
        all_elements = []
        for i, line in enumerate(lines):
            if i > 0:
                all_elements.append({"text_run": {"content": "\n", "text_element_style": {}}})
            all_elements.extend(parse_analysis_text(line))
        return text_block(all_elements)
    elif t == "table":
        # 表格渲染为代码块（Feishu原生表格限制9列）
        table_text = render_table_as_text(section["headers"], section["rows"])
        return {
            "block_type": 14,  # code block
            "code": {
                "elements": [{"text_run": {"content": table_text}}],
                "style": {"language": 1}  # 1 = PlainText
            }
        }
    return None


def insert_table_blocks(token: str, document_id: str, parent_block_id: str,
                        index: int, headers: list, rows: list):
    """插入飞书原生表格块。index=-1 表示自动追加到末尾"""
    n_cols = len(headers)
    n_rows = len(rows) + 1  # 含表头

    # 1. 创建表格块
    table_def = {
        "block_type": 31,
        "table": {
            "property": {
                "row_size": n_rows,
                "column_size": n_cols,
                "header_row": True,
            }
        }
    }

    api_headers = {"Authorization": f"Bearer {token}"}
    payload = {"children": [table_def]}
    if index >= 0:
        payload["index"] = index

    resp = requests.post(
        f"{FEISHU_BASE}/docx/v1/documents/{document_id}/blocks/{parent_block_id}/children",
        headers=api_headers,
        json=payload
    )

    # 如果是 400 错误，打印详细信息
    if resp.status_code == 400:
        error_detail = resp.json()
        raise Exception(f"创建表格失败 (HTTP 400): {error_detail}")

    resp.raise_for_status()
    data = resp.json()
    if data.get("code") != 0:
        raise Exception(f"创建表格失败: {data}")

    children = data["data"].get("children", [])
    if not children:
        return

    table_block_id = children[0]["block_id"]
    time.sleep(0.5)

    # 2. 获取表格的所有单元格（cell 块）
    resp = requests.get(
        f"{FEISHU_BASE}/docx/v1/documents/{document_id}/blocks/{table_block_id}/children",
        headers=api_headers,
        params={"page_size": 500}
    )
    resp.raise_for_status()
    cell_data = resp.json()
    cell_blocks = cell_data["data"].get("items", [])

    # 飞书表格的 cell 顺序是按行优先排列的（先第一行所有列，再第二行...）
    # 每个 cell 块下面包含一个 text 块作为内容
    # 整理：先表头，后数据行
    all_data = list(headers) + [v for row in rows for v in row]

    # 3. 对每个 cell 内的 text 块进行更新
    update_requests = []
    for idx, cell in enumerate(cell_blocks):
        if idx >= len(all_data):
            break
        cell_children = cell.get("children", [])
        if not cell_children:
            continue
        text_block_id = cell_children[0]
        content = str(all_data[idx]) if all_data[idx] is not None else ""
        is_header = idx < n_cols

        # 给负数标红、增长标绿（针对内容是百分比或数字带正负号的）
        color = None
        if not is_header:
            content_clean = content.strip()
            if content_clean.startswith("-") and content_clean != "-":
                # 负数：检查是否真的是负值
                try:
                    val_str = content_clean.rstrip("%").replace(",", "")
                    val = float(val_str)
                    if val < 0:
                        color = 1  # 红色
                except (ValueError, AttributeError):
                    pass

        update_requests.append({
            "block_id": text_block_id,
            "update_text_elements": {
                "elements": [text_run(content, bold=is_header, color=color)]
            }
        })

    # 4. 批量更新单元格内容
    if update_requests:
        batch_update_blocks(token, document_id, update_requests)

    return index + 1


def write_to_doc(token: str, document_id: str, sections: list):
    """把 sections 写入文档（使用代码块渲染表格）"""
    # 1. 清空文档
    print("  清空文档...")
    all_items = get_doc_blocks(token, document_id)
    root_children = [b for b in all_items if b.get("parent_id") == document_id]
    if root_children:
        print(f"  删除 {len(root_children)} 个已有块")
        chunk = 50
        for i in range(0, len(root_children), chunk):
            end = min(i + chunk, len(root_children))
            delete_blocks(token, document_id, document_id, 0, end - i)
            time.sleep(0.5)

    # 2. 批量插入所有块（表格用代码块，无需单独处理）
    print(f"  构建 {len(sections)} 个块...")
    blocks = []
    for section in sections:
        block = section_to_block(section)
        if block:
            blocks.append(block)

    print(f"  批量插入 {len(blocks)} 个块...")
    chunk_size = 50
    for i in range(0, len(blocks), chunk_size):
        chunk = blocks[i:i+chunk_size]
        append_blocks(token, document_id, document_id, chunk)
        time.sleep(0.5)

    table_count = sum(1 for s in sections if s["type"] == "table")
    print(f"  完成：共 {len(blocks)} 个块（含 {table_count} 个表格代码块）")


def insert_table_blocks_native(token: str, document_id: str, parent_block_id: str,
                                index: int, headers: list, rows: list):
    """插入飞书原生表格块（使用动态索引修复 1770001 错误）"""
    n_cols = len(headers)
    n_rows = len(rows) + 1  # 含表头

    # 1. 创建表格块
    table_def = {
        "block_type": 31,
        "table": {
            "property": {
                "row_size": n_rows,
                "column_size": n_cols,
                "header_row": True,
            }
        }
    }

    api_headers = {"Authorization": f"Bearer {token}"}
    resp = requests.post(
        f"{FEISHU_BASE}/docx/v1/documents/{document_id}/blocks/{parent_block_id}/children",
        headers=api_headers,
        json={"children": [table_def], "index": index}
    )

    if resp.status_code != 200:
        error_detail = resp.json()
        raise Exception(f"创建表格失败 (HTTP {resp.status_code}): {error_detail}")

    data = resp.json()
    if data.get("code") != 0:
        raise Exception(f"创建表格失败: {data}")

    children = data["data"].get("children", [])
    if not children:
        return

    table_block_id = children[0]["block_id"]
    time.sleep(0.5)

    # 2. 获取表格的所有单元格
    resp = requests.get(
        f"{FEISHU_BASE}/docx/v1/documents/{document_id}/blocks/{table_block_id}/children",
        headers=api_headers,
        params={"page_size": 500}
    )
    resp.raise_for_status()
    cell_data = resp.json()
    cell_blocks = cell_data["data"].get("items", [])

    # 3. 填充单元格内容
    all_data = list(headers) + [v for row in rows for v in row]
    update_requests = []

    for idx, cell in enumerate(cell_blocks):
        if idx >= len(all_data):
            break
        cell_children = cell.get("children", [])
        if not cell_children:
            continue

        text_block_id = cell_children[0]
        content = str(all_data[idx]) if all_data[idx] is not None else ""
        is_header = idx < n_cols

        # 负数标红
        color = None
        if not is_header:
            content_clean = content.strip()
            if content_clean.startswith("-") and content_clean != "-":
                try:
                    val_str = content_clean.rstrip("%").replace(",", "")
                    val = float(val_str)
                    if val < 0:
                        color = 1  # 红色
                except (ValueError, AttributeError):
                    pass

        update_requests.append({
            "block_id": text_block_id,
            "update_text_elements": {
                "elements": [text_run(content, bold=is_header, color=color)]
            }
        })

    # 4. 批量更新单元格内容
    if update_requests:
        batch_update_blocks(token, document_id, update_requests)


# ============ 主流程 ============

def main():
    print("=" * 60)
    print("飞书文档发布 v2")
    print("=" * 60)

    cfg = load_config()
    print(f"\n[1] 获取 access token...")
    token = get_tenant_access_token(cfg["app_id"], cfg["app_secret"])

    print(f"\n[2] 查找/创建文件夹: {cfg['folder_name']}")
    folder_token = find_or_create_folder(token, cfg["folder_name"])

    today = datetime.now().date()
    date_str = today.strftime("%Y%m%d")
    title = cfg["doc_title_format"].format(date=date_str)
    print(f"\n[3] 文档标题: {title}")

    doc_id = find_doc_by_title(token, folder_token, title)
    if doc_id:
        print(f"  找到现有文档: {doc_id}（将覆盖更新）")
    else:
        print(f"  创建新文档...")
        doc_id = create_doc(token, folder_token, title)
        print(f"  新文档 ID: {doc_id}")

    print(f"\n[4] 收集周报数据...")
    data = collect_data()

    print(f"\n[5] 构建文档内容...")
    yesterday = today - timedelta(days=1)
    sections = build_sections(data, yesterday.strftime("%Y年%m月%d日"))
    print(f"  共 {len(sections)} 个章节")
    table_count = sum(1 for s in sections if s["type"] == "table")
    print(f"  其中表格 {table_count} 个")

    print(f"\n[6] 写入飞书文档...")
    write_to_doc(token, doc_id, sections)

    doc_url = f"https://feishu.cn/docx/{doc_id}"
    print(f"\n{'='*60}")
    print(f"✅ 完成！")
    print(f"📄 文档链接: {doc_url}")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
