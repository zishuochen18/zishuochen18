"""
飞书文档发布脚本

用法：
    python publish_to_feishu.py            # 推送当天周报到飞书
    python publish_to_feishu.py --new       # 强制新建（不检查同日文档）

功能：
    1. 用 generate_weekly_report.py 中的数据生成飞书文档
    2. 文档存放到"港澳地区周报归档"文件夹
    3. 当天若已存在同名文档，覆盖更新内容
    4. 返回飞书文档链接

依赖：
    pip install requests
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
        print(f"   请复制 feishu_config.example.json → feishu_config.local.json 并填入凭证")
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
    """在我的空间根目录查找或创建文件夹，返回 folder_token"""
    headers = {"Authorization": f"Bearer {token}"}

    # 1) 查找根目录下的文件夹
    # GET /open-apis/drive/v1/files?folder_token=root
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
                print(f"  找到现有文件夹: {folder_name} (token={f['token']})")
                return f["token"]

    # 2) 不存在，创建新文件夹
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
    """在指定文件夹下按标题查找文档，找到返回 document_id，找不到返回 None"""
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
    """在指定文件夹下创建新文档，返回 document_id"""
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
    resp = requests.get(
        f"{FEISHU_BASE}/docx/v1/documents/{document_id}/blocks",
        headers=headers,
        params={"page_size": 500}
    )
    resp.raise_for_status()
    data = resp.json()
    if data.get("code") != 0:
        raise Exception(f"获取块失败: {data}")
    return data["data"].get("items", [])


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
    """批量追加子块"""
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


# ============ 飞书文档 Block 构造 ============

def text_run(text: str, bold=False, color=None):
    """构造一个文本片段"""
    style = {}
    if bold:
        style["bold"] = True
    if color is not None:
        style["text_color"] = color  # 1=red, 2=orange...
    return {
        "text_run": {
            "content": text,
            "text_element_style": style
        }
    }


def heading_block(text: str, level: int = 1):
    """构造标题块（1-9级）"""
    block_type = 3 + (level - 1)  # heading1=3, heading2=4, heading3=5...
    key = f"heading{level}"
    return {
        "block_type": block_type,
        key: {
            "elements": [text_run(text)],
            "style": {}
        }
    }


def text_block(text: str, bold=False):
    """构造普通文本段落"""
    return {
        "block_type": 2,  # text
        "text": {
            "elements": [text_run(text, bold=bold)],
            "style": {}
        }
    }


def callout_block(text: str, emoji: str = "📌"):
    """构造高亮块"""
    return {
        "block_type": 19,  # callout
        "callout": {
            "background_color": 1,
            "border_color": 0,
            "text_color": 0,
            "emoji_id": emoji,
        },
    }


def table_block(headers: list, rows: list):
    """
    构造表格块（飞书 docx 表格）
    headers: ["列1", "列2", ...]
    rows: [["a", "b"], ["c", "d"]]
    返回的是包含 table 块和所有 cell 块的列表
    """
    n_cols = len(headers)
    n_rows = len(rows) + 1  # 含表头

    table = {
        "block_type": 31,  # table
        "table": {
            "property": {
                "row_size": n_rows,
                "column_size": n_cols,
                "column_width": [120] * n_cols,
                "header_row": True,
            }
        }
    }
    return table, headers, rows


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


# ============ 内容生成 ============

def fmt_pct(v):
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


def build_content_blocks(data: dict, date_str: str):
    """构建飞书文档的内容块列表（按顺序追加）

    返回的是简化的 (block, table_data) 列表
    block 是一个 dict 描述这个段落要插入的内容
    """
    sections = []

    # 标题
    sections.append({"type": "h1", "text": f"港澳商务周报 - {date_str}"})
    sections.append({"type": "text", "text": f"📅 数据截至 {date_str}"})

    # ============ 1. 港澳商务流速 ============
    if data["flow"]:
        sections.append({"type": "h2", "text": "1. 港澳商务流速"})

        # KOL 模块
        kol_names = ["钟嘉欣图片", "周家蔚", "钟嘉欣视频1", "其他汇总"]
        sc_names = ["代理商场", "书展", "代理人汇总"]
        kol_rows = [r for r in data["flow"] if r["name"] in kol_names]
        sc_rows = [r for r in data["flow"] if r["name"] in sc_names]
        kol_summary = next((r for r in data["flow"] if r["name"] == "KOL-汇总"), None)
        sc_summary = next((r for r in data["flow"] if r["name"] == "商超&社群合计"), None)
        total_row = next((r for r in data["flow"] if r["name"] == "合计"), None)

        sections.append({"type": "h3", "text": "1.1 KOL 模块"})
        headers = ["供应商", "MTD目标", "实际例子", "例子gap", "达成率", "约课目标", "实际约课", "约课成本"]
        rows = []
        for r in kol_rows + ([kol_summary] if kol_summary else []):
            rate = r["actual"] / r["mtd_target"] * 100 if r["mtd_target"] else 0
            rows.append([
                r["name"],
                fmt_num(r["mtd_target"]),
                fmt_num(r["actual"]),
                fmt_num(r["gap"]),
                f"{rate:.1f}%",
                fmt_num(r["lesson_target"]),
                fmt_num(r["lesson_actual"]),
                fmt_num(r["cost"], decimal=1) if r["cost"] else "-",
            ])
        sections.append({"type": "table", "headers": headers, "rows": rows})

        sections.append({"type": "h3", "text": "1.2 商超&社群"})
        rows = []
        for r in sc_rows + ([sc_summary] if sc_summary else []):
            rate = r["actual"] / r["mtd_target"] * 100 if r["mtd_target"] else 0
            rows.append([
                r["name"],
                fmt_num(r["mtd_target"]),
                fmt_num(r["actual"]),
                fmt_num(r["gap"]),
                f"{rate:.1f}%",
                fmt_num(r["lesson_target"]),
                fmt_num(r["lesson_actual"]),
                fmt_num(r["cost"], decimal=1) if r["cost"] else "-",
            ])
        sections.append({"type": "table", "headers": headers, "rows": rows})

        if total_row:
            sections.append({"type": "h3", "text": "1.3 总计"})
            rate = total_row["actual"] / total_row["mtd_target"] * 100 if total_row["mtd_target"] else 0
            sections.append({"type": "text",
                "text": f"目标 {fmt_num(total_row['mtd_target'])}，实际 {fmt_num(total_row['actual'])}，"
                f"达成率 {rate:.1f}%（gap {fmt_num(total_row['gap'])}）"})

    # ============ 2. KOL 转化 ============
    if data["kol"]:
        sections.append({"type": "h2", "text": "2. 本月KOL转化数据"})
        kol_all = data["kol"].get("all_rows", [])

        # 提取本月行
        today = datetime.now()
        cur_month = today.month
        prev_month = cur_month - 1 if cur_month > 1 else 12

        kolhk_cur = next((r for r in kol_all if f"{cur_month}月-KOLHK汇总" in r.get("name", "")), None)
        kolhk_prev = next((r for r in kol_all if f"{prev_month}月-KOLHK汇总" in r.get("name", "")), None)
        zjx_cur = next((r for r in kol_all if f"{cur_month}月钟嘉欣" in r.get("name", "")), None)
        zjx_prev = next((r for r in kol_all if f"{prev_month}月钟嘉欣" in r.get("name", "")), None)
        zjw_cur = next((r for r in kol_all if f"{cur_month}月周家蔚" in r.get("name", "")), None)
        zjw_prev = next((r for r in kol_all if f"{prev_month}月周家蔚" in r.get("name", "")), None)

        headers = ["主体", "例子数", "约课数", "例子约课率", "约课到课率", "到课转化率", "滚动转化率", "滚动ROI2"]
        rows = []
        for label, r in [
            ("KOLHK 整体（本月）", kolhk_cur),
            ("KOLHK 整体（上月）", kolhk_prev),
            ("钟嘉欣（本月）", zjx_cur),
            ("钟嘉欣（上月）", zjx_prev),
            ("周家蔚（本月）", zjw_cur),
            ("周家蔚（上月）", zjw_prev),
        ]:
            if r:
                rows.append([
                    label,
                    fmt_num(r.get("例子数")),
                    fmt_num(r.get("约课数")),
                    fmt_pct(r.get("例子约课率")),
                    fmt_pct(r.get("约课到课率")),
                    fmt_pct(r.get("到课转化率")),
                    fmt_pct(r.get("滚动转化率")),
                    f"{r['滚动ROI2']:.2f}" if isinstance(r.get("滚动ROI2"), (int, float)) else "-",
                ])
        if rows:
            sections.append({"type": "table", "headers": headers, "rows": rows})

    # ============ 3. TMK 做工 ============
    if data["tmk"]:
        sections.append({"type": "h2", "text": "3. TMK 做工周报"})
        params = data["tmk"].get("params", {})
        overall = data["tmk"].get("overall", {})
        period = f"{params.get('start_date', '?')} ~ {params.get('end_date', '?')}"
        sections.append({"type": "text", "text": f"📅 周期：{period}"})
        if "日均通次" in overall:
            sections.append({"type": "text", "text": f"团队日均通次 {int(overall['日均通次'])} 次"})
        if "日均通次环比" in overall:
            v = overall["日均通次环比"]
            direction = "↑" if v > 0 else "↓"
            sections.append({"type": "text", "text": f"环比{direction} {abs(v) * 100:.1f}%"})
        if "日均通时（分）" in overall:
            sections.append({"type": "text", "text": f"团队日均通时 {overall['日均通时（分）']:.1f} 分钟"})
        if "生均跟进时效达成率" in overall:
            sections.append({"type": "text", "text": f"生均跟进时效达成率 {overall['生均跟进时效达成率'] * 100:.1f}%"})

        alerts = data["tmk"].get("alerts", [])
        if alerts:
            sections.append({"type": "h3", "text": f"个人异常情况（{len(alerts)} 项）"})
            from collections import defaultdict
            by_name = defaultdict(list)
            for a in alerts:
                by_name[a["name"]].append(a["desc"])
            for name, descs in by_name.items():
                sections.append({"type": "text", "text": f"⚠️ {name}：{' ；'.join(descs)}"})

    # ============ 4. 书展 ============
    if data["book_fair"]:
        sections.append({"type": "h2", "text": "4. 书展数据复盘"})
        for fair_name in data["book_fair"]["current_names"]:
            fair = data["book_fair"]["current"].get(fair_name)
            if fair and fair.get("total"):
                t = fair["total"]
                sections.append({"type": "h3", "text": fair_name})
                ex = fmt_num(t.get("例子数"))
                yk = fmt_num(t.get("约课数"))
                roi = t.get("滚动ROI2")
                roi_str = f"{roi:.2f}" if isinstance(roi, (int, float)) else "-"
                sections.append({"type": "text",
                    "text": f"例子 {ex} | 约课 {yk} | 例子约课率 {fmt_pct(t.get('例子约课率'))} | "
                    f"到课转化率 {fmt_pct(t.get('到课转化率'))} | 滚动ROI2 {roi_str}"})

    # ============ 5. 线下商超 ============
    if data["shangchao"]:
        sections.append({"type": "h2", "text": "5. 线下商超复盘"})
        cur = data["shangchao"]["current_month"]
        prev = data["shangchao"]["prev_month"]
        if cur:
            sections.append({"type": "h3", "text": f"{cur['name']} 整体数据"})
            sections.append({"type": "text",
                "text": f"出摊 {int(cur.get('天数', 0))} 天，天均 {cur.get('天均', 0):.1f} 例子"})
            sections.append({"type": "text",
                "text": f"例子 {fmt_num(cur.get('例子数'))} | 约课 {fmt_num(cur.get('约课数'))} | "
                f"滚动消耗 {fmt_num(cur.get('滚动消耗'))} | 滚动GMV {fmt_num(cur.get('滚动GMV'))}"})
            sections.append({"type": "text",
                "text": f"例子约课率 {fmt_pct(cur.get('例子约课率'))} | 滚动ROI2 "
                f"{cur['滚动ROI2']:.2f}" if isinstance(cur.get('滚动ROI2'), (int, float)) else "-"})

        # 当月分场次
        valid_venues = [v for v in data["shangchao"]["current_venues"]
                        if (v["total"].get("例子数") or 0) > 0]
        if valid_venues:
            sections.append({"type": "h3", "text": "当月分场次明细"})
            headers = ["场次", "天数", "天均", "例子数", "约课数", "滚动消耗", "滚动GMV", "滚动ROI2"]
            rows = []
            for v in valid_venues:
                t = v["total"]
                roi = t.get("滚动ROI2")
                rows.append([
                    t["name"],
                    fmt_num(t.get("天数")),
                    f"{t.get('天均', 0):.1f}" if t.get("天均") else "-",
                    fmt_num(t.get("例子数")),
                    fmt_num(t.get("约课数")),
                    fmt_num(t.get("滚动消耗")),
                    fmt_num(t.get("滚动GMV")),
                    f"{roi:.2f}" if isinstance(roi, (int, float)) else "-",
                ])
            sections.append({"type": "table", "headers": headers, "rows": rows})

    # ============ 6. 转介绍打卡 ============
    if data["referral"]:
        sections.append({"type": "h2", "text": "6. 转介绍打卡"})

        sections.append({"type": "h3", "text": f"6.1 后端非手推达成情况（本期 {data['referral']['cur_period']}）"})
        headers = ["类型", "例子数", "目标", "MTD达成率", "约课人数", "到课人数", "成单数", "GMV"]
        rows = []
        for r in data["referral"]["cur_rows"]:
            rows.append([
                r["name"],
                str(r["examples"]),
                str(r["target"]),
                fmt_pct(r["mtd_rate"]),
                str(r["bookings"]),
                str(r["attendances"]),
                str(r["signups"]),
                fmt_num(r["gmv"]),
            ])
        sections.append({"type": "table", "headers": headers, "rows": rows})

        if data["punch"]:
            sections.append({"type": "h3", "text": f"6.2 打卡链路数据（本期 {data['punch']['cur_period']}）"})
            cur = data["punch"]["cur"]
            prev = data["punch"]["prev"]
            headers = ["指标", f"本期({data['punch']['cur_period']})", f"上期({data['punch']['prev_period']})"]
            rows = [
                ["可打卡学员", fmt_num(cur["can_punch"]), fmt_num(prev["can_punch"])],
                ["打卡人数", fmt_num(cur["punched"]), fmt_num(prev["punched"])],
                ["打卡次数", fmt_num(cur["punch_count"]), fmt_num(prev["punch_count"])],
                ["打卡率", fmt_pct(cur["punch_rate"]), fmt_pct(prev["punch_rate"])],
                ["人均打卡次数", f"{cur['avg_punch']:.2f}", f"{prev['avg_punch']:.2f}"],
                ["例子数", fmt_num(cur["examples"]), fmt_num(prev["examples"])],
                ["打卡裂变率", fmt_pct(cur["split_rate"]), fmt_pct(prev["split_rate"])],
                ["约课数", fmt_num(cur["bookings"]), fmt_num(prev["bookings"])],
                ["约课率", fmt_pct(cur["book_rate"]), fmt_pct(prev["book_rate"])],
                ["到课数", fmt_num(cur["attendances"]), fmt_num(prev["attendances"])],
                ["约课到课率", fmt_pct(cur["attend_rate"]), fmt_pct(prev["attend_rate"])],
                ["GMV", fmt_num(cur["gmv"]), fmt_num(prev["gmv"])],
            ]
            sections.append({"type": "table", "headers": headers, "rows": rows})

    return sections


# ============ 写入飞书文档 ============

def sections_to_blocks(sections: list) -> list:
    """把 sections 转换为飞书文档块（不含表格 - 表格需要单独处理）"""
    blocks = []
    table_specs = []  # 收集表格信息，后续单独写入

    for i, sec in enumerate(sections):
        t = sec["type"]
        if t == "h1":
            blocks.append(heading_block(sec["text"], level=1))
        elif t == "h2":
            blocks.append(heading_block(sec["text"], level=2))
        elif t == "h3":
            blocks.append(heading_block(sec["text"], level=3))
        elif t == "text":
            blocks.append(text_block(sec["text"]))
        elif t == "table":
            # 占位段落，记录表格位置
            table_specs.append({
                "block_index_in_blocks": len(blocks),
                "headers": sec["headers"],
                "rows": sec["rows"],
            })
            # 留一个空文本块作为占位
            blocks.append(text_block("（表格如下）", bold=True))

    return blocks, table_specs


def insert_table(token: str, document_id: str, parent_block_id: str,
                  insert_index: int, headers: list, rows: list):
    """在文档中插入一个表格"""
    n_cols = len(headers)
    n_rows = len(rows) + 1  # 含表头

    # 1. 创建表格块
    table_block_def = {
        "block_type": 31,  # table
        "table": {
            "property": {
                "row_size": n_rows,
                "column_size": n_cols,
                "header_row": True,
            }
        }
    }

    headers_api = {"Authorization": f"Bearer {token}"}
    resp = requests.post(
        f"{FEISHU_BASE}/docx/v1/documents/{document_id}/blocks/{parent_block_id}/children",
        headers=headers_api,
        json={"children": [table_block_def], "index": insert_index}
    )
    resp.raise_for_status()
    data = resp.json()
    if data.get("code") != 0:
        raise Exception(f"创建表格失败: {data}")

    # 取出表格的 block_id 和单元格 block_id
    children = data["data"].get("children", [])
    if not children:
        return

    table_block_id = children[0]["block_id"]
    # 飞书 API 创建表格后，表格内部已自动生成 cells
    # 需要再次查询表格的子块（每个 cell 是一个 table_cell 块，里面再有 text 块）
    time.sleep(0.5)
    resp = requests.get(
        f"{FEISHU_BASE}/docx/v1/documents/{document_id}/blocks/{table_block_id}/children",
        headers=headers_api,
        params={"page_size": 500}
    )
    resp.raise_for_status()
    cell_data = resp.json()
    cell_blocks = cell_data["data"].get("items", [])

    # 现在每个 cell 块下应该有一个 text 块作为子块
    # 先获取每个 cell 的子块（text block）
    all_data = headers + [v for row in rows for v in row]

    # 批量更新文本：飞书 API 中，需要 update text in each cell's text block
    # 我们改用 batch_update API
    requests_list = []
    for idx, cell in enumerate(cell_blocks):
        cell_id = cell["block_id"]
        # cell 块内已经自带一个 text 块
        cell_children = cell.get("children", [])
        if not cell_children:
            continue
        text_block_id = cell_children[0]
        if idx >= len(all_data):
            break
        content = str(all_data[idx]) if all_data[idx] is not None else ""
        is_header = idx < n_cols
        requests_list.append({
            "block_id": text_block_id,
            "update_text_elements": {
                "elements": [text_run(content, bold=is_header)]
            }
        })

    # batch update
    if requests_list:
        resp = requests.patch(
            f"{FEISHU_BASE}/docx/v1/documents/{document_id}/blocks/batch_update",
            headers=headers_api,
            json={"requests": requests_list}
        )
        resp.raise_for_status()
        data = resp.json()
        if data.get("code") != 0:
            print(f"  ⚠️ 批量更新单元格失败: {data}")


def write_to_doc(token: str, document_id: str, sections: list):
    """把 sections 写入文档"""
    headers = {"Authorization": f"Bearer {token}"}

    # 1. 清空文档（先获取所有顶层子块，全部删除）
    resp = requests.get(
        f"{FEISHU_BASE}/docx/v1/documents/{document_id}/blocks",
        headers=headers,
        params={"page_size": 500}
    )
    resp.raise_for_status()
    data = resp.json()

    # 文档的 root block_id 等于 document_id
    items = data["data"].get("items", [])
    # 找出根块的直接子块 - 通过 parent_id 判断
    root_children = [b for b in items if b.get("parent_id") == document_id]

    if root_children:
        print(f"  清空文档已有 {len(root_children)} 个块")
        delete_blocks(token, document_id, document_id, 0, len(root_children))

    # 2. 把 sections 转换为 blocks，先全部插入文本块
    print("  转换内容块...")

    # 简化方案：把所有内容拍平为文本块，表格也用文本块（管道分隔）
    flat_blocks = []
    for sec in sections:
        t = sec["type"]
        if t == "h1":
            flat_blocks.append(heading_block(sec["text"], level=1))
        elif t == "h2":
            flat_blocks.append(heading_block(sec["text"], level=2))
        elif t == "h3":
            flat_blocks.append(heading_block(sec["text"], level=3))
        elif t == "text":
            flat_blocks.append(text_block(sec["text"]))
        elif t == "table":
            # 表格处理为 codeblock 文本，飞书会渲染为对齐文本
            n_cols = len(sec["headers"])
            # 计算每列最大宽度
            col_widths = [len(str(h)) for h in sec["headers"]]
            for row in sec["rows"]:
                for i, v in enumerate(row):
                    col_widths[i] = max(col_widths[i], len(str(v)))

            def fmt_row(row):
                return " | ".join(f"{str(v):<{col_widths[i]}}" for i, v in enumerate(row))

            text_lines = []
            text_lines.append(fmt_row(sec["headers"]))
            text_lines.append("-+-".join("-" * w for w in col_widths))
            for row in sec["rows"]:
                text_lines.append(fmt_row(row))

            # 用代码块写入（保持对齐）
            code_block = {
                "block_type": 14,  # code
                "code": {
                    "elements": [text_run("\n".join(text_lines))],
                    "style": {"language": 1}  # plain text
                }
            }
            flat_blocks.append(code_block)

    # 3. 批量追加到文档
    print(f"  追加 {len(flat_blocks)} 个块到文档...")
    # 飞书 API 单次批量 ≤ 50 个块
    chunk = 30
    for i in range(0, len(flat_blocks), chunk):
        batch = flat_blocks[i:i+chunk]
        append_blocks(token, document_id, document_id, batch)
        time.sleep(0.5)


# ============ 主流程 ============

def main():
    print("=" * 60)
    print("飞书文档发布")
    print("=" * 60)

    cfg = load_config()
    print(f"\n[1] 获取 access token...")
    token = get_tenant_access_token(cfg["app_id"], cfg["app_secret"])
    print(f"  token: {token[:20]}...")

    print(f"\n[2] 查找/创建文件夹: {cfg['folder_name']}")
    folder_token = find_or_create_folder(token, cfg["folder_name"])

    today = datetime.now().date()
    yesterday = today - timedelta(days=1)
    date_str = yesterday.strftime("%Y%m%d")
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
    sections = build_content_blocks(data, yesterday.strftime("%Y年%m月%d日"))
    print(f"  共 {len(sections)} 个章节")

    print(f"\n[6] 写入飞书文档...")
    write_to_doc(token, doc_id, sections)

    doc_url = f"https://feishu.cn/docx/{doc_id}"
    print(f"\n{'='*60}")
    print(f"✅ 完成！")
    print(f"📄 文档链接: {doc_url}")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
