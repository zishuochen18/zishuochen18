"""
生成整合的 HTML 周报页面 v3
包含：1. 港澳商务流速 2. 本月KOL转化数据 3. TMK 做工周报
带左侧目录导航
"""
import sys
from datetime import datetime, timedelta
from pathlib import Path
from openpyxl import load_workbook

sys.stdout.reconfigure(encoding="utf-8")

# 路径配置
BASE_DIR = Path(__file__).parent
FLOW_OUTPUT = BASE_DIR / "港澳流速" / "output" / "2026年6月港澳市场流速-初稿5.21.xlsx"
KOL_CONV_OUTPUT = BASE_DIR / "本月KOL转化数据汇总" / "output" / "本月KOL转化链路数据汇总.xlsx"
OUTPUT_DIR = BASE_DIR / "output"

# 把 TMK 周报目录加入 import 路径
sys.path.insert(0, str(BASE_DIR / "TMK周报"))
sys.path.insert(0, str(BASE_DIR / "书展内容"))
sys.path.insert(0, str(BASE_DIR / "线下商超内容"))
sys.path.insert(0, str(BASE_DIR / "转介绍打卡内容"))
import process_tmk  # noqa: E402
import process_book_fair  # noqa: E402
import process_shangchao  # noqa: E402
import process_referral  # noqa: E402

# KOL 转化重点指标
KEY_METRICS = ["例子约课率", "约课到课率", "到课转化率", "滚动转化率", "滚动ROI2"]


def extract_flow_data():
    """提取流速数据"""
    from pathlib import Path

    # 检查缓存是否需要重生：文件不存在，或底表比输出更新
    source = BASE_DIR / "港澳流速" / "sample" / "海外港澳商务_各渠道主辅投数据 (2).xlsx"
    output = Path(FLOW_OUTPUT)

    needs_regen = (
        not output.exists()
        or (source.exists() and source.stat().st_mtime > output.stat().st_mtime)
    )

    if needs_regen:
        print("  [重新生成] 流速底表已更新或缓存缺失，调用 process_flow...")
        from 港澳流速 import process_flow
        process_flow.main()

    wb = load_workbook(FLOW_OUTPUT, data_only=True)
    ws = wb["香港市场目标"]

    rows_data = []
    heji_count = 0
    for row_idx in range(2, 14):
        b_val = ws.cell(row_idx, 2).value
        if b_val is None:
            continue
        b_val = str(b_val).strip()

        ao = ws.cell(row_idx, 41).value
        ap = ws.cell(row_idx, 42).value
        ar = ws.cell(row_idx, 44).value
        as_val = ws.cell(row_idx, 45).value
        au = ws.cell(row_idx, 47).value

        mtd_target = ao if isinstance(ao, (int, float)) else 0
        actual = ap if isinstance(ap, (int, float)) else 0
        lesson_actual = as_val if isinstance(as_val, (int, float)) else 0

        # 区分两个"合计"行：第一个=商超&社群合计，第二个=全渠道汇总
        if b_val == "合计":
            heji_count += 1
            if heji_count == 1:
                b_val = "商超&社群合计"
            else:
                b_val = "全渠道汇总"

        if isinstance(ar, (int, float)):
            lesson_target = ar
        else:
            if b_val in ("代理商场", "书展"):
                lesson_target = mtd_target * 0.78
            elif b_val == "代理人汇总":
                lesson_target = mtd_target * 0.6
            elif b_val in ("KOL-汇总", "商超&社群合计", "全渠道汇总"):
                lesson_target = 0
            else:
                lesson_target = mtd_target * 0.55

        rows_data.append({
            "name": b_val,
            "mtd_target": mtd_target,
            "actual": actual,
            "gap": actual - mtd_target,
            "lesson_target": round(lesson_target, 1),
            "lesson_actual": lesson_actual,
            "lesson_gap": round(lesson_actual - lesson_target, 1),
            "cost": au if isinstance(au, (int, float)) else None,
        })

    kol_names = ["钟嘉欣图片", "钟嘉欣视频1", "钟嘉欣视频", "其他汇总"]
    dlz_names = ["出席礼品测试", "edm"]
    shangchao_names = ["代理商场", "书展", "澳門展会", "代理人汇总"]

    kol_items = [r for r in rows_data if r["name"] in kol_names]
    dlz_items = [r for r in rows_data if r["name"] in dlz_names]
    shangchao_items = [r for r in rows_data if r["name"] in shangchao_names]

    for r in rows_data:
        if r["name"] == "KOL-汇总":
            r["mtd_target"] = sum(x["mtd_target"] for x in kol_items)
            r["actual"] = sum(x["actual"] for x in kol_items)
            r["gap"] = r["actual"] - r["mtd_target"]
            r["lesson_target"] = round(r["mtd_target"] * 0.55, 1)
            r["lesson_actual"] = sum(x["lesson_actual"] for x in kol_items)
            r["lesson_gap"] = round(r["lesson_actual"] - r["lesson_target"], 1)
        elif r["name"] == "商超&社群合计":
            r["mtd_target"] = sum(x["mtd_target"] for x in shangchao_items)
            r["actual"] = sum(x["actual"] for x in shangchao_items)
            r["gap"] = r["actual"] - r["mtd_target"]
            r["lesson_target"] = sum(x["lesson_target"] for x in shangchao_items)
            r["lesson_actual"] = sum(x["lesson_actual"] for x in shangchao_items)
            r["lesson_gap"] = round(r["lesson_actual"] - r["lesson_target"], 1)
        elif r["name"] == "全渠道汇总":
            all_items = kol_items + dlz_items + shangchao_items
            r["mtd_target"] = sum(x["mtd_target"] for x in all_items)
            r["actual"] = sum(x["actual"] for x in all_items)
            r["gap"] = r["actual"] - r["mtd_target"]
            r["lesson_target"] = sum(x["lesson_target"] for x in all_items)
            r["lesson_actual"] = sum(x["lesson_actual"] for x in all_items)
            r["lesson_gap"] = round(r["lesson_actual"] - r["lesson_target"], 1)

    return rows_data


def extract_kol_data():
    """提取【本月汇总数据】sheet 完整数据"""
    from pathlib import Path

    # 如果缓存文件不存在，先重新生成
    if not Path(KOL_CONV_OUTPUT).exists():
        print("  [重新生成] KOL转化数据缓存不存在，调用process_kol_conversion...")
        from 本月KOL转化数据汇总 import process_kol_conversion
        process_kol_conversion.main()

    wb = load_workbook(KOL_CONV_OUTPUT, data_only=True)
    ws = wb["本月汇总数据"]

    # 读取标题行（行2）
    headers = []
    for col in range(2, ws.max_column + 1):
        h = ws.cell(2, col).value
        if h:
            headers.append((col, str(h).strip()))

    # 读取所有数据行（行3-行20，去除空行）
    all_rows = []
    for row in range(3, ws.max_row + 1):
        name = ws.cell(row, 1).value
        if not name:
            continue
        name_str = str(name).strip()
        row_data = {"name": name_str}
        for col, h in headers:
            val = ws.cell(row, col).value
            row_data[h] = val
        all_rows.append(row_data)

    # 动态识别当月和上月（只从2026年KOLHK汇总行名提取）
    import re
    months = []
    for r in all_rows:
        if "KOLHK汇总" in r["name"] and "2026年" in r["name"]:
            match = re.search(r'(\d+)月', r["name"])
            if match:
                months.append(int(match.group(1)))

    if len(months) >= 2:
        months_sorted = sorted(set(months), reverse=True)
        current_month = months_sorted[0]
        last_month = months_sorted[1]
    else:
        # 默认值（如果识别失败）
        current_month = 6
        last_month = 5

    month_label = f"{current_month}月 vs {last_month}月"

    # 同样读取【钟嘉欣图片&视频数据】
    ws_detail = wb["钟嘉欣图片&视频数据"]
    detail_headers = []
    for col in range(2, ws_detail.max_column + 1):
        h = ws_detail.cell(2, col).value
        if h:
            detail_headers.append((col, str(h).strip()))

    detail_rows = []
    seen_headers = set()
    detail_headers_unique = []
    for col, h in detail_headers:
        if h not in seen_headers:
            detail_headers_unique.append((col, h))
            seen_headers.add(h)

    for row in range(3, 6):  # 行3=图片，行4=视频1，行5=视频
        name = ws_detail.cell(row, 1).value
        if not name:
            continue
        row_data = {"name": str(name).strip()}
        for col, h in detail_headers_unique:
            val = ws_detail.cell(row, col).value
            row_data[h] = val
        detail_rows.append(row_data)

    # 读取【钟嘉欣视频渠道明细】sheet
    video_channel_rows = []
    if "钟嘉欣视频渠道明细" in wb.sheetnames:
        ws_video = wb["钟嘉欣视频渠道明细"]
        for row in range(2, ws_video.max_row + 1):
            ch_name = ws_video.cell(row, 1).value
            if not ch_name:
                continue
            video_channel_rows.append({
                "渠道简称": ch_name,
                "例子数": ws_video.cell(row, 2).value,
                "约课数": ws_video.cell(row, 3).value,
                "滚动消耗": ws_video.cell(row, 4).value,
                "例子约课率": ws_video.cell(row, 5).value,
                "约课到课率": ws_video.cell(row, 6).value,
                "到课转化率": ws_video.cell(row, 7).value,
                "滚动转化率": ws_video.cell(row, 8).value,
                "滚动GMV": ws_video.cell(row, 9).value,
                "滚动ROI2": ws_video.cell(row, 10).value,
            })

    return {
        "all_rows": all_rows,
        "headers": [h for _, h in headers],
        "detail_rows": detail_rows,
        "detail_headers": [h for _, h in detail_headers_unique],
        "video_channel_rows": video_channel_rows,
        "current_month": current_month,
        "last_month": last_month,
        "month_label": month_label,
    }


def calc_rate(actual, target):
    if target and target > 0:
        return round(actual / target * 100, 1)
    return 0


def get_action_for_supplier(name):
    """根据供应商名称映射补缺口的具体行动"""
    if "钟嘉欣图片" in name or "钟嘉欣视频" in name:
        return "核对原素材投放出价 + 增 1 条 A/B 测试素材"
    elif "代理商场" in name:
        return "补排本周末场次 + 排查到场转化漏斗"
    elif "代理人汇总" in name:
        return "催代理人当周播报数据，确认是否归因延迟"
    elif "edm" in name or "出席礼品" in name:
        return "检查投放是否暂停 + 确认昨日发送量"
    else:
        return "TL 与负责人 1on1 复盘 gap 原因"


def module_overview(module_name, data):
    """生成模块顶部的「整体观察」段（亮点 + 风险）"""
    if module_name == "flow":
        total = data.get("total_row")
        behind = data.get("behind", [])
        ahead = data.get("ahead", [])
        if not total:
            return ""
        rate = calc_rate(total["actual"], total["mtd_target"])
        top_behind = behind[0] if behind else None
        top_ahead = ahead[0] if ahead else None

        lines = [f'<strong>全渠道 MTD 达成率 {rate}%</strong>']
        if top_ahead:
            lines.append(f'亮点：{top_ahead["name"]} 达成率 {calc_rate(top_ahead["actual"], top_ahead["mtd_target"])}%')
        if top_behind:
            lines.append(f'风险：{top_behind["name"]} 缺口 {abs(int(top_behind["gap"]))} 例子（达成率 {calc_rate(top_behind["actual"], top_behind["mtd_target"])}%）')

    elif module_name == "kol":
        zjx_may = data.get("zjx_may")
        if not zjx_may:
            return ""
        lines = [f'<strong>钟嘉欣本月例子数 {int(zjx_may.get("例子数", 0))} 个</strong>']
        asp = zjx_may.get("滚动ASP")
        roi2 = zjx_may.get("滚动ROI2")
        roi2_str = f'{roi2:.2f}' if isinstance(roi2, (int, float)) else "-"
        if isinstance(asp, (int, float)):
            lines.append(f'亮点：客单价 {asp:.0f}，滚动 ROI2 {roi2_str}')
        else:
            lines.append(f'亮点：存量积累，跨月对比持续优化')
        book_rate = zjx_may.get("约课到课率")
        book_rate_str = f'{book_rate*100:.1f}%' if isinstance(book_rate, (int, float)) else "-"
        lines.append(f'风险：约课到课率 {book_rate_str}（需关注转化漏斗）')

    elif module_name == "tmk":
        rows = data.get("rows", [])
        anomalies = [r for r in rows if r.get("异常")]
        if not rows:
            return ""
        lines = [f'<strong>TMK 做工 {len(rows)} 人，异常项 {len(anomalies)} 个</strong>']
        if anomalies:
            worst = max(anomalies, key=lambda x: x.get("异常", 0))
            lines.append(f'亮点：基础数据已齐全')
            lines.append(f'风险：{worst.get("TMK", "个人")} 异常值最高（需 TL 排查）')
        else:
            lines.append(f'亮点：本周异常项为零')
            lines.append(f'风险：暂无')

    elif module_name == "bookfair":
        current = data.get("current_names", [])
        current_count = len([v for v in data.get("current_month_venues", []) if v["total"].get("例子数", 0) > 0])
        history_count = len(data.get("history", []))
        lines = [f'<strong>书展本月 {current_count} 场，历史参考 {history_count} 场</strong>']

        # 统计 ROI2
        current_venues = data.get("current_month_venues", [])
        if current_venues:
            roi2_vals = [v["total"].get("滚动ROI2") for v in current_venues if isinstance(v["total"].get("滚动ROI2"), (int, float))]
            if roi2_vals:
                avg_roi2 = sum(roi2_vals) / len(roi2_vals)
                lines.append(f'亮点：本月 ROI2 平均 {avg_roi2:.2f}（目标 0.80，' +
                            ('已达标' if avg_roi2 >= 0.8 else f'差 {(0.8 - avg_roi2):.2f}') + '）')
            else:
                lines.append(f'亮点：数据已就绪')
        else:
            lines.append(f'亮点：历史数据完整')

        if current_count == 0:
            lines.append(f'风险：本月暂无书展场次')
        else:
            lines.append(f'风险：当月场次仍待后续补充，关注 ROI2 达成')

    elif module_name == "shangchao":
        cur_month = data.get("current_month")
        prev_month = data.get("prev_month")
        if not cur_month:
            return ""
        ex_cur = cur_month.get("例子数", 0)
        ex_prev = prev_month.get("例子数", 0) if prev_month else 0
        trend = "↑" if ex_cur > ex_prev else "↓" if ex_cur < ex_prev else "→"
        lines = [f'<strong>线下商超本月例子 {int(ex_cur)} 个</strong>']
        lines.append(f'亮点：{cur_month.get("name", "本月")}出摊 {int(cur_month.get("天数", 0))} 天')

        # ROI2 描述
        roi2 = cur_month.get("滚动ROI2")
        roi2_str = f'{roi2:.2f}' if isinstance(roi2, (int, float)) else "-"
        if isinstance(roi2, (int, float)):
            roi2_status = '已达标' if roi2 >= 0.8 else f'差 {(0.8 - roi2):.2f}'
            roi2_desc = f'ROI2 {roi2_str}（目标 0.80，{roi2_status}）'
        else:
            roi2_desc = "ROI2 数据待完善"

        if prev_month and ex_cur != ex_prev:
            lines.append(f'风险：环比上月 {trend} {abs(int(ex_cur - ex_prev))} 个例子，{roi2_desc}')
        else:
            lines.append(f'风险：{roi2_desc}，需关注场次频次')

    elif module_name == "referral":
        cur = data.get("cur")
        prev = data.get("prev")
        if not cur or not prev:
            return ""
        cur_rate = cur.get("total_conv_rate", 0) * 100
        prev_rate = prev.get("total_conv_rate", 0) * 100
        book_rate = cur.get("book_rate", 0) * 100
        lines = [f'<strong>转介绍本期打卡裂变率 {cur.get("split_rate", 0)*100:.2f}%</strong>']
        lines.append(f'亮点：约课人数 {cur.get("bookings", 0)} 人，约课率 {book_rate:.1f}%')
        if cur_rate > prev_rate:
            lines.append(f'风险：注册转化率 {cur_rate:.2f}%（环比 ↑ {cur_rate - prev_rate:.2f}%，需保持）')
        else:
            lines.append(f'风险：注册转化率 {cur_rate:.2f}%（环比 ↓ {prev_rate - cur_rate:.2f}%，需排查漏斗）')

    else:
        return ""

    if not lines:
        return ""
    return f'<div class="analysis-block">{"<br>".join(lines)}</div>'


def fmt_by_field(v, field_name):
    """根据字段名格式化数值"""
    if v is None or v == "":
        return "-"
    if not isinstance(v, (int, float)):
        return str(v)

    # 滚动ROI2总成本 → 千分位整数
    if "滚动ROI2总成本" in field_name or field_name == "滚动ROI2总成本":
        return f"{v:,.0f}"

    # 滚动ROI2 → 保留 2 位小数（不带百分号）
    if field_name.strip() == "滚动ROI2":
        return f"{v:.2f}"

    # 百分比字段（含"率"或"占比"）
    if any(k in field_name for k in ["率", "占比"]):
        # 数据底表里的比率默认是 0~1 的小数
        if abs(v) <= 1.5:
            return f"{v * 100:.1f}%"
        return f"{v:.1f}%"

    # 整数字段（例子/分发/约课等）
    if any(k in field_name for k in ["例子数", "分发数", "约课数", "应到课数", "到课数", "成交数", "人数", "人次"]):
        return f"{int(v):,}"

    # 金额字段（千分位整数）
    if any(k in field_name for k in ["消耗", "成本", "GMV", "费用", "ASP", "CPS", "CPT"]):
        if abs(v) >= 1000:
            return f"{v:,.0f}"
        return f"{v:.2f}"

    # 默认
    if abs(v) >= 1000:
        return f"{v:,.0f}"
    return f"{v:.2f}"


def generate_kol_analysis(kol_data):
    """生成 KOL 转化数据的文字分析"""
    rows = kol_data["all_rows"]
    current_month = kol_data["current_month"]
    last_month = kol_data["last_month"]

    # 获取关键行（使用动态月份）
    zjx_may = next((r for r in rows if f"{current_month}月钟嘉欣" in r["name"]), None)
    zjx_apr = next((r for r in rows if f"{last_month}月钟嘉欣" in r["name"]), None)
    zjx_sum = next((r for r in rows if r["name"] == "钟嘉欣-汇总"), None)
    kolhk_may = next((r for r in rows if f"{current_month}月-KOLHK汇总" in r["name"]), None)
    kolhk_apr = next((r for r in rows if f"{last_month}月-KOLHK汇总" in r["name"]), None)

    analysis = []

    # 钟嘉欣分析
    if zjx_may and zjx_apr and zjx_sum:
        zjx_lines = ["<strong>钟嘉欣</strong>："]
        for metric in KEY_METRICS:
            v_may = zjx_may.get(metric)
            v_apr = zjx_apr.get(metric)
            v_sum = zjx_sum.get(metric)
            if isinstance(v_may, (int, float)) and isinstance(v_apr, (int, float)) and v_apr > 0:
                change = (v_may - v_apr) / v_apr * 100
                trend = "↑" if change > 0 else "↓"
                color = "positive" if change > 0 else "negative"
                v_may_disp = f"{v_may * 100:.1f}%" if v_may <= 1 else f"{v_may:.2f}"
                v_apr_disp = f"{v_apr * 100:.1f}%" if v_apr <= 1 else f"{v_apr:.2f}"
                v_sum_disp = f"{v_sum * 100:.1f}%" if isinstance(v_sum, (int, float)) and v_sum <= 1 else (f"{v_sum:.2f}" if isinstance(v_sum, (int, float)) else "-")
                zjx_lines.append(
                    f'· {metric}：{current_month}月 <strong>{v_may_disp}</strong>，{last_month}月 {v_apr_disp}（环比 <span class="{color}">{trend} {abs(change):.1f}%</span>），1-{current_month}月汇总 {v_sum_disp}'
                )
        analysis.append("<br>".join(zjx_lines))

    # KOLHK 整体
    if kolhk_may and kolhk_apr:
        kolhk_lines = ["<strong>KOLHK 整体</strong>："]
        for metric in KEY_METRICS:
            v_may = kolhk_may.get(metric)
            v_apr = kolhk_apr.get(metric)
            if isinstance(v_may, (int, float)) and isinstance(v_apr, (int, float)) and v_apr > 0:
                change = (v_may - v_apr) / v_apr * 100
                trend = "↑" if change > 0 else "↓"
                color = "positive" if change > 0 else "negative"
                v_may_disp = f"{v_may * 100:.1f}%" if v_may <= 1 else f"{v_may:.2f}"
                v_apr_disp = f"{v_apr * 100:.1f}%" if v_apr <= 1 else f"{v_apr:.2f}"
                kolhk_lines.append(
                    f'· {metric}：{current_month}月 <strong>{v_may_disp}</strong>，{last_month}月 {v_apr_disp}（环比 <span class="{color}">{trend} {abs(change):.1f}%</span>）'
                )
        analysis.append("<br>".join(kolhk_lines))

    # 综合判断
    judgements = []
    if zjx_may and kolhk_may:
        # 看钟嘉欣的滚动ROI2 vs KOLHK整体
        zjx_roi = zjx_may.get("滚动ROI2")
        kol_roi = kolhk_may.get("滚动ROI2")
        if isinstance(zjx_roi, (int, float)) and isinstance(kol_roi, (int, float)):
            if zjx_roi > kol_roi:
                judgements.append(f'✅ 钟嘉欣 {current_month}月滚动ROI2 ({zjx_roi*100:.1f}%) 高于 KOLHK 整体 ({kol_roi*100:.1f}%)，表现优异')
            else:
                judgements.append(f'⚠️ 钟嘉欣 {current_month}月滚动ROI2 ({zjx_roi*100:.1f}%) 低于 KOLHK 整体 ({kol_roi*100:.1f}%)，需要关注')

    return analysis, judgements


def generate_detail_analysis(kol_data):
    """生成钟嘉欣图片&视频明细的文字分析"""
    rows = kol_data["detail_rows"]
    if len(rows) < 2:
        return []

    pic = rows[0] if rows[0]["name"] == "钟嘉欣图片" else rows[1]
    v1 = rows[1] if rows[1]["name"] == "钟嘉欣视频1" else rows[0]

    lines = []
    for metric in KEY_METRICS:
        v_pic = pic.get(metric)
        v_v1 = v1.get(metric)
        if isinstance(v_pic, (int, float)) and isinstance(v_v1, (int, float)):
            v_pic_disp = f"{v_pic * 100:.1f}%" if v_pic <= 1 else f"{v_pic:.2f}"
            v_v1_disp = f"{v_v1 * 100:.1f}%" if v_v1 <= 1 else f"{v_v1:.2f}"
            better = "图片" if v_pic > v_v1 else "视频1"
            lines.append(f'· {metric}：图片 <strong>{v_pic_disp}</strong> vs 视频1 <strong>{v_v1_disp}</strong>，<span class="positive">{better}更优</span>')

    return lines


def row_html_flow(r, is_summary=False):
    rate = calc_rate(r["actual"], r["mtd_target"])
    lesson_rate = calc_rate(r["lesson_actual"], r["lesson_target"])
    gap_class = "negative" if r["gap"] < 0 else "positive" if r["gap"] > 0 else ""
    lgap_class = "negative" if r["lesson_gap"] < 0 else "positive" if r["lesson_gap"] > 0 else ""
    bold = "font-weight:700;" if is_summary else ""
    cost_str = f'{r["cost"]:.1f}' if r["cost"] else "-"

    # 异常行标记：达成率 < 80% 且非汇总行时整行变浅红
    row_class = ""
    if not is_summary and rate < 80:
        row_class = ' class="row-anomaly"'

    return f"""<tr{row_class} style="{bold}">
<td>{r["name"]}</td>
<td>{int(r["mtd_target"])}</td><td>{int(r["actual"])}</td>
<td class="{gap_class}">{int(r["gap"])}</td><td>{rate}%</td>
<td>{int(r["lesson_target"]) if r["lesson_target"] else '-'}</td><td>{int(r["lesson_actual"])}</td>
<td class="{lgap_class}">{int(r["lesson_gap"]) if r["lesson_gap"] else '-'}</td><td>{lesson_rate}%</td>
<td>{cost_str}</td>
</tr>"""


def kol_table_html(rows, headers, highlight_metrics=None):
    """生成 KOL 数据表格 HTML（完整列 + 首列冻结）"""
    if highlight_metrics is None:
        highlight_metrics = KEY_METRICS

    html = '<div class="table-wrap"><table class="kol-table">\n<thead><tr><th class="sticky-col">项目</th>'
    for h in headers:
        cls = "highlight-col" if h in highlight_metrics else ""
        html += f'<th class="{cls}">{h}</th>'
    html += "</tr></thead>\n<tbody>"

    for r in rows:
        is_summary = "汇总" in r["name"]
        cls = "summary-row" if is_summary else ""
        html += f'<tr class="{cls}"><td class="sticky-col"><strong>{r["name"]}</strong></td>'
        for h in headers:
            v = r.get(h)
            cls_td = "highlight-col" if h in highlight_metrics else ""
            html += f'<td class="{cls_td}">{fmt_by_field(v, h)}</td>'
        html += "</tr>"

    html += "\n</tbody></table></div>"
    return html


def metric_block(title, may_row, prev_row, sum_row=None, sum_label="汇总"):
    """生成某主体的关键指标环比分析块"""
    if not may_row:
        return ""

    lines = [f'<strong>{title}</strong>：']
    for metric in KEY_METRICS:
        v_may = may_row.get(metric)
        v_prev = prev_row.get(metric) if prev_row else None
        v_sum = sum_row.get(metric) if sum_row else None

        def disp(v):
            if not isinstance(v, (int, float)):
                return "-"
            # 滚动ROI2 用小数
            if metric == "滚动ROI2":
                return f"{v:.2f}"
            # 比率
            if v <= 1.5:
                return f"{v * 100:.1f}%"
            return f"{v:.2f}"

        v_may_disp = disp(v_may)
        v_prev_disp = disp(v_prev)
        v_sum_disp = disp(v_sum) if v_sum is not None else None

        # 环比
        change_html = ""
        if isinstance(v_may, (int, float)) and isinstance(v_prev, (int, float)) and v_prev != 0:
            change = (v_may - v_prev) / abs(v_prev) * 100
            trend = "↑" if change > 0 else "↓"
            color = "positive" if change > 0 else "negative"
            change_html = f'（环比 <span class="{color}">{trend} {abs(change):.1f}%</span>）'

        sum_part = f"，{sum_label} {v_sum_disp}" if v_sum_disp else ""
        lines.append(f'· {metric}：本月 <strong>{v_may_disp}</strong>，上月 {v_prev_disp}{change_html}{sum_part}')

    return f'<div class="analysis-block">{"<br>".join(lines)}</div>'


def detail_compare_block(detail_rows):
    """生成钟嘉欣图片 vs 视频1 vs 视频 的对比分析块"""
    if not detail_rows:
        return ""

    by_name = {r["name"]: r for r in detail_rows}
    items = []
    for n in ["钟嘉欣图片", "钟嘉欣视频1", "钟嘉欣视频"]:
        if n in by_name:
            items.append((n, by_name[n]))

    if len(items) < 2:
        return ""

    def disp(v, metric):
        if not isinstance(v, (int, float)):
            return "-"
        if metric == "滚动ROI2":
            return f"{v:.2f}"
        if v <= 1.5:
            return f"{v * 100:.1f}%"
        return f"{v:.2f}"

    lines = ["<strong>钟嘉欣 图片 vs 视频对比</strong>："]
    for metric in KEY_METRICS:
        parts = []
        valid_items = []
        for name, row in items:
            v = row.get(metric)
            parts.append(f'{name.replace("钟嘉欣", "")} <strong>{disp(v, metric)}</strong>')
            if isinstance(v, (int, float)):
                valid_items.append((name, v))

        if len(valid_items) >= 2:
            best = max(valid_items, key=lambda x: x[1])
            best_label = best[0].replace("钟嘉欣", "")
            lines.append(
                f'· {metric}：{" vs ".join(parts)} → <span class="positive">{best_label}更优</span>'
            )
        else:
            lines.append(f'· {metric}：{" vs ".join(parts)}')

    return f'<div class="analysis-block">{"<br>".join(lines)}</div>'


# ============ TMK 周报相关函数 ============

def fmt_tmk(val, kind):
    """TMK 数据格式化"""
    import pandas as pd
    if val is None or pd.isna(val):
        return "-"
    try:
        v = float(val)
    except (ValueError, TypeError):
        return str(val)
    if kind == "pct":
        return f"{v * 100:.1f}%"
    elif kind == "int":
        return f"{int(round(v)):,}"
    else:
        return f"{v:.2f}"


def tmk_table_html(tmk_data):
    """生成 TMK 周报数据表格（带分组表头 + 首列冻结）"""
    rows = tmk_data["rows"]
    groups_order = tmk_data["groups_order"]
    group_to_cols = tmk_data["group_to_cols"]

    # 计算每个分组的列数
    html = ['<div class="table-scroll-tall"><table class="kol-table tmk-table">']

    # 第一行：分组表头
    html.append('<thead>')
    html.append('<tr>')
    html.append('<th class="sticky-col" rowspan="2">TMK小组</th>')
    html.append('<th class="sticky-col-2" rowspan="2">TMK</th>')
    for g in groups_order:
        n = len(group_to_cols[g])
        html.append(f'<th colspan="{n}">{g}</th>')
    html.append('</tr>')

    # 第二行：列名
    html.append('<tr>')
    for g in groups_order:
        for col_name, _ in group_to_cols[g]:
            html.append(f'<th>{col_name}</th>')
    html.append('</tr>')
    html.append('</thead>')

    # 数据行
    html.append('<tbody>')
    for r in rows:
        tmk_val = str(r.get("TMK", "")).strip()
        group_val = str(r.get("TMK小组", "")).strip()
        is_grand_total = tmk_val == "总计" and group_val == "总计"
        is_team_total = tmk_val == "团队汇总"

        cls = ""
        if is_grand_total:
            cls = "tmk-grand-total"
        elif is_team_total:
            cls = "tmk-team-total"

        html.append(f'<tr class="{cls}">')
        html.append(f'<td class="sticky-col">{group_val}</td>')
        html.append(f'<td class="sticky-col-2">{tmk_val}</td>')

        for g in groups_order:
            for col_name, kind in group_to_cols[g]:
                v = r.get(f"{g}__{col_name}")
                html.append(f'<td>{fmt_tmk(v, kind)}</td>')
        html.append('</tr>')
    html.append('</tbody>')
    html.append('</table></div>')
    return "\n".join(html)


def tmk_overview_block(tmk_data):
    """TMK 整体数据概览块"""
    overall = tmk_data["overall"]
    params = tmk_data["params"]

    start_date = params.get("start_date", "?")
    end_date = params.get("end_date", "?")

    lines = []
    if "日均通次" in overall:
        lines.append(f'· 团队日均通次 <strong>{int(overall["日均通次"])}</strong> 次')

    if "日均通次环比" in overall:
        v = overall["日均通次环比"]
        trend = "↑" if v > 0 else "↓"
        color = "positive" if v > 0 else "negative"
        direction = "提升" if v > 0 else "下降"
        lines.append(f'  环比{direction} <span class="{color}">{trend} {abs(v) * 100:.1f}%</span>')

    if "日均通时（分）" in overall:
        lines.append(f'· 团队日均通时 <strong>{overall["日均通时（分）"]:.1f}</strong> 分钟')

    if "生均跟进时效达成率" in overall:
        rate = overall["生均跟进时效达成率"] * 100
        color = "positive" if rate >= 50 else ("negative" if rate < 30 else "")
        lines.append(f'· 生均跟进时效达成率 <strong class="{color}">{rate:.1f}%</strong>')

    return f"""<div class="analysis-block">
<strong>📅 周期</strong>：{start_date} ~ {end_date}<br>
{"<br>".join(lines)}
</div>"""


def tmk_alert_block(tmk_data):
    """TMK 个人异常情况块"""
    alerts = tmk_data["alerts"]

    if not alerts:
        return '<p style="font-size:13px;color:#666;">本周无明显异常</p>'

    # 按人聚合
    from collections import defaultdict
    by_name = defaultdict(list)
    for a in alerts:
        by_name[a["name"]].append(a["desc"])

    html = ['<ul style="font-size:13px; line-height:1.8; padding-left:20px;">']
    for name, descs in by_name.items():
        html.append(f'<li>⚠️ <strong>{name}</strong>：{" ；".join(descs)}</li>')
    html.append('</ul>')
    return "\n".join(html)


# ============ 书展复盘相关函数 ============

def book_fair_table_html(rows, headers, key_metrics):
    """生成书展数据表（首列冻结 + 重点指标高亮 + 总计行加粗）"""
    # 隐藏"滚动消耗(不含赠课成本)"列
    headers = [h for h in headers if h != "滚动消耗(不含赠课成本)"]

    html = ['<div class="table-scroll-tall"><table class="kol-table">']
    html.append('<thead><tr><th class="sticky-col">渠道名称</th>')
    for h in headers:
        cls = "highlight-col" if h in key_metrics else ""
        html.append(f'<th class="{cls}">{h}</th>')
    html.append('</tr></thead>')

    html.append('<tbody>')
    for r in rows:
        is_total = r.get("_is_total", False)
        cls = "summary-row" if is_total else ""
        html.append(f'<tr class="{cls}">')
        html.append(f'<td class="sticky-col"><strong>{r["name"]}</strong></td>')
        for h in headers:
            v = r.get(h)
            cls_td = "highlight-col" if h in key_metrics else ""
            html.append(f'<td class="{cls_td}">{fmt_by_field(v, h)}</td>')
        html.append('</tr>')
    html.append('</tbody></table></div>')
    return "\n".join(html)


def book_fair_compare_block(book_fair_data):
    """生成两个本月书展的对比分析块"""
    current = book_fair_data["current"]
    history = book_fair_data["history"]
    key_metrics = book_fair_data["key_metrics"]

    fair_names = book_fair_data["current_names"]
    if len(fair_names) < 1:
        return ""

    def disp(v, metric):
        if not isinstance(v, (int, float)):
            return "-"
        if metric == "滚动ROI2":
            return f"{v:.2f}"
        if v <= 1.5:
            return f"{v * 100:.1f}%"
        return f"{v:.2f}"

    # 取本月两个书展的总计 + 同期历史均值
    blocks = []

    # 1) 本月两个书展对比
    fair_a = current.get(fair_names[0], {}).get("total")
    fair_b = current.get(fair_names[1], {}).get("total") if len(fair_names) > 1 else None

    if fair_a and fair_b:
        lines = [f'<strong>本月双展对比 — {fair_names[0]} vs {fair_names[1]}</strong>：']
        for metric in key_metrics:
            va = fair_a.get(metric)
            vb = fair_b.get(metric)
            if isinstance(va, (int, float)) and isinstance(vb, (int, float)):
                better = fair_names[0] if va > vb else fair_names[1]
                short_better = better.replace("26年5月", "")
                lines.append(
                    f'· {metric}：{disp(va, metric)} vs {disp(vb, metric)} → '
                    f'<span class="positive">{short_better}更优</span>'
                )
        blocks.append(f'<div class="analysis-block">{"<br>".join(lines)}</div>')

    # 2) 与历史均值对比（重点指标）
    if history:
        hist_means = {}
        for metric in key_metrics:
            vals = [h.get(metric) for h in history if isinstance(h.get(metric), (int, float))]
            if vals:
                hist_means[metric] = sum(vals) / len(vals)

        if hist_means:
            for fair_name in fair_names:
                cur = current.get(fair_name, {}).get("total")
                if not cur:
                    continue
                lines = [f'<strong>{fair_name} vs 历史{len(history)}场均值</strong>：']
                for metric in key_metrics:
                    v = cur.get(metric)
                    h = hist_means.get(metric)
                    if isinstance(v, (int, float)) and isinstance(h, (int, float)) and h != 0:
                        diff = (v - h) / h * 100
                        trend = "↑" if diff > 0 else "↓"
                        color = "positive" if diff > 0 else "negative"
                        lines.append(
                            f'· {metric}：本场 <strong>{disp(v, metric)}</strong>，'
                            f'历史均值 {disp(h, metric)}（<span class="{color}">{trend} {abs(diff):.1f}%</span>）'
                        )
                blocks.append(f'<div class="analysis-block">{"<br>".join(lines)}</div>')

    return "\n".join(blocks)


# ============ 线下商超相关函数 ============

def shangchao_analysis_block(sc_data):
    """生成商超当月分析块（环比 + 重点指标）"""
    cur = sc_data["current_month"]
    prev = sc_data["prev_month"]
    total = sc_data["total"]
    key_metrics = sc_data["key_metrics"]

    if not cur:
        return ""

    def disp(v, metric):
        if not isinstance(v, (int, float)):
            return "-"
        if metric == "滚动ROI2":
            return f"{v:.2f}"
        if v <= 1.5:
            return f"{v * 100:.1f}%"
        return f"{v:.2f}"

    lines = [f'<strong>{cur["name"]}商超数据概览</strong>：']
    days = cur.get("天数")
    avg = cur.get("天均")
    if days:
        lines.append(f'· 出摊天数 <strong>{int(days)}</strong> 天，天均例子 <strong>{avg:.1f}</strong>')
    ex = cur.get("例子数")
    yk = cur.get("约课数")
    if ex:
        lines.append(f'· 例子数 <strong>{int(ex)}</strong>，约课数 <strong>{int(yk) if yk else "-"}</strong>')

    blocks = [f'<div class="analysis-block">{"<br>".join(lines)}</div>']

    # 环比上月
    if prev:
        lines2 = [f'<strong>{cur["name"]} vs {prev["name"]}（环比）</strong>：']
        for metric in key_metrics:
            vc = cur.get(metric)
            vp = prev.get(metric)
            if isinstance(vc, (int, float)) and isinstance(vp, (int, float)) and vp != 0:
                change = (vc - vp) / abs(vp) * 100
                trend = "↑" if change > 0 else "↓"
                color = "positive" if change > 0 else "negative"
                lines2.append(
                    f'· {metric}：本月 <strong>{disp(vc, metric)}</strong>，'
                    f'上月 {disp(vp, metric)}（<span class="{color}">{trend} {abs(change):.1f}%</span>）'
                )
        blocks.append(f'<div class="analysis-block">{"<br>".join(lines2)}</div>')

    # vs 历史均值
    if total:
        lines3 = [f'<strong>{cur["name"]} vs 历史整体均值</strong>：']
        for metric in key_metrics:
            vc = cur.get(metric)
            vt = total.get(metric)
            if isinstance(vc, (int, float)) and isinstance(vt, (int, float)) and vt != 0:
                change = (vc - vt) / abs(vt) * 100
                trend = "↑" if change > 0 else "↓"
                color = "positive" if change > 0 else "negative"
                lines3.append(
                    f'· {metric}：本月 <strong>{disp(vc, metric)}</strong>，'
                    f'历史均值 {disp(vt, metric)}（<span class="{color}">{trend} {abs(change):.1f}%</span>）'
                )
        blocks.append(f'<div class="analysis-block">{"<br>".join(lines3)}</div>')

    return "\n".join(blocks)


def shangchao_monthly_table(sc_data):
    """生成商超分月趋势表"""
    monthly = sc_data["monthly"]
    key_metrics = sc_data["key_metrics"]
    display_fields = [
        "天数", "天均", "CPS消耗", "CPT消耗", "滚动消耗", "例子数", "当月成交数", "约课数",
        "到课数", "滚动到课数", "滚动成交数", "滚动GMV", "滚动ASP",
        "例子约课率", "约课到课率", "到课转化率", "滚动转化率", "注册转化率",
        "滚动ROI2总成本", "滚动ROI2"
    ]

    html = ['<div class="table-scroll-tall"><table class="kol-table">']
    html.append('<thead><tr><th class="sticky-col">月份</th>')
    for f in display_fields:
        cls = "highlight-col" if f in key_metrics else ""
        html.append(f'<th class="{cls}">{f}</th>')
    html.append('</tr></thead><tbody>')

    for r in monthly:
        is_total = r["name"] == "汇总"
        cls = "summary-row" if is_total else ""
        html.append(f'<tr class="{cls}"><td class="sticky-col"><strong>{r["name"]}</strong></td>')
        for f in display_fields:
            v = r.get(f)
            cls_td = "highlight-col" if f in key_metrics else ""
            html.append(f'<td class="{cls_td}">{fmt_by_field(v, f)}</td>')
        html.append('</tr>')

    html.append('</tbody></table></div>')
    return "\n".join(html)


def shangchao_last_month_venue_table(sc_data):
    """生成上月场次表（5月场次 + 承担人明细，含跨月累计数据）"""
    last_month_venues = sc_data.get("last_month_venues", [])
    key_metrics = sc_data["key_metrics"]

    if not last_month_venues:
        return '<p style="font-size:13px;color:#666;">上月无场次数据</p>'

    display_fields = [
        "天数", "天均", "CPS成本", "CPT成本", "滚动消耗", "例子数", "分发数", "约课数",
        "到课数", "滚动到课数", "当月成交数", "滚动成交数", "滚动GMV", "滚动ASP",
        "例子约课率", "约课到课率", "到课转化率", "滚动转化率", "注册转化率",
        "滚动ROI2总成本", "滚动ROI2"
    ]

    html = ['<div class="table-scroll-tall"><table class="kol-table">']
    html.append('<thead><tr><th class="sticky-col">场次/承担人</th>')
    for f in display_fields:
        cls = "highlight-col" if f in key_metrics else ""
        html.append(f'<th class="{cls}">{f}</th>')
    html.append('</tr></thead><tbody>')

    for venue in last_month_venues:
        t = venue["total"]
        html.append(f'<tr class="summary-row"><td class="sticky-col"><strong>{t["name"]}</strong></td>')
        for f in display_fields:
            v = t.get(f)
            cls_td = "highlight-col" if f in key_metrics else ""
            html.append(f'<td class="{cls_td}">{fmt_by_field(v, f)}</td>')
        html.append('</tr>')

        for d in venue["details"]:
            html.append(f'<tr><td class="sticky-col" style="padding-left:24px!important;">{d["name"]}</td>')
            for f in display_fields:
                v = d.get(f)
                cls_td = "highlight-col" if f in key_metrics else ""
                html.append(f'<td class="{cls_td}">{fmt_by_field(v, f)}</td>')
            html.append('</tr>')

    html.append('</tbody></table></div>')
    return "\n".join(html)


def shangchao_current_month_venue_table(sc_data):
    """生成当月场次表（6月场次 + 承担人明细，从数据底表实时构建）"""
    current_month_venues = sc_data.get("current_month_venues", [])
    key_metrics = sc_data["key_metrics"]

    if not current_month_venues:
        return '<p style="font-size:13px;color:#666;">当月无场次数据</p>'

    # 只展示例子数>0 的场次（即本月有效场次）
    valid_venues = []
    for v in current_month_venues:
        ex = v["total"].get("例子数") or 0
        if ex > 0:
            valid_venues.append(v)

    display_fields = [
        "天数", "天均", "CPS成本", "CPT成本", "滚动消耗", "例子数", "分发数", "约课数",
        "到课数", "滚动到课数", "滚动成交数", "滚动GMV", "滚动ASP",
        "例子约课率", "约课到课率", "到课转化率", "滚动转化率",
        "滚动ROI2总成本", "滚动ROI2"
    ]

    html = ['<div class="table-scroll-tall"><table class="kol-table">']
    html.append('<thead><tr><th class="sticky-col">场次/承担人</th>')
    for f in display_fields:
        cls = "highlight-col" if f in key_metrics else ""
        html.append(f'<th class="{cls}">{f}</th>')
    html.append('</tr></thead><tbody>')

    for venue in valid_venues:
        t = venue["total"]
        html.append(f'<tr class="summary-row"><td class="sticky-col"><strong>{t["name"]}</strong></td>')
        for f in display_fields:
            v = t.get(f)
            cls_td = "highlight-col" if f in key_metrics else ""
            html.append(f'<td class="{cls_td}">{fmt_by_field(v, f)}</td>')
        html.append('</tr>')

        for d in venue["details"]:
            html.append(f'<tr><td class="sticky-col" style="padding-left:24px!important;">{d["name"]}</td>')
            for f in display_fields:
                v = d.get(f)
                cls_td = "highlight-col" if f in key_metrics else ""
                html.append(f'<td class="{cls_td}">{fmt_by_field(v, f)}</td>')
            html.append('</tr>')

    html.append('</tbody></table></div>')
    return "\n".join(html)


def shangchao_venue_table(sc_data):
    """生成当月分场次表（只展示当月场次 + 承担人明细，从数据底表实时构建）"""
    current_venues = sc_data.get("current_venues", [])
    key_metrics = sc_data["key_metrics"]

    if not current_venues:
        return '<p style="font-size:13px;color:#666;">当月无场次数据</p>'

    # 只展示例子数>0 的场次（即本月有效场次）
    valid_venues = []
    for v in current_venues:
        ex = v["total"].get("例子数") or 0
        if ex > 0:
            valid_venues.append(v)

    display_fields = [
        "天数", "天均", "CPS成本", "CPT成本", "滚动消耗", "例子数", "分发数", "约课数",
        "到课数", "滚动到课数", "滚动成交数", "滚动GMV", "滚动ASP",
        "例子约课率", "约课到课率", "到课转化率", "滚动转化率",
        "滚动ROI2总成本", "滚动ROI2"
    ]

    html = ['<div class="table-wrap"><table class="kol-table">']
    html.append('<thead><tr><th class="sticky-col">场次/承担人</th>')
    for f in display_fields:
        cls = "highlight-col" if f in key_metrics else ""
        html.append(f'<th class="{cls}">{f}</th>')
    html.append('</tr></thead><tbody>')

    for venue in valid_venues:
        t = venue["total"]
        html.append(f'<tr class="summary-row"><td class="sticky-col"><strong>{t["name"]}</strong></td>')
        for f in display_fields:
            v = t.get(f)
            cls_td = "highlight-col" if f in key_metrics else ""
            html.append(f'<td class="{cls_td}">{fmt_by_field(v, f)}</td>')
        html.append('</tr>')

        for d in venue["details"]:
            html.append(f'<tr><td class="sticky-col" style="padding-left:24px!important;">{d["name"]}</td>')
            for f in display_fields:
                v = d.get(f)
                cls_td = "highlight-col" if f in key_metrics else ""
                html.append(f'<td class="{cls_td}">{fmt_by_field(v, f)}</td>')
            html.append('</tr>')

    html.append('</tbody></table></div>')
    return "\n".join(html)


def shangchao_repeat_venue_block(sc_data):
    """生成重复场次对比块"""
    repeats = sc_data.get("repeat_venues", {})
    key_metrics = sc_data["key_metrics"]

    if not repeats:
        return '<p style="font-size:13px;color:#666;">未发现重复场次</p>'

    display_fields = [
        "天数", "天均", "滚动消耗", "例子数", "约课数",
        "例子约课率", "约课到课率", "到课转化率", "滚动转化率",
        "滚动ROI2总成本", "滚动ROI2"
    ]

    html = []
    for base, group in sorted(repeats.items()):
        # 构建该组的对比表
        html.append(f'<div class="card" style="background:#f8fafc; box-shadow:none; padding:14px 18px; margin:12px 0;">')
        html.append(f'<h3 style="font-size:14px; margin-bottom:10px;">📍 {base}（{len(group)} 次场次）</h3>')

        html.append('<div class="table-scroll-tall"><table class="kol-table">')
        html.append('<thead><tr><th class="sticky-col">场次</th>')
        for f in display_fields:
            cls = "highlight-col" if f in key_metrics else ""
            html.append(f'<th class="{cls}">{f}</th>')
        html.append('</tr></thead><tbody>')

        for v in group:
            t = v["total"]
            html.append(f'<tr><td class="sticky-col"><strong>{t["name"]}</strong></td>')
            for f in display_fields:
                val = t.get(f)
                cls_td = "highlight-col" if f in key_metrics else ""
                html.append(f'<td class="{cls_td}">{fmt_by_field(val, f)}</td>')
            html.append('</tr>')

        html.append('</tbody></table></div>')

        # 对比分析（最优场次）
        analysis_lines = []
        for metric in key_metrics:
            vals = [(v["total"]["name"], v["total"].get(metric)) for v in group
                    if isinstance(v["total"].get(metric), (int, float))]
            if len(vals) >= 2:
                best = max(vals, key=lambda x: x[1])
                worst = min(vals, key=lambda x: x[1])
                if best[0] != worst[0]:
                    def disp(v):
                        if metric == "滚动ROI2":
                            return f"{v:.2f}"
                        if v <= 1.5:
                            return f"{v * 100:.1f}%"
                        return f"{v:.2f}"
                    analysis_lines.append(
                        f'· {metric}：<span class="positive">{best[0]} {disp(best[1])} 最优</span>，'
                        f'{worst[0]} {disp(worst[1])} 最差'
                    )

        if analysis_lines:
            html.append(f'<div style="margin-top:8px; font-size:13px; line-height:1.8;">{"<br>".join(analysis_lines)}</div>')
        html.append('</div>')

    return "\n".join(html)


def shangchao_new_venues_alert(sc_data):
    """生成新增场次提醒（提示用户补充天数）"""
    current_venues = sc_data.get("current_venues", [])
    all_venues = sc_data.get("all_venues", [])
    cur_month = sc_data.get("current_month")
    if not cur_month:
        return ""

    new_venues = process_shangchao.find_new_venues(
        current_venues, all_venues, cur_month["name"]
    )
    if not new_venues:
        return ""

    lines = [f'⚠️ 发现 {len(new_venues)} 个本月新增场次（需要补充天数信息以计算天均）：']
    for nv in new_venues:
        lines.append(
            f'· {nv["suggested_name"]}: 例子={int(nv["examples"])}, '
            f'滚动消耗={nv["cost"]:.0f} 元'
        )
    lines.append('💡 请在 process_shangchao.py 的 VENUE_DAYS_MAP 中配置天数')

    return f'''<div class="alert-card warn" style="margin-top:12px;">
<h3 style="font-size:14px; margin-bottom:8px;">新增场次提醒</h3>
<div style="font-size:13px; line-height:1.8;">{"<br>".join(lines)}</div>
</div>'''


def generate_html(flow_data, kol_data, tmk_data=None, book_fair_data=None, sc_data=None, referral_data=None, punch_data=None):
    today = datetime.now()
    yesterday = today - timedelta(days=1)
    date_str = yesterday.strftime("%Y年%m月%d日")

    # 流速
    kol_names = ["钟嘉欣图片", "钟嘉欣视频1", "钟嘉欣视频", "其他汇总"]
    dlz_names = ["出席礼品测试", "edm"]
    shangchao_names = ["代理商场", "书展", "澳門展会", "代理人汇总"]

    kol_rows = [r for r in flow_data if r["name"] in kol_names]
    dlz_rows = [r for r in flow_data if r["name"] in dlz_names]
    kol_summary = next((r for r in flow_data if r["name"] == "KOL-汇总"), None)
    shangchao_rows = [r for r in flow_data if r["name"] in shangchao_names]
    shangchao_summary = next((r for r in flow_data if r["name"] == "商超&社群合计"), None)
    total_row = next((r for r in flow_data if r["name"] == "全渠道汇总"), None)

    behind = [r for r in kol_rows + dlz_rows + shangchao_rows if r["gap"] < 0]
    ahead = [r for r in kol_rows + dlz_rows + shangchao_rows if r["gap"] > 0]

    # KOL 转化分析（综合判断仍保留）
    _, judgements = generate_kol_analysis(kol_data)

    # 分组数据
    kol_all_rows = kol_data["all_rows"]
    kol_headers = kol_data["headers"]

    zjx_rows = [r for r in kol_all_rows if "钟嘉欣" in r["name"] and "图片" not in r["name"] and "视频" not in r["name"]]
    kolhk_rows = [r for r in kol_all_rows if "KOL汇总" in r["name"] or "KOLHK汇总" in r["name"]]

    # 找各模块的本月/上月/汇总行（用于动态月份）
    current_month = kol_data.get("current_month", 6)
    last_month = kol_data.get("last_month", 5)
    zjx_may = next((r for r in kol_all_rows if f"{current_month}月钟嘉欣" in r["name"]), None)
    zjx_apr = next((r for r in kol_all_rows if f"{last_month}月钟嘉欣" in r["name"]), None)
    zjx_sum = next((r for r in kol_all_rows if r["name"] == "钟嘉欣-汇总"), None)
    kolhk_may = next((r for r in kol_all_rows if f"{current_month}月-KOLHK汇总" in r["name"]), None)
    kolhk_apr = next((r for r in kol_all_rows if f"{last_month}月-KOLHK汇总" in r["name"]), None)

    html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>港澳商务周报 - 截至{date_str}</title>
<style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ font-family: -apple-system, "Microsoft YaHei", sans-serif; background:#f5f7fa; color:#333; display:flex; }}
.sidebar {{ width:240px; background:#2c3e50; color:#ecf0f1; height:100vh; position:fixed; padding:20px 0; overflow-y:auto; }}
.sidebar h2 {{ font-size:16px; padding:0 20px 16px; border-bottom:1px solid #34495e; margin-bottom:16px; }}
.sidebar nav a {{ display:block; padding:8px 20px; color:#ecf0f1; text-decoration:none; transition:background 0.2s; font-size:13px; }}
.sidebar nav a:hover {{ background:#34495e; }}
.sidebar nav a.subitem {{ padding-left:36px; font-size:12px; color:#bdc3c7; }}
.main {{ margin-left:240px; padding:24px 32px; flex:1; max-width:calc(100vw - 240px); }}
h1 {{ font-size:22px; margin-bottom:6px; }}
.subtitle {{ color:#666; font-size:14px; margin-bottom:24px; }}
.card {{ background:#fff; border-radius:10px; box-shadow:0 2px 8px rgba(0,0,0,0.06); padding:20px 24px; margin-bottom:20px; }}
.card h2 {{ font-size:16px; margin-bottom:12px; color:#1a1a1a; border-left:4px solid #4f8cff; padding-left:10px; }}
.card h3 {{ font-size:14px; margin: 14px 0 10px; color:#1a1a1a; }}
table {{ width:100%; border-collapse:collapse; font-size:12px; }}
.kol-table {{ font-size:11px; min-width:max-content; }}
.kol-table th, .kol-table td {{ padding:5px 4px; }}
.table-wrap {{ overflow-x:auto; max-width:100%; border:1px solid #e8edf3; border-radius:6px; }}
.table-scroll-tall {{ max-height:70vh; overflow:auto; border:1px solid #e8edf3; border-radius:6px; }}
.table-scroll-tall thead th {{ position:sticky; top:0; z-index:5; box-shadow:0 1px 0 0 #dde4f0; }}
.table-scroll-tall thead th.sticky-col {{ z-index:6; }}
th {{ background:#f0f4ff; padding:8px 6px; text-align:center; font-weight:600; border-bottom:2px solid #dde4f0; white-space:nowrap; }}
td {{ padding:7px 6px; text-align:center; border-bottom:1px solid #eee; white-space:nowrap; }}
tr:hover {{ background:#fafbff; }}
.summary-row {{ background:#f8fafc; font-weight:700; }}
.highlight-col {{ background:#fff8e1 !important; }}
.summary-row.highlight-col, .summary-row td.highlight-col {{ background:#ffe8b3 !important; }}
/* 首列冻结 */
.sticky-col {{
  position:sticky;
  left:0;
  z-index:2;
  background:#fff;
  border-right:2px solid #c5cfe0;
  min-width:140px;
  text-align:left !important;
  padding-left:10px !important;
  box-shadow:2px 0 4px rgba(0,0,0,0.04);
}}
thead th.sticky-col {{ z-index:3; background:#f0f4ff; }}
.summary-row .sticky-col {{ background:#f8fafc; }}
/* TMK 表格的第二个冻结列（TMK 个人名） */
.sticky-col-2 {{
  position:sticky;
  left:140px;
  z-index:2;
  background:#fff;
  border-right:2px solid #c5cfe0;
  min-width:100px;
  text-align:left !important;
  padding-left:10px !important;
}}
thead th.sticky-col-2 {{ z-index:3; background:#f0f4ff; }}
.tmk-grand-total td {{ background:#fff4cc !important; font-weight:700; }}
.tmk-grand-total .sticky-col, .tmk-grand-total .sticky-col-2 {{ background:#fff4cc !important; }}
.tmk-team-total td {{ background:#e0e8f0 !important; font-weight:700; }}
.tmk-team-total .sticky-col, .tmk-team-total .sticky-col-2 {{ background:#e0e8f0 !important; }}
.negative {{ color:#e53935; font-weight:700; }}
.positive {{ color:#2e7d32; font-weight:700; }}
.alert-section {{ display:flex; gap:16px; flex-wrap:wrap; }}
.alert-card {{ flex:1; min-width:280px; padding:16px; border-radius:8px; }}
.alert-card.warn {{ background:#fff3e0; border:1px solid #ffe0b2; }}
.alert-card.good {{ background:#e8f5e9; border:1px solid #c8e6c9; }}
.alert-card.todo {{ background:#fff7ed; border:1px solid #fed7aa; }}
.alert-card h3 {{ font-size:14px; margin-bottom:8px; }}
.alert-card ul {{ list-style:none; padding:0; font-size:13px; }}
.alert-card li {{ padding:3px 0; }}
.tag {{ display:inline-block; padding:2px 8px; border-radius:4px; font-size:12px; margin-left:6px; }}
.tag-red {{ background:#ffcdd2; color:#c62828; }}
.tag-green {{ background:#c8e6c9; color:#2e7d32; }}
.analysis-block {{ background:#f8fafc; border-left:3px solid #4f8cff; padding:14px 18px; margin:12px 0; border-radius:4px; line-height:1.8; font-size:13px; }}
.row-anomaly:not(.summary-row) {{ background:#fef2f2; }}
.judgement {{ background:#e3f2fd; padding:12px 16px; border-radius:6px; margin:8px 0; font-size:13px; }}
section {{ scroll-margin-top:20px; }}
</style>
</head>
<body>
<div class="sidebar">
<h2>📊 周报目录</h2>
<nav>
<a href="#flow">1. 港澳商务流速</a>
<a href="#flow-key" class="subitem">— 重点关注</a>
<a href="#flow-kol" class="subitem">— KOL 模块</a>
<a href="#flow-dlz" class="subitem">— 独立站HK</a>
<a href="#flow-sc" class="subitem">— 商超&社群</a>
<a href="#flow-total" class="subitem">— 全渠道汇总</a>
<a href="#kol-conversion">2. 本月KOL转化数据</a>
<a href="#kol-hk" class="subitem">— KOLHK 整体趋势</a>
<a href="#kol-zjx" class="subitem">— 钟嘉欣（1月至今）</a>
<a href="#kol-detail" class="subitem">— 钟嘉欣图片&视频</a>
<a href="#kol-judge" class="subitem">— 综合判断</a>
<a href="#tmk">3. TMK 做工周报</a>
<a href="#tmk-overview" class="subitem">— 整体数据概览</a>
<a href="#tmk-alerts" class="subitem">— 个人异常情况</a>
<a href="#tmk-table" class="subitem">— 个人做工数据</a>
<a href="#book-fair">4. 书展数据复盘</a>
<a href="#book-fair-last-month" class="subitem">— 上月书展数据</a>
<a href="#book-fair-current-month" class="subitem">— 本月书展数据</a>
<a href="#book-fair-history" class="subitem">— 历史书展对比</a>
<a href="#shangchao">5. 线下商超复盘</a>
<a href="#sc-analysis" class="subitem">— 当月分析</a>
<a href="#sc-monthly" class="subitem">— 分月趋势</a>
<a href="#sc-venue" class="subitem">— 当月分场次</a>
<a href="#sc-repeat" class="subitem">— 重复场次对比</a>
<a href="#referral">6. 转介绍打卡</a>
<a href="#ref-achieve" class="subitem">— 后端非手推达成</a>
<a href="#ref-punch" class="subitem">— 打卡链路数据</a>
</nav>
</div>
<div class="main">
<h1>港澳商务周报</h1>
<p class="subtitle">数据截至 {date_str}（MTD累计）</p>

<section id="flow">
<section id="flow-key" class="card">
<h2>1.1 重点关注</h2>
{module_overview("flow", {"total_row": total_row, "behind": behind, "ahead": ahead})}
<div class="alert-section">
<div class="alert-card warn">
<h3>未达进度（例子）</h3>
<ul>"""

    for r in sorted(behind, key=lambda x: x["gap"]):
        rate = calc_rate(r["actual"], r["mtd_target"])
        html += f'\n<li>{r["name"]} <span class="tag tag-red">gap {int(r["gap"])}</span> 达成率 {rate}%</li>'

    html += """
</ul>
</div>
<div class="alert-card good">
<h3>超预期达成</h3>
<ul>"""

    for r in sorted(ahead, key=lambda x: -x["gap"]):
        rate = calc_rate(r["actual"], r["mtd_target"])
        html += f'\n<li>{r["name"]} <span class="tag tag-green">+{int(r["gap"])}</span> 达成率 {rate}%</li>'

    html += """
</ul>
</div>
<div class="alert-card todo">
<h3>📋 本周待办</h3>
<ul>"""

    if behind and any(r["gap"] < -5 for r in behind):
        for r in sorted([x for x in behind if x["gap"] < -5], key=lambda x: x["gap"]):
            rate = calc_rate(r["actual"], r["mtd_target"])
            action = get_action_for_supplier(r["name"])
            html += f'\n<li>{r["name"]} 缺口 {abs(int(r["gap"]))} 例子（达成率 {rate}%）→ {action}</li>'
    else:
        html += '\n<li>本周节奏正常，关注下周转化漏斗即可</li>'

    html += f"""
</ul>
</div>
</div>
</section>

<section id="flow-kol" class="card">
<h2>1.2 KOL 模块</h2>
<table>
<tr><th>供应商</th><th>MTD目标</th><th>实际例子</th><th>例子gap</th><th>例子达成率</th><th>约课目标</th><th>实际约课</th><th>约课gap</th><th>约课达成率</th><th>约课成本</th></tr>
{''.join(row_html_flow(r) for r in kol_rows)}
{row_html_flow(kol_summary, True) if kol_summary else ''}
</table>
</section>

<section id="flow-dlz" class="card">
<h2>1.3 独立站HK 模块</h2>
<table>
<tr><th>供应商</th><th>MTD目标</th><th>实际例子</th><th>例子gap</th><th>例子达成率</th><th>约课目标</th><th>实际约课</th><th>约课gap</th><th>约课达成率</th><th>约课成本</th></tr>
{''.join(row_html_flow(r) for r in dlz_rows)}
</table>
</section>

<section id="flow-sc" class="card">
<h2>1.4 商超&社群 模块</h2>
<table>
<tr><th>供应商</th><th>MTD目标</th><th>实际例子</th><th>例子gap</th><th>例子达成率</th><th>约课目标</th><th>实际约课</th><th>约课gap</th><th>约课达成率</th><th>约课成本</th></tr>
{''.join(row_html_flow(r) for r in shangchao_rows)}
{row_html_flow(shangchao_summary, True) if shangchao_summary else ''}
</table>
</section>

<section id="flow-total" class="card">
<h2>1.5 全渠道汇总</h2>
<table>
<tr><th>模块</th><th>MTD目标</th><th>实际例子</th><th>例子gap</th><th>例子达成率</th><th>约课目标</th><th>实际约课</th><th>约课gap</th><th>约课达成率</th><th>约课成本</th></tr>
{row_html_flow(kol_summary) if kol_summary else ''}
{row_html_flow({"name": "独立站HK", "mtd_target": sum(r["mtd_target"] for r in dlz_rows), "actual": sum(r["actual"] for r in dlz_rows), "gap": sum(r["actual"] for r in dlz_rows) - sum(r["mtd_target"] for r in dlz_rows), "lesson_target": sum(r["lesson_target"] for r in dlz_rows), "lesson_actual": sum(r["lesson_actual"] for r in dlz_rows), "lesson_gap": sum(r["lesson_actual"] for r in dlz_rows) - sum(r["lesson_target"] for r in dlz_rows), "cost": None}) if dlz_rows else ''}
{row_html_flow(shangchao_summary) if shangchao_summary else ''}
{row_html_flow(total_row, True) if total_row else ''}
</table>
</section>
</section>

<section id="kol-conversion">
<section id="kol-hk" class="card">
<h2>2.1 KOLHK 整体月度趋势</h2>
{module_overview("kol", {"zjx_may": zjx_may})}
<p style="font-size:13px; color:#666; margin-bottom:10px;">📌 高亮列为重点指标：例子约课率、约课到课率、到课转化率、滚动转化率、滚动ROI2</p>
{metric_block(f"KOLHK 整体（{kol_data['month_label']}）", kolhk_may, kolhk_apr)}
{kol_table_html(kolhk_rows, kol_headers)}
</section>

<section id="kol-zjx" class="card">
<h2>2.2 钟嘉欣转化数据（1月至今）</h2>
{metric_block(f"钟嘉欣（{kol_data['month_label']}）", zjx_may, zjx_apr, zjx_sum, sum_label=f"1-{current_month}月汇总")}
{kol_table_html(zjx_rows, kol_headers)}
</section>

<section id="kol-detail" class="card">
<h2>2.3 钟嘉欣图片&视频明细数据</h2>
{detail_compare_block(kol_data["detail_rows"])}
{kol_table_html(kol_data["detail_rows"], kol_data["detail_headers"])}

<h3>钟嘉欣视频渠道明细</h3>
<div class="table-wrap">
<table class="kol-table">
<thead>
<tr>
<th>渠道简称</th>
<th>例子数</th>
<th>约课数</th>
<th>滚动消耗</th>
<th>例子约课率</th>
<th>约课到课率</th>
<th>到课转化率</th>
<th>滚动转化率</th>
<th>滚动GMV</th>
<th>滚动ROI2</th>
</tr>
</thead>
<tbody>
{''.join([f'''
<tr>
<td style="text-align:left;">{ch["渠道简称"]}</td>
<td>{ch["例子数"] if ch["例子数"] else "-"}</td>
<td>{ch["约课数"] if ch["约课数"] else "-"}</td>
<td>{round(ch["滚动消耗"], 2) if isinstance(ch["滚动消耗"], (int, float)) else "-"}</td>
<td>{f"{ch['例子约课率']*100:.1f}%" if isinstance(ch["例子约课率"], (int, float)) else "-"}</td>
<td>{f"{ch['约课到课率']*100:.1f}%" if isinstance(ch["约课到课率"], (int, float)) else "-"}</td>
<td>{f"{ch['到课转化率']*100:.1f}%" if isinstance(ch["到课转化率"], (int, float)) else "-"}</td>
<td>{f"{ch['滚动转化率']*100:.1f}%" if isinstance(ch["滚动转化率"], (int, float)) else "-"}</td>
<td>{f"{ch['滚动GMV']:,.0f}" if isinstance(ch["滚动GMV"], (int, float)) else "-"}</td>
<td>{f"{ch['滚动ROI2']:.2f}" if isinstance(ch["滚动ROI2"], (int, float)) else "-"}</td>
</tr>
''' for ch in kol_data["video_channel_rows"]])}
</tbody>
</table>
</div>
</section>

<section id="kol-judge" class="card">
<h2>2.4 综合判断</h2>
"""

    if judgements:
        for j in judgements:
            html += f'<div class="judgement">{j}</div>\n'
    else:
        html += '<p style="font-size:13px;color:#666;">暂无综合判断</p>\n'

    html += """
</section>
</section>
"""

    # ============ TMK 周报模块 ============
    if tmk_data:
        params = tmk_data.get("params", {})
        excel_name = f"个人做工数据_{params.get('start_date', '?')}_{params.get('end_date', '?')}.xlsx"
        excel_link = f"TMK周报/output/{excel_name}"

        html += f"""
<section id="tmk">
<section id="tmk-overview" class="card">
<h2>3.1 TMK 做工 - 整体数据概览</h2>
{module_overview("tmk", tmk_data)}
{tmk_overview_block(tmk_data)}
</section>

<section id="tmk-alerts" class="card">
<h2>3.2 TMK 做工 - 个人异常情况</h2>
{tmk_alert_block(tmk_data)}
</section>

<section id="tmk-table" class="card">
<h2>3.3 TMK 做工 - 个人做工数据</h2>
<p style="font-size:13px; color:#666; margin-bottom:10px;">
📊 完整数据表（带分组表头 + 冻结）：<a href="{excel_link}" target="_blank" style="color:#4f8cff;">下载 Excel</a>
</p>
{tmk_table_html(tmk_data)}
</section>
</section>
"""

    # ============ 书展数据复盘模块 ============
    if book_fair_data:
        # 构建上月书展表格行
        last_month_rows = []
        if book_fair_data["last_month_fairs"]:
            for fair in book_fair_data["last_month_fairs"]["fairs"]:
                r = dict(fair)
                # 识别汇总行（名称为供应商名的行）
                if fair.get("name") in book_fair_data["last_month_fairs"]["fairs"]:
                    pass  # 不特别标记，由表格逻辑判断
                last_month_rows.append(r)

        # 构建本月书展表格行（如有）
        current_month_rows = []
        has_current_month = book_fair_data["current_month_fairs"] is not None
        if has_current_month:
            for fair in book_fair_data["current_month_fairs"]["fairs"]:
                r = dict(fair)
                current_month_rows.append(r)

        # 历史书展行
        history_rows = []
        for h in book_fair_data["history"]:
            r = dict(h)
            r["_is_total"] = True  # 历史每条都是汇总
            history_rows.append(r)

        # 获取日期信息
        last_month_label = book_fair_data["last_month_fairs"]["period_label"] if book_fair_data["last_month_fairs"] else ""
        last_month_range = book_fair_data["last_month_fairs"]["period_range"] if book_fair_data["last_month_fairs"] else ""

        # 动态章节编号
        history_section = "4.3" if has_current_month else "4.2"


        html += f"""
<section id="book-fair">
<section id="book-fair-last-month" class="card">
<h2>4.1 上月书展数据（{last_month_label}）</h2>
{module_overview("bookfair", book_fair_data)}
<p style="font-size:13px; color:#666; margin-bottom:10px;">统计周期：{last_month_range} | 📌 高亮列为重点指标</p>
"""
        if last_month_rows:
            html += f"{book_fair_table_html(last_month_rows, book_fair_data['summary_fields'], book_fair_data['key_metrics'])}\n"
        else:
            html += "<p style=\"color:#999;\">上月无书展数据</p>\n"
        html += "</section>\n"

        # 本月书展子模块（条件显示）
        if has_current_month:
            current_month_label = book_fair_data["current_month_fairs"]["period_label"]
            current_month_range = book_fair_data["current_month_fairs"]["period_range"]
            html += f"""
<section id="book-fair-current-month" class="card">
<h2>4.2 本月书展数据（{current_month_label}）</h2>
<p style="font-size:13px; color:#666; margin-bottom:10px;">统计周期：{current_month_range}</p>
{book_fair_table_html(current_month_rows, book_fair_data["summary_fields"], book_fair_data["key_metrics"])}
</section>
"""

        # 历史书展
        html += f"""
<section id="book-fair-history" class="card">
<h2>{history_section} 历史书展数据汇总</h2>
<p style="font-size:13px; color:#666; margin-bottom:10px;">{len(history_rows)} 场历史书展，可作为参照</p>
{book_fair_table_html(history_rows, book_fair_data["summary_fields"], book_fair_data["key_metrics"])}
</section>
</section>
"""


    # ============ 线下商超复盘模块 ============
    if sc_data:
        cur_name = sc_data["current_month"]["name"] if sc_data["current_month"] else "本月"
        periods = sc_data.get("periods", {})
        last_prefix = periods.get("last_month", {}).get("prefix", "5月")
        cur_prefix = periods.get("current_month", {}).get("prefix", "6月")
        yesterday = periods.get("yesterday")

        # 日期范围（动态计算）
        if not yesterday:
            yesterday = datetime.now().date() - timedelta(days=1)
        last_range = f"{5}.1-{yesterday.month}.{yesterday.day}"
        cur_range = f"{yesterday.month}.1-{yesterday.month}.{yesterday.day}"

        html += f"""
<section id="shangchao">
<section id="sc-analysis" class="card">
<h2>5.1 线下商超 - 当月分析</h2>
{module_overview("shangchao", sc_data)}
<p style="font-size:13px; color:#666; margin-bottom:10px;">📌 高亮列为重点指标：例子约课率、约课到课率、到课转化率、滚动转化率、滚动ROI2</p>
{shangchao_analysis_block(sc_data)}
</section>

<section id="sc-monthly" class="card">
<h2>5.2 线下商超 - 分月趋势</h2>
<p style="font-size:13px; color:#666; margin-bottom:10px;">9月至今的整体月度趋势</p>
{shangchao_monthly_table(sc_data)}
</section>

<section id="sc-last-month" class="card">
<h2>5.3 线下商超 - {last_prefix}场次（{last_range}，含{cur_prefix}延续）</h2>
<p style="font-size:13px; color:#666; margin-bottom:10px;">{last_prefix}场次完整链路滚动展示，滚动消耗已累计{cur_prefix}延续部分</p>
{shangchao_last_month_venue_table(sc_data)}
</section>

<section id="sc-venue" class="card">
<h2>5.4 线下商超 - {cur_name}分场次明细</h2>
<p style="font-size:13px; color:#666; margin-bottom:10px;">本月各商场场次 + 承担人明细（实时从数据底表读取）</p>
{shangchao_new_venues_alert(sc_data)}
{shangchao_current_month_venue_table(sc_data)}
</section>

<section id="sc-repeat" class="card">
<h2>5.5 线下商超 - 重复场次对比</h2>
<p style="font-size:13px; color:#666; margin-bottom:10px;">同一商场多次出摊的对比分析（自动识别基础名称相同的场次）</p>
{shangchao_repeat_venue_block(sc_data)}
</section>
</section>
"""

    # ============ 转介绍打卡模块 ============
    if referral_data and punch_data:
        def ref_row_html(rows, period_label):
            """生成后端非手推达成表格行"""
            h = ""
            for r in rows:
                is_total = r["name"] == "合计"
                cls = "summary-row" if is_total else ""
                h += f'<tr class="{cls}">'
                h += f'<td class="sticky-col"><strong>{r["name"]}</strong></td>'
                h += f'<td>{r["examples"]}</td>'
                h += f'<td>{r["target"]}</td>'
                h += f'<td>{r["mtd_rate"]*100:.1f}%</td>'
                h += f'<td>{r["example_pct"]*100:.1f}%</td>'
                h += f'<td>{r["bookings"]}</td>'
                h += f'<td>{r["attendances"]}</td>'
                h += f'<td>{r["signups"]}</td>'
                h += f'<td>{r["gmv"]:,.0f}</td>'
                h += f'<td>{r["book_rate"]*100:.1f}%</td>'
                h += f'<td>{r["attend_rate"]*100:.1f}%</td>'
                h += f'<td>{r["conv_rate"]*100:.1f}%</td>'
                h += f'<td>{r["total_conv_rate"]*100:.2f}%</td>'
                h += f'<td>{r["asp"]:,.0f}</td>'
                h += f'<td>{r["gmv_pct"]*100:.1f}%</td>'
                h += '</tr>'
            return h

        ref_headers = '<tr><th class="sticky-col">后端</th><th>例子数</th><th>目标例子数</th><th>MTD达成率</th><th>例子占比</th><th>约课人数</th><th>到课人数</th><th>正式课成单数</th><th>正式课GMV</th><th>约课率</th><th>约课到课率</th><th>到课转化率</th><th>转化率</th><th>ASP</th><th>正式课GMV占比</th></tr>'

        def punch_row_html(d, label):
            """生成打卡链路单行"""
            return f"""<tr>
<td class="sticky-col"><strong>{label}</strong></td>
<td>{d['can_punch']:,}</td>
<td>{d['punched']:,}</td>
<td>{d['punch_count']:,}</td>
<td>{d['punch_rate']*100:.2f}%</td>
<td>{d['avg_punch']:.2f}</td>
<td>{d['examples']}</td>
<td>{d['split_rate']*100:.4f}%</td>
<td>{d['bookings']}</td>
<td>{d['book_rate']*100:.2f}%</td>
<td>{d['attendances']}</td>
<td>{d['attend_rate']*100:.2f}%</td>
<td>{d['signups']}</td>
<td>{d['conv_rate']*100:.2f}%</td>
<td>{d['total_conv_rate']*100:.2f}%</td>
<td>{d['gmv']:,.0f}</td>
<td>{d['asp']:,.0f}</td>
</tr>"""

        # 环比行
        def calc_env(cur_val, prev_val):
            if prev_val and prev_val != 0:
                return (cur_val - prev_val) / abs(prev_val)
            return None

        punch_cur = punch_data["cur"]
        punch_prev = punch_data["prev"]
        env_lines = []
        metrics_env = [
            ("可打卡学员", punch_cur["can_punch"], punch_prev["can_punch"]),
            ("打卡人数", punch_cur["punched"], punch_prev["punched"]),
            ("打卡率", punch_cur["punch_rate"], punch_prev["punch_rate"]),
            ("例子数", punch_cur["examples"], punch_prev["examples"]),
            ("打卡裂变率", punch_cur["split_rate"], punch_prev["split_rate"]),
            ("约课率", punch_cur["book_rate"], punch_prev["book_rate"]),
            ("约课到课率", punch_cur["attend_rate"], punch_prev["attend_rate"]),
        ]
        for name, cv, pv in metrics_env:
            env = calc_env(cv, pv)
            if env is not None:
                trend = "↑" if env > 0 else "↓"
                color = "positive" if env > 0 else "negative"
                env_lines.append(f'· {name}：<span class="{color}">{trend} {abs(env)*100:.1f}%</span>')

        env_html = f'<div class="analysis-block"><strong>环比上期（{punch_data["cur_period"]} vs {punch_data["prev_period"]}）</strong><br>{"<br>".join(env_lines)}</div>' if env_lines else ""

        html += f"""
<section id="referral">
<section id="ref-achieve" class="card">
<h2>6.1 转介绍 - 后端非手推达成情况</h2>
{module_overview("referral", {"cur": punch_data.get("cur"), "prev": punch_data.get("prev")})}
<h3>{referral_data['cur_period']}（本期）</h3>
<div class="table-wrap"><table class="kol-table">
{ref_headers}
{ref_row_html(referral_data['cur_rows'], referral_data['cur_period'])}
</table></div>
<h3>{referral_data['prev_period']}（上期）</h3>
<div class="table-wrap"><table class="kol-table">
{ref_headers}
{ref_row_html(referral_data['prev_rows'], referral_data['prev_period'])}
</table></div>
</section>

<section id="ref-punch" class="card">
<h2>6.2 转介绍 - 打卡链路数据</h2>
{env_html}
<div class="table-wrap"><table class="kol-table">
<thead>
<tr><th class="sticky-col">日期</th><th>可打卡学员</th><th>打卡人数</th><th>打卡次数</th><th>打卡率</th><th>人均打卡次数</th><th>例子数</th><th>打卡裂变率</th><th>约课数</th><th>约课率</th><th>到课数</th><th>约课到课率</th><th>转化例子数</th><th>到课转化率</th><th>注册转化率</th><th>GMV</th><th>ASP</th></tr>
</thead>
<tbody>
{punch_row_html(punch_data['cur'], punch_data['cur_period'])}
{punch_row_html(punch_data['prev'], punch_data['prev_period'])}
</tbody>
</table></div>
</section>
</section>
"""

    html += """
</div>
</body>
</html>"""

    return html


def write_run_log(metrics: dict):
    """把本次运行的关键指标追加写入 runs.log（CSV 格式）"""
    import csv
    log_path = BASE_DIR / "runs.log"
    is_new = not log_path.exists()

    headers = [
        "运行时间", "本期",
        # 流速
        "流速_总例子目标", "流速_总实际例子", "流速_总达成率",
        # KOL
        "KOL_钟嘉欣本月例子", "KOL_KOLHK本月例子",
        # TMK
        "TMK_行数", "TMK_异常数",
        # 书展
        "书展_本月场次", "书展_历史场次",
        # 商超
        "商超_本月有效场次", "商超_本月例子数", "商超_本月CPS消耗",
        # 转介绍
        "转介绍_手推例子", "转介绍_非手推例子", "转介绍_合计达成率",
        "打卡_可打卡学员", "打卡_打卡人数", "打卡_例子数",
    ]

    with open(log_path, "a", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        if is_new:
            writer.writerow(headers)
        writer.writerow([metrics.get(h, "") for h in headers])

    print(f"\n[日志] 已写入: {log_path.name}")


def collect_run_metrics(flow_data, kol_data, tmk_data, book_fair_data, sc_data,
                         referral_data, punch_data):
    """从各模块数据中收集关键指标用于日志记录"""
    from datetime import datetime
    metrics = {
        "运行时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }

    # 流速
    if flow_data:
        total_row = next((r for r in flow_data if r.get("name") == "合计"), None)
        if total_row:
            metrics["本期"] = f"{datetime.now().strftime('%Y-%m')}"
            metrics["流速_总例子目标"] = total_row.get("mtd_target")
            metrics["流速_总实际例子"] = total_row.get("actual")
            target = total_row.get("mtd_target") or 0
            if target:
                metrics["流速_总达成率"] = f"{total_row.get('actual', 0) / target * 100:.1f}%"

    # KOL
    if kol_data:
        for r in kol_data.get("all_rows", []):
            if "5月钟嘉欣" in r.get("name", "") or "6月钟嘉欣" in r.get("name", ""):
                metrics["KOL_钟嘉欣本月例子"] = r.get("例子数")
            if "KOLHK汇总" in r.get("name", "") and ("5月" in r.get("name", "") or "6月" in r.get("name", "")):
                metrics["KOL_KOLHK本月例子"] = r.get("例子数")

    # TMK
    if tmk_data:
        metrics["TMK_行数"] = len(tmk_data.get("rows", []))
        metrics["TMK_异常数"] = len(tmk_data.get("alerts", []))

    # 书展
    if book_fair_data:
        # 获取本月书展场次
        current_month_count = 0
        if book_fair_data.get("current_month_fairs"):
            current_month_count = len(book_fair_data["current_month_fairs"].get("fairs", []))
        metrics["书展_本月场次"] = current_month_count
        metrics["书展_历史场次"] = len(book_fair_data.get("history", []))

    # 商超
    if sc_data:
        valid_count = sum(
            1 for v in sc_data.get("current_venues", [])
            if (v["total"].get("例子数") or 0) > 0
        )
        metrics["商超_本月有效场次"] = valid_count
        cur_month = sc_data.get("current_month")
        if cur_month:
            metrics["商超_本月例子数"] = cur_month.get("例子数")
            cps = cur_month.get("CPS消耗")
            if cps:
                metrics["商超_本月CPS消耗"] = f"{cps:,.0f}"

    # 转介绍
    if referral_data:
        for r in referral_data.get("cur_rows", []):
            if r["name"] == "手推链接":
                metrics["转介绍_手推例子"] = r["examples"]
            elif r["name"] == "非手推链接":
                metrics["转介绍_非手推例子"] = r["examples"]
            elif r["name"] == "合计":
                metrics["转介绍_合计达成率"] = f"{r['mtd_rate']*100:.1f}%"

    if punch_data:
        cur = punch_data.get("cur", {})
        metrics["打卡_可打卡学员"] = cur.get("can_punch")
        metrics["打卡_打卡人数"] = cur.get("punched")
        metrics["打卡_例子数"] = cur.get("examples")

    return metrics


def main():
    print("[0] 校验数据新鲜度（强制更新到昨天）...")

    # 导入所有下载模块（任一失败立即抛异常，全局硬中断）
    from utils.smartbi_helper import get_periods, assert_smartbi_credentials
    from 港澳流速.download_hk_flow import run_download as dl_hk_flow
    from 书展数据复盘.download_bookfair import download_bookfair_data, verify_bookfair_freshness, get_bookfair_periods
    from 本月KOL转化数据汇总.download_kol import run_download as dl_kol
    from TMK周报.download_tmk import run_download as dl_tmk
    from 线下商超内容.download_shangchao import run_download as dl_shangchao
    from 转介绍打卡内容.download_referral import run_download as dl_referral

    # 确保凭证
    assert_smartbi_credentials()

    # 获取日期窗口
    periods = get_periods()
    print(f"  目标快照日期: {periods['yesterday']}")

    # 按顺序下载各模块数据（任一失败立即中止）
    print("  [0.1] 下载共享底表...")
    dl_hk_flow(periods)

    print("  [0.2] 下载书展数据...")
    bf_paths = download_bookfair_data(periods)
    verify_bookfair_freshness(bf_paths, periods)

    print("  [0.3] 校验 KOL...")
    dl_kol(periods)

    print("  [0.4] 下载 TMK...")
    dl_tmk(periods)

    print("  [0.5] 校验商超...")
    dl_shangchao(periods)

    print("  [0.6] 下载转介绍...")
    dl_referral(periods)

    print(f"  ✓ 所有模块数据已就绪（快照={periods['yesterday']}）")

    # 清除各模块的中间缓存，强制重新读取最新底表
    print("  [清除缓存] 清除中间缓存文件，确保使用最新数据...")
    import glob
    from pathlib import Path

    cache_patterns = [
        "港澳流速/output/*.xlsx",
        "线下商超内容/output/*.xlsx",
        "书展内容/output/*.xlsx",
        "本月KOL转化数据汇总/output/*.xlsx",
        "TMK周报/output/*.xlsx",
        "转介绍打卡内容/output/*.xlsx",
    ]

    for pattern in cache_patterns:
        for cache_file in glob.glob(pattern):
            try:
                Path(cache_file).unlink()
                print(f"    [删除] {Path(cache_file).name}")
            except:
                pass

    print()

    print("[1] 提取流速数据...")
    flow_data = extract_flow_data()

    print("[2] 提取KOL转化数据...")
    kol_data = extract_kol_data()
    print(f"  汇总数据行数: {len(kol_data['all_rows'])}")
    print(f"  明细数据行数: {len(kol_data['detail_rows'])}")

    print("[3] 提取TMK做工数据...")
    tmk_data = process_tmk.extract_tmk_data()
    if tmk_data:
        print(f"  TMK 行数: {len(tmk_data['rows'])}")
        print(f"  异常项: {len(tmk_data['alerts'])}")
    else:
        print("  ⚠️ 未找到 TMK 数据，跳过该模块")

    print("[4] 提取书展数据复盘...")
    book_fair_data = process_book_fair.extract_book_fair_data()
    if book_fair_data:
        # 获取本月和上月书展场次
        last_month_count = len(book_fair_data.get("last_month_fairs", {}).get("fairs", [])) if book_fair_data.get("last_month_fairs") else 0
        current_month_count = len(book_fair_data.get("current_month_fairs", {}).get("fairs", [])) if book_fair_data.get("current_month_fairs") else 0
        print(f"  上月书展: {last_month_count} 场")
        print(f"  本月书展: {current_month_count} 场")
        print(f"  历史书展: {len(book_fair_data['history'])} 场")
    else:
        print("  ⚠️ 未找到 书展 数据，跳过该模块")

    print("[5] 提取线下商超数据...")
    sc_data = process_shangchao.extract_shangchao_data()
    if sc_data:
        print(f"  分月数据: {len(sc_data['monthly'])} 行")
        print(f"  当月场次: {len(sc_data['current_venues'])} 个")
        print(f"  历史场次: {len(sc_data['all_venues'])} 个")
        print(f"  重复场次组: {len(sc_data['repeat_venues'])} 组")
    else:
        print("  ⚠️ 未找到 线下商超 数据，跳过该模块")

    print("[6] 提取转介绍打卡数据...")
    referral_data = None
    punch_data = None
    try:
        referral_data = process_referral.extract_referral_data()
        punch_data = process_referral.extract_punch_card_data()
        if referral_data:
            print(f"  后端非手推: 本期{referral_data['cur_period']}, 上期{referral_data['prev_period']}")
        if punch_data:
            print(f"  打卡链路: 本期打卡人数{punch_data['cur']['punched']}, 例子{punch_data['cur']['examples']}")
    except Exception as e:
        print(f"  ⚠️ 转介绍数据提取失败: {e}")

    print("[7] 生成HTML...")
    html = generate_html(flow_data, kol_data, tmk_data, book_fair_data, sc_data, referral_data, punch_data)

    OUTPUT_DIR.mkdir(exist_ok=True)
    output_html = OUTPUT_DIR / "港澳商务周报.html"
    output_html.write_text(html, encoding="utf-8")
    print(f"\n[完成] 输出: {output_html}")

    # 产物时间戳验证（确保本次运行刷新了所有数据）
    print("\n[产物时间戳] 数据新鲜度检查（应为本次运行时间）")
    from datetime import datetime
    freshness_files = [
        ("流速", FLOW_OUTPUT),
        ("KOL", KOL_CONV_OUTPUT),
        ("商超", BASE_DIR / "线下商超内容/output/线下商超内容汇总.xlsx"),
        ("HTML", output_html),
    ]
    for label, path in freshness_files:
        p = Path(path)
        if p.exists():
            mt = datetime.fromtimestamp(p.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
            print(f"  {label:6s}  {mt}  {p.name}")
        else:
            print(f"  {label:6s}  ⚠️ 文件不存在")

    # 记录本次运行的关键指标
    metrics = collect_run_metrics(
        flow_data, kol_data, tmk_data, book_fair_data, sc_data,
        referral_data, punch_data
    )
    write_run_log(metrics)


if __name__ == "__main__":
    main()
