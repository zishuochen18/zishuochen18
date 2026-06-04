"""
港澳商务流速处理脚本
1. 读取流速表【香港市场目标】sheet
2. 读取数据底表【海外港澳商务_各渠道主辅投数据】
3. 处理 MTD例子目标（从当月1日到昨天的日列求和）
4. 从数据底表匹配实际例子达成、实际约课达成、约课成本
5. 负数 gap 标红
"""
"""
港澳商务流速处理脚本（模块 1）

═══════════════════════════════════════════════════════════════
功能：将数据底表按渠道汇总后，写入流速表的"实际例子达成"等列
═══════════════════════════════════════════════════════════════

【输入】
- sample/海外港澳商务_各渠道主辅投数据.xlsx（数据底表，BI 导出）
- sample/香港大v Amy20XX年X月渠道结算单.xlsx（流速模板，手工维护，本月日维度数据已填）

【输出】
- output/同名结算单.xlsx（带 KOL/商超&社群 实际达成数据）

【数据流】
数据底表（线下HK/KOLHK/社群HK）→ 按供应商汇总 → 写入流速表对应行

【关键列索引】（数据底表 海外港澳商务_各渠道主辅投数据.xlsx）
- 第6行（header=5）：列标题行
- 关键字段：渠道组 / 供应商 / 渠道名称 / 例子数 / 约课数 / 滚动消耗

【关键参数】
- BOOK_FAIR_SUPPLIERS：本月书展供应商列表（每月需更新）
  例：["26年5月第十二屆兒童書展", "26年5月課外活動展"]

【KOL 模块计算逻辑】
- 流速表 B 列指定 KOL（钟嘉欣图片/钟嘉欣视频1/钟嘉欣视频2/周家蔚 等）
- "其他汇总" = KOLHK 渠道组中除上述指定 KOL 外的所有供应商之和
- 约课成本 = 滚动消耗 / 约课数

【商超&社群模块计算逻辑】
- 代理商场 = 线下HK 渠道组中的全部商超供应商
- 书展 = 线下HK 渠道组中的展会供应商（按 BOOK_FAIR_SUPPLIERS 匹配）
- 代理人汇总 = 社群HK 渠道组的全部供应商

【MTD 例子目标计算】
从流速表的日维度数据列（J列起）累加至昨天日期对应的列

【公式（自动嵌套）】
- AQ列 例子gap = AP - AO
- AT列 约课例子gap = AS - AR
- 负数标红
"""
import sys
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

sys.stdout.reconfigure(encoding="utf-8")

SAMPLE_DIR = Path(__file__).parent / "sample"
FLOW_FILE = SAMPLE_DIR / "2026年5月港澳市场流速-初稿4.21.xlsx"
DATA_FILE = SAMPLE_DIR / "海外港澳商务_各渠道主辅投数据.xlsx"
OUTPUT_DIR = Path(__file__).parent / "output"

# 本月书展供应商（每月更新）
BOOK_FAIR_SUPPLIERS = ["26年5月第十二屆兒童書展", "26年5月課外活動展"]

# 流速表中已列出的 KOL 供应商（B列，行1-3）
# 脚本会自动从流速表读取，不需要硬编码

# 流速表关键列索引（0-based）
COL_B_SUPPLIER = 1       # B列：供应商名
COL_DAY_START = 9        # J列开始：每日流速数据
COL_AO_MTD_TARGET = 40   # AO列：MTD例子目标
COL_AP_ACTUAL = 41       # AP列：实际例子达成
COL_AQ_GAP = 42          # AQ列：例子gap
COL_AR_MTD_LESSON = 43   # AR列：MTD约课例子目标
COL_AS_LESSON_ACTUAL = 44  # AS列：实际约课达成
COL_AT_LESSON_GAP = 45   # AT列：约课例子gap
COL_AU_COST = 46         # AU列：约课成本

# 流速表数据行（0-based，对应 openpyxl 行号需要+1）
ROW_DATA_START = 1  # 行1开始是数据（行0是表头）

RED_FILL = PatternFill(start_color="FF6666", end_color="FF6666", fill_type="solid")


def load_data_table() -> dict:
    """
    读取数据底表，返回供应商级别的汇总数据
    返回: {
        'kol': {供应商名: {'例子数': x, '约课数': y, '滚动消耗': z}},
        'shequ': {供应商名: {...}},
        'xiaxian': {供应商名: {...}},
        'kol_total': {'例子数': x, '约课数': y, '滚动消耗': z},
        'shequ_total': {...},
        'xiaxian_total': {...},
    }
    """
    df = pd.read_excel(DATA_FILE, header=5)
    df["渠道组"] = df["渠道组"].ffill()
    df["供应商"] = df["供应商"].ffill()

    # 过滤掉口径说明行
    df = df[df["渠道组"].isin(["KOLHK", "主页HK", "独立站HK", "社群HK", "线下HK", "总计"])]

    # 找到每个供应商的总计行
    supplier_totals = df[df["渠道名称"].str.contains("总计", na=False)].copy()

    result = {"kol": {}, "shequ": {}, "xiaxian": {}}

    # KOLHK
    kol_rows = supplier_totals[supplier_totals["渠道组"] == "KOLHK"]
    for _, row in kol_rows.iterrows():
        name = row["供应商"]
        if name == "总计":
            result["kol_total"] = {
                "例子数": row["例子数"], "约课数": row["约课数"], "滚动消耗": row["滚动消耗"]
            }
        else:
            result["kol"][name] = {
                "例子数": row["例子数"], "约课数": row["约课数"], "滚动消耗": row["滚动消耗"]
            }

    # 社群HK
    sq_rows = supplier_totals[supplier_totals["渠道组"] == "社群HK"]
    for _, row in sq_rows.iterrows():
        name = row["供应商"]
        if name == "总计":
            result["shequ_total"] = {
                "例子数": row["例子数"], "约课数": row["约课数"], "滚动消耗": row["滚动消耗"]
            }
        else:
            result["shequ"][name] = {
                "例子数": row["例子数"], "约课数": row["约课数"], "滚动消耗": row["滚动消耗"]
            }

    # 线下HK
    xx_rows = supplier_totals[supplier_totals["渠道组"] == "线下HK"]
    for _, row in xx_rows.iterrows():
        name = row["供应商"]
        if name == "总计":
            result["xiaxian_total"] = {
                "例子数": row["例子数"], "约课数": row["约课数"], "滚动消耗": row["滚动消耗"]
            }
        else:
            result["xiaxian"][name] = {
                "例子数": row["例子数"], "约课数": row["约课数"], "滚动消耗": row["滚动消耗"]
            }

    return result


def get_days_this_month_until_yesterday() -> int:
    """返回当月1日到昨天的天数"""
    today = datetime.now()
    yesterday = today - timedelta(days=1)
    return yesterday.day


def process_flow_sheet(data: dict, book_fair_suppliers: list):
    """
    处理流速表【香港市场目标】sheet
    """
    OUTPUT_DIR.mkdir(exist_ok=True)
    output_file = OUTPUT_DIR / FLOW_FILE.name
    import shutil
    shutil.copy(FLOW_FILE, output_file)

    wb = load_workbook(output_file)
    ws = wb["香港市场目标"]
    print(f"[流速处理] 打开 sheet: 香港市场目标, 行数={ws.max_row}, 列数={ws.max_column}")

    # --- Step a: 计算 MTD例子目标（AO列）---
    # 日列从 COL_DAY_START+1 (openpyxl 1-based = 列10) 开始
    # 需要求和从当月1日到昨天的天数对应的列
    days = get_days_this_month_until_yesterday()
    day_col_start = COL_DAY_START + 1  # openpyxl 1-based
    day_col_end = day_col_start + days - 1
    print(f"[Step a] MTD目标求和: 列{day_col_start}~{day_col_end} (当月1日到昨天，共{days}天)")

    # 读取流速表的供应商列表（B列，从第2行开始，到"合计"行为止）
    flow_suppliers = []  # KOL 部分的供应商名
    data_rows = []  # (行号, 供应商名, 分类)

    for row_idx in range(2, ws.max_row + 1):
        b_val = ws.cell(row_idx, COL_B_SUPPLIER + 1).value  # openpyxl 1-based
        if b_val is None:
            continue
        b_val = str(b_val).strip()
        if b_val in ("", "KOL历史", "累计例子", "平均成本", "日均例子", "5月目标预计",
                     "KOL", "展会", "代理人", "商超", "线下HK历史达成数据"):
            break  # 到了下方的统计区域，停止

        data_rows.append((row_idx, b_val))

        # 收集 KOL 供应商名（排除汇总行和商超社群行）
        if b_val not in ("其他汇总", "KOL-汇总", "代理商场", "书展", "代理人汇总", "合计", "商超&社群合计"):
            flow_suppliers.append(b_val)

    print(f"[Step a] 流速表数据行: {len(data_rows)} 行")
    print(f"[Step a] KOL供应商: {flow_suppliers}")

    # 对每个数据行计算 MTD 目标求和
    for row_idx, supplier_name in data_rows:
        if supplier_name in ("KOL-汇总", "合计"):
            continue  # 汇总行用公式，不手动计算

        total = 0
        for col in range(day_col_start, day_col_end + 1):
            val = ws.cell(row_idx, col).value
            if val is not None and isinstance(val, (int, float)):
                total += val
        ws.cell(row_idx, COL_AO_MTD_TARGET + 1).value = total

    print(f"[Step a] MTD例子目标已更新")

    # --- Step c: 从数据底表匹配实际例子达成 / 约课达成 / 约课成本 ---
    # 记录各部分的滚动消耗和约课数，用于汇总行计算约课成本
    kol_total_consumption = 0
    kol_total_lessons = 0
    shangchao_total_consumption = 0
    shangchao_total_lessons = 0

    for row_idx, supplier_name in data_rows:
        if supplier_name in ("KOL-汇总", "合计", "商超&社群合计"):
            continue  # 汇总行单独处理

        examples = 0
        lessons = 0
        cost_consumption = 0

        if supplier_name == "其他汇总":
            # KOLHK 中不在 flow_suppliers 里的供应商求和
            for s_name, s_data in data["kol"].items():
                if s_name not in flow_suppliers:
                    examples += s_data["例子数"] or 0
                    lessons += s_data["约课数"] or 0
                    cost_consumption += s_data["滚动消耗"] or 0

        elif supplier_name == "代理人汇总":
            # 社群HK 总计
            t = data.get("shequ_total", {})
            examples = t.get("例子数", 0) or 0
            lessons = t.get("约课数", 0) or 0
            cost_consumption = t.get("滚动消耗", 0) or 0

        elif supplier_name == "书展":
            # 线下HK 中指定的书展供应商求和
            for bf in book_fair_suppliers:
                if bf in data["xiaxian"]:
                    examples += data["xiaxian"][bf]["例子数"] or 0
                    lessons += data["xiaxian"][bf]["约课数"] or 0
                    cost_consumption += data["xiaxian"][bf]["滚动消耗"] or 0

        elif supplier_name == "代理商场":
            # 线下HK 中除书展外的所有供应商求和
            for s_name, s_data in data["xiaxian"].items():
                if s_name not in book_fair_suppliers:
                    examples += s_data["例子数"] or 0
                    lessons += s_data["约课数"] or 0
                    cost_consumption += s_data["滚动消耗"] or 0

        else:
            # 普通 KOL 供应商：直接从 KOLHK 匹配
            if supplier_name in data["kol"]:
                s_data = data["kol"][supplier_name]
                examples = s_data["例子数"] or 0
                lessons = s_data["约课数"] or 0
                cost_consumption = s_data["滚动消耗"] or 0
            else:
                print(f"  [警告] 供应商 '{supplier_name}' 在数据底表中未找到")

        # 累计到对应的汇总
        if supplier_name in ("代理商场", "书展", "代理人汇总"):
            shangchao_total_consumption += cost_consumption
            shangchao_total_lessons += lessons
        else:
            kol_total_consumption += cost_consumption
            kol_total_lessons += lessons

        # 写入 AP列：实际例子达成
        ws.cell(row_idx, COL_AP_ACTUAL + 1).value = examples
        # 写入 AS列：实际约课达成
        ws.cell(row_idx, COL_AS_LESSON_ACTUAL + 1).value = lessons
        # 写入 AU列：约课成本 = 滚动消耗 / 约课数
        if lessons > 0:
            ws.cell(row_idx, COL_AU_COST + 1).value = cost_consumption / lessons
        else:
            ws.cell(row_idx, COL_AU_COST + 1).value = None

        print(f"  {supplier_name:15} → 例子={examples}, 约课={lessons}, "
              f"成本={round(cost_consumption/lessons, 2) if lessons > 0 else '-'}")

    # --- 处理汇总行的约课成本 ---
    for row_idx, supplier_name in data_rows:
        if supplier_name == "KOL-汇总":
            # KOL 汇总约课成本 = KOL 总滚动消耗 / KOL 总约课数
            if kol_total_lessons > 0:
                ws.cell(row_idx, COL_AU_COST + 1).value = kol_total_consumption / kol_total_lessons
            print(f"  KOL-汇总          → 约课成本={round(kol_total_consumption/kol_total_lessons, 2) if kol_total_lessons > 0 else '-'}")

        elif supplier_name == "合计" and ws.cell(row_idx, COL_AO_MTD_TARGET + 1).value and \
                "AO7" in str(ws.cell(row_idx, COL_AO_MTD_TARGET + 1).value):
            # 行10：商超&社群合计（原公式引用 AO7:AO9）
            ws.cell(row_idx, COL_B_SUPPLIER + 1).value = "商超&社群合计"
            if shangchao_total_lessons > 0:
                ws.cell(row_idx, COL_AU_COST + 1).value = shangchao_total_consumption / shangchao_total_lessons
            print(f"  商超&社群合计        → 约课成本={round(shangchao_total_consumption/shangchao_total_lessons, 2) if shangchao_total_lessons > 0 else '-'}")

        elif supplier_name == "合计":
            # 行11：总合计
            all_consumption = kol_total_consumption + shangchao_total_consumption
            all_lessons = kol_total_lessons + shangchao_total_lessons
            if all_lessons > 0:
                ws.cell(row_idx, COL_AU_COST + 1).value = all_consumption / all_lessons
            print(f"  合计（总）          → 约课成本={round(all_consumption/all_lessons, 2) if all_lessons > 0 else '-'}")

    # --- Step b: 标红负数 gap ---
    for row_idx, supplier_name in data_rows:
        # AQ列：例子gap
        gap_val = ws.cell(row_idx, COL_AQ_GAP + 1).value
        if gap_val is not None and isinstance(gap_val, (int, float)) and gap_val < 0:
            ws.cell(row_idx, COL_AQ_GAP + 1).fill = RED_FILL

        # AT列：约课例子gap
        lesson_gap_val = ws.cell(row_idx, COL_AT_LESSON_GAP + 1).value
        if lesson_gap_val is not None and isinstance(lesson_gap_val, (int, float)) and lesson_gap_val < 0:
            ws.cell(row_idx, COL_AT_LESSON_GAP + 1).fill = RED_FILL

    print(f"[Step b] 负数 gap 已标红")

    wb.save(output_file)
    print(f"\n[完成] 输出: {output_file}")
    return output_file


def main():
    print("=" * 60)
    print("港澳商务流速处理")
    print("=" * 60)

    print(f"\n[配置] 书展供应商: {BOOK_FAIR_SUPPLIERS}")
    print(f"[配置] 流速表: {FLOW_FILE.name}")
    print(f"[配置] 数据底表: {DATA_FILE.name}")

    print(f"\n--- 加载数据底表 ---")
    data = load_data_table()
    print(f"  KOLHK 供应商: {len(data['kol'])} 个")
    print(f"  社群HK 供应商: {len(data['shequ'])} 个")
    print(f"  线下HK 供应商: {len(data['xiaxian'])} 个")

    print(f"\n--- 处理流速表 ---")
    process_flow_sheet(data, BOOK_FAIR_SUPPLIERS)


if __name__ == "__main__":
    main()
