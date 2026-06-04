"""
书展数据复盘处理
1. 从数据底表 海外港澳商务_各渠道主辅投数据.xlsx 提取本月两个书展的明细 + 汇总
2. 写入书展数据汇总.xlsx 中（保留历史书展不动）
3. 提供 extract_book_fair_data() 接口给主 HTML 使用
"""
"""
书展数据复盘处理脚本（模块 4）

═══════════════════════════════════════════════════════════════
功能：从数据底表提取本月书展数据，与历史书展对比
═══════════════════════════════════════════════════════════════

【输入】
- sample/书展数据汇总.xlsx（手工维护，行 2-18 是历史书展）
- ../港澳流速/sample/海外港澳商务_各渠道主辅投数据.xlsx（数据底表）

【输出】
- output/书展数据汇总.xlsx（更新本月书展数据）

【关键参数】
- BOOK_FAIR_SUPPLIERS：本月书展供应商列表（每月需更新）
  例：["26年5月課外活動展", "26年5月第十二屆兒童書展"]

【数据流】
1. 从数据底表提取本月每个书展的承担人明细 + 总计
2. 写入【书展数据汇总.xlsx】（保留历史书展行 2-18 不动）
3. 行 19 起重写本月书展数据

【承担人简称提取规则】（extract_short_name）
- 优先：从 "-shukei-XX-思维" 中取 XX，前缀加 "-"（如 "-Amy-阿珍"）
- shukei 后紧跟"思维/主投/美术"时：取 "9.9小课包" 后的字符（如"李忠熹"）

【字段范围】
仅保留到第一个【滚动ROI2】列，后续辅投/汇总区域忽略

【对外接口】
- extract_book_fair_data() - 返回本月书展 + 历史书展数据

【注意事项】
- 历史书展行（2-18）保持不动，每月只更新本月行（19+）
- 重点指标：例子约课率、约课到课率、到课转化率、滚动转化率、滚动ROI2
"""
import sys
import shutil
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = Path(__file__).parent
SAMPLE_DIR = BASE_DIR / "sample"
OUTPUT_DIR = BASE_DIR / "output"
DATA_FILE = BASE_DIR.parent / "港澳流速" / "sample" / "海外港澳商务_各渠道主辅投数据.xlsx"
SUMMARY_FILE = SAMPLE_DIR / "书展数据汇总.xlsx"

# 本月需要更新的书展供应商
CURRENT_BOOK_FAIRS = ["26年5月課外活動展", "26年5月第十二屆兒童書展"]

# 字段顺序（从数据底表读，写入汇总表）。汇总表只保留到第一个"滚动ROI2"
SUMMARY_FIELDS = [
    "滚动消耗", "滚动消耗(不含赠课成本)", "单例子成本", "例子数", "分发数",
    "分发数\n(剔除毛例子)", "约课数", "应到课数", "滚动应到课数",
    "到课数", "滚动到课数", "当月成交数", "滚动成交数",
    "当月GMV", "滚动GMV", "海外GMV占比", "滚动ASP",
    "例子分发率", "分发约课率", "例子约课率", "约课到课率",
    "应到课率", "滚动应到课率", "到课转化率", "注册转化率",
    "滚动转化率", "滚动ROI2总成本", "滚动ROI2"
]

# 重点关注指标（HTML 中高亮）
KEY_METRICS = ["例子约课率", "约课到课率", "到课转化率", "滚动转化率", "滚动ROI2"]


def extract_short_name(channel_name: str) -> str:
    """
    从数据底表的"渠道名称"中提取承担人简称
    规则：
      1. 优先取 "-shukei-" 到 "-思维" 之间的内容（前面加"-"），但若 shukei 后紧接"思维"则跳过
      2. 否则取 "9.9小课包" 后的字符
      3. 否则原样返回
    """
    if not isinstance(channel_name, str):
        return ""

    if "-shukei-" in channel_name:
        idx = channel_name.find("-shukei-") + len("-shukei-")
        rest = channel_name[idx:]
        # 若 shukei 后紧接"思维"或"主投"，说明没承担人，跳过
        if not rest.startswith(("思维", "主投", "美术")) and "-思维" in rest:
            middle = rest.split("-思维")[0].strip()
            if middle:
                return f"-{middle}"

    if "9.9小课包" in channel_name:
        return channel_name.split("9.9小课包")[-1].strip()

    # 兜底：取最后一段
    parts = channel_name.split("-")
    return parts[-1].strip() if parts else channel_name


def load_book_fair_details(supplier_name: str) -> dict:
    """从数据底表提取指定书展的明细行 + 总计行"""
    df = pd.read_excel(DATA_FILE, header=5)
    df["渠道组"] = df["渠道组"].ffill()
    df["供应商"] = df["供应商"].ffill()

    sub = df[df["供应商"] == supplier_name].copy()

    details = []
    total = None
    for _, row in sub.iterrows():
        chan = row["渠道名称"]
        if not isinstance(chan, str):
            continue

        rec = {"_full_channel": chan}
        for f in SUMMARY_FIELDS:
            v = row.get(f)
            rec[f] = v if pd.notna(v) else None

        if chan == "总计":
            total = rec
            total["short_name"] = supplier_name  # 汇总行就用供应商名
        else:
            rec["short_name"] = extract_short_name(chan)
            details.append(rec)

    return {"total": total, "details": details}


def write_to_summary(book_fairs_data: dict):
    """
    把本月两个书展的数据写入 书展数据汇总.xlsx。
    保留行 1 (表头) 和行 2-18 (历史书展)，从行 19 开始重写本月数据。
    """
    OUTPUT_DIR.mkdir(exist_ok=True)
    out_file = OUTPUT_DIR / SUMMARY_FILE.name
    shutil.copy(SUMMARY_FILE, out_file)

    wb = load_workbook(out_file)
    ws = wb["Sheet1"]

    # 清空行 19 起的所有数据（保险：清空 19~50 行）
    for row in range(19, 60):
        for col in range(1, ws.max_column + 1):
            ws.cell(row, col).value = None

    # 列映射：列1=渠道名称(short_name)，列2-29=数据列（对应 SUMMARY_FIELDS）
    write_row = 19
    for fair_name in ["26年5月課外活動展", "26年5月第十二屆兒童書展"]:
        data = book_fairs_data.get(fair_name)
        if not data:
            continue

        # 写汇总行
        if data["total"]:
            ws.cell(write_row, 1).value = fair_name
            for i, f in enumerate(SUMMARY_FIELDS, start=2):
                v = data["total"].get(f)
                ws.cell(write_row, i).value = v
            write_row += 1

        # 写明细行
        for d in data["details"]:
            ws.cell(write_row, 1).value = d["short_name"]
            for i, f in enumerate(SUMMARY_FIELDS, start=2):
                v = d.get(f)
                ws.cell(write_row, i).value = v
            write_row += 1

    wb.save(out_file)
    print(f"  [书展汇总] 已写入 {write_row - 19} 行（含汇总+明细）")
    return out_file


def extract_book_fair_data() -> dict:
    """供 HTML 周报使用的接口"""
    # 1. 历史书展（直接读 sample 模板的行 2-18）
    wb = load_workbook(SUMMARY_FILE, data_only=True)
    ws = wb["Sheet1"]

    headers = ["渠道名称"] + SUMMARY_FIELDS

    history_rows = []
    for row in range(3, 20):  # 行 2-18 (openpyxl 1-based: 3-19)
        name = ws.cell(row, 1).value
        if not name:
            continue
        rec = {"name": str(name).strip()}
        for i, f in enumerate(SUMMARY_FIELDS, start=2):
            v = ws.cell(row, i).value
            rec[f] = v
        history_rows.append(rec)

    # 2. 本月两个书展（从数据底表实时计算）
    current_data = {}
    for fair_name in CURRENT_BOOK_FAIRS:
        d = load_book_fair_details(fair_name)
        # 把 short_name 改成 name 字段
        current_data[fair_name] = {
            "total": d["total"],
            "details": [{"name": x["short_name"], **{f: x.get(f) for f in SUMMARY_FIELDS}}
                        for x in d["details"]]
        }
        if d["total"]:
            current_data[fair_name]["total"] = {
                "name": fair_name,
                **{f: d["total"].get(f) for f in SUMMARY_FIELDS}
            }

    return {
        "headers": headers,
        "summary_fields": SUMMARY_FIELDS,
        "key_metrics": KEY_METRICS,
        "history": history_rows,
        "current": current_data,
        "current_names": CURRENT_BOOK_FAIRS,
    }


def main():
    print("=" * 60)
    print("书展数据复盘")
    print("=" * 60)

    print("\n[1] 从数据底表提取本月书展数据...")
    book_fairs = {}
    for fair_name in CURRENT_BOOK_FAIRS:
        d = load_book_fair_details(fair_name)
        book_fairs[fair_name] = d
        if d["total"]:
            t = d["total"]
            print(f"  {fair_name}: 例子={t.get('例子数')}, 约课={t.get('约课数')}, 滚动消耗={t.get('滚动消耗')}")
            print(f"    明细 {len(d['details'])} 行：")
            for x in d["details"]:
                print(f"      - {x['short_name']:<15} 例子={x.get('例子数')}, 约课={x.get('约课数')}")
        else:
            print(f"  ⚠️ {fair_name}: 未找到")

    print("\n[2] 写入书展数据汇总.xlsx...")
    out_file = write_to_summary(book_fairs)
    print(f"\n[完成] 输出: {out_file}")


if __name__ == "__main__":
    main()
