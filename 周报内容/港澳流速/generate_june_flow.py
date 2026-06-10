#!/usr/bin/env python3
"""
生成6月港澳流速HTML报表
从 sample/ 目录自动查找最新的流速数据文件
"""
import sys
from pathlib import Path
from openpyxl import load_workbook
from datetime import datetime

sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = Path(__file__).parent
SAMPLE_DIR = BASE_DIR / "sample"
OUTPUT_DIR = BASE_DIR / "output"


def find_latest_excel(directory: Path, pattern: str = "*.xlsx") -> Path:
    """找到目录中最新的 Excel 文件"""
    files = list(directory.glob(pattern))
    if not files:
        raise FileNotFoundError(f"未找到Excel文件: {directory}")
    # 按修改时间排序，返回最新的
    return max(files, key=lambda p: p.stat().st_mtime)


def extract_flow_data():
    """提取流速数据"""
    # 自动查找最新文件
    excel_file = find_latest_excel(SAMPLE_DIR)
    print(f"[信息] 使用文件: {excel_file.name}")

    wb = load_workbook(excel_file, data_only=True)
    ws = wb["Sheet1"]

    print(f"[信息] 工作表: {wb.sheetnames}")

    # 读取日期范围（第1-2行）
    start_date = ws.cell(1, 3).value
    end_date = ws.cell(2, 3).value
    print(f"[信息] 数据时间范围: {start_date} 至 {end_date}")

    # 表头在第6行，数据从第7行开始
    rows_data = []
    current_channel = ""  # 记录当前渠道

    for row_idx in range(7, 100):  # 读取到第100行
        channel = ws.cell(row_idx, 2).value  # 渠道名称
        supplier = ws.cell(row_idx, 3).value  # 供应商
        subchannel = ws.cell(row_idx, 4).value  # 二级渠道

        # 更新当前渠道
        if channel:
            current_channel = str(channel)

        # 如果供应商和二级渠道都为空，可能到了数据末尾
        if not supplier and not subchannel:
            break

        # 跳过"合计"行（但不是最后的总合计）
        if subchannel == "合计" and supplier is None:
            continue

        # 提取关键数据
        cost = ws.cell(row_idx, 9).value  # 消耗成本
        example_num = ws.cell(row_idx, 11).value  # 发放数
        lead_num = ws.cell(row_idx, 12).value  # 引流数
        reservation = ws.cell(row_idx, 14).value  # 约课数
        attendance = ws.cell(row_idx, 15).value  # 应到数

        row_dict = {
            "channel": current_channel,
            "supplier": str(supplier) if supplier else "",
            "subchannel": str(subchannel) if subchannel else "",
            "cost": float(cost) if cost and isinstance(cost, (int, float)) else 0,
            "example_num": float(example_num) if example_num and isinstance(example_num, (int, float)) else 0,
            "lead_num": float(lead_num) if lead_num and isinstance(lead_num, (int, float)) else 0,
            "reservation": float(reservation) if reservation and isinstance(reservation, (int, float)) else 0,
            "attendance": float(attendance) if attendance and isinstance(attendance, (int, float)) else 0,
        }
        rows_data.append(row_dict)

    return rows_data, excel_file.name, start_date, end_date


def generate_html(rows_data, source_file, start_date, end_date):
    """生成HTML报表"""
    html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>6月港澳流速数据</title>
    <style>
        body {{
            font-family: "Microsoft YaHei", Arial, sans-serif;
            margin: 20px;
            background-color: #f5f5f5;
        }}
        .container {{
            max-width: 1400px;
            margin: 0 auto;
            background-color: white;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }}
        h1 {{
            color: #333;
            border-bottom: 3px solid #4CAF50;
            padding-bottom: 10px;
        }}
        .info {{
            color: #666;
            margin-bottom: 20px;
            font-size: 14px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
            font-size: 13px;
        }}
        th, td {{
            border: 1px solid #ddd;
            padding: 10px 8px;
            text-align: center;
        }}
        th {{
            background-color: #4CAF50;
            color: white;
            font-weight: bold;
        }}
        tr:nth-child(even) {{
            background-color: #f9f9f9;
        }}
        tr:hover {{
            background-color: #f1f1f1;
        }}
        .text-left {{
            text-align: left !important;
        }}
        .num {{
            text-align: right !important;
        }}
        .footer {{
            margin-top: 30px;
            text-align: center;
            color: #999;
            font-size: 12px;
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1>6月港澳流速数据</h1>
        <div class="info">
            <p><strong>数据来源：</strong>{source_file}</p>
            <p><strong>数据时间：</strong>{start_date} 至 {end_date}</p>
            <p><strong>生成时间：</strong>{datetime.now().strftime("%Y年%m月%d日 %H:%M:%S")}</p>
        </div>

        <table>
            <thead>
                <tr>
                    <th>序号</th>
                    <th>渠道</th>
                    <th>供应商</th>
                    <th>二级渠道</th>
                    <th>消耗成本</th>
                    <th>发放数</th>
                    <th>引流数</th>
                    <th>约课数</th>
                    <th>应到数</th>
                </tr>
            </thead>
            <tbody>
"""

    for idx, row in enumerate(rows_data, 1):
        html += f"""                <tr>
                    <td>{idx}</td>
                    <td class="text-left">{row['channel']}</td>
                    <td class="text-left">{row['supplier']}</td>
                    <td class="text-left">{row['subchannel']}</td>
                    <td class="num">{row['cost']:.2f}</td>
                    <td class="num">{row['example_num']:.0f}</td>
                    <td class="num">{row['lead_num']:.0f}</td>
                    <td class="num">{row['reservation']:.0f}</td>
                    <td class="num">{row['attendance']:.0f}</td>
                </tr>
"""

    html += """            </tbody>
        </table>

        <div class="footer">
            <p>自动生成于 BI 报表系统</p>
        </div>
    </div>
</body>
</html>
"""

    return html


def main():
    print("=" * 60)
    print("6月港澳流速数据报表生成")
    print("=" * 60)
    print()

    # 提取数据
    rows_data, source_file, start_date, end_date = extract_flow_data()
    print(f"\n[信息] 提取了 {len(rows_data)} 行数据")

    # 生成HTML
    html = generate_html(rows_data, source_file, start_date, end_date)

    # 保存文件
    OUTPUT_DIR.mkdir(exist_ok=True)
    output_file = OUTPUT_DIR / f"6月港澳流速_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"\n[成功] HTML报表已生成: {output_file}")
    print(f"[成功] 请用浏览器打开查看")
    print("=" * 60)


if __name__ == "__main__":
    main()
