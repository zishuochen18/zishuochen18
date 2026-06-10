"""
单独生成书展模块 HTML（用于验证）
"""
import sys
from pathlib import Path

# 添加父目录到 sys.path
parent_dir = Path(__file__).parent.parent
sys.path.insert(0, str(parent_dir))
sys.stdout.reconfigure(encoding="utf-8")

# 导入模块
from 书展数据复盘 import process_bookfair
from 书展内容 import process_book_fair

print("="*60)
print("书展模块单独验证")
print("="*60)
print()

# Step 1: 下载 + 整合数据
print("[1] 数据整合（含自动下载）...")
process_bookfair.merge_bookfair_data()
print()

# Step 2: 提取数据
print("[2] 提取书展数据...")
book_fair_data = process_book_fair.extract_book_fair_data()
if not book_fair_data:
    print("  ⚠️ 未找到书展数据")
    sys.exit(1)

print(f"  上月书展: {len(book_fair_data.get('last_month_fairs', {}).get('rows', []))} 行")
print(f"  本月书展: {len(book_fair_data.get('current_month_fairs', {}).get('rows', []))} 行")
print(f"  历史书展: {len(book_fair_data.get('history', {}).get('rows', []))} 行")
print()

# Step 3: 生成 HTML
print("[3] 生成 HTML...")
from jinja2 import Template

# 简化版 HTML 模板（仅书展部分）
html_template = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>书展数据复盘（单模块验证）</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #333; }
        h2 { color: #666; margin-top: 30px; }
        table { border-collapse: collapse; width: 100%; margin: 20px 0; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        .summary { background-color: #e8f4f8; padding: 15px; margin: 20px 0; border-radius: 5px; }
    </style>
</head>
<body>
    <h1>书展数据复盘（验证版）</h1>
    <p>生成时间: {{ periods.today }}</p>

    {% if last_month_fairs %}
    <h2>4.1 上月书展（{{ periods.last_month.prefix }}）</h2>
    <div class="summary">
        <p>统计周期: {{ periods.last_month_to_now[0] }} ~ {{ periods.last_month_to_now[1] }}</p>
        <p>场次: {{ last_month_fairs.summary.fair_count }} 场</p>
    </div>
    <table>
        <tr>
            <th>供应商</th>
            <th>渠道名称</th>
            <th>滚动消耗</th>
            <th>滚动ROI2总成本</th>
            <th>滚动GMV</th>
            <th>滚动ROI2</th>
            <th>例子数</th>
            <th>滚动到课数</th>
        </tr>
        {% for row in last_month_fairs.rows %}
        <tr>
            <td>{{ row.供应商 }}</td>
            <td>{{ row.渠道名称 }}</td>
            <td>{{ "%.2f"|format(row.滚动消耗) }}</td>
            <td>{{ "%.2f"|format(row.滚动ROI2总成本) }}</td>
            <td>{{ "%.2f"|format(row.滚动GMV) }}</td>
            <td>{{ "%.4f"|format(row.滚动ROI2) }}</td>
            <td>{{ row.例子数 }}</td>
            <td>{{ row.滚动到课数 }}</td>
        </tr>
        {% endfor %}
    </table>
    {% endif %}

    {% if current_month_fairs %}
    <h2>4.2 本月书展（{{ periods.current_month.prefix }}）</h2>
    <div class="summary">
        <p>统计周期: {{ periods.current_month_to_now[0] }} ~ {{ periods.current_month_to_now[1] }}</p>
        <p>场次: {{ current_month_fairs.summary.fair_count }} 场</p>
    </div>
    <table>
        <tr>
            <th>供应商</th>
            <th>渠道名称</th>
            <th>滚动消耗</th>
            <th>滚动GMV</th>
            <th>滚动ROI2</th>
            <th>例子数</th>
        </tr>
        {% for row in current_month_fairs.rows %}
        <tr>
            <td>{{ row.供应商 }}</td>
            <td>{{ row.渠道名称 }}</td>
            <td>{{ "%.2f"|format(row.滚动消耗) }}</td>
            <td>{{ "%.2f"|format(row.滚动GMV) }}</td>
            <td>{{ "%.4f"|format(row.滚动ROI2) }}</td>
            <td>{{ row.例子数 }}</td>
        </tr>
        {% endfor %}
    </table>
    {% endif %}

    {% if history %}
    <h2>4.3 历史书展对比</h2>
    <table>
        <tr>
            <th>供应商</th>
            <th>滚动消耗</th>
            <th>滚动GMV</th>
            <th>滚动ROI2</th>
            <th>例子数</th>
        </tr>
        {% for row in history.rows %}
        <tr>
            <td>{{ row.供应商 }}</td>
            <td>{{ "%.2f"|format(row.滚动消耗) }}</td>
            <td>{{ "%.2f"|format(row.滚动GMV) }}</td>
            <td>{{ "%.4f"|format(row.滚动ROI2) }}</td>
            <td>{{ row.例子数 }}</td>
        </tr>
        {% endfor %}
    </table>
    {% endif %}
</body>
</html>
"""

template = Template(html_template)
html = template.render(
    periods=book_fair_data['periods'],
    last_month_fairs=book_fair_data.get('last_month_fairs'),
    current_month_fairs=book_fair_data.get('current_month_fairs'),
    history=book_fair_data.get('history')
)

output_dir = Path(__file__).parent / "output"
output_dir.mkdir(exist_ok=True)
output_file = output_dir / "书展数据复盘_单模块验证.html"

with open(output_file, 'w', encoding='utf-8') as f:
    f.write(html)

print(f"  ✅ HTML 已生成: {output_file}")
print()
print("="*60)
print("完成！请打开 HTML 文件检查以下关键数据：")
print("  1. 上月书展场次是否正确（应该有 2 场：第十二屆兒童書展、課外活動展）")
print("  2. 每场书展的\"总计\"行：")
print("     - 滚动消耗 应该 > 5月完整月的消耗（因为累加了6月）")
print("     - 滚动GMV 应该 > 5月完整月的GMV（因为6月有延续转化）")
print("     - 滚动ROI2 = 滚动GMV / 滚动ROI2总成本（整合后）")
print("="*60)
