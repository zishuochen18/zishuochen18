"""检查【本月汇总数据】sheet 中5月各行的标识"""
import sys
from pathlib import Path
from openpyxl import load_workbook

sys.stdout.reconfigure(encoding="utf-8")

CONV_FILE = Path(r"c:\Users\chenzishuo\Desktop\新建文件夹 (2)\周报内容\本月KOL转化数据汇总\sample\本月KOL转化链路数据汇总.xlsx")

wb = load_workbook(CONV_FILE)

print("【钟嘉欣图片&视频数据】sheet 标题行（行2）和数据行（行3-4）：")
ws2 = wb["钟嘉欣图片&视频数据"]
for row in [1, 2, 3, 4]:
    print(f"  行{row}:")
    for col in range(1, ws2.max_column + 1):
        v = ws2.cell(row, col).value
        if v is not None:
            print(f"    [{col}] {v}")

