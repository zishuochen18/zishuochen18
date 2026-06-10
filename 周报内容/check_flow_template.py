import sys
sys.stdout.reconfigure(encoding="utf-8")
import openpyxl

wb = openpyxl.load_workbook('港澳流速/sample/2026年6月港澳市场流速-初稿5.21.xlsx')
ws = wb['香港市场目标']

print("流速表模板B列供应商列表(行2-20):")
for row_idx in range(2, 21):
    b_val = ws.cell(row_idx, 2).value
    if b_val:
        print(f"  行{row_idx}: {b_val}")
    if str(b_val).strip() in ("合计", "KOL历史", ""):
        if str(b_val).strip() == "":
            continue
        else:
            break
