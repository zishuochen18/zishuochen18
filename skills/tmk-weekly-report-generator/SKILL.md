---
name: tmk-weekly-report-generator
description: Use when the user needs to automatically generate weekly TMK (telemarketing) work reports from BI-exported Excel files. Handles three multi-source spreadsheets, merges cross-file data, identifies underperforming team members, and produces both Markdown summary and styled Excel output.
---

# TMK Weekly Report Generator

Automate weekly TMK work report generation from three BI-exported Excel files.

## What this skill does

Given three weekly BI exports:
1. TMK 做工监控 (work monitoring) — main metrics: 日均通次/通时, 跟进时效, 新生跟进, 老生跟进, 结果指标
2. TMK 未邀约做工监控播报 — 进线非勿扰时段相关指标
3. TMK 做工勿扰情况汇总 — 勿扰跟进、约课率、首次接通邀约率

Produces:
- `周报.md` — overview with overall metrics + auto-flagged personal anomalies (week-over-week drop > 10%)
- `个人做工数据_<period>.xlsx` — full data table with grouped headers (merged cells), 大盘总计 + 团队汇总 + 每位 TMK 数据

## Key design decisions

- Excel-style filter rows are skipped via header row detection by group/column name positions, not the first numeric data row (which can be sparse for inactive members).
- Cross-file naming differences are handled by `find_cross_file_row`: e.g., 做工监控 uses "团队汇总" while 勿扰汇总 uses "总计" for the same team aggregate row.
- 总计 / 团队汇总 / 个人 rows are merged into one ordered output dataframe and visually distinguished in Excel via highlight fills.
- Metric source group inference (`infer_source_group`) maps standard column names to the correct group within each source file (e.g., 首次接通邀约率 lives in 结果指标 group of the 勿扰 file, not 勿扰情况).
- Output uses HTML inside Markdown for multi-row group headers fallback, plus a separate openpyxl-styled .xlsx for direct copy-paste into 飞书 docs.

## Usage

```bash
pip install pandas openpyxl
python generate_weekly_report.py <data-folder>
```

The data folder must contain three .xlsx files whose names include keywords: `做工监控`, `未邀约`, `勿扰`.
