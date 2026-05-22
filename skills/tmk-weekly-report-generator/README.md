# tmk-weekly-report-generator

TMK 做工周报自动生成脚本。读取本周下载的 3 个 BI Excel 文件，自动产出 Markdown 周报 + 带分组表头的 Excel 数据表。

## 依赖

```bash
pip install pandas openpyxl
```

Python 3.10+。

## 运行

```bash
python generate_weekly_report.py <数据文件夹路径>
```

例子：

```bash
python generate_weekly_report.py 5.22测试
```

## 输入文件

将本周下载的 3 个 Excel 放到一个子文件夹下（文件名包含以下关键词即可被自动识别）：

| 关键词 | 报表 | 主要字段 |
|--------|------|---------|
| `做工监控` | TMK 做工监控 | 日均通次/通时, 跟进时效, 新生/老生跟进, 结果指标 |
| `未邀约` | TMK 未邀约做工监控播报 | 进线非勿扰时段相关 |
| `勿扰` | TMK 做工勿扰情况汇总 | 勿扰跟进、约课率、首次接通邀约率 |

## 输出

在同一文件夹下生成：

| 文件 | 说明 |
|------|------|
| `周报.md` | Markdown 周报，含整体概览 + 个人异常自动点出 |
| `个人做工数据_<起止日期>.xlsx` | 完整数据表，带分组表头、合并单元格、总计/团队汇总高亮 |

## 周报结构

1. **整体数据概览**：团队日均通次（含环比）、日均通时、生均跟进时效达成率
2. **个人做工数据**（链接到 Excel 文件）：
   - 大盘总计（黄色高亮）
   - 海外夜班 TMK 组团队汇总（蓝色高亮） + 个人
   - 海外白班 TMK 组团队汇总（蓝色高亮） + 个人
3. **个人异常情况**：环比下降超 10% 自动点出，含日均通次、日均通时、跟进时效达成率三个维度

## 调整阈值

```python
ALERT_THRESHOLD = -0.10  # 环比下降超过 10% 视为异常
```

修改 `generate_weekly_report.py` 顶部的 `ALERT_THRESHOLD` 即可调整。

## 调整数据列

`STANDARD_COLUMNS` 列表定义了输出 Excel 的列结构（分组、列名、来源文件、数据格式）。每个元组：

```python
(分组名, 列名, 来源文件关键词, 数据格式)
```

- 数据格式：`int`（整数）/ `num`（保留 2 位小数）/ `pct`（百分比）

新增/删减列只需修改这个列表，无需改动业务逻辑。

## 适配说明

脚本针对**带筛选行 + 多级表头 + 列名跨文件命名差异**的 BI 导出 Excel 做了通用化处理：

| 函数 | 用途 |
|------|------|
| `parse_excel_with_groups` | 通用：根据分组行/列名行解析 BI Excel |
| `find_cross_file_row` | 处理跨文件命名差异（如做工监控里"团队汇总"≈勿扰文件里"总计"） |
| `infer_source_group` | 标准列分组映射到源 Excel 实际分组 |
| `merge_person_data` | 合并三个数据源到一张完整表 |
| `analyze_alerts` | 自动识别个人异常 |

## 文件结构

```
tmk-weekly-report-generator/
  generate_weekly_report.py    # 主脚本
  README.md
  SKILL.md
```

## 注意事项

- 输入 Excel 中的真实业务数据（人员姓名、具体数值）属于敏感信息，不会进入版本库
- 数据子文件夹（含 .xlsx）需放到本目录之外，或确保被 `.gitignore` 排除
- 如果 BI 导出格式发生变化（列顺序变化、新增字段），调整 `STANDARD_COLUMNS` 即可
