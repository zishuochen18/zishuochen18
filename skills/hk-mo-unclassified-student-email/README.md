# 港澳地区未分班超14天明细同步教务

自动处理未分班学员 Excel 数据，筛选符合条件的记录，并通过腾讯企业邮箱发送给教务团队。

## 功能

1. **Excel 筛选**：从 BI 导出的未分班学员明细中，按学员阶段、渠道分类、等待时长三重条件筛选
2. **邮件草稿**：自动登录腾讯企业邮箱，填写收件人/抄送/主题/正文（含 HTML 表格）/附件，保存为草稿

## 依赖

```
pip install pandas openpyxl playwright
playwright install chromium
```

## 使用方式

1. 将最新的 Excel 文件放到 `sample/海外未分班学员明细.xlsx`
2. 运行脚本：

```bash
cd 港澳未分班
python process_and_email.py
```

首次运行需要手动登录邮箱（输入账号密码），登录状态会缓存到 `email_auth_state.json`。

## 筛选逻辑

| 列名 | 条件 |
|------|------|
| 学员阶段 | = 组班中 |
| 渠道一级分类 | = 海外港澳商务 |
| 未退费等待时长 | > 14 天 |

## 输出

- Excel 文件：`output/【港澳海外商务】等班14天以上学员明细+{日期}.xlsx`
- 邮件草稿：含正文表格 + Excel 附件

## 配置

修改脚本顶部常量即可自定义：
- `RECIPIENTS_TO` / `RECIPIENTS_CC`：收件人和抄送人
- `FILTER_WAIT_THRESHOLD`：等待天数阈值（默认 14）
- `SOURCE_SHEET`：源 Excel 的 sheet 名称
