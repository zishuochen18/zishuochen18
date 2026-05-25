# zishuochen18

个人技能库（Skills Repository）

收录可复用的自动化脚本、工具和流程模板。

## 目录结构

```
skills/
  crm-channel-promotion-automation/   # CRM 渠道 + 推广自动化（Playwright + 飞书集成）
  tmk-weekly-report-generator/        # TMK 做工周报自动生成（pandas + openpyxl）
  hk-mo-unclassified-student-email/   # 港澳未分班学员明细筛选 + 邮件同步教务（Playwright + pandas）
```

## 使用规范

- 每个 skill 有独立的 `README.md` 说明用法
- 敏感配置使用 `config.local.json`（本地，不进版本库）
- 模板配置使用 `config.example.json`（进版本库）
