---
name: hk-mo-unclassified-student-email
description: Filter unclassified students (>14 days) from BI export and draft notification email via Tencent Enterprise Mail
version: 1.0.0
tags: [excel, email-automation, playwright, pandas]
---

# HK/MO Unclassified Student Email

Automates the weekly workflow of filtering students who have been waiting for class assignment >14 days in the HK/Macau overseas business channel, then drafting a notification email to the academic affairs team via Tencent Enterprise Mail (exmail.qq.com).

## Technique

- pandas + openpyxl for Excel filtering (multi-condition: stage + channel + wait days)
- Playwright (sync_api) for browser automation on Tencent Enterprise Mail
- Cookie-based login persistence (email_auth_state.json)
- HTML table injection into rich-text email editor via iframe manipulation
- Frameset navigation: top frame → mainFrame → compose page → editor iframe
