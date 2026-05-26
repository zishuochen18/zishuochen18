---
name: hk-mo-coin-grant-oa
description: Process weekly sales detail Excel and auto-submit virtual coin grant application to internal OA (Landray-based) with SSO login, complex form filling including manifest address picker, radio buttons, file upload, and save-as-draft
version: 1.0.0
tags: [excel, oa-automation, playwright, pandas, sso, landray]
---

# HK/MO Channel Coin Grant OA Auto-Submission

Automates the weekly workflow of submitting virtual coin (豌豆币) grant applications for HK/Macau overseas business channel transactions:

1. Read weekly sales detail Excel → filter & transform student data into the company's coin grant template
2. Login to internal OA via SSO → navigate to coin grant application form → fill 13 form fields → upload generated Excel → save as draft

## Technical Highlights

- **pandas + openpyxl** for Excel filtering with explicit text format on ID/order columns (prevents scientific notation on 23-digit order numbers)
- **Playwright** for browser automation across SSO + multi-iframe OA portal + popup compose page
- **Cookie-based SSO persistence** with login state verification before operations (refreshes session if expired)
- **Three field interaction patterns** for the Landray OA form:
  - Text/textarea: by `name` attribute pattern with JS fallback for hidden fields
  - Radio: direct `[value="X"]` selector with `force=True` click + onclick dispatch
  - Manifest address picker (jQuery plugin): click container → type → keyboard `ArrowDown`+`Enter` to select first result
- **iframe retry logic**: OA portal renders portlets in nested iframes that load asynchronously
- **JS-based link click** in iframe to bypass Playwright's text matching strictness with whitespace/newlines around Chinese text
- **Dialog mask removal** before clicking save-draft button (file upload progress overlay can intercept clicks)

## Date Logic

Output filename uses last week's Mon-Sun range, calculated from current date:
```
{YY}年{M.D-M.D}港澳商务渠道成交用户赠送4节课豌豆币.xlsx
```

## Usage

```bash
pip install pandas openpyxl playwright
playwright install chromium

python process_and_oa.py
```

Configure `OA_LOGIN_URL` and `OA_PORTAL_URL` for your environment. Place weekly sales export at `sample/海外用户销售明细_末次渠道.xlsx`.
