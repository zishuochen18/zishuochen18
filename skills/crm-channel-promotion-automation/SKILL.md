---
name: crm-channel-promotion-automation
description: Use when the user needs to automate CRM channel creation and promotion creation workflows on an Element UI based CRM platform using Playwright browser automation.
---

# CRM Channel & Promotion Automation

This skill automates two sequential operations on an Element UI based CRM system:
1. **Create Channel** — navigate to channel management, fill in channel details, select business line, configure sub-channel, set ratings and storage defaults, and save.
2. **Create Promotion** — navigate to promotion list, configure H5 promotion with page type, link domain, product, platform channel, landing page, WeChat account, SMS signature, and publish.

## Usage

Before running, copy `config.example.json` to `config.local.json` and fill in the real business parameters. The script reads all environment-specific values (CRM URL, IDs, names) from this config file — no hardcoded business data in the script.

```bash
python crm_create_channel.py
```

## Key design decisions

- All sensitive/business-specific values are externalized to `config.local.json` (gitignored).
- Five generic Element UI helper functions handle dropdown selection, label-based input, placeholder-based input, and button clicks — robust against minor DOM changes.
- Login is handled by the operator (QR code scan); the script waits for confirmation before proceeding.
- New tab detection uses `context.expect_page()` to correctly follow CRM flows that open promotion editing in a new browser tab.
