# crm-channel-promotion-automation

Playwright 自动化脚本：在 Element UI 风格的 CRM 系统中，一键完成**创建渠道 + 创建推广**两段操作，并输出最终推广链接。

## 依赖

```bash
pip install playwright
playwright install chromium
```

Python 3.10+。

## 配置

```bash
# 复制模板，填入你公司 CRM 的真实参数
cp config.example.json config.local.json
```

`config.local.json` 已被 `.gitignore` 排除，不会进入版本库。

| 字段 | 说明 |
|------|------|
| `crm.home_url` | CRM 平台首页 URL |
| `channel.channel_no` | 渠道编号（搜索用） |
| `channel.business_line` | 业务线名称 |
| `channel.sub_channel_id` | 子渠道 ID |
| `channel.operator_name` | 所属市场/运营姓名 |
| `channel.default_storage` | 默认入库类型 |
| `promotion.page_type` | 推广页面类型（如 H5） |
| `promotion.page_config` | 页面配置选项 |
| `promotion.link_domain` | 自定义链接域名 |
| `promotion.product_id` | 关联商品 ID |
| `promotion.landing_page_id` | 落地页 ID |
| `promotion.wechat_account` | 静默授权公众号名称 |
| `promotion.sms_signature` | 验证码短信签名 |

## 运行

```bash
python crm_create_channel.py
```

1. 输入【渠道名称】（每次唯一，作为渠道和推广的共同名称）
2. 浏览器自动打开 CRM → 扫码登录 → 终端按回车
3. 脚本自动完成创建渠道 13 步 + 创建推广 11 步
4. 终端输出推广链接

## 适配说明

脚本针对 **Element UI** 框架封装了四个通用工具函数：

| 函数 | 用途 |
|------|------|
| `fill_filterable_dropdown` | 可搜索下拉框（键盘输入触发筛选） |
| `fill_input_by_label` | 按 label 文字定位输入框 |
| `fill_input_by_placeholder` | 按 placeholder 关键词定位输入框 |
| `select_dropdown_option` | 普通下拉框选择 |
| `click_button_with_text` | 兼容"确 定"空格风格的按钮点击 |

若 CRM 升级后出现操作步骤卡住，通常只需调整 `config.local.json` 中对应的 ID / 名称，无需改动脚本逻辑。

## 文件结构

```
crm-channel-promotion-automation/
  crm_create_channel.py    # 主脚本
  config.example.json      # 配置模板（可上传）
  config.local.json        # 真实配置（gitignore，不上传）
  README.md
```
