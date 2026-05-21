# crm-channel-promotion-automation

Playwright 自动化脚本：在 Element UI 风格的 CRM 系统中，一键完成**创建渠道 + 创建推广 + 创建推广二维码**三段操作，并输出推广链接和二维码文件。

支持与**飞书多维表格**集成：在表格中填入渠道名称，后台自动执行全流程并回填结果。

## 依赖

```bash
pip install playwright requests
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
| `promotion.link_domain` | 自定义链接域名 |
| `promotion.product_id` | 关联商品 ID |
| `promotion.landing_page_id` | 落地页 ID |
| `promotion.wechat_account` | 静默授权公众号名称 |
| `promotion.sms_signature` | 验证码短信签名 |
| `qrcode.workbench_url` | 工作台 URL |
| `qrcode.short_link_group` | 短链分组名称 |
| `qrcode.short_link_domain` | 短链域名 |

## 运行（交互式）

```bash
python crm_create_channel.py
```

1. 输入【渠道名称】
2. 浏览器自动打开 CRM → 首次需扫码登录（后续自动复用登录状态）
3. 脚本自动完成：创建渠道 → 创建推广 → 创建二维码
4. 终端输出推广链接，二维码保存到脚本目录

## 运行（飞书多维表格集成）

```bash
# 1. 配置飞书应用凭证
cp feishu_config.example.json feishu_config.local.json

# 2. 启动 worker（30秒轮询一次）
python feishu_bitable_worker.py
```

### 飞书配置要求

1. 在飞书开放平台创建自建应用，获取 `app_id` / `app_secret`
2. 应用权限：`bitable:app`（多维表格读写）+ `drive:drive`（上传附件）
3. 将应用添加为多维表格的**协作者**（可编辑权限）
4. `app_token` 需使用飞书 API 返回的**真实 token**（非 URL 中的短链 token）

### 多维表格结构

| 列名 | 字段类型 | 说明 |
|------|----------|------|
| 渠道名称 | 文本 | 输入：用户填写 |
| 推广链接 | 文本/URL | 输出：脚本回填 |
| 二维码图片 | 附件 | 输出：脚本上传 JPG |
| 执行状态 | 单选 | 待执行/执行中/成功/失败 |

## 适配说明

脚本针对 **Element UI** 框架封装了通用工具函数：

| 函数 | 用途 |
|------|------|
| `fill_filterable_dropdown` | 可搜索下拉框（键盘输入触发筛选） |
| `fill_input_by_label` | 按 label 文字定位输入框 |
| `fill_input_by_placeholder` | 按 placeholder 关键词定位输入框 |
| `select_dropdown_option` | 普通下拉框选择 |
| `click_button_with_text` | 兼容"确 定"空格风格的按钮点击 |

## 文件结构

```
crm-channel-promotion-automation/
  crm_create_channel.py        # 主脚本（三段自动化 + 登录状态管理）
  feishu_bitable_worker.py     # 飞书多维表格轮询 worker
  config.example.json          # CRM 配置模板
  feishu_config.example.json   # 飞书配置模板
  config.local.json            # 真实 CRM 配置（gitignore）
  feishu_config.local.json     # 真实飞书配置（gitignore）
  auth_state.json              # 登录状态缓存（gitignore）
  README.md
```
