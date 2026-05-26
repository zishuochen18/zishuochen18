# 港澳渠道成交用户赠送豌豆币 OA 申请自动化

每周自动处理销售明细，生成豌豆币发放表格，并提交到 OA 系统暂存草稿。

## 功能

1. **Excel 表格处理**：读取销售明细，提取学员ID + 首签订单号，按豌豆币模板格式生成发放表
2. **OA 表单填写**：登录公司 OA → 进入豌豆币添加申请 → 自动填写 13 个表单字段 → 上传附件 → 暂存草稿

## 依赖

```
pip install pandas openpyxl playwright
playwright install chromium
```

## 使用方式

1. 将本周的销售明细 Excel 放到 `sample/海外用户销售明细_末次渠道.xlsx`
2. 修改脚本顶部的 `OA_LOGIN_URL` 和 `OA_PORTAL_URL` 为公司 OA 地址
3. 运行：

```bash
cd 港澳赠课豌豆币
python process_and_oa.py
```

首次运行需要在弹出的浏览器中完成 SSO 登录，登录状态会缓存到 `oa_auth_state.json`。

## 表格处理规则

| 列 | 来源 |
|---|---|
| 用户ID | 销售明细 B 列（学员ID） |
| 学科品类 | 固定 `VIP_WanDou` |
| 发放数量 | 固定 32000 |
| 关联订单号 | 销售明细 DH 列（首签订单号） |
| 其他字段 | 固定值 |

输出文件名：`{年}年{上周日期范围}港澳商务渠道成交用户赠送4节课豌豆币.xlsx`

A 列（学员ID）和 I 列（订单号）会强制设为文本格式，避免 Excel 自动转科学计数法。

## OA 表单字段映射

脚本针对兰的 (Landray) OA 系统设计，字段通过 `name="extendDataFormInfo.value(fd_xxx)"` 模式定位。三种字段类型分别处理：

| 字段类型 | 处理方式 |
|---|---|
| 普通文本/textarea | `fill_text(field_id, value)` |
| 单选按钮 (radio) | `check_radio(field_id, value)` |
| 地址选择器 (manifest) | 点击容器 → 输入 → 键盘 ArrowDown+Enter |

## 注意事项

- 文档名包含上周日期范围，每周一运行时会自动生成
- 销售明细 Excel 表头在第 8 行（`SALES_HEADER_ROW = 7`）
- OA 表单的字段 ID（`fd_xxx`）与具体 OA 模板绑定，模板变更时需要重新探测
- "积分成本归属部门" 用 jQuery manifest 地址选择器，搜索结果用键盘 ArrowDown + Enter 选择最稳
