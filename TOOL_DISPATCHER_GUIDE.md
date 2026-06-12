# 商务工具统一调度系统 - 使用指南

## 📋 概述

这是一个统一的工具调度系统，用于管理和运行陈梓烁的 5 个商务自动化工具。系统支持从根目录一键调度，同时保持每个子项目的独立性。

**系统位置**：`c:\Users\chenzishuo\Desktop\新建文件夹 (2)\`

---

## 🛠️ 工具清单

| 工具名 | 类型 | 频率 | 位置 | 入口脚本 |
|--------|------|------|------|---------|
| `agent-settlement` | Tool | 月度 | `代理对账结算/` | `process_settlement.py` |
| `tmk-weekly` | Skill | 周度 | `周报内容/TMK周报/` | `process_tmk.py` |
| `hk-mo-coin` | Tool | 周度 | `港澳赠课豌豆币/` | `process_and_oa.py` |
| `hk-mo-email` | Tool | 周度 | `港澳未分班/` | `process_and_email.py` |
| `crm-channel` | Skill | 按需 | 根目录 | `crm_create_channel.py` |
| `crm-feishu` | Skill | 守护进程 | 根目录 | `feishu_bitable_worker.py` |

---

## 📖 使用方式

### 1️⃣ 查看所有可用工具

```bash
python run_tool.py --list
```

**输出**：列出所有工具及其频率和描述。

### 2️⃣ 查看工具运行状态

```bash
python run_tool.py --status
```

**输出**：显示每个工具的上次运行时间、执行状态、耗时等。

示例：
```
[Status] Tool Status:
====================================================================================================
Tool                 Freq       Last Run             Status          Duration
----------------------------------------------------------------------------------------------------
agent-settlement     monthly    2026-06-12 18:16     [OK]            45s
tmk-weekly           weekly     2026-06-10 09:30     [OK]            12s
hk-mo-coin           weekly     -                    [PENDING]       -
...
```

### 3️⃣ 运行单个工具

#### 基础运行（无参数）

```bash
python run_tool.py tmk-weekly
python run_tool.py hk-mo-email
```

#### 带参数运行

```bash
python run_tool.py agent-settlement Amy 0.35
```

参数会原样透传给子工具的入口脚本。

**日志输出**：自动保存在 `logs/工具名_时间戳.log`

### 4️⃣ 批量运行（按频率）

#### 运行所有周度工具

```bash
python run_tool.py --weekly
```

依次运行：`tmk-weekly` → `hk-mo-coin` → `hk-mo-email`

#### 运行所有月度工具

```bash
python run_tool.py --monthly
```

运行：`agent-settlement`（如需参数会提示跳过）

### 5️⃣ 获取帮助

```bash
python run_tool.py --help
python run_tool.py -h
```

---

## 📂 目录结构

```
c:\Users\chenzishuo\Desktop\新建文件夹 (2)\
├── run_tool.py                    # 调度器主程序
├── tools_registry.json            # 工具注册表（配置文件）
├── run_history.json               # 运行历史（自动生成）
├── logs/                          # 日志目录
│   ├── tmk-weekly_20260612_093000.log
│   ├── agent-settlement_20260612_181600.log
│   └── ...
│
├── 代理对账结算/
│   ├── process_settlement.py      # 入口脚本
│   ├── sample/
│   └── output/
│
├── 周报内容/TMK周报/
│   ├── process_tmk.py
│   ├── sample/
│   └── output/
│
├── 港澳赠课豌豆币/
│   ├── process_and_oa.py
│   ├── sample/
│   └── output/
│
├── 港澳未分班/
│   ├── process_and_email.py
│   ├── sample/
│   └── output/
│
├── crm_create_channel.py          # 根目录工具
├── feishu_bitable_worker.py       # 根目录工具
└── ...
```

---

## 🔍 运行历史追踪

每次运行都会自动记录在 `run_history.json`，包含：

- 上次运行时间
- 执行状态（success/failed/timeout/error）
- 耗时
- 退出码
- 日志文件路径

### 查看历史记录

```bash
cat run_history.json
```

示例：
```json
{
  "agent-settlement": {
    "last_run": "2026-06-12 18:16:00",
    "last_status": "success",
    "duration": "0:00:45.123456",
    "exit_code": 0,
    "log_file": "logs/agent-settlement_20260612_181600.log"
  },
  "tmk-weekly": {
    "last_run": "2026-06-10 09:30:15",
    "last_status": "success",
    "duration": "0:00:12.654321",
    "exit_code": 0,
    "log_file": "logs/tmk-weekly_20260610_093015.log"
  }
}
```

---

## ⚙️ 工具配置（tools_registry.json）

如需修改工具配置（路径、频率、入口脚本等），编辑 `tools_registry.json`：

```json
{
  "agent-settlement": {
    "dir": "代理对账结算",           # 相对于根目录的路径
    "entry": "process_settlement.py", # 入口脚本名
    "freq": "monthly",                # 频率：monthly/weekly/ondemand/daemon
    "args_template": ["{agent_name}", "{cps_rate}"],  # 参数模板
    "description": "代理对账结算（支持 RMB/HKD）"      # 描述
  }
}
```

---

## 🚀 常见场景

### 场景 1：周一早上自动运行所有周度工具

```bash
# Windows 任务计划程序
# 新建任务 → 触发器设为「每周一 09:00」
# 操作：python run_tool.py --weekly
```

### 场景 2：月初自动运行月度结算

```bash
# 假设在 Windows 任务计划程序中配置
# 触发器：每月 1 号 9:00
# 操作：python run_tool.py agent-settlement Amy 0.35
```

### 场景 3：手动调试单个工具

直接 cd 到子目录运行（不使用调度器）：

```bash
cd "代理对账结算"
python process_settlement.py Amy 0.35
```

这完全不受影响，调度器和直接运行可以混用。

### 场景 4：查看最近失败的工具日志

```bash
# 查看运行历史
python run_tool.py --status

# 根据日志文件路径查看详细日志
type logs\agent-settlement_20260612_181600.log
```

---

## ❌ 错误处理

### 工具不存在

```bash
$ python run_tool.py invalid-tool
[ERROR] Unknown tool: invalid-tool

Available tools:
  - agent-settlement
  - tmk-weekly
  - hk-mo-coin
  - hk-mo-email
  - crm-channel
  - crm-feishu
```

### 入口脚本找不到

```bash
[ERROR] Entry script not found: /path/to/script.py
```

**解决**：检查 `tools_registry.json` 中的 `dir` 和 `entry` 配置是否正确。

### 工具执行失败

```bash
[FAIL] agent-settlement: FAILED (Duration: 0:00:05.123456)
```

**查看日志**：
```bash
type logs\agent-settlement_20260612_181600.log
```

---

## 📝 日志管理

### 日志位置

所有工具的日志都保存在 `logs/` 目录，文件名格式：

```
{工具名}_{YYYYMMDD}_{HHMMSS}.log
```

示例：
- `agent-settlement_20260612_181600.log`
- `tmk-weekly_20260610_093015.log`

### 清理旧日志

```bash
# 删除 30 天前的日志（可选）
forfiles /S /D +30 /C "cmd /c del @path"
```

---

## 🔐 安全性注意

1. **配置文件**：`tools_registry.json` 是明文，不要提交敏感信息
2. **日志文件**：`logs/` 可能包含敏感数据（如 cookie、API key），定期清理
3. **run_history.json**：记录完整的命令行参数，避免提交到公开仓库

建议在 `.gitignore` 中添加：
```
run_history.json
logs/
*.log
run_tool.py  # 可选，如果不想版本控制此脚本
```

---

## 🎯 下一步（可选扩展）

目前系统已支持：
- ✅ 单工具运行
- ✅ 批量运行（按频率）
- ✅ 状态查询
- ✅ 自动日志记录

未来可扩展：
- ⏳ Windows 任务计划程序集成
- ⏳ 失败告警（飞书/企微）
- ⏳ Web UI 监控面板
- ⏳ 并发执行支持
- ⏳ 邮件通知

---

## 📞 故障排查

| 问题 | 解决方案 |
|------|--------|
| `ModuleNotFoundError: No module named 'xxx'` | 检查子工具的依赖是否安装（pandas、openpyxl、playwright 等） |
| 日志文件中文乱码 | 确保 Python 版本 >= 3.7，或手动设置 `PYTHONIOENCODING=utf-8` |
| 工具超时（>1小时） | 系统会自动中止并标记为 `timeout`，查看日志确认原因 |
| `run_history.json` 损坏 | 删除该文件，下次运行会自动重建 |

---

**最后更新**：2026-06-12
