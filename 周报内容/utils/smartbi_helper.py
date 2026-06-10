"""
SmartBI 自动下载公共工具库

抽取书展模块的下载逻辑，供 5 个模块共用：
- 港澳流速（共享底表）
- 书展
- KOL
- TMK
- 线下商超
- 转介绍
"""
import sys
import json
import subprocess
import os
import glob
import shutil
from pathlib import Path
from datetime import datetime, timedelta
from calendar import monthrange

sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = Path(__file__).parent.parent
SMARTBI_CLI = BASE_DIR / "smartbi-data-cli-internal-20260526" / "smartbi-data-cli-internal-20260526" / "scripts" / "smartbi_cli.py"


def get_periods():
    """计算上月/本月/昨天的日期窗口（通用，不耦合任何模块）"""
    today = datetime.now().date()
    yesterday = today - timedelta(days=1)

    current_year = today.year
    current_month = today.month

    if current_month == 1:
        last_year = current_year - 1
        last_month = 12
    else:
        last_year = current_year
        last_month = current_month - 1

    last_month_start = datetime(last_year, last_month, 1).date()
    last_month_end = datetime(last_year, last_month, monthrange(last_year, last_month)[1]).date()
    current_month_start = today.replace(day=1)

    return {
        'today': today,
        'yesterday': yesterday,
        'last_month': {
            'year': last_year,
            'month': last_month,
            'year_short': str(last_year)[2:],
            'prefix': f"{str(last_year)[2:]}年{last_month}月",
            'prefix_short': f"{last_month}月",
        },
        'current_month': {
            'year': current_year,
            'month': current_month,
            'year_short': str(current_year)[2:],
            'prefix': f"{str(current_year)[2:]}年{current_month}月",
            'prefix_short': f"{current_month}月",
        },
        'last_month_full': (last_month_start, last_month_end),
        'last_month_to_now': (last_month_start, yesterday),
        'current_month_to_now': (current_month_start, yesterday),
        'is_month_start': today.day == 1,
    }


def assert_smartbi_credentials():
    """检查 SmartBI 凭证环境变量"""
    if not os.environ.get("SMARTBI_USERNAME"):
        raise RuntimeError("缺少环境变量 SMARTBI_USERNAME，请参考 [SmartBI Credentials] 设置")


def render_runtime_config(template_path: Path, period_subs: dict = None) -> Path:
    """
    读取模板配置，渲染日期占位符，输出临时配置文件

    参数:
        template_path: 模板配置文件路径（JSON）
        period_subs: 日期替换映射字典（可选）。如不提供，则用 get_periods() 自动生成默认替换。
                    支持的占位符格式：{{last_month_start}}、{{yesterday}} 等
                    也支持直接修改 JSON 对象中的值（通过嵌套键）

    返回:
        渲染后的临时配置文件路径（放在 ./_runtime/ 下，gitignore）
    """
    if period_subs is None:
        periods = get_periods()
        period_subs = {
            "last_month_start": periods['last_month_full'][0].strftime('%Y-%m-%d'),
            "last_month_end": periods['last_month_full'][1].strftime('%Y-%m-%d'),
            "yesterday": periods['yesterday'].strftime('%Y-%m-%d'),
        }

    with open(template_path, 'r', encoding='utf-8') as f:
        config = json.load(f)

    def replace_placeholders(obj):
        """递归替换 JSON 中的占位符"""
        if isinstance(obj, dict):
            for k, v in obj.items():
                if isinstance(v, str):
                    for placeholder, value in period_subs.items():
                        pattern = "{{" + placeholder + "}}"
                        obj[k] = v.replace(pattern, str(value))
                        v = obj[k]
                else:
                    replace_placeholders(v)
        elif isinstance(obj, list):
            for item in obj:
                replace_placeholders(item)

    replace_placeholders(config)

    runtime_dir = template_path.parent / "_runtime"
    runtime_dir.mkdir(exist_ok=True)

    output_path = runtime_dir / (template_path.stem + "_runtime.json")
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)

    return output_path


def run_smartbi_cli_task(config_path: Path, task_key: str, output_dir: Path = None, max_retries: int = 1) -> Path:
    """
    调用 SmartBI CLI 下载单个 task

    参数:
        config_path: 渲染后的运行时配置文件路径
        task_key: task 的 key（如 "bookfair_may_full"）
        output_dir: 输出目录（可选，用于验证）
        max_retries: 重试次数

    返回:
        下载产出的文件路径（第一个发现的 xlsx）

    异常:
        RuntimeError: CLI 执行失败或找不到产出文件
    """
    username = os.environ.get("SMARTBI_USERNAME")
    password = os.environ.get("SMARTBI_PASSWORD", "")

    for attempt in range(max_retries + 1):
        try:
            print(f"  [调用 SmartBI CLI] task={task_key}, 尝试 {attempt+1}/{max_retries+1}")

            cmd = [
                sys.executable,
                str(SMARTBI_CLI),
                "run",
                "--username", username,
                "--password", password,
                "--config", str(config_path),
                "--task", task_key,
                "--overwrite",
            ]

            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=300,
            )

            if result.returncode != 0:
                err_msg = result.stderr or result.stdout or "未知错误"
                if attempt < max_retries:
                    print(f"    ⚠️ CLI 返回非零状态码 {result.returncode}，准备重试...")
                    continue
                else:
                    raise RuntimeError(f"SmartBI CLI 执行失败 (task={task_key}): {err_msg}")

            # 查找产出文件
            if output_dir:
                pattern = str(output_dir / "*.xlsx")
            else:
                pattern = "*.xlsx"

            files = glob.glob(pattern, recursive=False)
            if not files:
                if attempt < max_retries:
                    print(f"    ⚠️ 未找到产出文件，准备重试...")
                    continue
                else:
                    raise RuntimeError(f"SmartBI CLI 下载后未找到产出文件 (task={task_key}, pattern={pattern})")

            downloaded_file = max(files, key=os.path.getmtime)
            print(f"    ✓ 下载成功: {downloaded_file}")
            return Path(downloaded_file)

        except (subprocess.TimeoutExpired, Exception) as e:
            if attempt < max_retries and isinstance(e, subprocess.TimeoutExpired):
                print(f"    ⚠️ CLI 超时，准备重试...")
                continue
            else:
                raise RuntimeError(f"SmartBI CLI 失败 (task={task_key}): {str(e)}")

    raise RuntimeError(f"SmartBI CLI 达到最大重试次数 (task={task_key})")


def rename_downloaded_file(src: Path, target_name: str) -> Path:
    """
    重命名下载的文件，避免被后续 task 覆盖

    参数:
        src: 原文件路径
        target_name: 目标文件名

    返回:
        重命名后的文件路径
    """
    target = src.parent / target_name
    if src != target:
        shutil.move(str(src), str(target))
    return target


def run_smartbi_browser_task(
    report_id: str,
    filters: list = None,
    output_path: Path = None,
    *,
    max_rows: int = 10000,
    headless: bool = True,
) -> Path:
    """
    走 Playwright 浏览器下载 SIMPLE_REPORT

    参数:
        report_id: SmartBI 报表 ID
        filters: 过滤器列表 [(key, value, displayValue), ...]，key 支持 alias（中文）或 name（英文）
        output_path: 输出文件路径
        max_rows: 最大行数（默认 100000）
        headless: 是否无头模式（默认 True）

    返回:
        输出文件路径

    异常:
        RuntimeError: 浏览器导出失败
    """
    import asyncio
    import sys as _sys

    assert_smartbi_credentials()

    if not output_path:
        raise ValueError("output_path 必须指定")

    if filters is None:
        filters = []

    username = os.environ.get("SMARTBI_USERNAME")
    password = os.environ.get("SMARTBI_PASSWORD", "")
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        smartbi_scripts = BASE_DIR / "smartbi-data-cli-internal-20260526" / "smartbi-data-cli-internal-20260526" / "scripts"
        if str(smartbi_scripts) not in sys.path:
            sys.path.insert(0, str(smartbi_scripts))

        from smartbi_browser_export import export_simple_report_with_browser

        print(f"  [浏览器导出] report_id={report_id}, filters={len(filters)} 条")
        result = asyncio.run(export_simple_report_with_browser(
            username=username,
            password=password,
            report_id=report_id,
            output_path=output_path,
            max_rows=max_rows,
            filters=filters,
            headless=headless,
        ))

        # 诊断：打印 SmartBI 返回的结果字段
        row_count = result.get("rowCount", "?")
        applied = result.get("applied", [])
        panel_values = result.get("panelValues", {})

        print(f"    [SmartBI 诊断] rowCount={row_count}, applied={len(applied)} 个 filter")
        if applied:
            for f in applied:
                print(f"      ├─ {f.get('alias', '?')} = {f.get('value', '')}")

        if row_count == 0:
            print(f"    [面板参数实际值] (检查哪些默认值在限制数据)")
            import json as _json
            panel_str = _json.dumps(panel_values, ensure_ascii=False, indent=6)
            for line in panel_str.split('\n')[:50]:  # 只打前50行避免太长
                print(f"      {line}")
            if len(panel_values) > 50:
                print(f"      ... (共 {len(panel_values)} 个参数)")
            raise RuntimeError(
                f"SmartBI 下载到 0 行（被默认 filter 过滤掉）。"
                f"请对照上面的 panelValues 找出需要清空的默认参数（如学员区域/报名时间等）"
            )

        if not output_path.exists() or output_path.stat().st_size < 1024:
            raise RuntimeError(f"浏览器导出未生成有效文件: {output_path}")

        file_size_kb = output_path.stat().st_size / 1024
        print(f"    ✓ 下载成功: {output_path.name} ({file_size_kb:.1f} KB, {row_count} 行)")
        return output_path

    except Exception as e:
        raise RuntimeError(f"SmartBI 浏览器导出失败 (report_id={report_id}): {str(e)}")


def verify_file_freshness(file_path: Path, expected_date: datetime, tolerance_minutes: int = 15) -> None:
    """
    校验文件新鲜度（mtime + 可读性）

    参数:
        file_path: 文件路径
        expected_date: 期望快照日期
        tolerance_minutes: 容差分钟数（允许 mtime 偏离 ±tolerance_minutes）

    异常:
        FileNotFoundError: 文件不存在
        ValueError: 文件 mtime 偏离过大或不可读
    """
    if not file_path.exists():
        raise FileNotFoundError(f"文件不存在: {file_path}")

    file_size = file_path.stat().st_size
    if file_size < 1024:  # 小于 1KB 认为无效
        raise ValueError(f"文件过小（{file_size} bytes）: {file_path}")

    mtime = datetime.fromtimestamp(file_path.stat().st_mtime)
    expected_datetime = datetime.combine(expected_date, datetime.min.time())
    time_diff = abs((mtime - expected_datetime).total_seconds() / 60)

    if time_diff > tolerance_minutes:
        print(f"    ⚠️ 文件 mtime 偏离：期望 {expected_date}，实际 {mtime.date()}（差异 {time_diff:.1f} 分钟）")

    try:
        import pandas as pd
        df = pd.read_excel(file_path, nrows=5)
        if df.empty:
            raise ValueError("Excel 内容为空")
    except Exception as e:
        raise ValueError(f"Excel 文件不可读: {e}")
