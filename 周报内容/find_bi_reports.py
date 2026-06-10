#!/usr/bin/env python3
"""
BI 报表查找工具
从 report_profiles_v2.json 中根据关键词搜索报表
"""
import sys
import json
from pathlib import Path

BASE_DIR = Path(__file__).parent
SMARTBI_ROOT = BASE_DIR / "smartbi-data-cli-internal-20260526" / "smartbi-data-cli-internal-20260526"
REPORT_PROFILES = SMARTBI_ROOT / "report_profiles_v2(1).json"


def search_reports(keywords: list[str], profiles_path: Path) -> list:
    """在 report_profiles_v2.json 中搜索包含关键词的报表"""
    with open(profiles_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    matches = []
    for name, report in data['reports'].items():
        # 检查报表名称或路径是否包含任意关键词
        path_str = ' / '.join(report['identity'].get('path', []))
        text = f"{name} {path_str}".lower()

        if any(kw.lower() in text for kw in keywords):
            export_info = report.get('export', {})
            matches.append({
                'name': name,
                'id': report['identity'].get('smartbi_report_id', ''),
                'path': report['identity'].get('path', []),
                'path_text': path_str,
                'type': report['identity'].get('smartbi_resource_type', ''),
                'export_supported': export_info.get('supported', False),
                'open_supported': export_info.get('open_supported', False),
                'last_status': export_info.get('last_status', 'unknown'),
                'direct_cli_eligible': export_info.get('direct_cli', {}).get('eligible', False),
            })

    return matches


def print_results(matches: list):
    """打印搜索结果"""
    if not matches:
        print("未找到匹配的报表")
        return

    print(f"\n找到 {len(matches)} 个匹配的报表:\n")
    print("=" * 100)

    for i, m in enumerate(matches, 1):
        print(f"\n[{i}] {m['name']}")
        print(f"    报表ID: {m['id']}")
        print(f"    路径: {m['path_text']}")
        print(f"    类型: {m['type']}")
        print(f"    导出支持: {'[是]' if m['export_supported'] else '[否]'}")
        print(f"    CLI直接导出: {'[是]' if m['direct_cli_eligible'] else '[否]'}")
        print(f"    最近状态: {m['last_status']}")

    print("\n" + "=" * 100)


def main():
    if len(sys.argv) < 2:
        print("用法: python find_bi_reports.py <关键词1> [关键词2] ...")
        print("\n示例:")
        print("  python find_bi_reports.py 港澳 流速")
        print("  python find_bi_reports.py KOL 转化")
        sys.exit(1)

    keywords = sys.argv[1:]
    print(f"搜索关键词: {', '.join(keywords)}")

    if not REPORT_PROFILES.exists():
        print(f"❌ 配置文件不存在: {REPORT_PROFILES}")
        sys.exit(1)

    matches = search_reports(keywords, REPORT_PROFILES)
    print_results(matches)

    # 输出 JSON 格式（供脚本调用）
    if matches:
        print("\nJSON 输出（用于脚本调用）:")
        print(json.dumps(matches, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
