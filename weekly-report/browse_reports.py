"""
浏览 619 个 BI 报表的分类概况
帮用户挑选要做哪个周报
"""
import json
import sys
from collections import Counter, defaultdict
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).parent
KB_FILE = ROOT / "kb_bi_business_data_map(1).json"
PROFILE_FILE = ROOT / "report_profiles_v2(1).json"


def show_modules():
    """展示 14 个业务模块的概况"""
    with open(KB_FILE, encoding="utf-8") as f:
        data = json.load(f)

    print("=" * 80)
    print("📊 BI 报表业务模块分布（按知识图谱分类）")
    print("=" * 80)
    for m in data["modules"]:
        report_count = len(m.get("reports", []))
        print(f"\n📁 {m['id']} - {m['title']}")
        print(f"   报表数: {report_count}")
        if m.get("questions"):
            print(f"   核心问题: {' | '.join(m['questions'][:3])}")
        if m.get("metrics"):
            print(f"   关键指标: {', '.join(m['metrics'][:8])}")


def show_top_paths():
    """按一级目录展示报表分布"""
    with open(PROFILE_FILE, encoding="utf-8") as f:
        data = json.load(f)

    top_dirs = Counter()
    second_dirs = defaultdict(Counter)

    for name, report in data["reports"].items():
        path = report.get("identity", {}).get("path", [])
        if path:
            top_dirs[path[0]] += 1
            if len(path) >= 2:
                second_dirs[path[0]][path[1]] += 1

    print("\n" + "=" * 80)
    print(f"📂 619 个报表的目录分布（前 8 个）")
    print("=" * 80)
    for top, count in top_dirs.most_common(8):
        print(f"\n📁 {top}（共 {count} 个报表）")
        for sub, sub_count in second_dirs[top].most_common(8):
            print(f"   ├─ {sub}：{sub_count}")


def search_keyword(kw: str):
    """按关键词搜索报表"""
    with open(PROFILE_FILE, encoding="utf-8") as f:
        data = json.load(f)

    matched = []
    for name, report in data["reports"].items():
        path = report.get("identity", {}).get("path", [])
        path_text = " / ".join(path)
        if kw in name or kw in path_text:
            matched.append((name, path_text, report.get("identity", {}).get("smartbi_resource_type", ""),
                            report.get("identity", {}).get("status", "")))

    print(f"\n🔍 关键词 '{kw}' 命中 {len(matched)} 个报表")
    for name, path, rtype, status in matched[:30]:
        flag = "✅" if status == "pass" else ("❌" if status == "fail" else "❓")
        print(f"  {flag} [{rtype:18}] {path}")
    if len(matched) > 30:
        print(f"  ... 还有 {len(matched) - 30} 个未列出")


if __name__ == "__main__":
    if len(sys.argv) > 1:
        # 关键词搜索
        for kw in sys.argv[1:]:
            search_keyword(kw)
    else:
        # 默认：展示模块和目录分布
        show_modules()
        show_top_paths()
