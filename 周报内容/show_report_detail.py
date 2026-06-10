"""读取指定报表的完整画像"""
import json
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
ROOT = Path(__file__).parent
PROFILE_FILE = ROOT / "report_profiles_v2(1).json"
KB_FILE = ROOT / "kb_bi_business_data_map(1).json"


def show_report_detail(name_keyword: str):
    with open(PROFILE_FILE, encoding="utf-8") as f:
        data = json.load(f)

    matched = []
    for name, report in data["reports"].items():
        if name_keyword in name:
            matched.append((name, report))

    if not matched:
        print(f"未找到包含 '{name_keyword}' 的报表")
        return

    for name, report in matched:
        print("=" * 80)
        print(f"📊 报表名：{name}")
        print("=" * 80)
        identity = report.get("identity", {})
        print(f"路径：{' / '.join(identity.get('path', []))}")
        print(f"报表ID：{identity.get('smartbi_report_id', '')}")
        print(f"类型：{identity.get('smartbi_resource_type', '')}")
        print(f"状态：{identity.get('status', '')}")

        # 筛选项
        filters = report.get("filters", [])
        if filters:
            print(f"\n📌 筛选项（{len(filters)} 个）：")
            for f_item in filters:
                label = f_item.get("label", "")
                ftype = f_item.get("control_type", "")
                semantic = f_item.get("semantic", "")
                default = f_item.get("default_value", "")
                required = "*" if f_item.get("required") else " "
                print(f"   {required} [{ftype:10}|{semantic:12}] {label:20} 默认值: {default}")

        # 字段（从 schema 读）
        schema = report.get("schema", {})
        if schema:
            cols = schema.get("columns", [])
            if cols:
                print(f"\n📋 字段（{len(cols)} 列）：")
                for col in cols[:50]:
                    if isinstance(col, dict):
                        print(f"   - {col.get('name', col)}")
                    else:
                        print(f"   - {col}")
                if len(cols) > 50:
                    print(f"   ... 还有 {len(cols) - 50} 列")
        print()


def show_kb_report_detail(name_keyword: str):
    """从 kb_bi_business_data_map 读字段分组"""
    with open(KB_FILE, encoding="utf-8") as f:
        data = json.load(f)

    found = False
    for module in data.get("modules", []):
        for r in module.get("reports", []):
            if name_keyword in r.get("name", ""):
                print(f"\n📚 知识图谱字段分组（来自模块「{module['title']}」）")
                print(f"   报表：{r.get('name')}")
                print(f"   字段总数：{r.get('column_count', 0)}")

                samples = r.get("sample_columns", [])
                if samples:
                    print(f"\n   示例字段（前 {len(samples)}）:")
                    for col in samples:
                        print(f"     - {col}")

                metric_hits = r.get("metric_hits", [])
                if metric_hits:
                    print(f"\n   命中指标：{', '.join(metric_hits)}")

                groups = r.get("field_groups", {})
                if groups:
                    print(f"\n   字段分组：")
                    for g_name, g_cols in groups.items():
                        print(f"     [{g_name}] ({len(g_cols)} 列)")
                        for col in g_cols[:5]:
                            print(f"       · {col}")
                        if len(g_cols) > 5:
                            print(f"       · ... 还有 {len(g_cols) - 5} 列")

                found = True
                break
        if found:
            break


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法：python show_report_detail.py 报表名关键词")
        sys.exit(1)
    kw = sys.argv[1]
    show_report_detail(kw)
    show_kb_report_detail(kw)
