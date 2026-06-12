"""
验证脚本：检查完整流程是否可以正确执行
- 检查所有必需的函数是否存在
- 检查文件和目录结构
- 检查依赖包是否已安装
"""
import sys
from pathlib import Path

print("\n" + "="*70)
print("APP打卡海报日维度更新 - 流程验证")
print("="*70 + "\n")

# 获取脚本所在目录
SCRIPT_DIR = Path(__file__).parent

# 1. 检查必需的文件
print("[1/5] 检查必需的文件...")
required_files = [
    "run_complete_flow.py",
    "test_step5_to_8.py",
    "poster_update.py",
    "README_完整流程.md",
    "IMPLEMENTATION_SUMMARY.md",
    "start.bat",
]

missing_files = []
for file in required_files:
    file_path = SCRIPT_DIR / file
    if file_path.exists():
        print(f"  [OK] {file}")
    else:
        print(f"  [NG] {file} - 未找到")
        missing_files.append(file)

if missing_files:
    print(f"\n[WARNING] 缺少 {len(missing_files)} 个文件")
else:
    print(f"[OK] 所有必需文件都存在\n")

# 2. 检查目录结构
print("[2/5] 检查目录结构...")
required_dirs = ["output", "data"]
missing_dirs = []

for dir_name in required_dirs:
    dir_path = SCRIPT_DIR / dir_name
    if dir_path.exists():
        print(f"  [OK] {dir_name}/")
    else:
        print(f"  [INFO] {dir_name}/ - 不存在（将在运行时创建）")
        missing_dirs.append(dir_name)

print(f"[OK] 目录结构检查完成\n")

# 3. 检查依赖包
print("[3/5] 检查依赖包...")
required_packages = [
    ("playwright", "Playwright 浏览器自动化"),
    ("pandas", "Pandas 数据处理"),
    ("openpyxl", "Excel 文件处理"),
]

missing_packages = []
for package_name, description in required_packages:
    try:
        __import__(package_name)
        print(f"  [OK] {package_name} ({description})")
    except ImportError:
        print(f"  [NG] {package_name} - 未安装")
        missing_packages.append(package_name)

if missing_packages:
    print(f"\n[WARNING] 缺少 {len(missing_packages)} 个依赖包")
    print("\n安装缺少的包：")
    for pkg in missing_packages:
        print(f"  pip install {pkg}")
else:
    print(f"[OK] 所有依赖包都已安装\n")

# 4. 检查函数
print("[4/5] 检查所有步骤函数...")
sys.path.insert(0, str(SCRIPT_DIR))

required_functions = [
    ("step1_login_sensors", "登录 Sensors"),
    ("step2_enter_bookmark", "进入书签"),
    ("step3_export_data", "导出数据"),
    ("step4_calculate_conversion_rate", "计算转换率"),
    ("step5_login_bizcenter", "登录 BizCenter"),
    ("step6_search_poster_group", "查询海报组"),
    ("step7_remove_business_group", "删除旧海报"),
    ("step8_add_sorted_posters", "添加新海报"),
    ("step8_5_verify_posters", "验证海报"),
    ("step9_save", "保存更新"),
]

missing_functions = []
try:
    import poster_update
    for func_name, description in required_functions:
        if hasattr(poster_update, func_name):
            print(f"  [OK] {func_name} ({description})")
        else:
            print(f"  [NG] {func_name} - 未找到")
            missing_functions.append(func_name)
except Exception as e:
    print(f"  [NG] 导入 poster_update 失败: {e}")

if missing_functions:
    print(f"\n[WARNING] 缺少 {len(missing_functions)} 个函数")
else:
    print(f"[OK] 所有步骤函数都存在\n")

# 5. 检查配置
print("[5/5] 检查配置...")
try:
    with open(SCRIPT_DIR / "run_complete_flow.py", "r", encoding="utf-8") as f:
        content = f.read()

    # 检查海报组配置
    if "siweidaka" in content and "tongyongzhouzhoudaka" in content:
        print(f"  [OK] 海报组配置: siweidaka, tongyongzhouzhoudaka")
    else:
        print(f"  [WARNING] 海报组配置可能不完整")

    # 检查 URL 配置
    if "bizcenter-h5-cms.61info.cn" in content:
        print(f"  [OK] BizCenter URL 配置正确")
    else:
        print(f"  [WARNING] BizCenter URL 配置可能有问题")

    print(f"[OK] 配置检查完成\n")
except Exception as e:
    print(f"  [NG] 配置检查失败: {e}\n")

# 总结
print("\n" + "="*70)
print("验证结果")
print("="*70 + "\n")

total_issues = len(missing_files) + len(missing_packages) + len(missing_functions)

if total_issues == 0:
    print("[SUCCESS] 所有检查通过！系统已准备就绪")
    print("\n可以运行完整流程：")
    print("  1. 双击 start.bat，或")
    print("  2. python run_complete_flow.py")
    print("\n" + "="*70)
    sys.exit(0)
else:
    print(f"[ERROR] 发现 {total_issues} 个问题，需要修复后才能运行\n")

    if missing_files:
        print(f"缺少文件: {missing_files}")
    if missing_packages:
        print(f"\n缺少包，请运行: pip install {' '.join(missing_packages)}")
    if missing_functions:
        print(f"缺少函数: {missing_functions}")

    print("\n" + "="*70)
    sys.exit(1)
