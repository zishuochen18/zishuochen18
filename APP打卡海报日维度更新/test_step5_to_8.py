"""
完整流程：两个海报组的海报更新
- 第一个海报组：siweidaka
- 第二个海报组：tongyongzhouzhoudaka

对每个海报组执行：Step 6 查询 → Step 7 删除 → Step 8 添加 → Step 8.5 验证 → Step 9 保存
"""
import sys
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding="utf-8")

# 配置
BIZCENTER_URL = "https://bizcenter-h5-cms.61info.cn/#/groupPoster"
POSTER_GROUPS = [
    "siweidaka",
    "tongyongzhouzhoudaka"
]
# 输出目录相对于脚本所在目录
SCRIPT_DIR = Path(__file__).parent
OUTPUT_DIR = SCRIPT_DIR / "output"

# 导入步骤函数
from poster_update import (
    step5_login_bizcenter,
    step6_search_poster_group,
    step7_remove_business_group,
    step8_add_sorted_posters,
    step8_5_verify_posters,
    step9_save
)

def process_poster_group(page, sorted_file, poster_group_code):
    """处理单个海报组：查询 → 删除 → 添加 → 验证 → 保存"""
    print(f"\n{'='*60}")
    print(f"开始处理海报组: {poster_group_code}")
    print(f"{'='*60}\n")

    # Step 6: 查询海报组
    print("[步骤 6] 查询海报组...")
    if step6_search_poster_group(page, poster_group_code):
        print("  ✓ 海报组查询完成，已进入编辑界面\n")
    else:
        print("  ❌ 查询失败\n")
        return False

    # Step 7: 删除业务分组下的旧海报
    print("[步骤 7] 删除旧海报...")
    if step7_remove_business_group(page):
        print("  ✓ 旧海报已删除\n")
    else:
        print("  ❌ 删除失败\n")

    # Step 8: 输入新海报ID
    print("[步骤 8] 输入新海报ID...")
    if step8_add_sorted_posters(page, sorted_file):
        print("  ✓ 新海报ID已输入\n")
    else:
        print("  ❌ 输入失败\n")

    # Step 8.5: 验证所有海报是否都已添加
    print("[步骤 8.5] 验证所有海报是否都已正确添加...")
    if step8_5_verify_posters(page, sorted_file):
        print("  ✓ 所有海报验证通过\n")
    else:
        print("  ❌ 验证失败，请检查海报是否都已添加\n")
        return False

    # Step 9: 保存
    print("[步骤 9] 保存更新...")
    if step9_save(page):
        print("  ✓ 已保存\n")
    else:
        print("  ❌ 保存失败\n")

    print(f"✓ 海报组 {poster_group_code} 处理完成\n")
    return True

def main():
    print("=" * 60)
    print("APP打卡海报日维度更新 - 完整流程")
    print("=" * 60)
    print()

    # 查找排序文件
    sorted_files = list(OUTPUT_DIR.glob("*裂变率排序*.xlsx"))
    if not sorted_files:
        print("❌ 未找到排序文件，请先运行 Step 1-4")
        return

    sorted_file = sorted_files[0]
    print(f"[准备] 使用排序文件: {sorted_file.name}\n")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        try:
            # Step 5: 登录平台中心（仅一次）
            print("[步骤 5] 登录平台中心...")
            if step5_login_bizcenter(page):
                print("  ✓ 平台中心已登录\n")
            else:
                print("  ❌ 登录失败\n")
                return

            # 对每个海报组进行处理
            success_groups = []
            failed_groups = []

            for idx, poster_group_code in enumerate(POSTER_GROUPS):
                print(f"\n[处理进度] 第 {idx+1}/{len(POSTER_GROUPS)} 个海报组\n")

                # 处理这个海报组
                if process_poster_group(page, sorted_file, poster_group_code):
                    success_groups.append(poster_group_code)
                else:
                    failed_groups.append(poster_group_code)

                # 如果不是最后一个海报组，需要回到海报组列表页面
                if idx < len(POSTER_GROUPS) - 1:
                    print(f"[准备] 返回海报组管理页面...")
                    page.goto(BIZCENTER_URL, wait_until="load")
                    page.wait_for_timeout(3000)
                    print(f"✓ 已返回海报组管理页面\n")

            # 打印最终总结
            print("\n" + "=" * 60)
            print("✓ 所有海报组处理完成！")
            print("=" * 60)
            print(f"\n成功处理: {len(success_groups)} 个")
            for group in success_groups:
                print(f"  ✓ {group}")

            if failed_groups:
                print(f"\n失败处理: {len(failed_groups)} 个")
                for group in failed_groups:
                    print(f"  ❌ {group}")

            input("\n按 Enter 关闭浏览器...")

        except Exception as e:
            print(f"\n❌ 错误: {e}")
            import traceback
            traceback.print_exc()
        finally:
            context.close()
            browser.close()

if __name__ == "__main__":
    main()
