"""
完整端到端流程：从 Step 1 到 Step 9
========================================
Step 1-4: 数据导出和计算（Sensors → 导出 → 转换率计算）
Step 5-9: 海报更新（登录 → 查询 → 删除 → 添加 → 验证 → 保存）
         对两个海报组重复执行 Step 6-9

整个流程自动化，仅需手动：
- Step 1: 扫码登录 Sensors
- Step 5: 扫码登录 BizCenter
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

# 脚本所在目录
SCRIPT_DIR = Path(__file__).parent
SAMPLE_DIR = SCRIPT_DIR / "sample"

# 导入所有步骤函数
from poster_update import (
    step1_login_sensors,
    step2_enter_bookmark,
    step3_export_data,
    step4_calculate_conversion_rate,
    step5_login_bizcenter,
    step6_search_poster_group,
    step7_remove_business_group,
    step8_add_sorted_posters,
    step8_5_verify_posters,
    step9_save
)


def run_step1_to_4():
    """运行 Step 1-4：数据导出和计算

    返回排序后的 Excel 文件路径，失败返回 None
    """
    print("\n" + "="*70)
    print("数据导出阶段 (Step 1-4)")
    print("="*70 + "\n")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        try:
            # Step 1: 登录 Sensors
            print("[步骤 1] 登录 Sensors...")
            if step1_login_sensors(page):
                print("  ✓ Sensors 已登录\n")
            else:
                print("  ❌ 登录失败\n")
                return None

            # Step 2: 进入书签
            print("[步骤 2] 进入指定书签...")
            if step2_enter_bookmark(page):
                print("  ✓ 书签已进入\n")
            else:
                print("  ❌ 进入失败\n")
                return None

            # Step 3: 选择日期【过去 14 天】+ 等待 5 秒缓存 + 导出数据
            print("[步骤 3] 选择日期并导出数据...")
            if step3_export_data(page, context):
                print("  ✓ 数据已导出\n")
            else:
                print("  ❌ 导出失败\n")
                return None

            # Step 4: 计算转换率
            # 找最新导出的 Excel 文件
            print("[步骤 4] 计算裂变率并生成排序文件...")
            excel_files = list(SAMPLE_DIR.glob("*SensorsAnalytics.xlsx"))
            if not excel_files:
                print("  ❌ 未找到导出的 Excel 文件\n")
                return None

            latest_excel = max(excel_files, key=lambda p: p.stat().st_mtime)
            sorted_file = step4_calculate_conversion_rate(latest_excel)
            if sorted_file:
                print(f"  ✓ 排序文件已生成: {sorted_file.name}\n")
                return sorted_file
            else:
                print("  ❌ 计算失败\n")
                return None

        except Exception as e:
            print(f"\n❌ 错误: {e}")
            import traceback
            traceback.print_exc()
            return None
        finally:
            context.close()
            browser.close()


def process_poster_group(page, sorted_file, poster_group_code):
    """处理单个海报组：查询 → 删除 → 添加 → 验证 → 保存"""
    print(f"\n{'='*70}")
    print(f"开始处理海报组: {poster_group_code}")
    print(f"{'='*70}\n")

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

    # Step 8: 逐个添加新海报（每个海报一个对话框循环）
    print("[步骤 8] 按裂变率逐个添加海报...")
    if step8_add_sorted_posters(page, sorted_file):
        print("  ✓ 新海报已全部添加\n")
    else:
        print("  ⚠️ 部分海报添加失败\n")

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


def run_step5_to_9(sorted_file):
    """运行 Step 5-9：海报更新（两个海报组）"""
    print("\n" + "="*70)
    print("海报更新阶段 (Step 5-9)")
    print("="*70 + "\n")

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
                return False

            # 对每个海报组进行处理
            success_groups = []
            failed_groups = []

            for idx, poster_group_code in enumerate(POSTER_GROUPS):
                print(f"\n[处理进度] 第 {idx+1}/{len(POSTER_GROUPS)} 个海报组\n")

                if process_poster_group(page, sorted_file, poster_group_code):
                    success_groups.append(poster_group_code)
                else:
                    failed_groups.append(poster_group_code)

                # 如果不是最后一个海报组，返回海报组列表页面
                if idx < len(POSTER_GROUPS) - 1:
                    print(f"[准备] 返回海报组管理页面...")
                    page.goto(BIZCENTER_URL, wait_until="load")
                    page.wait_for_timeout(3000)
                    print(f"  ✓ 已返回海报组管理页面\n")

            # 打印最终总结
            print("\n" + "="*70)
            print("海报更新阶段完成")
            print("="*70)
            print(f"\n成功: {len(success_groups)} 个")
            for group in success_groups:
                print(f"  ✓ {group}")

            if failed_groups:
                print(f"\n失败: {len(failed_groups)} 个")
                for group in failed_groups:
                    print(f"  ❌ {group}")

            return len(failed_groups) == 0

        except Exception as e:
            print(f"\n❌ 错误: {e}")
            import traceback
            traceback.print_exc()
            return False
        finally:
            context.close()
            browser.close()


def main():
    print("\n" + "="*70)
    print("APP打卡海报日维度更新 - 完整端到端流程")
    print("="*70)
    print("\n此流程将自动执行：")
    print("  1. Step 1-4: 从 Sensors 导出数据并计算转换率")
    print("     - Step 3 自动选择【过去 14 天】+ 等待 5 秒缓存")
    print("  2. Step 5-9: 登录 BizCenter 并更新两个海报组的海报")
    print("     - Step 8 逐个海报打开对话框选择确认")
    print("     - Step 8.5 保存前验证所有海报")
    print("\n需要手动操作：")
    print("  - Step 1: 扫码登录 Sensors（会自动打开浏览器）")
    print("  - Step 5: 扫码登录 BizCenter（会自动打开浏览器）")
    print("\n" + "="*70 + "\n")

    input("按 Enter 开始完整流程...")

    # ============ Step 1-4: 数据导出 ============
    sorted_file = run_step1_to_4()
    if not sorted_file:
        print("\n❌ Step 1-4 失败，无法继续")
        input("\n按 Enter 退出...")
        return

    print(f"\n✓ 数据导出完成，排序文件: {sorted_file.name}")

    # ============ Step 5-9: 海报更新 ============
    input("\n按 Enter 开始海报更新阶段...")
    success = run_step5_to_9(sorted_file)

    # ============ 最终总结 ============
    print("\n" + "="*70)
    print("完整流程结束")
    print("="*70)
    if success:
        print("\n✓ 所有步骤成功完成！")
    else:
        print("\n⚠️ 部分步骤失败，请查看上面的错误信息")

    input("\n按 Enter 退出...")


if __name__ == "__main__":
    main()
