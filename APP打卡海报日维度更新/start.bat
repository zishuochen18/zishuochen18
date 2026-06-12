@echo off
chcp 65001 >nul
title APP打卡海报日维度更新 - 完整流程
cls

echo =====================================================
echo  APP打卡海报日维度更新 - 完整流程启动器
echo =====================================================
echo.
echo 此脚本将运行完整的端到端流程：
echo   Step 1-4: 从 Sensors 导出数据并计算转换率
echo   Step 5-9: 登录 BizCenter 并更新两个海报组
echo.
echo 需要的准备：
echo   ✓ 网络连接正常
echo   ✓ Python 3.7+ 已安装
echo   ✓ Playwright 已安装 (pip install playwright)
echo   ✓ pandas 已安装 (pip install pandas openpyxl)
echo.
echo =====================================================
echo.
pause

cd /d "%~dp0"
python run_complete_flow.py

echo.
pause
