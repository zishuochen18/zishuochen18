@echo off
REM 港澳流速报表测试导出脚本
REM 使用前请先设置环境变量：
REM   set SMARTBI_USERNAME=您的用户名
REM   set SMARTBI_PASSWORD=您的密码

echo ========================================
echo 港澳流速报表测试导出
echo ========================================
echo.

REM 检查环境变量
if "%SMARTBI_USERNAME%"=="" (
    echo [错误] 未设置 SMARTBI_USERNAME 环境变量
    echo 请运行: set SMARTBI_USERNAME=您的用户名
    exit /b 1
)

if "%SMARTBI_PASSWORD%"=="" (
    echo [错误] 未设置 SMARTBI_PASSWORD 环境变量
    echo 请运行: set SMARTBI_PASSWORD=您的密码
    exit /b 1
)

echo [信息] 凭据已设置
echo [信息] 用户名: %SMARTBI_USERNAME%
echo.

cd smartbi-data-cli-internal-20260526\smartbi-data-cli-internal-20260526

echo [步骤1] 执行 dry-run 检查配置...
python scripts\smartbi_cli.py run --config ..\..\configs\test_hk_flow.json --task hk_flow_test --dry-run --json
echo.

echo [步骤2] 开始实际导出...
python scripts\smartbi_cli.py run --config ..\..\configs\test_hk_flow.json --task hk_flow_test --json

echo.
echo ========================================
echo 导出完成
echo 请检查 smartbi-data-cli-internal-20260526\smartbi-data-cli-internal-20260526\outputs\test_export 目录
echo ========================================
