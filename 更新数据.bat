@echo off
chcp 65001 >nul
title 新股数据更新工具

echo.
echo ================================================================
echo                   🚀 新股市场数据更新工具
echo ================================================================
echo.

echo 📅 当前时间: %date% %time%
echo 📂 工作目录: %cd%
echo.

echo 🔍 检查必要文件...
if not exist "ipoinfo.py" (
    echo ❌ 错误: 找不到 ipoinfo.py 文件！
    echo 请确保在正确的目录运行此脚本。
    pause
    exit /b 1
)

if not exist "analyze_excel.py" (
    echo ❌ 错误: 找不到 analyze_excel.py 文件！
    echo 请确保在正确的目录运行此脚本。
    pause
    exit /b 1
)

echo ✅ 文件检查完成
echo.

echo ================================================
echo 第1步: 从Wind获取最新IPO数据
echo ================================================
echo 🔄 正在运行 ipoinfo.py...
echo 请稍候，这可能需要几分钟...
echo.

C:\ProgramData\anaconda3\python.exe ipoinfo.py

if %errorlevel% neq 0 (
    echo.
    echo ❌ ipoinfo.py 运行失败！
    echo 可能的原因：
    echo   - Wind数据库连接问题
    echo   - 网络连接问题
    echo   - Python环境问题
    echo.
    pause
    exit /b 1
)

echo.
echo ✅ 数据获取完成！
echo.

echo ================================================
echo 第2步: 转换数据格式为网站可用JSON
echo ================================================
echo 🔄 正在运行 analyze_excel.py...

C:\ProgramData\anaconda3\python.exe analyze_excel.py

if %errorlevel% neq 0 (
    echo.
    echo ❌ analyze_excel.py 运行失败！
    echo 可能的原因：
    echo   - Excel文件损坏或为空
    echo   - 权限问题
    echo   - Python库缺失
    echo.
    pause
    exit /b 1
)

echo.
echo ================================================
echo           🎉 数据更新完成！
echo ================================================
echo.
echo ✅ 新的Excel文件已生成
echo ✅ JSON数据文件已更新
echo 💡 现在可以刷新网站查看最新数据
echo.
echo 🌐 启动本地服务器？ (Y/N)
set /p choice="请选择: "

if /i "%choice%"=="Y" (
    echo.
    echo 🚀 正在启动服务器...
    C:\ProgramData\anaconda3\python.exe server.py
) else (
    echo.
    echo 💡 手动启动服务器命令:
    echo C:\ProgramData\anaconda3\python.exe server.py
    echo.
)

echo 感谢使用新股数据更新工具！
pause

