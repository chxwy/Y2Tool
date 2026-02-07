@echo off
chcp 65001 >nul
echo ====================================
echo      Y2订单辅助工具启动器
echo ====================================
echo.

echo 正在检查Python环境...
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到Python环境，请先安装Python 3.7+
    echo 下载地址：https://www.python.org/downloads/
    pause
    exit /b 1
)

echo Python环境检查通过
echo.

echo 正在安装依赖包...
pip install -r requirements.txt
if errorlevel 1 (
    echo [错误] 依赖包安装失败
    pause
    exit /b 1
)

echo 依赖包安装完成
echo.

echo 启动Y2订单辅助工具...
echo.
python image_organizer.py

if errorlevel 1 (
    echo.
    echo [错误] 程序运行失败
    pause
)