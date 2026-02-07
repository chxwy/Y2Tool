@echo off
chcp 65001 >nul
echo ========================================
echo 图片批量整理工具 - 安装包制作工具
echo ========================================
echo.

echo 正在检查必要文件...
if not exist "dist\图片批量整理工具.exe" (
    echo 错误：找不到 dist\图片批量整理工具.exe
    echo 请先运行打包脚本生成可执行文件
    pause
    exit /b 1
)

if not exist "logo.ico" (
    echo 错误：找不到 logo.ico 图标文件
    pause
    exit /b 1
)

if not exist "installer_script.iss" (
    echo 错误：找不到 installer_script.iss 安装脚本
    pause
    exit /b 1
)

echo 文件检查完成！
echo.

echo 请确保已安装 Inno Setup：
echo 1. 下载地址：https://jrsoftware.org/isdl.php
echo 2. 安装后需要下载中文语言包
echo 3. 语言包下载：https://jrsoftware.org/files/istrans/
echo 4. 将 ChineseSimplified.isl 放入 Inno Setup 的 Languages 文件夹
echo.

echo 按任意键继续制作安装包...
pause >nul

echo 正在启动 Inno Setup 编译器...
echo.

:: 尝试找到 Inno Setup 安装路径
set "INNO_PATH="
if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" (
    set "INNO_PATH=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
) else if exist "C:\Program Files\Inno Setup 6\ISCC.exe" (
    set "INNO_PATH=C:\Program Files\Inno Setup 6\ISCC.exe"
) else if exist "C:\Program Files (x86)\Inno Setup 5\ISCC.exe" (
    set "INNO_PATH=C:\Program Files (x86)\Inno Setup 5\ISCC.exe"
) else if exist "C:\Program Files\Inno Setup 5\ISCC.exe" (
    set "INNO_PATH=C:\Program Files\Inno Setup 5\ISCC.exe"
)

if "%INNO_PATH%"=="" (
    echo 未找到 Inno Setup 安装路径！
    echo 请手动运行 Inno Setup，然后打开 installer_script.iss 文件进行编译
    echo.
    echo 或者将 Inno Setup 的 ISCC.exe 路径添加到系统 PATH 环境变量中
    pause
    exit /b 1
)

echo 找到 Inno Setup：%INNO_PATH%
echo 正在编译安装包...
echo.

"%INNO_PATH%" "installer_script.iss"

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo 安装包制作成功！
    echo ========================================
    echo.
    echo 安装包位置：installer_output\图片批量整理工具_安装包.exe
    echo.
    echo 现在可以将安装包发送给朋友了！
    echo 朋友只需要双击安装包，按照向导提示即可完成安装。
    echo.
    
    :: 询问是否打开输出文件夹
    set /p "open_folder=是否打开安装包所在文件夹？(Y/N): "
    if /i "%open_folder%"=="Y" (
        if exist "installer_output" (
            explorer "installer_output"
        )
    )
) else (
    echo.
    echo 安装包制作失败！
    echo 请检查错误信息并重试。
)

echo.
echo 按任意键退出...
pause >nul