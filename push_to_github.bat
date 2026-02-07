@echo off
echo ========================================
echo 推送代码到 GitHub
echo ========================================
echo.

REM 添加远程仓库
git remote add origin https://github.com/chxwy/Y2Tool.git

REM 推送代码
git branch -M main
git push -u origin main

echo.
echo ========================================
echo 推送完成！
echo ========================================
pause
