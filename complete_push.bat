@echo off
echo ========================================
echo 完成推送
echo ========================================
echo.

echo [1/3] 提交合并...
git add .
git commit -m "Merge remote README"

echo.
echo [2/3] 推送到 GitHub...
git push -u origin main

echo.
echo ========================================
echo 推送完成！
echo ========================================
pause
