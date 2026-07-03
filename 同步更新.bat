@echo off
:: 强制使用 UTF-8 编码，防止中文显示乱码
chcp 65001 >nul
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo =================================
echo [1/3] 正在运行 Python 转换数据...
echo =================================
:: 运行 Python
python main.py
if %errorlevel% neq 0 (
    echo ❌ Python 转换出错，请检查 Excel 是否已关闭！
    pause
    exit /b
)

echo.
echo =================================
echo [2/3] 正在同步本地改动 (Git)...
echo =================================
git branch -M main
git add .
:: 检查是否有实际变动，没变动就不 commit
git diff --cached --quiet
if %errorlevel% neq 0 (
    git commit -m "Auto Update: %date% %time%"
) else (
    echo 💡 内容没有变化，跳过提交。
)

echo.
echo =================================
echo [3/3] 正在上传到指定仓库...
echo =================================
:: 这里直接写你要推送的仓库地址
:: 如果你是第一次在这个文件夹推送到这个新地址，建议加上 -u
git push https://github.com/27vyfjt27p-dot/Bolivia.git main

if %errorlevel% neq 0 (
    echo.
    echo ❌ 上传失败！
    echo 请检查网络或 Token 权限。
) else (
    echo.
    echo ✅ 同步成功！已上传至 Bolivia 仓库。
)

echo =================================
pause