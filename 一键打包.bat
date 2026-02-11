@echo off
chcp 65001 >nul
title PDF拆分工具 - 打包脚本

echo ============================================
echo    PDF拆分工具 - 一键打包为独立EXE程序
echo ============================================
echo.

echo [1/2] 正在安装所需的Python库...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo.
    echo 安装失败！请确认已安装Python并且pip可用。
    pause
    exit /b 1
)

echo.
echo [2/2] 正在打包为EXE文件，请稍候...
pyinstaller --onefile --windowed --name "PDF拆分工具" pdf_splitter.py
if %errorlevel% neq 0 (
    echo.
    echo 打包失败！
    pause
    exit /b 1
)

echo.
echo ============================================
echo    打包成功！
echo    EXE文件位于: dist\PDF拆分工具.exe
echo    把这个EXE文件复制给别人就能直接使用！
echo ============================================
echo.
pause
