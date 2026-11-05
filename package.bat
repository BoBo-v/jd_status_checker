@echo off
chcp 65001 >nul
title 一键打包 Windows 免安装绿色版
echo ========================================================
echo        正在开始打包 [京东商品状态巡检工具]...
echo        打包完成后无需安装 Python，任意电脑解压即用！
echo ========================================================
echo.

:: 清理旧的构建文件
if exist build rd /s /q build
if exist dist rd /s /q dist

:: 使用 PyInstaller 进行打包 (收集 DrissionPage 依赖)
pyinstaller --noconfirm ^
    --onedir ^
    --name "京东商品状态巡检工具" ^
    --clean ^
    --collect-all DrissionPage ^
    --copy-metadata DrissionPage ^
    --hidden-import openpyxl ^
    main.py

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================================
    echo [√] 打包成功！
    echo 生成的独立程序目录位于: dist\京东商品状态巡检工具
    echo.
    echo 复制 sample_jd_goods.xlsx 到输出目录供快速测试...
    copy sample_jd_goods.xlsx "dist\京东商品状态巡检工具\" >nul
    echo.
    echo 你只需把 dist\京东商品状态巡检工具 文件夹压缩为 .zip
    echo 直接发送给其他人，对方解压后双击即可直接使用！
    echo ========================================================
) else (
    echo.
    echo [X] 打包失败，请检查报错信息。
)

pause
