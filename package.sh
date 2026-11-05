#!/usr/bin/env bash
# -*- coding: utf-8 -*-
# 在 macOS 电脑上一键打包免 Python 依赖的独立应用

cd "$(dirname "$0")" || exit 1

echo "========================================================"
echo "    正在为 macOS 打包 [京东商品状态巡检工具]...         "
echo "========================================================"

# 安装 pyinstaller
python3 -m pip install pyinstaller -q

# 清理旧构建
rm -rf build dist

# 打包为 macOS 可执行目录
pyinstaller --noconfirm \
    --onedir \
    --name "京东商品状态巡检工具_Mac" \
    --clean \
    --collect-all DrissionPage \
    --copy-metadata DrissionPage \
    --hidden-import openpyxl \
    main.py

if [ $? -eq 0 ]; then
    cp sample_jd_goods.xlsx "dist/京东商品状态巡检工具_Mac/"
    echo ""
    echo "========================================================"
    echo "[√] Mac 版本打包成功！"
    echo "输出目录: dist/京东商品状态巡检工具_Mac"
    echo "只需将该文件夹压缩为 .zip 发给其他 Mac 用户即可直接运行！"
    echo "========================================================"
else
    echo "[X] 打包失败，请检查报错！"
fi
