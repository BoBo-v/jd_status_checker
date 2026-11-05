#!/usr/bin/env bash
# -*- coding: utf-8 -*-
# 京东商品状态巡检系统 - macOS 访达双击专属启动脚本 (.command)

cd "$(dirname "$0")" || exit 1

echo "========================================================"
echo "       京东商品状态自动化巡检系统 (macOS)       "
echo "========================================================"

# 检查 Python3
if ! command -v python3 &> /dev/null; then
    echo "[!] 未检测到 python3，请先前往 https://www.python.org/ 下载安装！"
    exit 1
fi

# 检查依赖项
echo "[*] 正在检查并补齐运行依赖..."
python3 -m pip install -r requirements.txt -q

echo "[*] 正在启动巡检系统..."
echo ""
python3 main.py

echo ""
echo "========================================================"
echo " 任务已结束，按回车键退出..."
echo "========================================================"
read -r
