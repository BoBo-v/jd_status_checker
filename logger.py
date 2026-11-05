# -*- coding: utf-8 -*-
"""
全局轻量日志模块
基于标准库 logging 与 RotatingFileHandler 实现自动轮转日志
单文件上限 5MB，最多保留 3 个备份（总占用永远不超过 15MB，零磁盘膨胀风险）
"""
import sys
import os
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path

# 获取程序所在根目录（兼容 PyInstaller 打包环境）
if getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).resolve().parent
else:
    APP_DIR = Path(__file__).resolve().parent

LOG_FILE = APP_DIR / "checker.log"

# 创建或获取全局 Logger
logger = logging.getLogger("JDStatusChecker")
logger.setLevel(logging.INFO)

# 防止重复添加 Handler
if not logger.handlers:
    # 1. 自动轮转文件 Handler (5MB 单文件上限, 最多保留 3 个备份)
    try:
        file_handler = RotatingFileHandler(
            filename=str(LOG_FILE),
            maxBytes=5 * 1024 * 1024,
            backupCount=3,
            encoding="utf-8"
        )
        file_handler.setLevel(logging.INFO)
        file_formatter = logging.Formatter(
            fmt="[%(asctime)s] [%(levelname)s] %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S"
        )
        file_handler.setFormatter(file_formatter)
        logger.addHandler(file_handler)
    except Exception as e:
        print(f"[!] 初始化日志文件失败: {e}")

    # 2. 控制台 Handler（仅输出 WARNING 及以上，避免控制台信息过载）
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.WARNING)
    console_formatter = logging.Formatter(
        fmt="[%(levelname)s] %(message)s"
    )
    console_handler.setFormatter(console_formatter)
    logger.addHandler(console_handler)


def get_log_file_path() -> Path:
    """返回日志文件的完整绝对路径"""
    return LOG_FILE
