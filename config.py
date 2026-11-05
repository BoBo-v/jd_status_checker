# -*- coding: utf-8 -*-
"""
全局配置模块
支持 Windows 与 macOS (Mac) 跨平台运行
"""
import os
import sys
from pathlib import Path

# 项目根目录（兼容源码运行与 PyInstaller 打包环境）
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent

# 是否开启异常截图（默认 False 关闭：全面使用文本日志记录，彻底避免占用磁盘）
ENABLE_SCREENSHOT = False

# 截图保存目录与生命周期管理（仅在手动开启截图时生效，防膨胀双保险）
SCREENSHOT_DIR = BASE_DIR / "error_screenshots"
SCREENSHOT_MAX_DAYS = 3       # 自动清理超过 3 天的历史截图
SCREENSHOT_MAX_COUNT = 50     # 最多保留最近 50 张截图

# 浏览器持久化用户数据目录（保留登录态 Cookie、缓存）
BROWSER_USER_DATA_DIR = BASE_DIR / "browser_data"
BROWSER_USER_DATA_DIR.mkdir(parents=True, exist_ok=True)

# 默认浏览器可执行文件备选路径 (兼容 Windows 与 Mac)
BROWSER_CANDIDATE_PATHS = [
    # Windows 路径
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Edge\Application\msedge.exe"),
    # Mac (macOS) 路径
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
    "/Applications/Chromium.app/Contents/MacOS/Chromium",
]

# 核心安全与大表保护模式：
# True: 原表 100% 只读（秒级流式加载，哪怕 250MB 大表也不卡顿，100% 绝对不会删除/破坏原表的商品图片！）
#       结果极速保存到独立的【原表名_巡检结果.xlsx】中，支持断点续传。
# False: 原表直接覆写（适用于几百KB没有图片的轻量小表）。
SAFE_EXPORT_MODE = True

# 是否强制直连国内网络（绕过本机 Clash / VPN 代理，防止被境外机房IP触发必死风控）
BYPASS_PROXY = True

# 核心防封逻辑：当检测到页面跳转提示登录时，是否自动关闭整个浏览器重新开启并重试
AUTO_RESTART_ON_LOGIN = True

# 登录拦截自动重启浏览器的最大重试次数
MAX_LOGIN_RESTART_RETRIES = 2

# 视口尺寸 (适合移动 H5 排版)
VIEWPORT_WIDTH = 414
VIEWPORT_HEIGHT = 896

# 默认是否无头运行 (False 表示弹出浏览器窗口)
DEFAULT_HEADLESS = False

# 遇安全验证时，是否在终端暂停并等待人工在浏览器中完成验证后继续
AUTO_PAUSE_ON_CAPTCHA = False

# 每次检测之间的随机延时范围（秒）
DELAY_MIN = 2.5
DELAY_MAX = 5.0

# 页面加载超时（秒）
PAGE_LOAD_TIMEOUT = 15.0

# 页面打开后等待 DOM 渲染时间（秒）
PAGE_WAIT_AFTER_LOAD = 2.0

# 失败重试次数
MAX_RETRIES = 2

# 无货/停售是否作为下架处理 (True=算下架填1; False=算在售填0)
OUT_OF_STOCK_AS_DELIST = False

# Excel 自动保存间隔（每处理 N 条数据保存一次，建议 1 即逐条保存防止断电/崩溃丢失）
SAVE_INTERVAL = 1

# Excel 列头模糊匹配关键词
EXCEL_COL_KEYWORDS = {
    "url": ["京东链接（必填）", "京东链接", "商品链接", "商品url", "jd链接", "url", "链接", "item_url", "sku", "商品sku"],
    "name": ["商品名称", "商品标题", "品名", "名称", "标题", "title", "name"],
    "status": ["是否已下架（1下架）", "是否已下架", "是否下架", "已下架", "下架状态", "下架标记", "状态"],
    "remark": ["检测备注/异常状态", "检测备注", "状态备注", "备注", "异常状态", "巡检备注"],
    "time": ["检测时间", "巡检时间", "更新时间", "核验时间"],
}
