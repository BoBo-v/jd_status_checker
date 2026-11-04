# -*- coding: utf-8 -*-
"""
全局配置模块
"""
import os
from pathlib import Path

# 项目根目录
BASE_DIR = Path(__file__).resolve().parent

# 截图保存目录（未知状态或异常时截图）
SCREENSHOT_DIR = BASE_DIR / "error_screenshots"
SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)

# 浏览器持久化用户数据目录（保留登录态 Cookie、缓存）
BROWSER_USER_DATA_DIR = BASE_DIR / "browser_data"
BROWSER_USER_DATA_DIR.mkdir(parents=True, exist_ok=True)

# 默认浏览器可执行文件备选路径
BROWSER_CANDIDATE_PATHS = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Edge\Application\msedge.exe"),
]

# 移动端伪装 User-Agent (iPhone 15 Safari / Android Chrome)
MOBILE_USER_AGENT = (
    "Mozilla/5.0 (iPhone; CPU iPhone OS 17_5 like Mac OS X) "
    "AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.5 Mobile/15E148 Safari/604.1"
)

# 视口尺寸 (模拟手机屏幕)
VIEWPORT_WIDTH = 390
VIEWPORT_HEIGHT = 844

# 默认是否无头运行 (False 表示弹出浏览器窗口，不易被风控且方便人工过滑块/扫码)
DEFAULT_HEADLESS = False

# 是否默认启用无痕模式（True=无痕模式不留数据; False=持久化保留历史与Cookie）
# 提示：京东对无任何历史Cookie的纯净无痕窗口风控更敏感，建议默认 False
DEFAULT_INCOGNITO = False

# 遇安全验证时，是否在终端暂停并等待人工在浏览器中完成验证后继续 (True=人机协作过验)
AUTO_PAUSE_ON_CAPTCHA = True

# 当触发风控被拦截时，是否自动清理当前会话的 Cookie 与缓存追踪标识
CLEAR_CACHE_ON_RISK = False

# 每次检测之间的随机延时范围（秒）
DELAY_MIN = 3.5
DELAY_MAX = 7.0

# 页面加载超时（秒）
PAGE_LOAD_TIMEOUT = 15.0

# 页面打开后等待 DOM 渲染时间（秒）
PAGE_WAIT_AFTER_LOAD = 2.5

# 失败重试次数
MAX_RETRIES = 2

# 无货/停售是否作为下架处理 (True=算下架填1; False=算在售填0)
OUT_OF_STOCK_AS_DELIST = False

# Excel 自动保存间隔（每处理 N 条数据保存一次，建议 1 即逐条保存防止断电/崩溃丢失）
SAVE_INTERVAL = 1

# Excel 列头模糊匹配关键词
EXCEL_COL_KEYWORDS = {
    "url": ["京东链接（必填）", "京东链接", "商品链接", "商品url", "jd链接", "url", "链接", "item_url"],
    "status": ["是否已下架（1下架）", "是否已下架", "是否下架", "已下架", "下架状态", "下架标记", "状态"],
    "remark": ["检测备注/异常状态", "检测备注", "状态备注", "备注", "异常状态", "巡检备注"],
    "time": ["检测时间", "巡检时间", "更新时间", "核验时间"],
}
