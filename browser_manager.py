# -*- coding: utf-8 -*-
"""
浏览器驱动与环境管理模块
基于 DrissionPage 实现移动端环境指纹模拟与真实浏览器驱动管理
"""
import os
import sys
from typing import Optional
from DrissionPage import ChromiumPage, ChromiumOptions

from config import (
    BROWSER_CANDIDATE_PATHS,
    BROWSER_USER_DATA_DIR,
    MOBILE_USER_AGENT,
    VIEWPORT_WIDTH,
    VIEWPORT_HEIGHT,
    DEFAULT_HEADLESS,
    PAGE_LOAD_TIMEOUT,
)


def find_browser_executable(custom_path: Optional[str] = None) -> Optional[str]:
    """
    自动查找系统中的 Chrome 或 Edge 浏览器路径
    """
    if custom_path and os.path.exists(custom_path):
        return custom_path

    for candidate in BROWSER_CANDIDATE_PATHS:
        if os.path.exists(candidate):
            return candidate

    return None


def create_browser_options(
    headless: bool = DEFAULT_HEADLESS,
    custom_browser_path: Optional[str] = None,
    user_data_path: Optional[str] = None,
) -> ChromiumOptions:
    """
    构造 DrissionPage 的 ChromiumOptions 配置
    注入移动端 UA、视口分辨率与免检参数
    """
    co = ChromiumOptions()

    # 1. 设置浏览器可执行文件路径
    browser_exe = find_browser_executable(custom_browser_path)
    if browser_exe:
        co.set_browser_path(browser_exe)

    # 2. 设置持久化用户数据目录，保留缓存与登录态 Cookie
    profile_dir = user_data_path or str(BROWSER_USER_DATA_DIR)
    co.set_user_data_path(profile_dir)

    # 3. 注入移动端 User-Agent
    co.set_user_agent(MOBILE_USER_AGENT)

    # 4. 配置防反爬与免自动化标识参数 (DrissionPage 本身无 webdriver 特征，可进一步强化)
    co.set_argument("--disable-blink-features=AutomationControlled")
    co.set_argument("--disable-infobars")
    co.set_argument("--no-first-run")
    co.set_argument("--no-default-browser-check")
    co.set_argument(f"--window-size={VIEWPORT_WIDTH + 80},{VIEWPORT_HEIGHT + 150}")

    # 5. 自动分配未被占用的端口
    co.auto_port()

    # 6. 设置无头模式
    co.headless(headless)

    # 7. 设置页面加载超时时间
    co.set_timeouts(base=PAGE_LOAD_TIMEOUT, page_load=PAGE_LOAD_TIMEOUT)

    return co


class BrowserManager:
    """
    浏览器生命周期管理器
    """

    def __init__(self, headless: bool = DEFAULT_HEADLESS, custom_browser_path: Optional[str] = None):
        self.headless = headless
        self.custom_browser_path = custom_browser_path
        self.page: Optional[ChromiumPage] = None

    def start(self) -> ChromiumPage:
        """
        启动并配置浏览器实例
        """
        if self.page is None:
            options = create_browser_options(
                headless=self.headless,
                custom_browser_path=self.custom_browser_path
            )
            self.page = ChromiumPage(options)
            try:
                # 调整页面窗口尺寸与移动端视口
                self.page.set.window.size(VIEWPORT_WIDTH, VIEWPORT_HEIGHT)
            except Exception:
                pass
        return self.page

    def stop(self):
        """
        安全退出浏览器
        """
        if self.page:
            try:
                self.page.quit()
            except Exception:
                pass
            finally:
                self.page = None

    def __enter__(self) -> ChromiumPage:
        return self.start()

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.stop()
