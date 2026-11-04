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
    DEFAULT_INCOGNITO,
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
    incognito: bool = DEFAULT_INCOGNITO,
    custom_browser_path: Optional[str] = None,
    user_data_path: Optional[str] = None,
) -> ChromiumOptions:
    """
    构造 DrissionPage 的 ChromiumOptions 配置
    移除引发警告的自动化命令行标志，注入移动端触屏支持与免检参数
    """
    co = ChromiumOptions()

    # 1. 设置浏览器可执行文件路径
    browser_exe = find_browser_executable(custom_browser_path)
    if browser_exe:
        co.set_browser_path(browser_exe)

    # 2. 设置无痕模式或持久化用户数据目录
    if incognito:
        co.incognito(True)
    else:
        profile_dir = user_data_path or str(BROWSER_USER_DATA_DIR)
        co.set_user_data_path(profile_dir)

    # 3. 注入移动端 User-Agent
    co.set_user_agent(MOBILE_USER_AGENT)

    # 4. 优化浏览器参数：
    # 注意：绝对不要添加 --disable-blink-features=AutomationControlled，否则会被浏览器弹窗告警且被风控识别！
    co.set_argument("--no-first-run")
    co.set_argument("--no-default-browser-check")
    co.set_argument(f"--window-size={VIEWPORT_WIDTH + 80},{VIEWPORT_HEIGHT + 150}")

    # 5. 注入触屏与移动视口支持 (模拟真实手机交互)
    co.set_argument("--enable-touch-events")
    co.set_argument("--touch-events=enabled")
    co.set_argument("--enable-viewport")

    # 6. 自动分配未被占用的端口
    co.auto_port()

    # 7. 设置无头模式
    co.headless(headless)

    # 8. 设置页面加载超时时间
    co.set_timeouts(base=PAGE_LOAD_TIMEOUT, page_load=PAGE_LOAD_TIMEOUT)

    return co


def clear_browser_cache(page: ChromiumPage):
    """
    清空当前会话的 Cookie、本地存储与缓存，清除风控追踪标识
    """
    try:
        page.clear_cache(session_storage=True, local_storage=True, cache=True, cookies=True)
        print("[*] 已成功清空当前浏览器会话的 Cookie 与缓存数据！")
    except Exception as e:
        print(f"[-] 清空缓存失败: {e}")


class BrowserManager:
    """
    浏览器生命周期管理器
    """

    def __init__(
        self,
        headless: bool = DEFAULT_HEADLESS,
        incognito: bool = DEFAULT_INCOGNITO,
        custom_browser_path: Optional[str] = None,
    ):
        self.headless = headless
        self.incognito = incognito
        self.custom_browser_path = custom_browser_path
        self.page: Optional[ChromiumPage] = None

    def start(self) -> ChromiumPage:
        """
        启动并配置浏览器实例
        """
        if self.page is None:
            options = create_browser_options(
                headless=self.headless,
                incognito=self.incognito,
                custom_browser_path=self.custom_browser_path,
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
