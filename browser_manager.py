# -*- coding: utf-8 -*-
"""
浏览器驱动与环境管理模块
基于 DrissionPage 实现原生真实指纹驱动与直连网络配置
"""
import os
import sys
import shutil
from typing import Optional
from DrissionPage import ChromiumPage, ChromiumOptions

from config import (
    BROWSER_CANDIDATE_PATHS,
    BROWSER_USER_DATA_DIR,
    VIEWPORT_WIDTH,
    VIEWPORT_HEIGHT,
    DEFAULT_HEADLESS,
    PAGE_LOAD_TIMEOUT,
    BYPASS_PROXY,
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


def clean_poisoned_browser_data():
    """
    清理可能被京东标记为黑名单会话的历史缓存目录
    """
    if BROWSER_USER_DATA_DIR.exists():
        try:
            shutil.rmtree(BROWSER_USER_DATA_DIR, ignore_errors=True)
            BROWSER_USER_DATA_DIR.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass


def create_browser_options(
    headless: bool = DEFAULT_HEADLESS,
    custom_browser_path: Optional[str] = None,
    user_data_path: Optional[str] = None,
    bypass_proxy: bool = BYPASS_PROXY,
) -> ChromiumOptions:
    """
    构造 DrissionPage 的 ChromiumOptions 配置
    使用浏览器原生 100% 真实指纹（避免伪造 UA 导致的 Sec-CH-UA 指纹矛盾），并强制直连避免境外代理风控
    """
    co = ChromiumOptions()

    # 1. 设置浏览器可执行文件路径
    browser_exe = find_browser_executable(custom_browser_path)
    if browser_exe:
        co.set_browser_path(browser_exe)

    # 2. 设置持久化用户数据目录
    profile_dir = user_data_path or str(BROWSER_USER_DATA_DIR)
    co.set_user_data_path(profile_dir)

    # 3. 绕过代理直接访问国内网络（彻底杜绝 VPN/Clash 导向境外机房触发假滑块死循环）
    if bypass_proxy:
        co.set_argument("--no-proxy-server")

    # 4. 保持真实原生指纹参数，不强行注入假 UA
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

    def __init__(
        self,
        headless: bool = DEFAULT_HEADLESS,
        custom_browser_path: Optional[str] = None,
        bypass_proxy: bool = BYPASS_PROXY,
    ):
        self.headless = headless
        self.custom_browser_path = custom_browser_path
        self.bypass_proxy = bypass_proxy
        self.page: Optional[ChromiumPage] = None

    def start(self) -> ChromiumPage:
        """
        启动并配置浏览器实例
        """
        if self.page is None:
            options = create_browser_options(
                headless=self.headless,
                custom_browser_path=self.custom_browser_path,
                bypass_proxy=self.bypass_proxy,
            )
            self.page = ChromiumPage(options)
            try:
                # 调整页面窗口尺寸以完美适配移动端 H5 布局
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
