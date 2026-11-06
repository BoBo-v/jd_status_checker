# -*- coding: utf-8 -*-
"""
浏览器驱动与环境管理模块
基于 DrissionPage 实现原生真实指纹驱动与直连网络配置
"""
import os
import sys
import time
import shutil
from typing import Optional
from DrissionPage import ChromiumPage, ChromiumOptions

from config import (
    CUSTOM_BROWSER_PATH,
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
    智能查找浏览器可执行文件：
    1. 优先使用命令行或参数传入的 custom_path；
    2. 其次使用 config.CUSTOM_BROWSER_PATH 中的自定义配置；
    3. 支持别名 "chrome" / "edge" 智能匹配系统对应浏览器；
    4. 支持任意真实存在的 Chromium 浏览器 .exe 绝对路径；
    5. 默认 (None) 自动寻找系统中已安装的 Chrome 或 Edge。
    """
    target = custom_path or CUSTOM_BROWSER_PATH
    if target:
        target_str = str(target).strip()
        target_lower = target_str.lower()

        # 别名快捷匹配：chrome
        if target_lower in ("chrome", "google", "google chrome", "google-chrome"):
            for p in BROWSER_CANDIDATE_PATHS:
                if "chrome" in p.lower() and os.path.exists(p):
                    return p
            print(f"[!] 提示：已配置优先使用 Chrome，但未在系统默认路径找到，将尝试自动匹配备选浏览器。")

        # 别名快捷匹配：edge
        elif target_lower in ("edge", "msedge", "microsoft edge", "microsoft-edge"):
            for p in BROWSER_CANDIDATE_PATHS:
                if ("edge" in p.lower() or "msedge" in p.lower()) and os.path.exists(p):
                    return p
            print(f"[!] 提示：已配置优先使用 Edge，但未在系统默认路径找到，将尝试自动匹配备选浏览器。")

        # 路径匹配：直接给定了真实存在的文件路径
        elif os.path.exists(target_str):
            return target_str
        else:
            print(f"[!] 警告：配置的浏览器路径不存在: {target_str}，将自动切换为系统可用浏览器。")

    # 默认回退：按候选列表顺序自动选取第一个存在的文件
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
            browser_exe = find_browser_executable(self.custom_browser_path)
            options = create_browser_options(
                headless=self.headless,
                custom_browser_path=browser_exe,
                bypass_proxy=self.bypass_proxy,
            )
            b_name = os.path.basename(browser_exe) if browser_exe else "系统推荐默认"
            print(f"[*] 启动浏览器驱动: {b_name} ({browser_exe})")
            self.page = ChromiumPage(options)
            try:
                # 调整页面窗口尺寸以完美适配移动端 H5 布局
                self.page.set.window.size(VIEWPORT_WIDTH, VIEWPORT_HEIGHT)
            except Exception:
                pass
        return self.page

    def is_alive(self) -> bool:
        """
        检查当前浏览器窗口/进程是否仍存活（未被用户手动关闭）
        """
        if not self.page:
            return False
        try:
            if hasattr(self.page, "states") and not self.page.states.is_alive:
                return False
            _ = self.page.url
            return True
        except Exception:
            return False

    def stop(self):
        """
        安全退出当前浏览器
        """
        if self.page:
            try:
                self.page.quit()
            except Exception:
                pass
            finally:
                self.page = None

    def restart(self, clean_session: bool = True) -> ChromiumPage:
        """
        核心恢复功能：完全退出当前浏览器进程，清理缓存后重新打开全新浏览器
        用于在遇到登录跳转提示时瞬间刷新环境状态
        """
        self.stop()
        if clean_session:
            clean_poisoned_browser_data()
        time.sleep(1.0)
        return self.start()

    def __enter__(self) -> ChromiumPage:
        return self.start()

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.stop()
