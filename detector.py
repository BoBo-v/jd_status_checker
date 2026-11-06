# -*- coding: utf-8 -*-
"""
页面状态多维判定模块（核心逻辑层）
针对京东移动端 H5 页面的特征与 DOM 进行多分支状态嗅探
"""
import time
import random
from datetime import datetime
from typing import NamedTuple, Optional
from DrissionPage import ChromiumPage

import config
from config import (
    PAGE_WAIT_AFTER_LOAD,
    OUT_OF_STOCK_AS_DELIST,
    ENABLE_SCREENSHOT,
    SCREENSHOT_DIR,
    SCREENSHOT_MAX_DAYS,
    SCREENSHOT_MAX_COUNT,
)
from logger import logger


class DetectionResult(NamedTuple):
    sku_id: str
    status: any              # 0=正常在售, 1=已下架, "待查"=需人工介入/重试, "异常"=非法链接等
    remark: str              # 状态详细说明
    screenshot_path: Optional[str] = None
    is_risk: bool = False    # 是否触发了风控或登录拦截
    is_captcha: bool = False # 是否触发了需要人机交互的安全验证
    is_login: bool = False   # 是否触发了提示登录/强制登录跳转


def clean_expired_screenshots():
    """
    自动清理过期与超出数量上限的历史截图（仅在开启截图时执行）
    """
    if not SCREENSHOT_DIR.exists():
        return
    try:
        now = time.time()
        max_age_seconds = SCREENSHOT_MAX_DAYS * 86400
        png_files = sorted(SCREENSHOT_DIR.glob("*.png"), key=lambda p: p.stat().st_mtime)

        # 1. 清理超过 N 天的文件
        for p in png_files:
            if now - p.stat().st_mtime > max_age_seconds:
                try:
                    p.unlink()
                except Exception:
                    pass

        # 2. 清理超出最大数量限制的早期文件
        remaining_files = sorted(SCREENSHOT_DIR.glob("*.png"), key=lambda p: p.stat().st_mtime)
        if len(remaining_files) > SCREENSHOT_MAX_COUNT:
            to_remove = remaining_files[:-SCREENSHOT_MAX_COUNT]
            for p in to_remove:
                try:
                    p.unlink()
                except Exception:
                    pass
    except Exception as e:
        logger.debug(f"清理过期截图时出错: {e}")


def take_error_screenshot(page: ChromiumPage, sku_id: str) -> str:
    """
    对异常页面进行屏幕截图并保存（受 ENABLE_SCREENSHOT 控制）
    默认不开启截图，改用日志记录，彻底杜绝磁盘膨胀！
    """
    if not ENABLE_SCREENSHOT:
        return ""

    try:
        SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)
        clean_expired_screenshots()

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        screenshot_filename = f"{sku_id}_{timestamp}.png"
        filepath = SCREENSHOT_DIR / screenshot_filename
        page.get_screenshot(path=str(SCREENSHOT_DIR), name=screenshot_filename, full_page=False)
        return str(filepath)
    except Exception as e:
        logger.warning(f"[-] 截图保存失败 [{sku_id}]: {e}")
        return ""


def simulate_human_micro_scroll(page: ChromiumPage, probability: Optional[float] = None):
    """
    轻量拟人微动滚动：
    - 随机概率触发（默认约 70%）：模拟真实人类有的看首屏直接离开，有的轻滑浏览；
    - 随机位移：滑动 160 ~ 380 像素，使用 smooth 平滑插值动画；
    - 拟人回弹：约 15% 几率产生 40 ~ 80 像素轻度回滚（模仿视线复位）。
    """
    prob = probability if probability is not None else config.MICRO_SCROLL_PROBABILITY
    if not config.ENABLE_MICRO_SCROLL or random.random() > prob:
        return

    try:
        # 1. 随机平滑下滑
        scroll_px = random.randint(160, 380)
        page.run_js(f"window.scrollBy({{top: {scroll_px}, behavior: 'smooth'}});")

        # 2. 拟人小细节：15% 几率轻微回弹看一眼首屏
        if random.random() < 0.15:
            time.sleep(round(random.uniform(0.15, 0.3), 2))
            scroll_back = random.randint(40, 80)
            page.run_js(f"window.scrollBy({{top: -{scroll_back}, behavior: 'smooth'}});")
    except Exception:
        try:
            page.scroll.down(200)
        except Exception:
            pass


def detect_product_status(
    page: ChromiumPage,
    sku_id: str,
    mobile_url: str,
    wait_time: float = PAGE_WAIT_AFTER_LOAD,
) -> DetectionResult:
    """
    核心判定函数：根据京东移动端 H5 页面多重特征判断商品上下架及风控状态
    """
    try:
        # 1. 打开商品移动端页面
        page.get(mobile_url)

        # 拟人交互：充分利用既有的 wait_time，全流程随机时间拟人化
        initial_wait = round(random.uniform(0.35, 0.65), 2)
        time.sleep(initial_wait)

        simulate_human_micro_scroll(page)

        remaining_wait = max(0.2, wait_time - initial_wait)
        time.sleep(remaining_wait)

        current_url = page.url or ""

        # -------------------------------------------------------------
        # 判定分支 1: 触发登录拦截 (plogin / passport 跳转或弹窗)
        # -------------------------------------------------------------
        is_login_url = (
            "plogin.m.jd.com" in current_url
            or "passport.m.jd.com" in current_url
            or "passport.jd.com" in current_url
        )

        # 提取页面全文（包含纯文本与源码）
        html_content = page.html or ""
        body_ele = page.ele("tag:body")
        page_text = body_ele.text if body_ele else ""

        is_login_modal = (
            "请登录后查看" in page_text
            or "立即登录" in page_text
            and ("账号密码登录" in page_text or "短信验证码登录" in page_text)
        )

        if is_login_url or is_login_modal:
            img_path = take_error_screenshot(page, sku_id)
            return DetectionResult(
                sku_id=sku_id,
                status="待查",
                remark="触发强制登录拦截",
                screenshot_path=img_path,
                is_risk=True,
                is_captcha=False,
                is_login=True,
            )

        # -------------------------------------------------------------
        # 判定分支 2: 触发 CFE 安全风控拦截页与滑块验证
        # -------------------------------------------------------------
        is_cfe_page = (
            "cfe.m.jd.com" in current_url
            or "验证一下，购物无忧" in page_text
            or "快速验证" in page_text
            or "遇到问题点我反馈" in page_text
        )

        risk_keywords = [
            "验证码",
            "安全验证",
            "访问过于频繁",
            "网络开小差",
            "滑动验证",
            "拖动滑块",
            "请按住滑块",
            "请完成下方验证",
        ]
        has_risk_text = any(kw in page_text or kw in html_content for kw in risk_keywords)

        if is_cfe_page or has_risk_text:
            img_path = take_error_screenshot(page, sku_id)
            remark_msg = "触发安全验证（CFE拦截）" if is_cfe_page else "触发验证码风控，需验证"
            return DetectionResult(
                sku_id=sku_id,
                status="待查",
                remark=remark_msg,
                screenshot_path=img_path,
                is_risk=True,
                is_captcha=True,
                is_login=False,  # 明确区分：这是安全验证/人机滑块，不是登录拦截！
            )

        # -------------------------------------------------------------
        # 判定分支 3: 404 或商品不存在 / 链接失效
        # -------------------------------------------------------------
        not_found_keywords = [
            "商品不存在",
            "您访问的商品不存在",
            "抱歉，没有找到相关页面",
            "404 Not Found",
            "商品已删除",
        ]
        if any(kw in page_text or kw in html_content for kw in not_found_keywords):
            return DetectionResult(
                sku_id=sku_id,
                status=1,
                remark="商品不存在或链接失效",
            )

        # 如果被重定向回移动端首页，通常意味着商品不存在或已失效
        if current_url.rstrip("/").endswith("m.jd.com") or "m.jd.com/index.html" in current_url:
            return DetectionResult(
                sku_id=sku_id,
                status=1,
                remark="商品不存在或重定向至首页",
            )

        # -------------------------------------------------------------
        # 判定分支 4: 下架状态 B（弹窗完全拦截下架）
        # 特征：页面弹窗显示“该商品已下架”，可能带“逛逛首页”
        # -------------------------------------------------------------
        if "该商品已下架" in page_text or "该商品已下架" in html_content:
            return DetectionResult(
                sku_id=sku_id,
                status=1,
                remark="已下架（弹窗完全下架）",
            )

        # -------------------------------------------------------------
        # 判定分支 5: 区域无货/暂无存货判定 (优先于展示相似商品，确保无存货算0)
        # 特征：页面提示“此商品在所选区域暂无存货”、“暂无存货”、“到货通知”等
        # -------------------------------------------------------------
        out_of_stock_keywords = [
            "此商品在所选区域暂无存货",
            "所选区域暂无存货",
            "所选地区暂无存货",
            "所选地区暂时无货",
            "该商品暂时无货",
            "暂时无货",
            "暂无存货",
            "暂无报价",
            "无存货",
            "到货通知",
        ]
        if any(kw in page_text or kw in html_content for kw in out_of_stock_keywords):
            val = 1 if OUT_OF_STOCK_AS_DELIST else 0
            return DetectionResult(
                sku_id=sku_id,
                status=val,
                remark="在售但当前区域无货",
            )

        # -------------------------------------------------------------
        # 判定分支 6: 停售与真正下架状态判定
        # -------------------------------------------------------------
        if "此商品已停售" in page_text or "此商品已停售" in html_content:
            return DetectionResult(
                sku_id=sku_id,
                status=1,
                remark="已停售",
            )

        delist_a_keywords = [
            "此商品已下架，推荐查看相似商品",
            "此商品已下架",
            "已下架，去看看其他",
            "商品已下架",
            "已下架",
        ]
        if any(kw in page_text or kw in html_content for kw in delist_a_keywords):
            return DetectionResult(
                sku_id=sku_id,
                status=1,
                remark="已下架（展示相似推荐）",
            )

        # -------------------------------------------------------------
        # 判定分支 7: 正常在售（底部包含购买/加购相关按钮）
        # -------------------------------------------------------------
        buy_button_keywords = [
            "立即购买",
            "加入购物车",
            "领券购买",
            "去APP购买",
            "去客户端购买",
            "立即预约",
            "立即抢购",
            "付定金",
            "开抢提醒",
        ]
        for kw in buy_button_keywords:
            if kw in page_text or kw in html_content:
                return DetectionResult(
                    sku_id=sku_id,
                    status=0,
                    remark="正常在售",
                )

        # 也可以尝试在 DOM 中定位底栏按钮元素
        buy_btn_ele = (
            page.ele("@text():立即购买")
            or page.ele("@text():加入购物车")
            or page.ele(".de_btn")
            or page.ele(".buy_btn")
        )
        if buy_btn_ele:
            return DetectionResult(
                sku_id=sku_id,
                status=0,
                remark="正常在售",
            )

        # -------------------------------------------------------------
        # 判定分支 8: 未匹配到已知模式，属于未知状态（截图保留证据）
        # -------------------------------------------------------------
        img_path = take_error_screenshot(page, sku_id)
        return DetectionResult(
            sku_id=sku_id,
            status="待查",
            remark="未知状态，请查看截图",
            screenshot_path=img_path,
        )

    except Exception as e:
        img_path = take_error_screenshot(page, sku_id)
        return DetectionResult(
            sku_id=sku_id,
            status="待查",
            remark=f"检测异常: {str(e)[:40]}",
            screenshot_path=img_path,
            is_risk=False,
        )
