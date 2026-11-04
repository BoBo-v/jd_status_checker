# -*- coding: utf-8 -*-
"""
页面状态多维判定模块（核心逻辑层）
针对京东移动端 H5 页面的特征与 DOM 进行多分支状态嗅探
"""
import time
from datetime import datetime
from typing import NamedTuple, Optional
from DrissionPage import ChromiumPage

from config import (
    PAGE_WAIT_AFTER_LOAD,
    OUT_OF_STOCK_AS_DELIST,
    SCREENSHOT_DIR,
)


class DetectionResult(NamedTuple):
    sku_id: str
    status: any              # 0=正常在售, 1=已下架, "待查"=需人工介入/重试, "异常"=非法链接等
    remark: str              # 状态详细说明
    screenshot_path: Optional[str] = None
    is_risk: bool = False    # 是否触发了风控或登录拦截
    is_captcha: bool = False # 是否触发了需要人机交互的安全验证（如 CFE 快速验证、滑块）


def take_error_screenshot(page: ChromiumPage, sku_id: str) -> str:
    """
    对异常页面进行屏幕截图并保存
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    screenshot_filename = f"{sku_id}_{timestamp}.png"
    filepath = SCREENSHOT_DIR / screenshot_filename
    try:
        page.get_screenshot(path=str(SCREENSHOT_DIR), name=screenshot_filename, full_page=False)
        return str(filepath)
    except Exception as e:
        print(f"[-] 截图保存失败 [{sku_id}]: {e}")
        return ""


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
        time.sleep(wait_time)

        current_url = page.url or ""

        # -------------------------------------------------------------
        # 判定分支 1: 触发登录拦截 (plogin / passport 跳转)
        # -------------------------------------------------------------
        if (
            "plogin.m.jd.com" in current_url
            or "passport.m.jd.com" in current_url
            or "passport.jd.com" in current_url
        ):
            img_path = take_error_screenshot(page, sku_id)
            return DetectionResult(
                sku_id=sku_id,
                status="待查",
                remark="触发强制登录拦截",
                screenshot_path=img_path,
                is_risk=True,
                is_captcha=False,
            )

        # 提取页面全文（包含纯文本与源码）
        html_content = page.html or ""
        body_ele = page.ele("tag:body")
        page_text = body_ele.text if body_ele else ""

        # -------------------------------------------------------------
        # 判定分支 2: 触发 CFE 安全风控拦截页与滑块验证
        # 特征：cfe.m.jd.com、"验证一下，购物无忧"、"快速验证" 等
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
            remark_msg = "触发安全验证（CFE购物无忧拦截）" if is_cfe_page else "触发验证码风控，需验证"
            return DetectionResult(
                sku_id=sku_id,
                status="待查",
                remark=remark_msg,
                screenshot_path=img_path,
                is_risk=True,
                is_captcha=True,  # 标记为人机验证，触发暂停等待人工处理逻辑
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
        # 判定分支 5: 下架状态 A（商品页留存，展示相似推荐）
        # 特征：黄色提示条“此商品已下架，推荐查看相似商品”，底部按钮“查看相似商品”
        # -------------------------------------------------------------
        delist_a_keywords = [
            "此商品已下架，推荐查看相似商品",
            "此商品已下架",
            "查看相似商品",
            "已下架，去看看其他",
        ]
        if any(kw in page_text or kw in html_content for kw in delist_a_keywords):
            return DetectionResult(
                sku_id=sku_id,
                status=1,
                remark="已下架（展示相似推荐）",
            )

        # -------------------------------------------------------------
        # 判定分支 6: 停售与区域无货状态判定
        # -------------------------------------------------------------
        if "此商品已停售" in page_text or "此商品已停售" in html_content:
            return DetectionResult(
                sku_id=sku_id,
                status=1,
                remark="已停售",
            )

        out_of_stock_keywords = ["所选地区暂时无货", "暂无报价", "暂时无货", "该商品暂时无货"]
        if any(kw in page_text or kw in html_content for kw in out_of_stock_keywords):
            val = 1 if OUT_OF_STOCK_AS_DELIST else 0
            return DetectionResult(
                sku_id=sku_id,
                status=val,
                remark="在售但当前区域无货",
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
