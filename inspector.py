# -*- coding: utf-8 -*-
"""
核心巡检调度模块
协调 Excel 读写、多工作表（分表）解析、浏览器自动化、URL 提取与多状态判定的完整巡检流程
"""
import sys
import time
import random
from typing import Optional, Dict, Union, List

from config import (
    DELAY_MIN,
    DELAY_MAX,
    SAVE_INTERVAL,
    DEFAULT_HEADLESS,
    AUTO_PAUSE_ON_CAPTCHA,
    AUTO_RESTART_ON_LOGIN,
    MAX_LOGIN_RESTART_RETRIES,
    BYPASS_PROXY,
)
from url_parser import to_mobile_h5_url
from browser_manager import BrowserManager
from detector import detect_product_status, DetectionResult
from excel_handler import ExcelHandler
from logger import logger, get_log_file_path


class JDStatusInspector:
    def __init__(
        self,
        excel_path: str,
        target_sheets: Optional[Union[str, List[str]]] = None,
        headless: bool = DEFAULT_HEADLESS,
        auto_pause_on_captcha: bool = AUTO_PAUSE_ON_CAPTCHA,
        auto_restart_on_login: bool = AUTO_RESTART_ON_LOGIN,
        max_login_retries: int = MAX_LOGIN_RESTART_RETRIES,
        bypass_proxy: bool = BYPASS_PROXY,
        skip_completed: bool = True,
        recheck_errors_only: bool = False,
        delay_range: tuple = (DELAY_MIN, DELAY_MAX),
        save_interval: int = SAVE_INTERVAL,
        custom_browser_path: Optional[str] = None,
    ):
        self.excel_path = excel_path
        self.target_sheets = target_sheets
        self.headless = headless
        self.auto_pause_on_captcha = auto_pause_on_captcha
        self.auto_restart_on_login = auto_restart_on_login
        self.max_login_retries = max(1, max_login_retries)
        self.bypass_proxy = bypass_proxy
        self.skip_completed = skip_completed
        self.recheck_errors_only = recheck_errors_only
        self.delay_min, self.delay_max = delay_range
        self.save_interval = max(1, save_interval)
        self.custom_browser_path = custom_browser_path

        # 统计计数
        self.stats = {
            "total": 0,
            "checked": 0,
            "skipped": 0,
            "in_stock": 0,
            "delisted": 0,
            "risk_or_manual": 0,
            "invalid_url": 0,
        }

    def run(self):
        """
        开始批量巡检主流程
        """
        print("\n" + "=" * 65)
        print("          京东商品状态自动化巡检系统 (DrissionPage)")
        print("=" * 65)
        print(f"[*] 目标表格: {self.excel_path}")
        print(f"[*] 原表保护: 100% 只读流式加载（绝不修改原表，商品主图 100% 完好无损！）")
        print(f"[*] 运行模式: {'无头后台模式' if self.headless else '有界面模式'}")
        print(f"[*] 登录自愈: {'已开启（遇到提示登录自动关闭浏览器并重启重试）' if self.auto_restart_on_login else '未开启'}")
        print(f"[*] 网络直连: {'强制国内直连 (已绕过VPN/代理)' if self.bypass_proxy else '系统默认网络'}")
        print(f"[*] 断点续传: {'开启（跳过已标记行）' if self.skip_completed else '关闭（重新检测全部）'}")
        print(f"[*] 请求随机延时: {self.delay_min} ~ {self.delay_max} 秒")
        print("=" * 65 + "\n")

        excel = ExcelHandler(self.excel_path, self.target_sheets)
        total_items = excel.count_total_items()
        self.stats["total"] = total_items

        # 预先筛选待处理行
        pending_rows = list(
            excel.iter_rows(
                skip_completed=self.skip_completed,
                recheck_errors_only=self.recheck_errors_only,
            )
        )
        self.stats["skipped"] = total_items - len(pending_rows)

        if not pending_rows:
            print("[√] 所有商品均已检测完毕，没有需要处理的记录！")
            excel.close()
            return self.stats

        print(f"[*] 本次实际需要检测商品数: {len(pending_rows)} 条 (已跳过 {self.stats['skipped']} 条)\n")

        # 启动浏览器驱动
        bm = BrowserManager(
            headless=self.headless,
            custom_browser_path=self.custom_browser_path,
            bypass_proxy=self.bypass_proxy,
        )
        try:
            print("[*] 正在启动真实指纹浏览器环境...")
            page = bm.start()
            print("[+] 浏览器初始化完毕，开始逐行巡检...\n")

            processed_count = 0

            for idx, (sheet_name, row_idx, raw_url, current_status, current_remark) in enumerate(pending_rows, start=1):
                sku_id, mobile_url = to_mobile_h5_url(raw_url)

                print(f"[{idx}/{len(pending_rows)}] 【{sheet_name}】行 {row_idx:3d} | 原始输入: {str(raw_url)[:30]}...")

                if not sku_id:
                    print(f"    [-] 链接无效，无法提取 SKU ID")
                    excel.update_row(sheet_name, row_idx, status="异常", remark="商品链接格式错误，无法提取SKU")
                    self.stats["invalid_url"] += 1
                else:
                    print(f"    [>] SKU: {sku_id} -> 访问: {mobile_url}")
                    res: DetectionResult = detect_product_status(page, sku_id, mobile_url)

                    # -------------------------------------------------------------
                    # 核心机制 1: 遇到安全验证（CFE 滑动验证码）-> 等待人工滑动 或 关窗重启
                    # -------------------------------------------------------------
                    if res.is_captcha:
                        print("\n" + "=" * 65)
                        print("  [!] 检测到京东安全验证（CFE 滑动验证码）！")
                        print("  [*] 请在弹出的浏览器窗口中【手动完成滑块验证】...")
                        print("  [操作指南]")
                        print("  1. 手动滑动成功后，程序将自动感应并继续后续检测；")
                        print("  2. 若您直接点击右上角【X】关闭浏览器，程序将自动为您重启新浏览器！")
                        print("=" * 65 + "\n")

                        # 每秒轮询感知人工操作（最多等待 180 秒）
                        for _ in range(180):
                            time.sleep(1.0)

                            # 探测 A：用户是否直接点 X 关闭了浏览器
                            if not bm.is_alive():
                                print("\n[*] 感应到您手动关闭了浏览器，正在自动重新打开全新浏览器并继续...")
                                page = bm.restart(clean_session=True)
                                time.sleep(1.5)
                                res = detect_product_status(page, sku_id, mobile_url)
                                break

                            # 探测 B：用户是否已在页面上滑过验证码
                            try:
                                curr_url = page.url or ""
                                body_ele = page.ele("tag:body")
                                page_text = body_ele.text if body_ele else ""
                                is_still_cfe = "cfe.m.jd.com" in curr_url or any(
                                    kw in page_text for kw in ["验证一下", "滑动验证", "拖动滑块", "请按住滑块", "快速验证"]
                                )
                                if not is_still_cfe:
                                    print("\n[√] 识别到人工滑块验证已完成！正在恢复检测...")
                                    time.sleep(1.5)
                                    res = detect_product_status(page, sku_id, mobile_url)
                                    break
                            except Exception:
                                if not bm.is_alive():
                                    print("\n[*] 感应到您手动关闭了浏览器，正在自动重新打开全新浏览器并继续...")
                                    page = bm.restart(clean_session=True)
                                    time.sleep(1.5)
                                    res = detect_product_status(page, sku_id, mobile_url)
                                    break
                        else:
                            print("\n[-] 人工验证等待超时（3分钟），自动跳过当前商品。")

                    # -------------------------------------------------------------
                    # 核心机制 2: 遇到提示登录/跳转拦截 (plogin / passport) -> 自动重启重试
                    # -------------------------------------------------------------
                    elif res.is_login and self.auto_restart_on_login:
                        for retry_idx in range(1, self.max_login_retries + 1):
                            print(f"    [!] 检测到提示登录/跳转拦截，执行【关闭浏览器并重新开启】恢复操作 ({retry_idx}/{self.max_login_retries})...")
                            page = bm.restart(clean_session=True)
                            time.sleep(1.5)
                            res = detect_product_status(page, sku_id, mobile_url)
                            if not res.is_login and not res.is_captcha:
                                status_text = "正常在售" if res.status == 0 else ("已下架" if res.status == 1 else str(res.status))
                                print(f"    [√] 重启浏览器后成功恢复！检测结果: 【{status_text}】(值={res.status})")
                                break
                        else:
                            print(f"    [-] 重启浏览器 {self.max_login_retries} 次后仍触发登录拦截，标记为待查并继续。")

                    # 格式化输出状态
                    status_text = "正常在售" if res.status == 0 else ("已下架" if res.status == 1 else str(res.status))
                    print(f"    [√] 检测结果: 【{status_text}】(值={res.status}) | 备注: {res.remark}")
                    logger.info(f"【{sheet_name}:行{row_idx:3d}】SKU:{sku_id} -> 【{status_text}】(值={res.status}) | 备注: {res.remark}")

                    if res.is_risk:
                        log_msg = f"触发风控或登录拦截: {res.remark}"
                        if res.screenshot_path:
                            print(f"    [!] {log_msg} (截图: {res.screenshot_path})")
                        else:
                            print(f"    [!] {log_msg}")
                        logger.warning(f"【{sheet_name}:行{row_idx:3d}】SKU:{sku_id} {log_msg} | URL: {mobile_url}")
                        self.stats["risk_or_manual"] += 1
                    elif res.status == 0:
                        self.stats["in_stock"] += 1
                    elif res.status == 1:
                        self.stats["delisted"] += 1
                    else:
                        self.stats["risk_or_manual"] += 1

                    # 回写数据至轻量结果表格
                    excel.update_row(sheet_name, row_idx, status=res.status, remark=res.remark)

                self.stats["checked"] += 1
                processed_count += 1

                # 实时保存机制
                if processed_count % self.save_interval == 0:
                    excel.save()

                # 随机防风控停顿（最后一个商品无需休眠）
                if idx < len(pending_rows):
                    sleep_sec = round(random.uniform(self.delay_min, self.delay_max), 2)
                    time.sleep(sleep_sec)

            # 最终保存
            excel.save()
            print("\n[+] 全部检测处理完毕！")
            print(f"[√] 巡检汇总结果已保存至：{excel.result_path.name}")
            print(f"    完整路径: {excel.result_path}")

        except KeyboardInterrupt:
            print("\n\n[!] 检测到用户手动中断 (Ctrl+C)！正在保存已完成的数据...")
            excel.save()
            print("[+] 数据保存完毕，安全退出！")
        except Exception as e:
            print(f"\n[-] 运行发生异常: {e}")
            try:
                excel.save()
                print("[+] 已紧急保存已处理的数据！")
            except Exception:
                pass
            raise e
        finally:
            bm.stop()
            excel.close()
            # 无论正常结束还是手动中断退出，自动将巡检结果写回原表（零损保护图片）
            try:
                excel.sync_to_original()
            except Exception as e:
                print(f"[!] 同步写回原表提示: {e}")

        # 打印巡检报告总结
        self._print_summary()
        return self.stats

    def _print_summary(self):
        print("\n" + "=" * 65)
        print("                      巡检任务统计报告")
        print("=" * 65)
        print(f"  总记录行数:        {self.stats['total']:4d}")
        print(f"  本次检测行数:      {self.stats['checked']:4d}")
        print(f"  已跳过行数:        {self.stats['skipped']:4d}")
        print(f"  -------------------------------------------")
        print(f"  [0] 正常在售:      {self.stats['in_stock']:4d}")
        print(f"  [1] 已下架:        {self.stats['delisted']:4d}")
        print(f"  [?] 风控/待查:     {self.stats['risk_or_manual']:4d}")
        print(f"  [!] 无效链接:      {self.stats['invalid_url']:4d}")
        print(f"  -------------------------------------------")
        print(f"  详细日志文件:      {get_log_file_path()}")
        print("=" * 65 + "\n")
