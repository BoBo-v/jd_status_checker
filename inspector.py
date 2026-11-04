# -*- coding: utf-8 -*-
"""
核心巡检调度模块
协调 Excel 读写、浏览器自动化、URL 提取与多状态判定的完整巡检流程
"""
import sys
import time
import random
from typing import Optional, Dict

from config import (
    DELAY_MIN,
    DELAY_MAX,
    SAVE_INTERVAL,
    DEFAULT_HEADLESS,
)
from url_parser import to_mobile_h5_url
from browser_manager import BrowserManager
from detector import detect_product_status, DetectionResult
from excel_handler import ExcelHandler


class JDStatusInspector:
    def __init__(
        self,
        excel_path: str,
        sheet_name: Optional[str] = None,
        headless: bool = DEFAULT_HEADLESS,
        skip_completed: bool = True,
        recheck_errors_only: bool = False,
        delay_range: tuple = (DELAY_MIN, DELAY_MAX),
        save_interval: int = SAVE_INTERVAL,
        custom_browser_path: Optional[str] = None,
    ):
        self.excel_path = excel_path
        self.sheet_name = sheet_name
        self.headless = headless
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
        print(f"[*] 运行模式: {'无头后台模式' if self.headless else '有界面模式（推荐，更低风控）'}")
        print(f"[*] 断点续传: {'开启（跳过已标记行）' if self.skip_completed else '关闭（重新检测全部）'}")
        print(f"[*] 请求随机延时: {self.delay_min} ~ {self.delay_max} 秒")
        print("=" * 65 + "\n")

        excel = ExcelHandler(self.excel_path, self.sheet_name)
        total_items = excel.count_total_items()
        self.stats["total"] = total_items

        print(f"[+] 成功加载表格，共包含 {total_items} 条待核验记录")
        print(f"[+] 识别到关键列位置: {excel.col_indices}\n")

        # 预先筛选待处理行
        pending_rows = list(
            excel.iter_rows(
                skip_completed=self.skip_completed,
                recheck_errors_only=self.recheck_errors_only
            )
        )
        self.stats["skipped"] = total_items - len(pending_rows)

        if not pending_rows:
            print("[√] 所有商品均已检测完毕，没有需要处理的记录！")
            excel.close()
            return self.stats

        print(f"[*] 本次实际需要检测商品数: {len(pending_rows)} 条 (已跳过 {self.stats['skipped']} 条)\n")

        # 启动浏览器驱动
        bm = BrowserManager(headless=self.headless, custom_browser_path=self.custom_browser_path)
        try:
            print("[*] 正在启动移动端仿真浏览器环境...")
            page = bm.start()
            print("[+] 浏览器初始化完毕，开始逐行巡检...\n")

            processed_count = 0

            for idx, (row_idx, raw_url, current_status, current_remark) in enumerate(pending_rows, start=1):
                sku_id, mobile_url = to_mobile_h5_url(raw_url)

                print(f"[{idx}/{len(pending_rows)}] 行 {row_idx:3d} | 原始输入: {str(raw_url)[:35]}...")

                if not sku_id:
                    # 链接无法解析出 SKU
                    print(f"    [-] 链接无效，无法提取 SKU ID")
                    excel.update_row(row_idx, status="异常", remark="商品链接格式错误，无法提取SKU")
                    self.stats["invalid_url"] += 1
                else:
                    print(f"    [>] SKU: {sku_id} -> 访问: {mobile_url}")
                    res: DetectionResult = detect_product_status(page, sku_id, mobile_url)

                    # 格式化输出状态
                    status_text = "正常在售" if res.status == 0 else ("已下架" if res.status == 1 else str(res.status))
                    print(f"    [√] 检测结果: 【{status_text}】(值={res.status}) | 备注: {res.remark}")

                    if res.is_risk:
                        print(f"    [!] 触发风控或登录拦截: {res.remark}，截图已存: {res.screenshot_path}")
                        self.stats["risk_or_manual"] += 1
                    elif res.status == 0:
                        self.stats["in_stock"] += 1
                    elif res.status == 1:
                        self.stats["delisted"] += 1
                    else:
                        self.stats["risk_or_manual"] += 1

                    # 回写数据至 Excel
                    excel.update_row(row_idx, status=res.status, remark=res.remark)

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
            print("\n[+] 全部检测处理完毕，表格数据已实时回写保存成功！")

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
        print("=" * 65 + "\n")
