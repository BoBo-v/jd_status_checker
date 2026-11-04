# -*- coding: utf-8 -*-
"""
Excel 数据读写与断点续传模块
基于 openpyxl 实现无损读写、自动识别列名、逐条回写与断点恢复
"""
import os
import shutil
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict, List, Tuple, Generator
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet

from config import EXCEL_COL_KEYWORDS


class ExcelHandler:
    def __init__(self, file_path: str, sheet_name: Optional[str] = None):
        self.file_path = Path(file_path).resolve()
        if not self.file_path.exists():
            raise FileNotFoundError(f"未找到指定的 Excel 文件: {self.file_path}")

        self.sheet_name = sheet_name
        self.wb = None
        self.ws: Optional[Worksheet] = None
        self.col_indices: Dict[str, int] = {}
        self.header_row: int = 1

        self._load()

    def _load(self):
        """
        加载 Excel 文件并解析列索引
        使用 data_only=False 保留原始公式与排版
        """
        # 加载工作簿
        self.wb = openpyxl.load_workbook(self.file_path, data_only=False)
        if self.sheet_name and self.sheet_name in self.wb.sheetnames:
            self.ws = self.wb[self.sheet_name]
        else:
            self.ws = self.wb.active

        # 解析表头列索引
        self._detect_headers()

    def _detect_headers(self):
        """
        在前 3 行中智能扫描匹配各关键列所在位置
        """
        detected = {}
        found_header_row = 1

        for r in range(1, min(4, self.ws.max_row + 1)):
            row_vals = [str(self.ws.cell(row=r, column=c).value or "").strip() for c in range(1, self.ws.max_column + 1)]
            # 检查是否包含 URL 关键词
            has_url = any(
                any(kw.lower() in v.lower() for kw in EXCEL_COL_KEYWORDS["url"])
                for v in row_vals
            )
            if has_url:
                found_header_row = r
                break

        self.header_row = found_header_row

        # 扫描该表头行的所有列
        for c in range(1, self.ws.max_column + 1):
            cell_val = str(self.ws.cell(row=self.header_row, column=c).value or "").strip().lower()
            if not cell_val:
                continue

            for col_type, keywords in EXCEL_COL_KEYWORDS.items():
                if col_type not in detected:
                    for kw in keywords:
                        if kw.lower() in cell_val:
                            detected[col_type] = c
                            break

        # 检查核心列是否匹配成功
        if "url" not in detected:
            raise ValueError(
                f"未能自动在表格表头识别到【商品链接】列！\n"
                f"支持的列名关键词包括: {EXCEL_COL_KEYWORDS['url']}"
            )

        # 如果没有“是否已下架”列，自动在最后一列追加
        if "status" not in detected:
            new_col = self.ws.max_column + 1
            self.ws.cell(row=self.header_row, column=new_col, value="是否已下架（1下架）")
            detected["status"] = new_col

        # 如果没有“检测备注/异常状态”列，自动追加
        if "remark" not in detected:
            new_col = self.ws.max_column + 1
            self.ws.cell(row=self.header_row, column=new_col, value="检测备注/异常状态")
            detected["remark"] = new_col

        # 如果没有“检测时间”列，自动追加
        if "time" not in detected:
            new_col = self.ws.max_column + 1
            self.ws.cell(row=self.header_row, column=new_col, value="检测时间")
            detected["time"] = new_col

        self.col_indices = detected

    def is_completed(self, status_val: any) -> bool:
        """
        判断该行是否已完成有效检测（0 或 1）
        """
        if status_val is None:
            return False
        s = str(status_val).strip()
        return s in ("0", "1", "0.0", "1.0")

    def iter_rows(
        self,
        skip_completed: bool = True,
        recheck_errors_only: bool = False,
    ) -> Generator[Tuple[int, any, any, any], None, None]:
        """
        遍历所有待检测行
        yields: (row_idx, raw_url, current_status, current_remark)
        """
        url_col = self.col_indices["url"]
        status_col = self.col_indices["status"]
        remark_col = self.col_indices["remark"]

        for row_idx in range(self.header_row + 1, self.ws.max_row + 1):
            raw_url = self.ws.cell(row=row_idx, column=url_col).value
            status_val = self.ws.cell(row=row_idx, column=status_col).value
            remark_val = self.ws.cell(row=row_idx, column=remark_col).value

            # 如果整行 URL 为空，跳过
            if raw_url is None or not str(raw_url).strip():
                continue

            # 断点续传过滤逻辑
            if skip_completed and not recheck_errors_only:
                if self.is_completed(status_val):
                    continue

            if recheck_errors_only:
                # 仅重跑异常或待查的行（既不是 0 也不是 1）
                if self.is_completed(status_val):
                    continue

            yield row_idx, raw_url, status_val, remark_val

    def count_total_items(self) -> int:
        """
        统计表格中有效商品行总数
        """
        url_col = self.col_indices["url"]
        count = 0
        for row_idx in range(self.header_row + 1, self.ws.max_row + 1):
            raw_url = self.ws.cell(row=row_idx, column=url_col).value
            if raw_url is not None and str(raw_url).strip():
                count += 1
        return count

    def update_row(
        self,
        row_idx: int,
        status: any,
        remark: str,
        check_time: Optional[str] = None,
    ):
        """
        回写指定行的状态、备注与检测时间
        """
        status_col = self.col_indices["status"]
        remark_col = self.col_indices["remark"]
        time_col = self.col_indices["time"]

        # 回写状态值 (0, 1, 待查, 异常等)
        self.ws.cell(row=row_idx, column=status_col, value=status)

        # 回写备注说明
        self.ws.cell(row=row_idx, column=remark_col, value=remark)

        # 回写检测时间
        t = check_time or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.ws.cell(row=row_idx, column=time_col, value=t)

    def save(self):
        """
        保存工作簿回原文件
        """
        try:
            self.wb.save(self.file_path)
        except PermissionError:
            raise PermissionError(
                f"\n[错误] 无法保存文件到: {self.file_path}！\n"
                f"该 Excel 文件可能正在被 WPS 或 Office 打开，请先关闭该 Excel 窗口后重试！"
            )

    def close(self):
        """
        关闭工作簿
        """
        if self.wb:
            try:
                self.wb.close()
            except Exception:
                pass
