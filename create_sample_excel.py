# -*- coding: utf-8 -*-
"""
生成测试 Excel 样例文档
包含多种不同类型的商品链接、已检测标记、异常链接等，用于本地快速联调
"""
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

def create_sample_excel(filename: str = "sample_jd_goods.xlsx") -> str:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "商品巡检表"

    # 表头定义
    headers = [
        "序号",
        "商品名称",
        "京东链接（必填）",
        "供货价",
        "是否已下架（1下架）",
        "检测备注/异常状态",
        "检测时间",
    ]

    # 测试样例数据
    rows = [
        (
            1,
            "正常在售测试商品 (PC端长链接)",
            "https://item.jd.com/100152311477.html#crumb-wrap",
            129.00,
            None,
            None,
            None,
        ),
        (
            2,
            "正常在售测试商品 (短链接带参数)",
            "https://item.jd.com/100152311477.html?dist=jd_mall",
            129.00,
            None,
            None,
            None,
        ),
        (
            3,
            "弹窗完全下架测试商品",
            "https://item.jd.com/1000000001.html",
            59.90,
            None,
            None,
            None,
        ),
        (
            4,
            "不存在/已失效商品",
            "https://item.m.jd.com/product/99999999999.html",
            99.00,
            None,
            None,
            None,
        ),
        (
            5,
            "纯SKU编号输入",
            "100152311477",
            129.00,
            None,
            None,
            None,
        ),
        (
            6,
            "断点续传测试（已检测为下架的行，应自动跳过）",
            "https://item.jd.com/1000000001.html",
            59.90,
            1,
            "已下架（历史检测）",
            "2026-11-04 12:00:00",
        ),
        (
            7,
            "错误格式链接测试",
            "https://item.jd.com/not_a_valid_sku.html",
            0.00,
            None,
            None,
            None,
        ),
    ]

    # 样式配置
    header_font = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    data_font = Font(name="微软雅黑", size=10)
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    # 写入表头
    ws.append(headers)
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align
        cell.border = thin_border

    # 写入行数据
    for row in rows:
        ws.append(list(row))

    # 应用数据单元格样式与列宽设置
    col_widths = {1: 8, 2: 32, 3: 50, 4: 12, 5: 22, 6: 28, 7: 22}

    for row_idx in range(2, ws.max_row + 1):
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = data_font
            cell.border = thin_border
            if col_idx in (1, 4, 5, 7):
                cell.alignment = center_align
            else:
                cell.alignment = left_align

    for col_idx, width in col_widths.items():
        ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = width

    out_path = Path(__file__).resolve().parent / filename
    wb.save(out_path)
    print(f"[√] 测试样例 Excel 已生成: {out_path}")
    return str(out_path)


if __name__ == "__main__":
    create_sample_excel()
