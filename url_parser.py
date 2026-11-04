# -*- coding: utf-8 -*-
"""
URL 处理与标准化模块
实现从各类京东 PC/H5/APP 链接或纯数字中提取商品 skuId，并拼接为移动端 H5 链接
"""
import re
from typing import Optional, Tuple


def extract_sku_id(raw_input: any) -> Optional[str]:
    """
    从单元格文本、URL、数字中提取京东商品 SKU ID。
    
    支持格式示例：
    - https://item.jd.com/100152311477.html#crumb-wrap
    - https://item.jd.com/100152311477.html?dist=jd
    - https://item.m.jd.com/product/100152311477.html
    - https://item.m.jd.com/ware/view.action?wareId=100152311477
    - https://npcitem.jd.hk/100152311477.html
    - 纯数字：100152311477 或 100152311477.0 (Excel 浮点数)
    """
    if raw_input is None:
        return None

    text = str(raw_input).strip()
    if not text:
        return None

    # 如果是从 Excel 读入的浮点数字符串（如 "100152311477.0"），先截断 .0
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]

    # 1. 尝试匹配常见的包含 sku.html 的 URL 模式
    # 覆盖 item.jd.com/123.html, item.m.jd.com/product/123.html, npcitem.jd.hk/123.html 等
    match_html = re.search(r"(?:item\.jd\.com/|product/|item\.jd\.hk/|cps/item/)(\d+)\.html", text, re.IGNORECASE)
    if match_html:
        return match_html.group(1)

    # 2. 尝试匹配 wareId=123 或 sku=123 或 skuId=123
    match_param = re.search(r"(?:wareId|sku|skuId|id)=(\d+)", text, re.IGNORECASE)
    if match_param:
        return match_param.group(1)

    # 3. 纯数字提取（通常京东 skuId 长度在 6 到 15 位数字之间）
    if re.fullmatch(r"\d{6,15}", text):
        return text

    # 4. 从文本中模糊抓取一串符合长度的数字
    match_digits = re.search(r"(?:^|[^\d])(\d{6,15})(?:$|[^\d])", text)
    if match_digits:
        return match_digits.group(1)

    return None


def to_mobile_h5_url(raw_input: any) -> Tuple[Optional[str], Optional[str]]:
    """
    将原始输入转换为移动端 H5 链接。
    
    返回:
        (sku_id, mobile_url)
        如果提取失败，则返回 (None, None)
    """
    sku_id = extract_sku_id(raw_input)
    if not sku_id:
        return None, None

    mobile_url = f"https://item.m.jd.com/product/{sku_id}.html"
    return sku_id, mobile_url


if __name__ == "__main__":
    # 单元测试验证
    test_cases = [
        "https://item.jd.com/100152311477.html#crumb-wrap",
        "https://item.jd.com/100152311477.html?dist=jd",
        "https://item.m.jd.com/product/100152311477.html",
        "https://item.m.jd.com/ware/view.action?wareId=100152311477",
        "https://npcitem.jd.hk/100152311477.html",
        "100152311477",
        "100152311477.0",
        "【京东】https://item.jd.com/100152311477.html 点击查看",
        "非法链接abc",
        "",
        None,
    ]

    for tc in test_cases:
        sku, m_url = to_mobile_h5_url(tc)
        print(f"输入: {tc!r:45} -> SKU: {sku!r:15} -> H5: {m_url}")
