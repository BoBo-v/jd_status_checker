# -*- coding: utf-8 -*-
"""
京东商品状态巡检系统 - 主程序入口
支持命令行参数调用与无参数双击交互式运行
全面支持多工作表（分表）选择与 100% 原表图片安全保护
"""
import sys
import os
import argparse
from pathlib import Path
from typing import List

from config import (
    DEFAULT_HEADLESS,
    AUTO_PAUSE_ON_CAPTCHA,
    AUTO_RESTART_ON_LOGIN,
    MAX_LOGIN_RESTART_RETRIES,
    BYPASS_PROXY,
    DELAY_MIN,
    DELAY_MAX,
    SAVE_INTERVAL,
)
from inspector import JDStatusInspector
from excel_handler import inspect_excel_sheets, parse_sheet_selection
from create_sample_excel import create_sample_excel
from ui_menu import select_menu, select_checkbox_menu


def scan_nearby_excels() -> List[Path]:
    """
    智能搜索当前目录、程序所在目录与桌面上的 .xlsx 文件，供用户快捷选择
    """
    found = []
    search_dirs = [Path.cwd(), Path(__file__).resolve().parent]
    if getattr(sys, "frozen", False):
        search_dirs.append(Path(sys.executable).resolve().parent)

    desktop_dir = Path.home() / "Desktop"
    if desktop_dir.exists():
        search_dirs.append(desktop_dir)

    for d in search_dirs:
        if d.exists():
            for p in d.glob("*.xlsx"):
                if not p.name.startswith("~$") and not p.name.endswith("_巡检结果.xlsx"):
                    if p not in found:
                        found.append(p)

    return found


def interactive_menu() -> dict:
    """
    交互式向导，支持使用 ↑ / ↓ 方向键上下移动选择，方便非命令行用户直接双击运行
    """
    print("\n" + "=" * 65)
    print("      欢迎使用 京东商品上下架/风控自动化巡检系统 (DrissionPage)")
    print("      [全新升级] 100% 保护原表主图不丢失 | 支持超大表格与多工作表")
    print("=" * 65)

    excels = scan_nearby_excels()
    selected_excel = None

    if excels:
        excel_choices = [
            f"📄 {f.name} ({f.stat().st_size / (1024 * 1024):.1f} MB)"
            for f in excels
        ]
        excel_choices.append("➕ 手动输入其他 Excel 文件的完整绝对路径")
        excel_choices.append("🧪 快速生成一份测试样例表格 (sample_jd_goods.xlsx)")

        idx = select_menu(
            "【步骤 1/4】扫描到以下表格，请选择要巡检的 Excel 文件：",
            excel_choices,
            default_idx=0
        )
        if idx < len(excels):
            selected_excel = str(excels[idx])
        elif idx == len(excels):
            while True:
                path_input = input("请输入 Excel 文件的完整绝对路径: ").strip().strip('"').strip("'")
                if Path(path_input).exists():
                    selected_excel = path_input
                    break
                else:
                    print("[!] 该文件不存在，请重新输入！")
        else:
            selected_excel = create_sample_excel()
    else:
        empty_choices = [
            "🧪 自动生成一份测试样例表格 (sample_jd_goods.xlsx) 并开始测试",
            "➕ 手动输入现有表格完整绝对路径"
        ]
        idx = select_menu("未在当前目录或桌面上自动发现表格，请选择：", empty_choices)
        if idx == 0:
            selected_excel = create_sample_excel()
        else:
            while True:
                path_input = input("请输入 Excel 文件的完整绝对路径: ").strip().strip('"').strip("'")
                if Path(path_input).exists():
                    selected_excel = path_input
                    break
                else:
                    print("[!] 该文件不存在，请重新输入！")

    # 探测该表格内部包含的工作表（分表）
    target_sheets = "ALL"
    sheet_infos = inspect_excel_sheets(selected_excel)
    all_sheet_names = [info["name"] for info in sheet_infos]

    if len(sheet_infos) > 1:
        sheet_mode_options = [
            f"🌟 巡检全部工作表（推荐：从第 1 张开始跑完全部 {len(sheet_infos)} 个分表）",
            "🚀 从指定分表开始（依次巡检选中的分表及后续所有分表至最后）",
            "🎯 仅巡检单张分表（只处理选中的单张分表）",
            "☑️  自定义自由勾选（使用方向键移动，按空格键自由勾选/反选任意分表）"
        ]

        mode_idx = select_menu(
            f"【步骤 2/4】检测到表格包含 {len(sheet_infos)} 个工作表（分表），请选择巡检方式：",
            sheet_mode_options,
            default_idx=0
        )

        if mode_idx == 0:
            target_sheets = "ALL"
        elif mode_idx == 1:
            # 从指定分表开始
            sheet_items = [
                f"[{i}] {info['name']} {'[含商品链接]' if info['has_goods'] else '[未发现链接]'}"
                for i, info in enumerate(sheet_infos, start=1)
            ]
            s_idx = select_menu(
                "请使用 ↑ / ↓ 键选择从哪一张分表开始（将从该分表依次巡检后续所有分表）：",
                sheet_items,
                default_idx=0
            )
            target_sheets = all_sheet_names[s_idx:]
            preview = " -> ".join([f"【{s}】" for s in target_sheets[:4]])
            if len(target_sheets) > 4:
                preview += f" ... 等共 {len(target_sheets)} 个分表"
            print(f"[√] 已设定从第 {s_idx + 1} 张分表开始巡检：{preview}\n")
        elif mode_idx == 2:
            # 仅单张分表
            sheet_items = [
                f"[{i}] {info['name']} {'[含商品链接]' if info['has_goods'] else '[未发现链接]'}"
                for i, info in enumerate(sheet_infos, start=1)
            ]
            s_idx = select_menu(
                "请使用 ↑ / ↓ 键选择要单独巡检的工作表：",
                sheet_items,
                default_idx=0
            )
            target_sheets = [all_sheet_names[s_idx]]
            print(f"[√] 已设定仅巡检分表：【{all_sheet_names[s_idx]}】\n")
        elif mode_idx == 3:
            # 自定义勾选
            cb_items = [
                f"{info['name']} {'[含商品链接]' if info['has_goods'] else '[未发现链接]'}"
                for info in sheet_infos
            ]
            default_chk = [i for i, info in enumerate(sheet_infos) if info["has_goods"]]
            chosen_indices = select_checkbox_menu(
                "请勾选需要巡检的分表：",
                cb_items,
                default_checked=default_chk
            )
            target_sheets = [all_sheet_names[i] for i in chosen_indices]

    # 步骤 3/4：浏览器运行模式
    browser_options = [
        "🖥️  有界面模式（推荐：弹出浏览器窗口，直观可见，安全度最高）",
        "👻 无头后台模式（静默后台运行，不弹窗）"
    ]
    b_idx = select_menu("【步骤 3/4】请选择浏览器运行模式：", browser_options, default_idx=0)
    headless = (b_idx == 1)

    # 步骤 4/4：断点续传
    resume_options = [
        "⚡ 启用断点续传（自动跳过已有检测结果的行，省时高效）",
        "🔄 重新全量检测（忽略历史结果，对全表重新巡检）"
    ]
    r_idx = select_menu("【步骤 4/4】请选择断点续传设置：", resume_options, default_idx=0)
    skip_completed = (r_idx == 0)

    return {
        "excel_path": selected_excel,
        "target_sheets": target_sheets,
        "headless": headless,
        "skip_completed": skip_completed,
    }


def main():
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8")
            sys.stderr.reconfigure(encoding="utf-8")
        except Exception:
            pass

    parser = argparse.ArgumentParser(description="京东商品上下架状态自动化巡检工具 (DrissionPage)")
    parser.add_argument("--excel", "-e", type=str, help="待检测的 Excel 文件绝对路径")
    parser.add_argument("--sheet", "-s", type=str, default="ALL", help="指定分表名称、序号范围（如 '3-' 从第3张表开始, '2-4', '1,3', 'ALL'）")
    parser.add_argument("--start-sheet", type=str, default=None, help="从指定分表开始巡检（可输入分表序号如 3 或分表名称，将从该分表一直跑至最后）")
    parser.add_argument("--headless", action="store_true", help="启用无头静默模式（默认显示浏览器窗口）")
    parser.add_argument("--no-bypass-proxy", action="store_true", help="不绕过代理（使用系统默认代理网络）")
    parser.add_argument("--no-auto-restart", action="store_true", help="禁用遇登录拦截时自动关闭并重启浏览器")
    parser.add_argument("--login-retries", type=int, default=MAX_LOGIN_RESTART_RETRIES, help="遇到登录拦截时自动重启浏览器的最大重试次数")
    parser.add_argument("--rerun-all", action="store_true", help="强制重新检测全部数据，不跳过已完成项")
    parser.add_argument("--recheck-errors", action="store_true", help="仅重跑状态为待查或异常的记录")
    parser.add_argument("--enable-screenshot", action="store_true", help="启用异常页面截图（默认关闭，采用纯文本日志记录以节省磁盘空间）")
    parser.add_argument("--no-micro-scroll", action="store_true", help="禁用拟人微动滑动（默认开启，隔几个商品随机微滑以防止风控滑块）")
    parser.add_argument("--delay-min", type=float, default=DELAY_MIN, help="最小停顿延时（秒）")
    parser.add_argument("--delay-max", type=float, default=DELAY_MAX, help="最大停顿延时（秒）")
    parser.add_argument("--save-interval", type=int, default=SAVE_INTERVAL, help="回写保存间隔（行数）")
    parser.add_argument("--generate-sample", action="store_true", help="在当前目录生成 sample_jd_goods.xlsx 测试文件并退出")

    args = parser.parse_args()

    if args.generate_sample:
        create_sample_excel()
        return

    if args.enable_screenshot:
        import config
        config.ENABLE_SCREENSHOT = True

    if args.no_micro_scroll:
        import config
        config.ENABLE_MICRO_SCROLL = False

    bypass_proxy = not args.no_bypass_proxy
    auto_restart = not args.no_auto_restart

    if args.excel:
        excel_path = args.excel
        if args.start_sheet:
            target_sheets = f"{args.start_sheet}-"
        else:
            target_sheets = args.sheet
        headless = args.headless
        skip_completed = not args.rerun_all
        recheck_errors = args.recheck_errors
    else:
        # 进入交互式向导
        config = interactive_menu()
        excel_path = config["excel_path"]
        target_sheets = config["target_sheets"]
        headless = config["headless"]
        skip_completed = config["skip_completed"]
        recheck_errors = False

    inspector = JDStatusInspector(
        excel_path=excel_path,
        target_sheets=target_sheets,
        headless=headless,
        auto_pause_on_captcha=AUTO_PAUSE_ON_CAPTCHA,
        auto_restart_on_login=auto_restart,
        max_login_retries=args.login_retries,
        bypass_proxy=bypass_proxy,
        skip_completed=skip_completed,
        recheck_errors_only=recheck_errors,
        delay_range=(args.delay_min, args.delay_max),
        save_interval=args.save_interval,
    )

    inspector.run()


if __name__ == "__main__":
    main()
