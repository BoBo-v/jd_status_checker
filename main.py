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
from excel_handler import inspect_excel_sheets
from create_sample_excel import create_sample_excel


def scan_nearby_excels() -> List[Path]:
    """
    智能搜索当前目录与桌面上的 .xlsx 文件，供用户快捷选择
    """
    found = []
    current_dir = Path(__file__).resolve().parent

    # 1. 当前目录及上级目录中的 xlsx
    for p in current_dir.glob("*.xlsx"):
        if not p.name.startswith("~$") and not p.name.endswith("_巡检结果.xlsx"):
            found.append(p)

    # 2. 桌面上的 xlsx
    desktop_dir = Path.home() / "Desktop"
    if desktop_dir.exists():
        for p in desktop_dir.glob("*.xlsx"):
            if not p.name.startswith("~$") and not p.name.endswith("_巡检结果.xlsx"):
                if p not in found:
                    found.append(p)

    return found


def interactive_menu() -> dict:
    """
    交互式向导，方便非命令行用户直接双击运行
    """
    print("\n" + "=" * 65)
    print("      欢迎使用 京东商品上下架/风控自动化巡检系统 (DrissionPage)")
    print("      [全新升级] 100% 保护原表主图不丢失 | 支持超大表格与多工作表")
    print("=" * 65)

    excels = scan_nearby_excels()
    selected_excel = None

    if excels:
        print("\n扫描到以下 Excel 表格文件：")
        for i, f in enumerate(excels, start=1):
            size_mb = f.stat().st_size / (1024 * 1024)
            print(f"  [{i}] {f.name} ({size_mb:.1f} MB)  ({f.parent})")
        print("  [S] 快速生成一份测试样例表格 (sample_jd_goods.xlsx)")
        print("  [0] 手动输入其他完整文件路径")

        while True:
            choice = input("\n请选择要处理的表格序号或选项 [默认 1]: ").strip()
            if not choice:
                selected_excel = str(excels[0])
                break
            elif choice.upper() == "S":
                sample_file = create_sample_excel()
                selected_excel = sample_file
                break
            elif choice == "0":
                path_input = input("请输入 Excel 文件的完整绝对路径: ").strip().strip('"').strip("'")
                if Path(path_input).exists():
                    selected_excel = path_input
                    break
                else:
                    print("[!] 该文件不存在，请重新输入！")
            elif choice.isdigit() and 1 <= int(choice) <= len(excels):
                selected_excel = str(excels[int(choice) - 1])
                break
            else:
                print("[!] 输入无效，请重新选择！")
    else:
        print("\n未在当前目录或桌面上自动发现 .xlsx 表格。")
        print("  [S] 自动生成一份测试样例表格并开始测试")
        print("  [0] 手工输入现有表格路径")
        choice = input("请选择 [S/0，默认 S]: ").strip()
        if choice == "0":
            path_input = input("请输入 Excel 文件的完整绝对路径: ").strip().strip('"').strip("'")
            selected_excel = path_input
        else:
            selected_excel = create_sample_excel()

    # 探测该表格内部包含的工作表（分表）
    target_sheets = "ALL"
    sheet_infos = inspect_excel_sheets(selected_excel)
    if len(sheet_infos) > 1:
        print(f"\n检测到该表格包含 {len(sheet_infos)} 个工作表（分表）：")
        print("  [A] 巡检所有包含商品的分表（合并处理，推荐）")
        for idx, info in enumerate(sheet_infos, start=1):
            tag = "[含商品链接]" if info["has_goods"] else "[未发现链接]"
            print(f"  [{idx}] {info['name']} {tag}")

        while True:
            s_choice = input(f"请选择要巡检的工作表序号 [A/{'/'.join(str(i) for i in range(1, len(sheet_infos)+1))}，默认 A]: ").strip().upper()
            if not s_choice or s_choice == "A":
                target_sheets = "ALL"
                break
            elif s_choice.isdigit() and 1 <= int(s_choice) <= len(sheet_infos):
                target_sheets = sheet_infos[int(s_choice) - 1]["name"]
                break
            else:
                print("[!] 输入无效，请重新选择！")

    # 询问是否无头运行
    hl_input = input("\n是否使用无头静默模式（后台运行不弹窗）？[y/N，默认 N (弹出窗口更安全)]: ").strip().lower()
    headless = True if hl_input in ("y", "yes") else False

    # 询问是否断点续传
    skip_input = input("是否启用断点续传（自动跳过已有检测结果的行）？[Y/n，默认 Y]: ").strip().lower()
    skip_completed = False if skip_input in ("n", "no") else True

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
    parser.add_argument("--sheet", "-s", type=str, default="ALL", help="指定分表名称或 ALL（默认 ALL 扫描所有包含链接的分表）")
    parser.add_argument("--headless", action="store_true", help="启用无头静默模式（默认显示浏览器窗口）")
    parser.add_argument("--no-bypass-proxy", action="store_true", help="不绕过代理（使用系统默认代理网络）")
    parser.add_argument("--no-auto-restart", action="store_true", help="禁用遇登录拦截时自动关闭并重启浏览器")
    parser.add_argument("--login-retries", type=int, default=MAX_LOGIN_RESTART_RETRIES, help="遇到登录拦截时自动重启浏览器的最大重试次数")
    parser.add_argument("--rerun-all", action="store_true", help="强制重新检测全部数据，不跳过已完成项")
    parser.add_argument("--recheck-errors", action="store_true", help="仅重跑状态为待查或异常的记录")
    parser.add_argument("--delay-min", type=float, default=DELAY_MIN, help="最小停顿延时（秒）")
    parser.add_argument("--delay-max", type=float, default=DELAY_MAX, help="最大停顿延时（秒）")
    parser.add_argument("--save-interval", type=int, default=SAVE_INTERVAL, help="回写保存间隔（行数）")
    parser.add_argument("--generate-sample", action="store_true", help="在当前目录生成 sample_jd_goods.xlsx 测试文件并退出")

    args = parser.parse_args()

    if args.generate_sample:
        create_sample_excel()
        return

    bypass_proxy = not args.no_bypass_proxy
    auto_restart = not args.no_auto_restart

    if args.excel:
        excel_path = args.excel
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
