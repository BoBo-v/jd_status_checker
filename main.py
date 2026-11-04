# -*- coding: utf-8 -*-
"""
京东商品状态巡检系统 - 主程序入口
支持命令行参数调用与无参数双击交互式运行
"""
import sys
import os
import argparse
from pathlib import Path
from typing import List

from config import (
    DEFAULT_HEADLESS,
    DEFAULT_INCOGNITO,
    AUTO_PAUSE_ON_CAPTCHA,
    CLEAR_CACHE_ON_RISK,
    DELAY_MIN,
    DELAY_MAX,
    SAVE_INTERVAL,
)
from inspector import JDStatusInspector
from create_sample_excel import create_sample_excel


def scan_nearby_excels() -> List[Path]:
    """
    智能搜索当前目录与桌面上的 .xlsx 文件，供用户快捷选择
    """
    found = []
    current_dir = Path(__file__).resolve().parent

    # 1. 当前目录及上级目录中的 xlsx
    for p in current_dir.glob("*.xlsx"):
        if not p.name.startswith("~$"):
            found.append(p)

    # 2. 桌面上的 xlsx
    desktop_dir = Path.home() / "Desktop"
    if desktop_dir.exists():
        for p in desktop_dir.glob("*.xlsx"):
            if not p.name.startswith("~$"):
                if p not in found:
                    found.append(p)

    return found


def interactive_menu() -> dict:
    """
    交互式向导，方便非命令行用户直接双击运行
    """
    print("\n" + "=" * 65)
    print("      欢迎使用 京东商品上下架/风控自动化巡检系统 (DrissionPage)")
    print("=" * 65)

    excels = scan_nearby_excels()
    selected_excel = None

    if excels:
        print("\n扫描到以下 Excel 表格文件：")
        for i, f in enumerate(excels, start=1):
            print(f"  [{i}] {f.name}  ({f.parent})")
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

    # 询问是否无头运行
    hl_input = input("\n是否使用无头模式（后台静默运行，不弹窗）？[y/N，默认 N (弹出窗口方便手动过验)]: ").strip().lower()
    headless = True if hl_input in ("y", "yes") else False

    # 询问是否开启无痕模式
    ic_input = input("是否启用无痕浏览器模式（不保留历史与Cookie）？[y/N，默认 N (持久化保留登录态)]: ").strip().lower()
    incognito = True if ic_input in ("y", "yes") else False

    # 询问是否断点续传
    skip_input = input("是否启用断点续传（自动跳过已有 0 或 1 结果的行）？[Y/n，默认 Y]: ").strip().lower()
    skip_completed = False if skip_input in ("n", "no") else True

    return {
        "excel_path": selected_excel,
        "headless": headless,
        "incognito": incognito,
        "skip_completed": skip_completed,
    }


def main():
    # 强制将标准输出转为 utf-8，解决 Windows CMD/PowerShell 乱码问题
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8")
            sys.stderr.reconfigure(encoding="utf-8")
        except Exception:
            pass

    parser = argparse.ArgumentParser(description="京东商品上下架状态自动化巡检工具 (DrissionPage)")
    parser.add_argument("--excel", "-e", type=str, help="待检测的 Excel 文件绝对路径")
    parser.add_argument("--sheet", "-s", type=str, default=None, help="指定工作表名称（默认活动表）")
    parser.add_argument("--headless", action="store_true", help="启用无头静默模式（默认显示浏览器窗口）")
    parser.add_argument("--incognito", action="store_true", help="启用无痕模式（不保留缓存与Cookie）")
    parser.add_argument("--no-pause", action="store_true", help="遇到安全验证不暂停，直接标记为待查并跳过")
    parser.add_argument("--clear-on-risk", action="store_true", help="触发风控被拦截时自动清空会话 Cookie 与缓存")
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

    if args.excel:
        excel_path = args.excel
        headless = args.headless
        incognito = args.incognito
        skip_completed = not args.rerun_all
        recheck_errors = args.recheck_errors
        auto_pause = not args.no_pause
        clear_cache = args.clear_on_risk
    else:
        # 进入交互式向导
        config = interactive_menu()
        excel_path = config["excel_path"]
        headless = config["headless"]
        incognito = config["incognito"]
        skip_completed = config["skip_completed"]
        recheck_errors = False
        auto_pause = AUTO_PAUSE_ON_CAPTCHA
        clear_cache = CLEAR_CACHE_ON_RISK

    inspector = JDStatusInspector(
        excel_path=excel_path,
        sheet_name=args.sheet,
        headless=headless,
        incognito=incognito,
        auto_pause_on_captcha=auto_pause,
        clear_cache_on_risk=clear_cache,
        skip_completed=skip_completed,
        recheck_errors_only=recheck_errors,
        delay_range=(args.delay_min, args.delay_max),
        save_interval=args.save_interval,
    )

    inspector.run()


if __name__ == "__main__":
    main()
