# -*- coding: utf-8 -*-
"""
终端交互式上下箭头菜单模块
纯 Python 原生标准库实现，零第三方依赖，兼容 Windows 10/11 与 macOS/Linux
支持 ↑ / ↓ 方向键移动光标、Enter 回车确认、空格键多选勾选与数字快捷键
"""
import sys
import os
import shutil
from typing import List, Optional, Tuple


def _enable_windows_ansi():
    """在 Windows 终端中开启 ANSI 虚拟终端序列支持"""
    if sys.platform == "win32":
        try:
            import ctypes
            kernel32 = ctypes.windll.kernel32
            hOut = kernel32.GetStdHandle(-11)  # STD_OUTPUT_HANDLE
            mode = ctypes.c_ulong()
            kernel32.GetConsoleMode(hOut, ctypes.byref(mode))
            # 0x0004 = ENABLE_VIRTUAL_TERMINAL_PROCESSING
            kernel32.SetConsoleMode(hOut, mode.value | 0x0004)
        except Exception:
            pass


def _hide_cursor():
    sys.stdout.write("\033[?25l")
    sys.stdout.flush()


def _show_cursor():
    sys.stdout.write("\033[?25h")
    sys.stdout.flush()


def _get_key() -> str:
    """读取单次按键输入"""
    if sys.platform == "win32":
        import msvcrt
        ch = msvcrt.getwch()
        if ch in ('\x00', '\xe0'):
            ch2 = msvcrt.getwch()
            if ch2 == 'H':
                return 'UP'
            elif ch2 == 'P':
                return 'DOWN'
            elif ch2 == 'K':
                return 'LEFT'
            elif ch2 == 'M':
                return 'RIGHT'
            return ''
        elif ch in ('\r', '\n'):
            return 'ENTER'
        elif ch == ' ':
            return 'SPACE'
        elif ch == '\x03':  # Ctrl+C
            raise KeyboardInterrupt
        elif ch == '\x1b':  # ESC
            return 'ESC'
        return ch
    else:
        import tty, termios
        fd = sys.stdin.fileno()
        old_settings = termios.tcgetattr(fd)
        try:
            tty.setraw(fd)
            ch = sys.stdin.read(1)
            if ch == '\x1b':
                ch2 = sys.stdin.read(1)
                if ch2 == '[':
                    ch3 = sys.stdin.read(1)
                    if ch3 == 'A':
                        return 'UP'
                    elif ch3 == 'B':
                        return 'DOWN'
                    elif ch3 == 'C':
                        return 'RIGHT'
                    elif ch3 == 'D':
                        return 'LEFT'
                return 'ESC'
            elif ch in ('\r', '\n'):
                return 'ENTER'
            elif ch == ' ':
                return 'SPACE'
            elif ch == '\x03':
                raise KeyboardInterrupt
            return ch
        finally:
            termios.tcsetattr(fd, termios.TCSADRAIN, old_settings)


def select_menu(
    title: str,
    options: List[str],
    default_idx: int = 0,
    help_text: str = "使用 ↑ / ↓ 方向键移动，Enter 回车确认选择"
) -> int:
    """
    单选菜单：支持使用 ↑ / ↓ 键上下移动光标，Enter 确认选择
    如果当前环境不支持交互式控制台，自动优雅降级为传统输入
    """
    _enable_windows_ansi()

    # 降级模式：如果不在标准交互式终端中
    if not sys.stdin.isatty():
        print(f"\n{title}")
        for i, opt in enumerate(options, start=1):
            print(f"  [{i}] {opt}")
        while True:
            choice = input(f"请输入选择序号 [1-{len(options)}, 默认 {default_idx + 1}]: ").strip()
            if not choice:
                return default_idx
            if choice.isdigit() and 1 <= int(choice) <= len(options):
                return int(choice) - 1
            print("[!] 输入无效，请重新输入！")

    current_idx = max(0, min(default_idx, len(options) - 1))
    term_width = shutil.get_terminal_size((80, 20)).columns

    print(f"\n{title}")
    if help_text:
        print(f"\033[90m{help_text}\033[0m")

    # 首次渲染选项
    for i, opt in enumerate(options):
        # 截断过长文字避免终端换行计算错位
        display_opt = opt if len(opt) < term_width - 10 else opt[:term_width - 13] + "..."
        if i == current_idx:
            print(f"  \033[1;36m❯ {display_opt}\033[0m")
        else:
            print(f"    {display_opt}")

    _hide_cursor()
    try:
        while True:
            key = _get_key()
            if key == 'UP':
                current_idx = (current_idx - 1) % len(options)
            elif key == 'DOWN':
                current_idx = (current_idx + 1) % len(options)
            elif key == 'ENTER':
                break
            elif key.isdigit() and 1 <= int(key) <= len(options):
                # 允许直接按数字键直接跳到对应项
                current_idx = int(key) - 1
                break
            else:
                continue

            # 光标上移重绘选项区域
            sys.stdout.write(f"\033[{len(options)}A")
            for i, opt in enumerate(options):
                display_opt = opt if len(opt) < term_width - 10 else opt[:term_width - 13] + "..."
                # 清除整行
                sys.stdout.write("\033[2K\r")
                if i == current_idx:
                    sys.stdout.write(f"  \033[1;36m❯ {display_opt}\033[0m\n")
                else:
                    sys.stdout.write(f"    {display_opt}\n")
            sys.stdout.flush()

    finally:
        _show_cursor()

    print(f"\033[32m[√] 已选择：{options[current_idx]}\033[0m\n")
    return current_idx


def select_checkbox_menu(
    title: str,
    options: List[str],
    default_checked: Optional[List[int]] = None,
    help_text: str = "使用 ↑ / ↓ 移动，[空格键 Space] 勾选/取消，[A 键] 全选/全不选，[Enter 回车] 确认"
) -> List[int]:
    """
    复选菜单：支持使用 ↑ / ↓ 键移动光标，空格键勾选/取消，Enter 确认
    """
    _enable_windows_ansi()

    if not sys.stdin.isatty():
        print(f"\n{title}")
        for i, opt in enumerate(options, start=1):
            print(f"  [{i}] {opt}")
        choice = input(f"请输入要选择的分表序号（逗号分隔，如 1,2,3，直接回车全选）: ").strip()
        if not choice:
            return list(range(len(options)))
        parts = [p.strip() for p in choice.replace("，", ",").split(",") if p.strip()]
        res = []
        for p in parts:
            if p.isdigit() and 1 <= int(p) <= len(options):
                res.append(int(p) - 1)
        return res if res else list(range(len(options)))

    checked = [False] * len(options)
    if default_checked:
        for idx in default_checked:
            if 0 <= idx < len(checked):
                checked[idx] = True
    else:
        # 默认全部勾选
        checked = [True] * len(options)

    current_idx = 0
    term_width = shutil.get_terminal_size((80, 20)).columns

    print(f"\n{title}")
    if help_text:
        print(f"\033[90m{help_text}\033[0m")

    # 首次渲染
    for i, opt in enumerate(options):
        mark = "\033[32m[√]\033[0m" if checked[i] else "[ ]"
        display_opt = opt if len(opt) < term_width - 15 else opt[:term_width - 18] + "..."
        if i == current_idx:
            print(f"  \033[1;36m❯ {mark} {display_opt}\033[0m")
        else:
            print(f"    {mark} {display_opt}")

    _hide_cursor()
    try:
        while True:
            key = _get_key()
            if key == 'UP':
                current_idx = (current_idx - 1) % len(options)
            elif key == 'DOWN':
                current_idx = (current_idx + 1) % len(options)
            elif key == 'SPACE':
                checked[current_idx] = not checked[current_idx]
            elif key in ('a', 'A'):
                # 快捷全选 / 全不选
                all_checked = all(checked)
                checked = [not all_checked] * len(options)
            elif key == 'ENTER':
                # 如果一项都没选，提示至少选择一项
                if not any(checked):
                    continue
                break
            else:
                continue

            # 光标上移重绘选项区域
            sys.stdout.write(f"\033[{len(options)}A")
            for i, opt in enumerate(options):
                mark = "\033[32m[√]\033[0m" if checked[i] else "[ ]"
                display_opt = opt if len(opt) < term_width - 15 else opt[:term_width - 18] + "..."
                sys.stdout.write("\033[2K\r")
                if i == current_idx:
                    sys.stdout.write(f"  \033[1;36m❯ {mark} {display_opt}\033[0m\n")
                else:
                    sys.stdout.write(f"    {mark} {display_opt}\n")
            sys.stdout.flush()

    finally:
        _show_cursor()

    selected_indices = [i for i, c in enumerate(checked) if c]
    selected_names = [options[i] for i in selected_indices]
    print(f"\033[32m[√] 已确认勾选 {len(selected_indices)} 个分表：{', '.join(selected_names)}\033[0m\n")
    return selected_indices
