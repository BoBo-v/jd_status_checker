# -*- coding: utf-8 -*-
"""
一键打包脚本 (Windows 绿色免安装版)
使用 Python 原生调用 PyInstaller，彻底避免 Windows .bat 批处理编码导致的语法解析崩溃
"""
import sys
import shutil
from pathlib import Path
import PyInstaller.__main__

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

def build():
    base_dir = Path(__file__).resolve().parent

    print("\n" + "=" * 65)
    print("    正在打包 [京东商品状态巡检工具] Windows 绿色免安装版...")
    print("=" * 65)

    build_dir = base_dir / "build"
    dist_dir = base_dir / "dist"

    # 清理旧构建缓存
    print("[*] 正在清理旧的构建缓存...")
    if build_dir.exists():
        shutil.rmtree(build_dir, ignore_errors=True)
    if dist_dir.exists():
        shutil.rmtree(dist_dir, ignore_errors=True)

    # 构造打包参数
    main_script = base_dir / "main.py"
    args = [
        str(main_script),
        "--name=京东商品状态巡检工具",
        "--onedir",
        "--noconfirm",
        "--clean",
        "--collect-all=DrissionPage",
        "--copy-metadata=DrissionPage",
        "--hidden-import=openpyxl",
        f"--distpath={dist_dir}",
        f"--workpath={build_dir}",
        f"--specpath={base_dir}",
    ]

    print("[*] 正在执行深度依赖打包（大约需要 30 秒，请稍候）...")
    try:
        PyInstaller.__main__.run(args)

        out_dir = dist_dir / "京东商品状态巡检工具"
        sample_file = base_dir / "sample_jd_goods.xlsx"
        if sample_file.exists() and out_dir.exists():
            shutil.copy2(sample_file, out_dir / "sample_jd_goods.xlsx")

        print("\n" + "=" * 65)
        print("[√] 打包成功！独立免安装绿色包已输出至：")
        print(f"    {out_dir}")
        print("\n【分发使用说明】：")
        print("1. 你只需把上面的【京东商品状态巡检工具】文件夹压缩为 .zip 发给其他人；")
        print("2. 对方电脑无需安装任何 Python 或运行库，解压后双击即可使用！")
        print("=" * 65 + "\n")
    except Exception as e:
        print(f"\n[X] 打包失败，错误信息: {e}")
        sys.exit(1)


if __name__ == "__main__":
    build()
