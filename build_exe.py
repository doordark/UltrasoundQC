#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_exe.py — Ultrasound QC Engine 一键打包脚本
=================================================
使用 PyInstaller 将 qc_engine.py 打包为单文件 exe（默认同时产出两个版本）:

    dist/UltrasoundQCEngine.exe        无控制台窗口版（HTTP 常驻服务，双击/开机自启用）
    dist/UltrasoundQCEngine-CLI.exe    控制台版（命令行批量检测，可看输出）

用法:
    python build_exe.py                     # 打包两个版本
    python build_exe.py --console-only      # 仅打包控制台版
    python build_exe.py --windowed-only     # 仅打包无控制台版
    python build_exe.py --icon logo.ico     # 指定图标
    python build_exe.py --distclean         # 打包前清理 build/dist/spec
"""

import argparse
import subprocess
import sys
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
EXCLUDES = ["pandas", "numpy", "openpyxl", "matplotlib", "tkinter", "PyQt5",
            "PySide2", "scipy", "requests", "PIL"]


def ensure_pyinstaller() -> None:
    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        print("[build] 未检测到 PyInstaller，正在自动安装 ...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])


def build(name: str, windowed: bool, icon: str | None, clean: bool) -> Path:
    args = [sys.executable, "-m", "PyInstaller", "--onefile", "--noconfirm",
            "--name", name]
    if windowed:
        args.append("--noconsole")
    if clean:
        args.append("--clean")
    if icon:
        args += ["--icon", icon]
    for mod in EXCLUDES:  # 引擎只用标准库，排除大件防止意外吸入，控制体积
        args += ["--exclude-module", mod]
    args.append(str(HERE / "qc_engine.py"))

    print(f"[build] 目标: {name}.exe ({'无控制台' if windowed else '控制台'})")
    print("[build] " + " ".join(args))
    subprocess.check_call(args, cwd=str(HERE))
    out = HERE / "dist" / f"{name}.exe"
    size_mb = out.stat().st_size / 1024 / 1024 if out.exists() else 0
    print(f"[build] 完成: {out} ({size_mb:.1f} MB)")
    return out


def main() -> None:
    ap = argparse.ArgumentParser(description="Ultrasound QC Engine 打包脚本")
    ap.add_argument("--console-only", action="store_true", help="仅打包控制台版")
    ap.add_argument("--windowed-only", action="store_true", help="仅打包无控制台版")
    ap.add_argument("--icon", default=None, help="exe 图标路径 (.ico)")
    ap.add_argument("--distclean", action="store_true", help="打包前删除 build/dist 与 spec 缓存")
    args = ap.parse_args()

    ensure_pyinstaller()
    bits = 64 if sys.maxsize > 2 ** 32 else 32
    print(f"[build] Python: {sys.version.split()[0]} ({bits}bit) | 开始时间 {datetime.now().strftime('%H:%M:%S')}")

    if args.distclean:
        import shutil
        for d in ("build", "dist"):
            shutil.rmtree(HERE / d, ignore_errors=True)
        (HERE / "UltrasoundQCEngine.spec").unlink(missing_ok=True)
        (HERE / "UltrasoundQCEngine-CLI.spec").unlink(missing_ok=True)
        print("[build] 已清理 build/dist/spec")

    outputs = []
    if not args.console_only:
        outputs.append(build("UltrasoundQCEngine", windowed=True, icon=args.icon, clean=True))
    if not args.windowed_only:
        outputs.append(build("UltrasoundQCEngine-CLI", windowed=False, icon=args.icon, clean=True))

    print("=" * 56)
    print("打包完成，产物列表:")
    for o in outputs:
        print(f"  {o}")
    print("提示: rules.json 请与 exe 放在同一目录（不打包进 exe，便于医院现场改规则免重编译）")
    print("提示: 需支持 Win7/32 位系统时，请改用 32 位 Python 3.8.10 + PyInstaller 5.13.2 重新打包")
    print("=" * 56)


if __name__ == "__main__":
    main()
