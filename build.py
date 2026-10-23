#!/usr/bin/env python3
"""打包成发给同事的"XL FreightPilot"文件夹和 zip（用 Nuitka 编译，同事那边不需要 Python）。

用法（只在你自己的电脑上）:
    pip install nuitka
    python build.py

结果:
    dist/XL FreightPilot/       整个文件夹就是发给同事的程序
    dist/XL FreightPilot.zip    同一份内容的压缩包
注意: 不会把你的登录信息（dachser_profile）打包进去。

为什么要用 D:\\liuxu_build:
    Nuitka 分析 DLL 依赖的工具不支持中文路径（用户名"刘旭"会变成乱码导致编译失败），
    所以把源码复制到纯英文路径、在那里单独建一个打包用的虚拟环境来编译。
    可以用环境变量 LIUXU_BUILD_DIR 换成别的纯英文路径。
"""
import os
import shutil
import subprocess
import sys
import zipfile
from importlib.metadata import version
from pathlib import Path

ROOT = Path(__file__).resolve().parent
NAME = "XL FreightPilot"
STAGE = Path(os.environ.get("LIUXU_BUILD_DIR") or (r"D:\liuxu_build" if Path("D:/").exists() else r"C:\liuxu_build"))
DIST = ROOT / "dist" / NAME
EXE_TMP = "LiuXuScript.exe"     # 编译时用英文名，打包完再改名
SOURCES = ["gui.py", "app_ui.py", "license_check.py", "dachser_fill.py", "create_orders.py", "match_orders.py",
           "pallet_plan.py"]
EXTRA_FILES = ["pallet_params.json", "使用说明.txt"]
PACKAGES = ["nuitka", f"playwright=={version('playwright')}", f"openpyxl=={version('openpyxl')}"]


def build_python():
    """在纯英文路径下准备一个打包专用的虚拟环境，返回它的 python.exe。"""
    if not str(STAGE).isascii():
        raise SystemExit(f"打包目录必须是纯英文路径: {STAGE}")
    py = STAGE / "venv" / "Scripts" / "python.exe"
    if not py.exists():
        base = getattr(sys, "_base_executable", sys.executable)
        print(f"创建打包用虚拟环境: {STAGE / 'venv'}（基于 {base}）")
        subprocess.run([base, "-m", "venv", str(STAGE / "venv")], check=True)
    subprocess.run([str(py), "-m", "pip", "install", "-q", *PACKAGES], check=True)
    return py


def compile_exe(py):
    src = STAGE / "src"
    if src.exists():
        shutil.rmtree(src)
    src.mkdir(parents=True)
    for f in SOURCES:
        shutil.copy2(ROOT / f, src / f)
    out = STAGE / "build"
    env = {**os.environ, "NUITKA_CACHE_DIR": str(STAGE / "cache")}   # 编译器等下载缓存也放英文路径
    cmd = [str(py), "-m", "nuitka",
           "--standalone", "--assume-yes-for-downloads",
           "--enable-plugin=tk-inter",
           "--windows-console-mode=disable",
           "--include-package=playwright",      # Nuitka 的 playwright 插件会自动带上 driver（node.exe 等）
           f"--output-dir={out}", f"--output-filename={EXE_TMP}",
           str(src / "gui.py")]
    print(" ".join(cmd))
    subprocess.run(cmd, check=True, cwd=src, env=env)
    return out / "gui.dist"


def assemble(built):
    if DIST.exists():
        try:
            shutil.rmtree(DIST)
        except PermissionError:
            raise SystemExit(f"\n删除旧版本失败：{DIST} 里的文件正在使用。\n"
                             f"请先关闭正在运行的 {NAME}.exe（以及打开的这个文件夹里的文件），再重新运行 python build.py")
    shutil.copytree(built, DIST)
    (DIST / EXE_TMP).rename(DIST / f"{NAME}.exe")
    for f in EXTRA_FILES:
        shutil.copy2(ROOT / f, DIST / f)
    (DIST / "危险品文件").mkdir(exist_ok=True)


def selftest():
    exe = DIST / f"{NAME}.exe"
    out = DIST / "selftest.txt"
    subprocess.run([str(exe), "--selftest"], timeout=120)
    text = out.read_text(encoding="utf-8") if out.exists() else "FAIL: 没有生成 selftest.txt"
    out.unlink(missing_ok=True)
    print("自检结果:", text.strip())
    return text.startswith("OK")


def make_zip():
    z = DIST.parent / f"{NAME}.zip"
    z.unlink(missing_ok=True)
    with zipfile.ZipFile(z, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in DIST.rglob("*"):
            zf.write(p, Path(NAME) / p.relative_to(DIST))
    return z


def main():
    built = compile_exe(build_python())
    assemble(built)
    ok = selftest()
    z = make_zip()
    print(f"\n完成: {DIST}\n压缩包: {z}")
    if not ok:
        print("⚠ 自检没通过，请看上面的自检结果")
        return 1


if __name__ == "__main__":
    sys.exit(main())
