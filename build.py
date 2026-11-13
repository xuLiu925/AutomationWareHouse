#!/usr/bin/env python3
"""打包成发给同事的"FreightPilot"文件夹和 zip（用 Nuitka 编译，同事那边不需要 Python）。

用法（只在你自己的电脑上）:
    pip install nuitka
    python build.py [更新说明]                 只打包
    python build.py [更新说明] --upload        打包后自动传到 R2（需要 wrangler 和 CLOUDFLARE_API_TOKEN）
    python build.py --upload-only              只上传上次打好的 dist/release

结果:
    dist/FreightPilot/          整个文件夹就是发给同事的程序
    dist/FreightPilot.zip       同一份内容的压缩包
    dist/FreightPilot-Setup-<版本>.exe  安装包（installer.iss，需要装 Inno Setup；没装就跳过）
注意: 不会把你的登录信息（dachser_profile）打包进去。

关于路径:
    Nuitka 分析 DLL 依赖的工具不支持中文路径（用户名"刘旭"会变成乱码导致编译失败）。
    项目在纯英文路径下（如 D:\\XLFreightPilot）时，直接用当前 Python 在项目的 build/ 里编译；
    否则把源码复制到纯英文路径 D:\\liuxu_build（环境变量 LIUXU_BUILD_DIR 可改），
    在那里单独建一个打包用的虚拟环境来编译。
"""
import os
import shutil
import subprocess
import sys
import hashlib
import json
import zipfile
from datetime import datetime
from importlib.metadata import version
from pathlib import Path

ROOT = Path(__file__).resolve().parent
NAME = "FreightPilot"
# 改名过渡（2026-11）：旧版的在线更新只认 "XL FreightPilot.exe"，更新包里再放一份同样的程序叫这个名字，
# 旧版才能装上新版；新版用旧名字启动时会改快捷方式、换成 FreightPilot.exe（updater.migrate_name）。
# 同事都更新过以后改成 None，更新包就小了。
LEGACY_EXE = "XL FreightPilot.exe"
IN_PLACE = str(ROOT).isascii() and not os.environ.get("LIUXU_BUILD_DIR")
STAGE = ROOT / "build" if IN_PLACE else Path(
    os.environ.get("LIUXU_BUILD_DIR") or (r"D:\liuxu_build" if Path("D:/").exists() else r"C:\liuxu_build"))
DIST = ROOT / "dist" / NAME
EXE_TMP = "LiuXuScript.exe"     # 编译时用英文名，打包完再改名
SOURCES = ["gui.py", "app_ui.py", "license_check.py", "dachser_fill.py", "ai_clean.py", "create_orders.py", "match_orders.py",
           "pallet_plan.py", "params_editor.py", "params_sync.py", "feishu_mail.py", "reconcile.py", "dachser_cost.py", "updater.py", "welcome.py"]
VERSION = datetime.now().strftime("%Y.%m.%d.%H%M")     # 在线更新按这个比较新旧（version.txt）
RELEASE = ROOT / "dist" / "release"                     # 上传到 R2 的两个文件
BUCKET = "xlfp-releases"                                # --upload 用 wrangler 传到这个 R2 存储桶
ICON = ROOT / "assets" / "xlfp.ico"             # exe 图标（tools/make_art.py 生成）
EXTRA_FILES = ["pallet_params.json", "xlfp-wms-sn.user.js", "assets/xlfp.ico", "assets/scene_morning.png",
               "assets/scene_day.png", "assets/scene_night.png"]     # pallet_params.json 不在 Git 里，打包前本机要有
PACKAGES = ["nuitka", f"playwright=={version('playwright')}", f"openpyxl=={version('openpyxl')}"]


def build_python():
    """返回用来编译的 python.exe：项目在英文路径就用当前 Python，否则在英文路径下准备一个打包专用的虚拟环境。"""
    if IN_PLACE:
        subprocess.run([sys.executable, "-m", "pip", "install", "-q", *PACKAGES], check=True)
        return Path(sys.executable)
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
    if IN_PLACE:
        src = ROOT
    else:
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
           f"--windows-icon-from-ico={ICON}",
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
    if LEGACY_EXE:
        shutil.copy2(DIST / f"{NAME}.exe", DIST / LEGACY_EXE)
    for f in EXTRA_FILES:
        (DIST / f).parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(ROOT / f, DIST / f)
    (DIST / "危险品文件").mkdir(exist_ok=True)
    (DIST / "version.txt").write_text(VERSION, encoding="utf-8")


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


def find_iscc():
    """Inno Setup 的编译器 ISCC.exe；没装返回 None（winget install JRSoftware.InnoSetup）。"""
    for p in (shutil.which("ISCC"),
              os.path.expandvars(r"%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"),
              r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe", r"C:\Program Files\Inno Setup 6\ISCC.exe"):
        if p and Path(p).exists():
            return Path(p)
    return None


def make_installer():
    """dist/FreightPilot-Setup-<版本>.exe：发给新同事的安装包（装到用户目录，桌面快捷方式）。
    在线更新还是用 zip，和安装包无关。"""
    iscc = find_iscc()
    if not iscc:
        print("⚠ 没装 Inno Setup，跳过安装包（winget install JRSoftware.InnoSetup）")
        return None
    r = subprocess.run([str(iscc), "/Q", f"/DAppVersion={VERSION}", f"/DSourceDir={DIST}",
                        f"/DOutputDir={DIST.parent}", str(ROOT / "installer.iss")], cwd=ROOT)
    out = DIST.parent / f"{NAME}-Setup-{VERSION}.exe"
    if r.returncode or not out.exists():
        print(f"⚠ 安装包生成失败（ISCC 返回 {r.returncode}）")
        return None
    for old in DIST.parent.glob(f"{NAME}-Setup-*.exe"):      # 只留这一版
        if old != out:
            old.unlink(missing_ok=True)
    return out


def make_release(z, notes):
    """在线更新用的两个文件：压缩包（文件名带版本号）和 latest.json。上传到 R2 见 worker/README.md。"""
    if RELEASE.exists():
        shutil.rmtree(RELEASE)
    RELEASE.mkdir(parents=True)
    name = f"FreightPilot-{VERSION}.zip"
    shutil.copy2(z, RELEASE / name)
    data = (RELEASE / name).read_bytes()
    info = {"version": VERSION, "file": name, "sha256": hashlib.sha256(data).hexdigest(), "size": len(data),
            "notes": notes}
    (RELEASE / "latest.json").write_text(json.dumps(info, ensure_ascii=False, indent=2), encoding="utf-8")
    return name


def _cf_env():
    """wrangler 用的环境变量；当前进程没有时从 Windows 用户环境变量（注册表）补上——刚设好、没重启的程序也能用。"""
    env = dict(os.environ)
    if sys.platform == "win32":
        import winreg
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment") as k:
                for name in ("CLOUDFLARE_API_TOKEN", "CLOUDFLARE_ACCOUNT_ID"):
                    if not env.get(name):
                        try:
                            env[name] = winreg.QueryValueEx(k, name)[0]
                        except OSError:
                            pass
        except OSError:
            pass
    if not env.get("CLOUDFLARE_API_TOKEN"):
        raise SystemExit("上传失败：没有环境变量 CLOUDFLARE_API_TOKEN（Cloudflare token，权限 Workers R2 Storage: Edit），"
                         "见 worker/README.md")
    return env


def _wrangler(args, env, **kw):
    exe = shutil.which("wrangler") or shutil.which("wrangler.cmd")
    if not exe:
        raise SystemExit("上传失败：找不到 wrangler，请先运行 npm install -g wrangler")
    return subprocess.run([exe, *args], env=env, **kw)


def upload():
    """把 dist/release 的两个文件传到 R2：先 zip 再 latest.json（zip 失败就不动 latest.json，同事不会看到新版本），
    然后读回 R2 上的 latest.json 核对版本，再问一次 Worker /update/latest。"""
    info = json.loads((RELEASE / "latest.json").read_text(encoding="utf-8"))
    zp = RELEASE / info["file"]
    if not zp.exists() or zp.stat().st_size != info["size"]:
        raise SystemExit(f"上传失败：{zp} 不存在或大小和 latest.json 不一致，请重新打包")
    env = _cf_env()
    for f, ctype in ((zp, "application/zip"), (RELEASE / "latest.json", "application/json")):
        print(f"上传 {f.name} ({f.stat().st_size / 1e6:.1f} MB) ...")
        r = _wrangler(["r2", "object", "put", f"{BUCKET}/{f.name}", "--file", str(f),
                       "--content-type", ctype, "--remote"], env)
        if r.returncode:
            raise SystemExit(f"上传 {f.name} 失败（wrangler 返回 {r.returncode}）"
                             + ("；latest.json 没动，同事那边还是旧版本" if f == zp else ""))
    r = _wrangler(["r2", "object", "get", f"{BUCKET}/latest.json", "--remote", "--pipe"], env,
                  capture_output=True, text=True, encoding="utf-8")
    try:
        remote = json.loads(r.stdout[r.stdout.index("{"):])
    except ValueError:
        remote = {}
    if remote.get("version") != info["version"] or remote.get("sha256") != info["sha256"]:
        raise SystemExit(f"上传后核对失败：R2 上的 latest.json 是 {remote.get('version')}，应为 {info['version']}")
    print(f"✓ 已上传到 R2：版本 {info['version']}")
    try:
        import license_check
        status, data = license_check.post("/update/latest", {}, 15)
        ver = data.get("version") if isinstance(data, dict) else None
        if status == 200 and ver == info["version"]:
            print(f"✓ 服务器 /update/latest 返回 {ver}，同事打开程序就会提示更新")
        else:
            print(f"⚠ 服务器 /update/latest 返回 HTTP {status}、版本 {ver}（服务开关关着时同事收不到更新）")
    except Exception as e:
        print(f"⚠ 没能问服务器确认（{e}），R2 上已是新版本")


def main():
    sys.stdout.reconfigure(line_buffering=True)    # 和 wrangler/nuitka 的输出按顺序显示（远程看日志时）
    args = sys.argv[1:]
    do_upload = "--upload" in args
    if "--upload-only" in args:                    # 只上传上次打好的 dist/release（上传失败后重试用）
        upload()
        return 0
    notes = " ".join(a for a in args if a != "--upload").strip()   # 可选：这次更新了什么（更新提示里显示）
    built = compile_exe(build_python())
    assemble(built)
    ok = selftest()
    z = make_zip()
    print(f"\n完成: {DIST}  版本 {VERSION}\n压缩包: {z}")
    if not ok:
        print("⚠ 自检没通过，请看上面的自检结果；没有生成在线更新文件")
        return 1
    setup = make_installer()
    if setup:
        print(f"安装包: {setup}（发给同事安装用）")
    name = make_release(z, notes)
    print(f"在线更新文件: {RELEASE}")
    if do_upload:
        upload()
    else:
        print(f"  上传到 R2 存储桶：先传 {name}，再传 latest.json（覆盖旧的），"
              f"或运行 python build.py --upload-only。同事打开程序就会提示更新。")


if __name__ == "__main__":
    sys.exit(main())
