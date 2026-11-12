"""在线更新：从开发者的服务器（Cloudflare Worker + R2，见 worker/README.md）下载新版本并替换程序文件。

- 版本号在打包时写进程序目录的 version.txt（build.py）；直接运行 .py（没有 version.txt）时不检查更新
- 服务开关开着才能查到和下载；下载完按 sha256 校验
- 程序运行中的文件不能覆盖：解压到 _update/new 后，启动一个 PowerShell 脚本，等程序退出再复制过去，
  然后重新打开程序。只覆盖更新包里有的文件：管理员码、飞书授权、邮件缓存、Dachser 登录信息、危险品文件
  都不在更新包里，不会动；本机已有的 pallet_params.json 也不覆盖（启动询价时会从服务器同步）
- 复制失败会写 _update/failed.txt，下次打开程序时提示
- 改名过渡：程序从 XL FreightPilot.exe 改名为 FreightPilot.exe。旧版只认旧名字，所以更新包里还带一份
  旧名字的同样程序（build.py LEGACY_EXE）；migrate_name() 把旧名字启动的换成新名字（见下）
"""
import hashlib
import os
import shutil
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.request
import zipfile
from pathlib import Path

import license_check
from app_ui import APP_DIR

EXE_NAME = "FreightPilot.exe"
LEGACY_EXE = "XL FreightPilot.exe"      # 改名前的 exe（旧版的快捷方式都指向它）
VERSION_FILE = APP_DIR / "version.txt"
UPDATE_DIR = APP_DIR / "_update"
FAILED_FILE = UPDATE_DIR / "failed.txt"
KEEP = ["pallet_params.json"]          # 本机已有就不覆盖

APPLY_PS1 = r"""param([int]$ProcId, [string]$Src, [string]$Dst, [string]$Exe)
$ErrorActionPreference = "Continue"
$upd = Join-Path $Dst "_update"
"$(Get-Date -Format s) waiting for $ProcId" | Out-File -LiteralPath (Join-Path $upd "update.log") -Encoding utf8
try { Wait-Process -Id $ProcId -Timeout 120 -ErrorAction SilentlyContinue } catch {}
Start-Sleep -Milliseconds 800
$xf = @("version.txt")      # 版本号最后写：中途失败时还是旧版本号，下次会再提示更新
foreach ($f in @(__KEEP__)) { if (Test-Path -LiteralPath (Join-Path $Dst $f)) { $xf += $f } }
$rcArgs = @($Src, $Dst, "/E", "/R:10", "/W:1", "/NP", "/NJH")
$rcArgs += "/XF"; $rcArgs += $xf
& robocopy @rcArgs | Out-File -LiteralPath (Join-Path $upd "update.log") -Encoding utf8 -Append
$rc = $LASTEXITCODE
if ($rc -lt 8) {
    try { Copy-Item -LiteralPath (Join-Path $Src "version.txt") -Destination $Dst -Force -ErrorAction Stop }
    catch { $rc = 99 }
}
if ($rc -ge 8) {
    "robocopy exit code $rc" | Out-File -LiteralPath (Join-Path $upd "failed.txt") -Encoding utf8
} else {
    Remove-Item -LiteralPath (Join-Path $upd "new") -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath (Join-Path $upd "package.zip") -Force -ErrorAction SilentlyContinue
}
Start-Process -FilePath $Exe -WorkingDirectory $Dst
"""


class UpdateError(Exception):
    pass


def current_version():
    """本机版本号；直接运行 .py（开发环境）时为空。"""
    try:
        return VERSION_FILE.read_text(encoding="utf-8").strip()
    except OSError:
        return ""


def check(timeout=10):
    """返回 (服务器上的版本信息或 None, 说明)。版本信息只在比本机新时返回。连不上抛异常。"""
    cur = current_version()
    if not cur:
        return None, "开发版本（直接运行 .py），不检查更新"
    status, data = license_check.post("/update/latest", {}, timeout)
    if status != 200 or not isinstance(data, dict):
        why = (data.get("message") or data.get("error")) if isinstance(data, dict) else ""
        raise UpdateError(f"服务器返回 HTTP {status}{'：' + why if why else ''}")
    ver = str(data.get("version") or "")
    if not ver:
        return None, data.get("message") or "服务器上还没有发布版本"
    if not (data.get("file") and data.get("sha256") and isinstance(data.get("size"), int)):
        raise UpdateError("服务器上的版本信息不完整")
    if ver <= cur:                     # 版本号是打包时间 YYYY.MM.DD.HHMM，按字符串比较即可
        return None, f"已经是最新版本（{cur}）"
    return data, f"有新版本 {ver}（当前 {cur}）"


def download(info, progress=None):
    """下载并校验更新包，返回本地 zip 路径。progress(已下载字节, 总字节)。"""
    UPDATE_DIR.mkdir(exist_ok=True)
    pkg = UPDATE_DIR / "package.zip"
    req = urllib.request.Request(
        license_check.server_url() + "/update/download", data=b"{}",
        headers=license_check.headers())
    h, n, size = hashlib.sha256(), 0, info["size"]
    try:
        with urllib.request.urlopen(req, timeout=60) as r, open(pkg, "wb") as f:
            while True:
                chunk = r.read(1 << 20)
                if not chunk:
                    break
                f.write(chunk)
                h.update(chunk)
                n += len(chunk)
                if progress:
                    progress(n, size)
    except urllib.error.HTTPError as e:
        raise UpdateError(f"下载失败：HTTP {e.code}") from None
    except OSError as e:
        raise UpdateError(f"下载失败：{e}") from None
    if n != size or h.hexdigest().lower() != str(info["sha256"]).lower():
        pkg.unlink(missing_ok=True)
        raise UpdateError("下载的更新包不完整（大小或校验不一致），请稍后再试")
    return pkg


def install(pkg):
    """解压更新包，启动替换脚本。调用后程序要马上退出（脚本等程序退出才开始复制）。"""
    new = UPDATE_DIR / "new"
    if new.exists():
        shutil.rmtree(new)
    new.mkdir(parents=True)
    root = new.resolve()
    with zipfile.ZipFile(pkg) as z:
        for m in z.namelist():
            if not (root / m).resolve().is_relative_to(root):
                raise UpdateError(f"更新包里有不安全的路径：{m}")
        z.extractall(new)
    exe = next(new.rglob(EXE_NAME), None)
    if exe is None:
        raise UpdateError(f"更新包里没有 {EXE_NAME}")
    src = exe.parent
    ps1 = UPDATE_DIR / "apply.ps1"
    keep = ", ".join(f'"{k}"' for k in KEEP)
    ps1.write_text(APPLY_PS1.replace("__KEEP__", keep), encoding="utf-8-sig")   # 带 BOM，PowerShell 5 才认 UTF-8
    FAILED_FILE.unlink(missing_ok=True)
    cmd = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
           "-File", str(ps1), "-ProcId", str(os.getpid()), "-Src", str(src), "-Dst", str(APP_DIR),
           "-Exe", str(APP_DIR / EXE_NAME)]
    # 隐藏的控制台（不能用 DETACHED_PROCESS：PowerShell 没有控制台会直接退出）+ 新进程组
    flags = 0x08000000 | 0x00000200
    try:    # 脱离本程序所在的作业对象，否则本程序退出时脚本可能被一起结束
        subprocess.Popen(cmd, creationflags=flags | 0x01000000, close_fds=True, cwd=str(UPDATE_DIR))
    except OSError:
        subprocess.Popen(cmd, creationflags=flags, close_fds=True, cwd=str(UPDATE_DIR))


def last_failure():
    """上次更新复制失败时返回说明（并清掉标记），否则 None。"""
    try:
        text = FAILED_FILE.read_text(encoding="utf-8-sig").strip()
    except OSError:
        return None
    FAILED_FILE.unlink(missing_ok=True)
    return text or "未知原因"


# ------------------------------------------------------------------ 改名过渡

FIX_SHORTCUTS_PS1 = r"""param([string]$Old, [string]$New, [string[]]$Dirs)
$sh = New-Object -ComObject WScript.Shell
if (-not $Dirs) {
    $Dirs = @([Environment]::GetFolderPath("Desktop"), [Environment]::GetFolderPath("CommonDesktopDirectory"),
              [Environment]::GetFolderPath("StartMenu"), [Environment]::GetFolderPath("CommonStartMenu"),
              (Join-Path $env:APPDATA "Microsoft\Internet Explorer\Quick Launch"))
}
foreach ($d in $Dirs) {
    if (-not $d -or -not (Test-Path -LiteralPath $d)) { continue }
    Get-ChildItem -LiteralPath $d -Filter *.lnk -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            $l = $sh.CreateShortcut($_.FullName)
            if ($l.TargetPath -ieq $Old) {
                $l.TargetPath = $New
                $l.IconLocation = "$New,0"
                $l.Save()
                if ($_.BaseName -ieq "XL FreightPilot") {
                    Rename-Item -LiteralPath $_.FullName -NewName "FreightPilot.lnk" -ErrorAction SilentlyContinue
                }
            }
        } catch {}
    }
}
"""


def fix_shortcuts(old, new, dirs=None):
    """把桌面、开始菜单、任务栏里指向旧 exe 的快捷方式改指新 exe（名字叫 XL FreightPilot 的顺便改名）。
    只动目标正好是旧 exe 的快捷方式；没权限改的（公共桌面）就算了。"""
    UPDATE_DIR.mkdir(exist_ok=True)
    ps1 = UPDATE_DIR / "fix_shortcuts.ps1"
    ps1.write_text(FIX_SHORTCUTS_PS1, encoding="utf-8-sig")
    cmd = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(ps1),
           "-Old", str(old), "-New", str(new)] + (["-Dirs", *map(str, dirs)] if dirs else [])
    try:
        subprocess.run(cmd, creationflags=0x08000000, timeout=60, cwd=str(UPDATE_DIR))
    except Exception:
        pass
    ps1.unlink(missing_ok=True)


def _remove_later(path, tries=30):
    """旧 exe 的进程刚退出时可能还删不掉，等一会儿再试。"""
    for _ in range(tries):
        try:
            path.unlink()
            return
        except FileNotFoundError:
            return
        except OSError:
            time.sleep(0.5)


def migrate_name(argv0=None):
    """程序启动时最先调用。返回 True = 这次是旧名字启动、已经换成新名字打开了，调用方直接退出。
    - 用 XL FreightPilot.exe 启动（旧快捷方式，或旧版在线更新后重启）且新 exe 在：改快捷方式，打开 FreightPilot.exe
    - 用 FreightPilot.exe 启动且旧 exe 还在：后台删掉旧 exe（每次更新包里又会带来一份，再删）
    直接运行 .py 时什么都不做。"""
    me = Path(argv0 or sys.argv[0]).name.lower()
    old, new = APP_DIR / LEGACY_EXE, APP_DIR / EXE_NAME
    if me == LEGACY_EXE.lower():
        if not new.exists():
            return False
        fix_shortcuts(old, new)
        subprocess.Popen([str(new), *sys.argv[1:]], cwd=str(APP_DIR), close_fds=True)
        return True
    if me == EXE_NAME.lower() and old.exists():
        threading.Thread(target=_remove_later, args=(old,), daemon=True).start()
    return False
