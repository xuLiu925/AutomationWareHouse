"""在线更新：从开发者的服务器（Cloudflare Worker + R2，见 worker/README.md）下载新版本并替换程序文件。

- 版本号在打包时写进程序目录的 version.txt（build.py）；直接运行 .py（没有 version.txt）时不检查更新
- 要有效的授权码才能查到和下载；下载完按 sha256 校验
- 程序运行中的文件不能覆盖：解压到 _update/new 后，启动一个 PowerShell 脚本，等程序退出再复制过去，
  然后重新打开程序。只覆盖更新包里有的文件：授权码、飞书授权、邮件缓存、Dachser 登录信息、危险品文件
  都不在更新包里，不会动；本机已有的 pallet_params.json 也不覆盖（启动询价时会从服务器同步）
- 复制失败会写 _update/failed.txt，下次打开程序时提示
"""
import hashlib
import os
import shutil
import subprocess
import urllib.error
import urllib.request
import zipfile
from pathlib import Path

import license_check
from app_ui import APP_DIR

EXE_NAME = "XL FreightPilot.exe"
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
    if not license_check.get_token():
        return None, "还没有填写授权码，不能检查更新"
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
        headers={"Content-Type": "application/json", "Authorization": f"Bearer {license_check.get_token()}",
                 "User-Agent": "xl-freightpilot"})
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
