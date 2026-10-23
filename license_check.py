"""授权开关：软件每次开始运行前，读取开发者放在 GitHub 上的授权文件。

授权文件内容（JSON）:
    {"enabled": true, "message": ""}
- enabled 为 true 才能使用；改成 false 就停用（不用重新打包）
- message 可选：停用时显示的提示，留空就用 DISABLED_MESSAGE
- 读不到授权文件（没网、GitHub 访问不了、文件格式不对）也不能使用

停用只是不再运行自动化，不会删除或修改任何文件。
"""
import json
import time
import urllib.request

# 授权文件放在单独的公开仓库 xuLiu925/liuxu-license 里（不要放在源码仓库里，源码仓库应设为私有）
LICENSE_URL = "https://raw.githubusercontent.com/xuLiu925/liuxu-license/main/license.json"

CONTACT = "xuliu00925@gmail.com"
DISABLED_MESSAGE = f"该软件已由开发者关闭使用，如需使用请联系开发者：{CONTACT}"
OFFLINE_MESSAGE = (f"无法连接授权服务器，暂时不能使用。\n请检查网络（需要能访问 GitHub）后重试。\n"
                   f"如一直无法使用，请联系开发者：{CONTACT}")
NOTICE = f"XL FreightPilot · Developed by Xu Liu（刘旭）· {CONTACT}"


def check(timeout=10):
    """返回 (能否使用, 不能使用时的提示)。"""
    url = f"{LICENSE_URL}?t={int(time.time())}"      # 加时间戳，尽量绕开缓存
    try:
        req = urllib.request.Request(url, headers={"Cache-Control": "no-cache", "User-Agent": "xl-freightpilot"})
        with urllib.request.urlopen(req, timeout=timeout) as r:
            data = json.loads(r.read().decode("utf-8"))
    except Exception:
        return False, OFFLINE_MESSAGE
    if isinstance(data, dict) and data.get("enabled") is True:
        return True, ""
    msg = data.get("message") if isinstance(data, dict) else ""
    return False, (msg or DISABLED_MESSAGE)
