"""授权：软件每次开始运行前，用本机的授权码问开发者的服务器（Cloudflare Worker，代码在 worker/）能不能用。

- 每个同事一个授权码，保存在程序目录的 xlfp_config.json 里（窗口里点"授权码"填写）
- 服务器上可以关掉总开关，或单独停用某个授权码（不用重新打包）
- 授权码无效、被停用、连不上服务器时都不能使用

停用只是不再运行自动化，不会删除或修改任何文件。
"""
import json
import os
import urllib.error
import urllib.request

from app_ui import APP_DIR

# 部署 Worker 后改成它的地址（见 worker/README.md）；测试时可用环境变量 XLFP_SERVER 覆盖
SERVER_URL = "https://xlfp.xuliu00925.workers.dev/"
CONFIG_FILE = APP_DIR / "xlfp_config.json"

CONTACT = "xuliu00925@gmail.com"
DISABLED_MESSAGE = f"该软件已由开发者关闭使用，如需使用请联系开发者：{CONTACT}"
NO_TOKEN_MESSAGE = f"还没有填写授权码，请点窗口里的\"授权码\"按钮填写。\n没有授权码请联系开发者：{CONTACT}"
OFFLINE_MESSAGE = (f"无法连接授权服务器，暂时不能使用。\n请检查网络后重试。\n"
                   f"如一直无法使用，请联系开发者：{CONTACT}")
NOTICE = f"XL FreightPilot · Developed by Xu Liu（刘旭）· {CONTACT}"


def server_url():
    return (os.environ.get("XLFP_SERVER", "").strip() or SERVER_URL).rstrip("/")


def _config():
    try:
        cfg = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        return cfg if isinstance(cfg, dict) else {}
    except Exception:
        return {}


def get_token():
    return str(_config().get("token", "")).strip()


def save_token(token):
    cfg = _config()
    cfg["token"] = token.strip()
    CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


def post(path, data, timeout):
    """带授权码 POST 到服务器，返回 (HTTP 状态码, JSON)。连不上时抛异常。"""
    req = urllib.request.Request(
        server_url() + path, data=json.dumps(data, ensure_ascii=False).encode("utf-8"),
        headers={"Content-Type": "application/json", "Authorization": f"Bearer {get_token()}",
                 "User-Agent": "xl-freightpilot"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return r.status, json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        try:
            return e.code, json.loads(e.read().decode("utf-8"))
        except Exception:
            return e.code, {}


def check(timeout=10):
    """返回 (能否使用, 不能使用时的提示)。"""
    if not get_token():
        return False, NO_TOKEN_MESSAGE
    try:
        status, data = post("/check", {}, timeout)
    except Exception:
        return False, OFFLINE_MESSAGE
    if not isinstance(data, dict):
        data = {}
    if status == 200 and data.get("enabled") is True:
        return True, ""
    if status in (401, 403):
        return False, data.get("message") or DISABLED_MESSAGE
    return False, OFFLINE_MESSAGE
