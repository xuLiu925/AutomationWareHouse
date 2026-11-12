"""服务开关：软件每次开始运行前问开发者的服务器（Cloudflare Worker，代码在 worker/）服务是否开着。

- 同事不用授权码。服务器上的总开关关掉，或连不上服务器时，都不能使用（窗口底部有说明）
- 请求头带 Windows 用户名，只用来显示谁提交了产品参数修改
- 程序目录有 xlfp_config.json（只有管理员的电脑上有）时带上里面的码，服务器据此认出管理员

暂停只是不再运行自动化，不会删除或修改任何文件。
"""
import getpass
import json
import os
import urllib.error
import urllib.parse
import urllib.request

from app_ui import APP_DIR

# 部署 Worker 后改成它的地址（见 worker/README.md）；测试时可用环境变量 XLFP_SERVER 覆盖
SERVER_URL = "https://xlfp.xuliu00925.workers.dev/"
CONFIG_FILE = APP_DIR / "xlfp_config.json"

CONTACT = "xuliu00925@gmail.com"
DISABLED_MESSAGE = f"服务暂停使用，请联系刘旭（{CONTACT}）"
OFFLINE_MESSAGE = (f"无法连接服务器，暂时不能使用。\n请检查网络后重试。\n"
                   f"如一直无法使用，请联系刘旭（{CONTACT}）")
NOTICE = f"使用有问题？请联系软件开发者邮箱：{CONTACT}"      # 窗口底部的一行字


def server_url():
    return (os.environ.get("XLFP_SERVER", "").strip() or SERVER_URL).rstrip("/")


def _config():
    try:
        cfg = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        return cfg if isinstance(cfg, dict) else {}
    except Exception:
        return {}


def get_token():
    """管理员码（只有管理员电脑上的 xlfp_config.json 里有），没有返回 ""。"""
    return str(_config().get("token", "")).strip()


def win_user():
    try:
        return getpass.getuser()
    except Exception:
        return os.environ.get("USERNAME", "")


def headers():
    h = {"Content-Type": "application/json", "User-Agent": "xl-freightpilot",
         "X-XLFP-User": urllib.parse.quote(win_user())}
    if get_token():
        h["Authorization"] = f"Bearer {get_token()}"
    return h


def post(path, data, timeout):
    """POST 到服务器，返回 (HTTP 状态码, JSON)。连不上时抛异常。"""
    req = urllib.request.Request(
        server_url() + path, data=json.dumps(data, ensure_ascii=False).encode("utf-8"), headers=headers())
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
