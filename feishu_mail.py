"""飞书邮件：授权 + 读自己邮箱里的邮件（对账功能用）。

- 每个人读的是自己的邮箱：第一次（或授权过期后）程序打开浏览器，本人在飞书页面点"同意授权"，
  浏览器跳回 http://localhost:8080/callback，程序在本机 8080 端口接收授权 code。
- App ID / App Secret 不在程序和仓库里：code 由开发者的服务器（worker/）换成 token。
- 换到的 user_access_token 只保存在本机程序目录的 feishu_token.json（不进 Git），约 2 小时有效；
  有 refresh_token（飞书后台开通 offline_access 后才有）时自动续期，否则过期后重新点一次授权。
- 邮件内容直接从飞书读取，不经过开发者服务器。
"""
import base64
import html
import json
import secrets
import socket
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
from http.server import BaseHTTPRequestHandler, HTTPServer

import app_ui
import license_check
from app_ui import APP_DIR

API = "https://open.feishu.cn/open-apis"
AUTHORIZE_URL = "https://accounts.feishu.cn/open-apis/authen/v1/authorize"
# 不要加 offline_access，除非飞书后台已开通（否则授权页报 20027）
SCOPES = ["mail:user_mailbox.message:readonly", "mail:user_mailbox.message.body:read",
          "mail:user_mailbox.message.subject:read", "mail:user_mailbox.message.address:read",
          "mail:user_mailbox.folder:read"]       # 报价邮件常被邮箱规则收进自建文件夹
TOKEN_FILE = APP_DIR / "feishu_token.json"
AUTH_TIMEOUT_S = 300
TOKEN_ERRORS = {99991668, 99991677, 99991663, 99991661, 20005, 20026}     # token 无效 / 过期
RATE_LIMIT = {99991400, 99991401}


class FeishuError(Exception):
    pass


# ------------------------------------------------------------------ token

def _server(path, data):
    try:
        status, res = license_check.post(path, data, 30)
    except Exception as e:
        raise FeishuError(f"连不上服务器：{e}") from None
    if status != 200 or not isinstance(res, dict):
        why = ""
        if isinstance(res, dict):
            why = res.get("error_description") or res.get("error") or res.get("message") or res.get("msg") or ""
        raise FeishuError(f"服务器返回 HTTP {status}{'：' + str(why) if why else ''}")
    return res


def _load():
    try:
        t = json.loads(TOKEN_FILE.read_text(encoding="utf-8"))
        return t if isinstance(t, dict) else {}
    except Exception:
        return {}


def _save(res):
    now = time.time()
    t = {"access_token": res["access_token"],
         "expires_at": now + int(res.get("expires_in") or 7200) - 300}
    if res.get("refresh_token"):
        t["refresh_token"] = res["refresh_token"]
        t["refresh_expires_at"] = now + int(res.get("refresh_token_expires_in") or 0) - 300
    TOKEN_FILE.write_text(json.dumps(t, ensure_ascii=False, indent=2), encoding="utf-8")
    return t["access_token"]


def clear_token():
    try:
        TOKEN_FILE.unlink()
    except OSError:
        pass


def code_from(s):
    """接受授权 code，或浏览器跳转后的整个地址（…/callback?code=xxx&state=…）。"""
    s = (s or "").strip()
    if "code=" in s:
        return urllib.parse.parse_qs(urllib.parse.urlsplit(s).query).get("code", [""])[0]
    return s


def get_access_token(code=None):
    """返回可用的 user_access_token：手动给的 code → 本机缓存 → refresh_token 续期 → 浏览器授权。
    code：不在这台电脑上授权时（比如在手机上点授权），把跳转后的地址或 code 交给程序。"""
    if code:
        return _save(_server("/feishu/token", {"code": code_from(code)}))
    t = _load()
    if t.get("access_token") and time.time() < t.get("expires_at", 0):
        return t["access_token"]
    if t.get("refresh_token") and time.time() < t.get("refresh_expires_at", 0):
        try:
            return _save(_server("/feishu/token", {"refresh_token": t["refresh_token"]}))
        except FeishuError as e:
            print(f"  飞书授权续期失败（{e}），需要重新授权")
    return _save(_server("/feishu/token", {"code": authorize()}))


def authorize():
    """打开浏览器让本人授权，在本机端口等飞书带着 code 跳回来。返回 code。"""
    cfg = _server("/feishu/config", {})
    redirect = urllib.parse.urlsplit(cfg["redirect_uri"])
    state = secrets.token_urlsafe(16)
    url = AUTHORIZE_URL + "?" + urllib.parse.urlencode({
        "client_id": cfg["app_id"], "redirect_uri": cfg["redirect_uri"],
        "scope": " ".join(SCOPES), "state": state}, quote_via=urllib.parse.quote)
    got = {}

    class Handler(BaseHTTPRequestHandler):
        def do_GET(self):
            u = urllib.parse.urlsplit(self.path)
            q = urllib.parse.parse_qs(u.query)
            if u.path != redirect.path:
                self.send_response(404)
                self.end_headers()
                return
            if q.get("state", [""])[0] != state:
                msg = "这个授权链接已经过期（程序重新开始过）。请使用运行记录里最新的授权链接。"
            elif q.get("code"):
                got["code"] = q["code"][0]
                msg = "授权成功，可以关闭这个页面，回到 FreightPilot。"
            else:
                got["error"] = q.get("error", ["未知错误"])[0]
                msg = f"授权没有完成：{html.escape(got['error'])}"
            body = f"<html><meta charset='utf-8'><body style='font:16px sans-serif;padding:40px'>{msg}</body></html>"
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(body.encode("utf-8"))

        def log_message(self, *a):
            pass

    # localhost 在有的电脑上先解析成 IPv6 (::1)，浏览器不一定退回 IPv4，所以两个都监听（只监听本机）
    port, servers = redirect.port or 80, []
    for host, family in (("127.0.0.1", socket.AF_INET), ("::1", socket.AF_INET6)):
        cls = type("Srv", (HTTPServer,), {"address_family": family})
        try:
            servers.append(cls((host, port), Handler))
        except OSError as e:
            if family == socket.AF_INET:
                raise FeishuError(f"本机 {port} 端口被占用，无法接收飞书授权（{e}）。"
                                  f"请关掉占用这个端口的程序后重试。") from None
    for s in servers:
        s.timeout = 0.5
    print("  请在打开的浏览器页面里点「授权」（用你自己的飞书账号）。如果浏览器没有自动打开，复制下面的链接打开：")
    print(f"  {url}")
    threading.Thread(target=webbrowser.open, args=(url,), daemon=True).start()
    deadline = time.time() + AUTH_TIMEOUT_S
    try:
        while not got and time.time() < deadline and not app_ui.stop_requested():
            for s in servers:
                s.handle_request()
    finally:
        for s in servers:
            s.server_close()
    if "code" in got:
        print("  飞书授权成功")
        return got["code"]
    if got.get("error"):
        raise FeishuError(f"飞书授权没有完成：{got['error']}")
    raise FeishuError("已停止" if app_ui.stop_requested() else f"{AUTH_TIMEOUT_S // 60} 分钟内没有完成飞书授权")


# ------------------------------------------------------------------ 邮件接口

def _get(path, token, params=None):
    url = API + path + ("?" + urllib.parse.urlencode(params) if params else "")
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"})
    for attempt in range(5):
        try:
            with urllib.request.urlopen(req, timeout=30) as r:
                res = json.loads(r.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            try:
                res = json.loads(e.read().decode("utf-8"))
            except Exception:
                res = {"code": e.code, "msg": f"HTTP {e.code}"}
            if e.code == 429:
                res.setdefault("code", 99991400)
        except (urllib.error.URLError, TimeoutError, OSError) as e:
            if attempt < 4:
                time.sleep(1 + attempt)
                continue
            raise FeishuError(f"连不上飞书：{e}") from None
        code = res.get("code")
        if code == 0:
            return res.get("data") or {}
        if code in RATE_LIMIT and attempt < 4:
            time.sleep(2 ** attempt)
            continue
        if code in TOKEN_ERRORS:
            clear_token()
            raise FeishuError(f"飞书授权已失效（{res.get('msg')}），请重新开始，会再次打开授权页面")
        raise FeishuError(f"飞书接口出错 {code}：{res.get('msg')}")
    raise FeishuError("飞书接口一直限流，请稍后再试")


SKIP_FOLDERS = {"DRAFT", "SPAM", "TRASH"}


def list_folders(token):
    """返回 [(文件夹 id, 名称)]：收件箱、已发送和自建文件夹（不含草稿、垃圾邮件、已删除）。
    没有文件夹权限（旧授权）时只返回收件箱和已发送。"""
    try:
        items = _get("/mail/v1/user_mailboxes/me/folders", token).get("items") or []
    except FeishuError as e:
        print(f"  读不到邮箱文件夹列表，只看收件箱和已发送（{str(e)[:80]}）")
        return [("INBOX", "收件箱"), ("SENT", "已发送")]
    out = [(str(i.get("id")), i.get("name") or str(i.get("id"))) for i in items if i.get("id")]
    out = [f for f in out if f[0].upper() not in SKIP_FOLDERS]
    have = {f[0] for f in out}
    return [f for f in (("INBOX", "收件箱"), ("SENT", "已发送")) if f[0] not in have] + out


def list_messages(token, folder_id, page_token=None, page_size=20):
    """返回 (邮件 id 列表, 下一页 token 或 None)。"""
    params = {"folder_id": folder_id, "page_size": page_size}
    if page_token:
        params["page_token"] = page_token
    d = _get("/mail/v1/user_mailboxes/me/messages", token, params)
    ids = [i if isinstance(i, str) else i.get("message_id") for i in d.get("items") or []]
    return [i for i in ids if i], (d.get("page_token") if d.get("has_more") else None)


def decode_body(s):
    """正文是 base64url 编码。"""
    if not s:
        return ""
    try:
        return base64.urlsafe_b64decode(s + "=" * (-len(s) % 4)).decode("utf-8", "replace")
    except Exception:
        return s


def get_message(token, message_id):
    """返回 {"date": 毫秒, "subject", "sender", "html", "text"}。"""
    m = _get(f"/mail/v1/user_mailboxes/me/messages/{urllib.parse.quote(message_id, safe='')}", token).get("message") or {}
    try:
        date = int(m.get("internal_date") or 0)
    except (TypeError, ValueError):
        date = 0
    return {"date": date, "subject": m.get("subject") or "",
            "sender": (m.get("head_from") or {}).get("mail_address", ""),
            "html": decode_body(m.get("body_html")), "text": decode_body(m.get("body_plain_text"))}
