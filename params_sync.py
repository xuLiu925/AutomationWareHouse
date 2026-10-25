"""产品托盘参数和服务器同步，以及修改申请 / 审核（服务器代码在 worker/worker.js）。

服务器上的参数为准：每次询价开始前 sync_local() 拉一次，版本变了就写到本机的
pallet_params.json（当缓存用，"_version" 记录版本号）。连不上时用本机缓存。
修改：管理员（user 记录里 "role": "admin"）直接保存；其他人提交申请，管理员审核通过后生效。
"""
from datetime import datetime

import license_check
import pallet_plan

STATUS = {"pending": "待审核", "approved": "已通过", "rejected": "已驳回", "withdrawn": "已撤回"}
ACTION = {"upsert": "修改", "delete": "删除", "defaults": "通用参数"}


class ServerError(Exception):
    def __init__(self, msg, status=0):
        super().__init__(msg)
        self.status = status


def call(path, data=None, timeout=15):
    """POST 到服务器，成功返回 JSON；任何失败抛 ServerError（文字给人看）。"""
    if not license_check.get_token():
        raise ServerError(license_check.NO_TOKEN_MESSAGE)
    try:
        status, res = license_check.post(path, data or {}, timeout)
    except Exception:
        raise ServerError(license_check.OFFLINE_MESSAGE)
    if not isinstance(res, dict):
        res = {}
    if status == 200:
        return res
    raise ServerError(res.get("error") or res.get("message") or f"服务器返回错误（{status}）", status)


def strip_local(raw):
    """去掉只在本机缓存里用的键。"""
    return {k: v for k, v in raw.items() if k != "_version"}


def local_version():
    try:
        return pallet_plan.load_raw().get("_version")
    except Exception:
        return None


def write_cache(data, version):
    pallet_plan.save_raw({"_version": version, **strip_local(data)})


def fetch():
    return call("/products/get")


def sync_local():
    """询价开始前调用：服务器版本和本机不同就写到本机。返回要显示的提示列表。失败抛 ServerError。"""
    r = fetch()
    msgs = []
    if r.get("data") is None:
        msgs.append("服务器上还没有产品参数，使用本机的 pallet_params.json")
    elif r.get("version") != local_version():
        write_cache(r["data"], r["version"])
        msgs.append(f"产品参数已更新到第 {r['version']} 版（{r.get('by', '')} {fmt_time(r.get('updated'))}）")
    if r.get("pending"):
        msgs.append(f"有 {r['pending']} 条产品参数修改申请待审核（点\"产品参数\"处理）")
    return msgs


def save(base_version, data, note=""):
    return call("/products/save", {"base_version": base_version, "data": strip_local(data), "note": note})


def submit(base_version, action, sku="", old_sku=None, data=None, note=""):
    return call("/proposals/submit", {"base_version": base_version, "action": action, "sku": sku,
                                      "old_sku": old_sku, "data": data, "note": note})


def list_proposals():
    return call("/proposals/list").get("items") or []


def approve(pid):
    return call("/proposals/approve", {"id": pid})


def reject(pid, reason):
    return call("/proposals/reject", {"id": pid, "reason": reason})


def withdraw(pid):
    return call("/proposals/withdraw", {"id": pid})


def history():
    return call("/products/history").get("versions") or []


def rollback(base_version, version):
    return call("/products/rollback", {"base_version": base_version, "version": version})


def fmt_time(iso):
    """服务器的 UTC 时间 -> 本机时间 'YYYY-MM-DD HH:MM'"""
    if not iso:
        return ""
    try:
        return datetime.fromisoformat(iso.replace("Z", "+00:00")).astimezone().strftime("%Y-%m-%d %H:%M")
    except ValueError:
        return iso
