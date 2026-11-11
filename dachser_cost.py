"""对账用：从 Dachser『My enquiries』读每张出库单的成本价（网站询价的 Price 就是成本）。

只读列表，不点任何按钮。邮件询价的成本（价格在附件里）不读。
匹配：Dachser 单号里有这张出库单号（建单时会把 +出库单号 加到 Order number 后面）→ 就是它；
否则按备注里的编号匹配，规则和建单相同（create_orders.find_hits），并去掉已经建给本次其他出库单的报价。
几条都对上时优先用已经 created 的；价格不一样就不填，交给人工。
"""
import argparse
import re
from datetime import datetime
from pathlib import Path

from dachser_fill import START_URL


def fetch_rows(since, html=None):
    """读 since 到今天的全部网站询价。html = 离线测试用的列表页文件。"""
    from playwright.sync_api import sync_playwright
    import create_orders as co
    with sync_playwright() as pw:
        if html:
            b = pw.chromium.launch()
            page = b.new_page()
            page.route("**/*", lambda r: r.abort())
            page.set_content(Path(html).read_text(encoding="utf-8", errors="ignore"), wait_until="domcontentloaded")
            rows = [co.to_dz(r) for r in (page.evaluate(co.JS_ROWS) or [])]
            b.close()
            return rows
        from dachser_fill import _launch, open_quotation
        ctx = _launch(pw, argparse.Namespace(profile="dachser_profile", chromium_path=None))
        try:
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
            open_quotation(page, START_URL)
            days = max(1, (datetime.now() - since).days + 1)
            return co.search(page, "", days)
        finally:
            try:
                ctx.close()
            except Exception:
                pass


def has_no(order, no):
    """Dachser 单号里有这个出库单号（前后不紧挨字母数字）。"""
    order, no = re.sub(r"\s+", "", order or ""), re.sub(r"\s+", "", no or "")
    return bool(no) and re.search(r"(?<![A-Za-z0-9])" + re.escape(no) + r"(?![A-Za-z0-9])", order, re.I) is not None


def costs(orders, rows):
    """返回 {出库单号: (成本价或 None, 说明)}。"""
    import create_orders as co
    from reconcile import fmt, parse_price
    nos = [o["no"] for o in orders]
    out = {}
    for o in orders:
        hits = [r for r in rows if has_no(r["order"], o["no"])]
        if not hits and co.remark_keys(o["remark"]):
            hits = [r for r in co.find_hits(o["remark"], rows)[0]
                    if not any(has_no(r["order"], n) for n in nos if n != o["no"])]
        if not hits:
            out[o["no"]] = (None, "Dachser 网站里没找到这单的询价（可能是邮件询价）")
            continue
        use = [r for r in hits if not r["can_create"]] or hits
        prices = {}
        for r in use:
            p = parse_price(re.sub(r"(?i)eur|€", "", r["price"]))
            if p and p[0] == "price":
                prices.setdefault(round(p[1], 2), []).append(r)
        if len(prices) == 1:
            out[o["no"]] = (next(iter(prices)), "")
        elif not prices:
            out[o["no"]] = (None, f"Dachser 的价格看不懂：{use[0]['price']}")
        else:
            out[o["no"]] = (None, "Dachser 有几条询价价格不同：" + "、".join(
                f"{rs[0]['order']}（{rs[0]['time']}）{fmt(p)}" for p, rs in prices.items()))
    return out
