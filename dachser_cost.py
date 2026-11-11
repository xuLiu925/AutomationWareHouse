"""对账用：从 Dachser『My enquiries』读每张出库单的成本价（网站询价的 Price 就是成本）。

只读列表，不点任何按钮。邮件询价的成本（价格在附件里）不读。
匹配：Dachser 单号里有这张出库单号（建单时会把 +出库单号 加到 Order number 后面）→ 就是它；
否则按备注里的编号匹配，规则和建单相同（create_orders.find_hits），并去掉已经建给本次其他出库单的报价。
几条都对上时优先用已经 created 的；价格不一样就不填，交给人工。
"""
import argparse
import re
from datetime import date, datetime, timedelta
from pathlib import Path

from dachser_fill import START_URL


# 读日期框上 jQuery UI 日期控件允许的最早日期（Dachser 只保留最近一段时间的询价，更早的日期框标红、搜索一直转圈）
JS_MIN_DATE = r"""() => {
    const a = document.querySelector("form.filter-form input[name='fromDate']");
    const $ = window.jQuery;
    if (!a || !$ || !$.datepicker) return null;
    try {
        const inst = $.data(a, 'datepicker');
        const d = inst && $.datepicker._getMinMaxDate(inst, 'min');
        if (d instanceof Date) return [d.getFullYear(), d.getMonth() + 1, d.getDate()];
    } catch (e) {}
    return null;
}"""
FALLBACK_DAYS = 30      # 读不到最早日期、搜索又失败时，退到最近这么多天再试一次
OLD_MARGIN_DAYS = 7     # 出库时间离最早可查日期不到这么多天：询价可能在可查范围之前，查不到不奇怪


def fetch_rows(since, html=None):
    """读 since（太早就从 Dachser 允许的最早日期）到今天的全部网站询价。
    返回 (行列表, 实际查询的开始日期 或 None)。html = 离线测试用的列表页文件。"""
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
            return rows, None
        from dachser_fill import _launch, open_quotation
        ctx = _launch(pw, argparse.Namespace(profile="dachser_profile", chromium_path=None))
        try:
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
            open_quotation(page, START_URL)
            co.open_history(page)
            today = date.today()
            start = since.date() if isinstance(since, datetime) else since
            try:
                m = page.evaluate(JS_MIN_DATE)
            except Exception:
                m = None
            if m and date(*m) > start:
                start = date(*m)
                print(f"  Dachser 网站只能查 {start:%Y-%m-%d} 之后的询价")
            try:
                return co.search(page, "", 0, start, today, strict=True), start
            except RuntimeError as e:
                fb = today - timedelta(days=FALLBACK_DAYS)
                if m or start >= fb:
                    raise
                print(f"  [WARN] {e}；改查最近 {FALLBACK_DAYS} 天")
                return co.search(page, "", 0, fb, today, strict=True), fb
        finally:
            try:
                ctx.close()
            except Exception:
                pass


def has_no(order, no):
    """Dachser 单号里有这个出库单号（前后不紧挨字母数字）。"""
    order, no = re.sub(r"\s+", "", order or ""), re.sub(r"\s+", "", no or "")
    return bool(no) and re.search(r"(?<![A-Za-z0-9])" + re.escape(no) + r"(?![A-Za-z0-9])", order, re.I) is not None


def costs(orders, rows, earliest=None):
    """返回 {出库单号: (成本价或 None, 说明)}。earliest = 实际查询的开始日期（Dachser 更早的询价查不到）。"""
    from reconcile import parse_time
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
            t = parse_time(o["time"])
            if earliest and t and t.date() < earliest + timedelta(days=OLD_MARGIN_DAYS):
                out[o["no"]] = (None, f"Dachser 网站只能查 {earliest:%Y-%m-%d} 之后的询价，这单的询价可能更早")
            else:
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
