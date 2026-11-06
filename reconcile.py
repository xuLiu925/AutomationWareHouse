#!/usr/bin/env python3
"""对账：领星出库单 Excel → 按备注里的编号（MR…）在飞书邮件里找我们给客户的报价 → 生成对账 Excel。

报价邮件：回复客户"出库询价"的邮件，正文第一张表有"编号"和"运费报价(欧元)"两列（如"320欧"）。
- 只看邮件自己写的部分，后面引用的原询价邮件（报价是空的）不看。
- 一个编号有多封报价邮件时取最新的一封；最新的写"待定"等 → 报价留空。
- 一张出库单对应一个报价：备注里的几个编号在报价表里共用同一个报价格子（同一格写多个编号，
  或报价格合并了几行）时只算一次。
- 表格读不出来的（没有表、一格里有多个数字…）交给 AI（开发者服务器 /quote），AI 的结果要能在原文里找到才用。
- 读过的邮件只把解析结果缓存在本机 feishu_mail_cache.json（不进 Git），下次不用重新下载。

用法:
  python reconcile.py 出库单.xlsx [--days 60] [-o 结果.xlsx]
  python reconcile.py 出库单.xlsx --eml 报价1.eml 报价2.eml     # 离线：用导出的邮件文件代替飞书
  python reconcile.py 出库单.xlsx --code "http://localhost:8080/callback?code=…"   # 在别的设备上授权后
"""
import argparse
import email
import email.utils
import html
import json
import re
import sys
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta
from email import policy
from html.parser import HTMLParser
from pathlib import Path

import openpyxl
from openpyxl.styles import Font

import app_ui
import license_check
from app_ui import APP_DIR

MR = re.compile(r"MR\d{8}(?!\d)")
TOKEN = re.compile(r"[A-Za-z0-9](?:[A-Za-z0-9\-_./]*[A-Za-z0-9])?")
CACHE_FILE = APP_DIR / "feishu_mail_cache.json"
PARSER_VERSION = 3           # 解析规则变了就加 1，旧缓存作废
PENDING = re.compile(r"待定|待报|待确认|未定|再定|TBD|TBC|pending", re.I)
QUOTE_SUBJECT = "报价"
AI_TEXT_LIMIT = 6000
MAX_MESSAGES = 5000          # 一个文件夹最多看多少封，防止无限翻页
PARALLEL = 4

HEADERS = ["出库时间", "编号", "出库单号", "报价(欧元)", "状态", "报价邮件时间", "报价邮件主题", "说明"]


# ------------------------------------------------------------------ 出库单

def read_outbound(path):
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
    it = ws.iter_rows(values_only=True)
    head = [str(h or "").strip() for h in next(it, ())]

    def col(prefix):
        i = next((i for i, h in enumerate(head) if h.startswith(prefix)), None)
        if i is None:
            if "编号" in head:
                raise SystemExit("这是询价表，不是领星导出的出库单。对账请选领星导出的出库单 Excel。")
            raise SystemExit(f"Excel 第 1 行找不到\"{prefix}\"列，看起来不是领星导出的出库单。")
        return i

    c_no, c_time, c_rm = col("Outbound Order No"), col("OutboundTime"), col("Remark")
    out = []
    for r in it:
        no = str(r[c_no] or "").strip() if c_no < len(r) else ""
        if not no:
            continue
        t = r[c_time] if c_time < len(r) else None
        remark = str(r[c_rm] or "") if c_rm < len(r) else ""
        out.append({"no": no, "time": t, "remark": remark.strip(), "codes": code_tokens(remark),
                    "mrs": list(dict.fromkeys(MR.findall(remark.upper())))})
    return out


def code_tokens(s):
    """格子或备注里像编号的片段（同时有字母和数字、至少 4 个字符），如 MR12345678、ZR-0101-1、MR12345678-aZR。"""
    out = [t.upper() for t in TOKEN.findall(s or "")
           if len(t) >= 4 and re.search(r"[A-Za-z]", t) and re.search(r"\d", t)]
    return list(dict.fromkeys(out))


def code_key(code):
    """比较编号时忽略大小写和分隔符：ZR0101-1 = ZR-0101-1，但 ZR-0101-1 ≠ ZR-0101-12。"""
    return re.sub(r"[^A-Z0-9]", "", code.upper())


def parse_time(v):
    if isinstance(v, datetime):
        return v
    try:
        return datetime.strptime(str(v or "").strip()[:19], "%Y-%m-%d %H:%M:%S")
    except ValueError:
        return None


def time_text(v):
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d %H:%M:%S")
    return str(v or "").strip()


# ------------------------------------------------------------------ 报价邮件解析

QUOTE_MARKERS = [re.compile(p, re.I) for p in (
    r'<div[^>]+class="[^"]*history-quote-wrapper',      # 飞书
    r"<blockquote", r'<div[^>]+class="[^"]*gmail_quote', r'<div[^>]+id="divRplyFwdMsg"',
    r'<div[^>]+id="appendonsend"', r"-----\s*(Original Message|原始邮件)")]
TEXT_QUOTE = re.compile(r"^(>|-----\s*(Original Message|原始邮件)|发件人[:：].*|From:.*)", re.M)


def own_html(h):
    """去掉回复里引用的旧邮件，只留这封邮件自己写的部分。"""
    cut = min((m.start() for p in QUOTE_MARKERS for m in [p.search(h)] if m), default=len(h))
    return h[:cut]


def own_text(t):
    m = TEXT_QUOTE.search(t or "")
    return (t[:m.start()] if m else t or "").strip()


class _Tables(HTMLParser):
    """把 HTML 里的表格读成 [表[行[(格子文字, 格子 id)]]]。colspan 展开成多格（后面的为空），
    rowspan（合并的几行）往下复制同一格子；格子 id = 原格子的 (行, 列)，用来认出"同一个报价格"。"""
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.tables, self.stack = [], []

    def handle_starttag(self, tag, attrs):
        if tag == "table":
            self.stack.append([])
        elif tag == "tr" and self.stack:
            self.stack[-1].append([])
        elif tag in ("td", "th") and self.stack and self.stack[-1]:
            a = dict(attrs)
            self.stack[-1][-1].append([a.get("colspan") or "1", a.get("rowspan") or "1"])
        elif tag in ("br", "div", "p") and self.stack and self.stack[-1] and self.stack[-1][-1]:
            self.stack[-1][-1][-1].append("\n")

    @staticmethod
    def _span(v):
        try:
            return max(1, min(int(v), 50))
        except ValueError:
            return 1

    def handle_endtag(self, tag):
        if tag == "table" and self.stack:
            rows, down = [], {}          # down: 列 → [还要往下占几行, 文字, id]
            for r, row in enumerate(self.stack.pop()):
                cells, src = [], iter(row)
                while True:
                    col = len(cells)
                    if col in down:
                        d = down[col]
                        cells.append((d[1], d[2]))
                        d[0] -= 1
                        if d[0] <= 0:
                            del down[col]
                        continue
                    c = next(src, None)
                    if c is None:
                        break
                    txt = re.sub(r"[ \t\xa0]+", " ", "".join(c[2:]))
                    txt = re.sub(r"\s*\n\s*", "\n", txt).strip()
                    cs, rs = self._span(c[0]), self._span(c[1])
                    for k in range(cs):
                        cells.append((txt if k == 0 else "", (r, col)))
                        if rs > 1:
                            down[col + k] = [rs - 1, txt if k == 0 else "", (r, col)]
                rows.append(cells)
            self.tables.append(rows)

    def handle_data(self, d):
        if self.stack and self.stack[-1] and self.stack[-1][-1]:
            self.stack[-1][-1][-1].append(d)


def html_to_text(h):
    h = re.sub(r"(?is)<(style|script|head)\b.*?</\1>", "", h)
    h = re.sub(r"(?i)<br\s*/?>|</(p|div|tr|li|h\d)>", "\n", h)
    h = re.sub(r"(?i)</t[dh]>", " | ", h)
    h = re.sub(r"<[^>]+>", "", h)
    t = html.unescape(h).replace("\xa0", " ")
    return re.sub(r"\n\s*\n+", "\n", re.sub(r"[ \t]+", " ", t)).strip()


def parse_price(raw):
    """报价格子 → ("price", 金额) / ("pending", None) / ("unclear", None)；空格子返回 None。"""
    t = (raw or "").strip()
    if not t:
        return None
    nums = re.findall(r"\d+(?:[.,]\d+)*", t)
    if not nums:
        return ("pending", None) if PENDING.search(t) or re.fullmatch(r"[-/—?？]+", t) else ("unclear", None)
    if len(nums) > 1 or PENDING.search(t):
        return ("unclear", None)
    n = nums[0]
    if re.fullmatch(r"\d{1,3}(,\d{3})+(\.\d+)?", n):        # 1,200 / 1,200.50
        n = n.replace(",", "")
    elif re.fullmatch(r"\d{1,3}(\.\d{3})+(,\d+)?", n):      # 1.200 / 1.200,50
        n = n.replace(".", "").replace(",", ".")
    else:
        n = n.replace(",", ".")
    try:
        return ("price", float(n))
    except ValueError:
        return ("unclear", None)


def quotes_from_tables(h):
    """{编号: (报价格子原文, 报价格子 id)}：表头有"编号"和"运费报价"的表格，每个编号取第一次出现且报价
    不为空的格子。几个编号共用一个报价格子（同一格写多个编号 / 报价格合并了几行）时 id 相同。"""
    p = _Tables()
    p.feed(h)
    out = {}
    for t, rows in enumerate(p.tables):
        c_no = c_price = None
        for row in rows:
            norm = [re.sub(r"\s+", "", c[0]) for c in row]
            if "编号" in norm and any(c.startswith("运费报价") for c in norm):
                c_no, c_price = norm.index("编号"), next(i for i, c in enumerate(norm) if c.startswith("运费报价"))
                continue
            if c_no is None or max(c_no, c_price) >= len(row):
                continue
            price, cid = row[c_price][0].strip(), row[c_price][1]
            for code in code_tokens(row[c_no][0]):
                if price and code not in out:
                    out[code] = (price, f"{t}:{cid[0]}:{cid[1]}")
    return out


def parse_mail(subject, sender, date_ms, html_body, text_body):
    """一封邮件 → 缓存记录。q: {编号: [状态, 金额, 原文]}；ai: 要交给 AI 的编号；t: AI 用的正文。"""
    rec = {"v": PARSER_VERSION, "d": date_ms, "s": subject, "f": sender, "q": {}}
    mine_html = own_html(html_body) if html_body else ""
    text = (html_to_text(mine_html) if mine_html else own_text(text_body))[:AI_TEXT_LIMIT]
    for mr, (raw, gid) in quotes_from_tables(mine_html).items() if mine_html else ():
        st = parse_price(raw)
        if st:
            rec["q"][mr] = [st[0], st[1], raw[:100], gid]
    need = [mr for mr, v in rec["q"].items() if v[0] == "unclear"]
    if QUOTE_SUBJECT in subject:          # 报价邮件里正文提到、表格里没读到报价的编号
        in_table = {mr for c in rec["q"] for mr in MR.findall(c)}     # 表格写 MR…-aZR 也算有了
        need += [mr for mr in dict.fromkeys(MR.findall(text.upper())) if mr not in in_table]
    if need:
        rec["ai"], rec["t"] = need, text
    if not rec["q"] and not need:
        return {"v": PARSER_VERSION, "d": date_ms, "skip": 1}
    return rec


def mail_from_eml(path):
    m = email.message_from_binary_file(open(path, "rb"), policy=policy.default)
    h = t = ""
    for part in m.walk():
        if part.get_filename():
            continue
        if part.get_content_type() == "text/html" and not h:
            h = part.get_content()
        elif part.get_content_type() == "text/plain" and not t:
            t = part.get_content()
    try:
        date_ms = int(email.utils.parsedate_to_datetime(m["Date"]).timestamp() * 1000)
    except Exception:
        date_ms = 0
    sender = email.utils.parseaddr(str(m["From"] or ""))[1]
    return parse_mail(str(m["Subject"] or ""), sender, date_ms, h, t)


# ------------------------------------------------------------------ 飞书

def _load_cache():
    try:
        c = json.loads(CACHE_FILE.read_text(encoding="utf-8"))
        return {k: v for k, v in c.items() if isinstance(v, dict) and v.get("v") == PARSER_VERSION}
    except Exception:
        return {}


def _save_cache(c):
    tmp = CACHE_FILE.with_suffix(".tmp")
    tmp.write_text(json.dumps(c, ensure_ascii=False), encoding="utf-8")
    tmp.replace(CACHE_FILE)


def mail_domain(addr):
    return (addr or "").rpartition("@")[2].strip().lower()


def fetch_feishu(since, code=None):
    """读飞书邮箱里 since 之后的邮件，返回 ({邮件 id: 记录}（含以前缓存的）, 自己的邮箱域名或 "")。
    自己的域名取自“已发送”里的发件人，不写死在代码里。"""
    import feishu_mail
    token = feishu_mail.get_access_token(code)
    cache = _load_cache()
    cutoff = since.timestamp() * 1000
    new, own = 0, ""

    def fetch(mid):
        m = feishu_mail.get_message(token, mid)
        return mid, parse_mail(m["subject"], m["sender"], m["date"], m["html"], m["text"])

    with ThreadPoolExecutor(PARALLEL) as pool:
        folders = feishu_mail.list_folders(token)
        print(f"  邮箱文件夹：{'、'.join(n for _, n in folders)}")
        for folder, fname in folders:
            page, seen = None, 0
            while True:
                if app_ui.stop_requested():
                    raise SystemExit("已停止")
                try:
                    ids, page = feishu_mail.list_messages(token, folder, page)
                except feishu_mail.FeishuError as e:
                    if folder == "INBOX":
                        raise
                    print(f"  读取文件夹「{fname}」失败，跳过：{e}")
                    break
                if folder.upper() == "SENT" and ids and not own:
                    own = mail_domain(cache.get(ids[0], {}).get("f")
                                      or feishu_mail.get_message(token, ids[0])["sender"])
                todo = [i for i in ids if i not in cache]
                for mid, rec in pool.map(fetch, todo):
                    cache[mid] = rec
                new += len(todo)
                seen += len(ids)
                if todo:
                    _save_cache(cache)
                print(f"  {fname}: 已看 {seen} 封邮件（新下载 {new} 封）", flush=True)
                # 邮件按时间倒序：一整页都早于起始日期就不用再往后翻
                if not page or not ids or all(cache[i]["d"] and cache[i]["d"] < cutoff for i in ids) \
                        or seen >= MAX_MESSAGES:
                    break
    return cache, own


# ------------------------------------------------------------------ AI 兜底

def run_ai(mails, wanted):
    """对需要 AI 的（邮件, 编号）调用服务器 /quote，结果写进 mails[id]["aires"]。返回是否改动了记录。"""
    jobs = [(mid, mr) for mid, r in mails.items() for mr in r.get("ai", [])
            if code_key(mr) in wanted and mr not in r.get("aires", {})]
    if not jobs:
        return False
    if not license_check.get_token():
        print("  没有授权码，表格读不出的报价不能交给 AI，只能人工确认")
        return False
    print(f"  {len(jobs)} 个报价表格读不出来，交给 AI 读邮件 ...")
    changed = False
    for i in range(0, len(jobs), 100):
        part = jobs[i:i + 100]
        rows = [{"row": n, "编号": mr, "邮件主题": mails[mid]["s"], "邮件正文": mails[mid]["t"]}
                for n, (mid, mr) in enumerate(part)]
        try:
            status, data = license_check.post("/quote", {"rows": rows}, 180)
        except Exception as e:
            print(f"  [AI] 连不上服务器（{e}），这些报价需要人工确认")
            return changed
        if status != 200 or not isinstance(data, dict) or not isinstance(data.get("results"), list):
            why = (data.get("error") or data.get("message")) if isinstance(data, dict) else ""
            print(f"  [AI] 调用失败（HTTP {status}{'：' + why if why else ''}），这些报价需要人工确认")
            return changed
        for it in data["results"]:
            try:
                mid, mr = part[int(it["row"])]
            except (KeyError, TypeError, ValueError, IndexError):
                continue
            mails[mid].setdefault("aires", {})[mr] = _check_ai(it, mails[mid]["t"])
            changed = True
    return changed


def _check_ai(it, text):
    """AI 给的金额必须在它引用的原文里，引用的原文必须在邮件里，否则不用。返回 [状态, 金额, 原文]。"""
    price, ev = it.get("price"), str(it.get("evidence") or "").strip()
    if price is None:
        return ["none", None, ""]
    squash = lambda s: re.sub(r"\s+", "", s)
    try:
        price = float(price)
    except (TypeError, ValueError):
        return ["unclear", None, ev]
    nums = [parse_price(n) for n in re.findall(r"\d+(?:[.,]\d+)*", ev)]
    if not ev or squash(ev) not in squash(text) or not any(n and n[1] == price for n in nums):
        return ["unclear", None, ev]
    return ["price", price, ev]


# ------------------------------------------------------------------ 对账

def build_index(mails):
    """报价邮件里出现过的所有编号：{比较用的 key: {原写法}}，以及 {MR 号: {key}}。"""
    by_key, by_mr = {}, {}
    for r in mails.values():
        for code in list(r.get("q", {})) + list(r.get("ai", [])):
            k = code_key(code)
            by_key.setdefault(k, set()).add(code)
            for mr in MR.findall(code):
                by_mr.setdefault(mr, set()).add(k)
    return by_key, by_mr


def match_order(o, by_key, by_mr):
    """出库单备注 → (对上的编号 key 列表, 模糊说明)。
    先找备注里和报价表编号完全一致的片段；没有再按 MR 号找（报价表写 MR12345678-aZR 这种）。
    一个 MR 号对上报价表里好几个编号时返回全部，并给出模糊说明：它们共用一个报价格就没问题，
    分开报价的话不知道这张出库单是哪几个，要人工确认（在 reconcile 里判断）。"""
    keys = [code_key(t) for t in o["codes"] if code_key(t) in by_key]
    if keys:
        return list(dict.fromkeys(keys)), ""
    why = ""
    for mr in o["mrs"]:
        ks = sorted(by_mr.get(mr, ()))
        keys += ks
        if len(ks) > 1:
            names = "、".join(sorted(c for k in ks for c in by_key[k]))
            why = f"报价邮件里 {mr} 有多条、分开报价：{names}，请人工确认"
    return list(dict.fromkeys(keys)), why


def lookup(k, mails):
    """这个编号最新的报价：(状态, 金额, 邮件记录, 原文, 报价格子)；没找到返回 None。
    报价格子 = (邮件 id, 格子 id)，几个编号共用同一个格子时相同。
    表格里空着的、AI 说没有报价的都不算（可能只是转发、回复里提到）。"""
    found = []
    for mid, r in mails.items():
        v = next((val for c, val in r.get("q", {}).items() if code_key(c) == k), None)
        cell = (mid, v[3]) if v and len(v) > 3 else (mid, k)
        ai_code = next((c for c in r.get("ai", []) if code_key(c) == k), None)
        a = r.get("aires", {}).get(ai_code) if ai_code else None
        if a and a[0] != "none":
            v = a
        elif a and a[0] == "none" and (not v or v[0] == "unclear"):
            v = ["unclear", None, v[2]] if v else None
        elif v is None and ai_code:
            v = ["unclear", None, ""]          # 没交给 AI（没授权码 / 调用失败）
        if v:
            found.append((r.get("d") or 0, v, r, cell))
    if not found:
        return None
    d, v, r, cell = max(found, key=lambda x: x[0])
    return v[0], v[1], r, v[2], cell


def reconcile(orders, mails):
    by_key, by_mr = build_index(mails)
    rows = []
    for o in orders:
        keys, why = match_order(o, by_key, by_mr)
        shown = " ".join(sorted(by_key[k])[0] for k in keys) or " ".join(o["mrs"] or o["codes"][:2])
        base = [time_text(o["time"]), shown, o["no"]]
        if not keys:
            st = "未找到报价邮件" if o["codes"] else "备注里没有编号"
            rows.append(base + [None, st, "", "", "" if o["codes"] else o["remark"][:200]])
            continue
        # 一张出库单一个报价：几个编号最新的报价落在同一个报价格子里（同一格写多个编号 / 合并格）只算一次
        cells, missing = {}, False
        for k in keys:
            hit = lookup(k, mails)
            if hit is None:
                missing = True
                continue
            st, price, rec, raw, cell = hit
            cells.setdefault(cell, [st, price, rec, raw, []])[4].append(sorted(by_key[k])[0])
        if why and len(cells) > 1:
            rows.append(base + [None, "需人工确认", "", "", why])
            continue
        if len(cells) > 1:          # 一张出库单应该只有一个报价，分开报价说明有问题
            parts = "；".join(f"{'+'.join(c[4])}={c[3]}" for c in cells.values())
            rows.append(base + [None, "需人工确认", "", "", f"多个编号分开报价，一张出库单应只有一个报价：{parts}"])
            continue
        prices, notes, srcs, states = [], [], [], ["未找到报价邮件"] if missing else []
        for st, price, rec, raw, names in cells.values():
            name = "+".join(names)
            srcs.append(rec)
            if st == "price":
                prices.append(price)
                notes.append(f"{name}={raw}")
                states.append("已找到")
            elif st == "pending":
                states.append("最新报价待定")
                notes.append(f"{name}={raw}")
            else:
                states.append("需人工确认")
                notes.append(f"{name}：报价写法看不懂{'（' + raw + '）' if raw else ''}")
        ok = bool(states) and all(s == "已找到" for s in states)
        status = "已找到" if ok else next(s for s in states if s != "已找到")
        price = round(sum(prices), 2) if ok else None
        if len(keys) > 1 and ok:
            notes = [f"多个编号同一个报价：{notes[0]}"]
        src = max(srcs, key=lambda r: r.get("d") or 0) if srcs else None
        when = datetime.fromtimestamp(src["d"] / 1000).strftime("%Y-%m-%d %H:%M") if src and src.get("d") else ""
        rows.append(base + [price, status, when, src["s"] if src else "",
                            "；".join(notes) if len(keys) > 1 or not ok else ""])
    return rows


def write_xlsx(rows, out):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "对账"
    ws.append(HEADERS)
    for r in rows:
        ws.append(r)
    for c, w in zip("ABCDEFGH", (20, 14, 18, 11, 14, 17, 50, 40)):
        ws.column_dimensions[c].width = w
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for r in ws.iter_rows(min_row=2, min_col=4, max_col=4):
        r[0].number_format = "0.00"
    ws.freeze_panes = "A2"
    while True:
        try:
            wb.save(out)
            return True
        except PermissionError:
            if app_ui.ask(f"保存失败：{out.name} 可能正在 Excel 里打开。关掉后点\"重试\"。",
                          [("r", "重试"), ("q", "不保存")]) == "q":
                return False


def run(args):
    path = Path(args.xlsx)
    orders = read_outbound(path)
    if not orders:
        raise SystemExit("Excel 里没有出库单")
    times = [t for t in (parse_time(o["time"]) for o in orders) if t]
    since = (min(times) if times else datetime.now() - timedelta(days=31)) - timedelta(days=args.days)
    print(f"出库单 {len(orders)} 张；搜索 {since:%Y-%m-%d} 之后的邮件")

    if args.eml:
        mails, own = {str(p): mail_from_eml(p) for p in args.eml}, ""
    else:
        mails, own = fetch_feishu(since, getattr(args, "code", None))
    quotes = {k: v for k, v in mails.items() if not v.get("skip")}
    # 报价 = 我们发给客户的：只认和自己同一个邮箱域名的发件人（客户的出货通知里也有报价表，可能抄的是旧价）
    if own:
        other = sum(1 for v in quotes.values() if mail_domain(v.get("f")) != own)
        quotes = {k: v for k, v in quotes.items() if mail_domain(v.get("f")) == own}
        print(f"找到 {len(quotes)} 封我们发出的报价邮件（@{own}；另有 {other} 封外部邮件带报价表，不算）")
    else:
        print(f"找到 {len(quotes)} 封有报价的邮件" + ("" if args.eml else "（没读到自己的发件地址，所有发件人都算）"))
    by_key, by_mr = build_index(quotes)
    wanted = {k for o in orders for k in match_order(o, by_key, by_mr)[0]}
    if run_ai(quotes, wanted) and not args.eml:
        _save_cache(mails)

    rows = reconcile(orders, quotes)
    out = (Path(args.o) if args.o else path.with_name(f"{path.stem}_对账结果.xlsx")).resolve()
    saved = write_xlsx(rows, out)
    count = {}
    for r in rows:
        count[r[4]] = count.get(r[4], 0) + 1
    print("，".join(f"{k} {v}" for k, v in count.items()))
    for r in rows:
        if r[4] != "已找到":
            print(f"  [{r[4]}] {r[2]} {r[1]} {r[7]}")
    print(f"\n对账结果已保存到：{out}" if saved else "\n对账结果没有保存")


def main():
    ap = argparse.ArgumentParser(description="按出库单在飞书邮件里找报价，生成对账 Excel")
    ap.add_argument("xlsx", help="领星导出的出库单 Excel")
    ap.add_argument("--days", type=int, default=60, help="从最早的出库时间再往前找多少天的邮件（默认 60）")
    ap.add_argument("-o", help="输出文件，默认 <Excel名>_对账结果.xlsx")
    ap.add_argument("--eml", nargs="+", help="离线：用导出的 .eml 邮件代替飞书")
    ap.add_argument("--code", help="不在这台电脑上授权时：飞书授权后跳转的地址（或其中的 code），几分钟内有效")
    run(ap.parse_args())


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    main()
