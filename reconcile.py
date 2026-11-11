#!/usr/bin/env python3
"""对账：领星出库单 Excel → 按备注里的编号（MR…）在飞书邮件里找我们给客户的报价 → 生成对账 Excel。

报价邮件：回复客户"出库询价"的邮件，正文第一张表有"编号"和"运费报价(欧元)"两列（如"320欧"）。
- 只看邮件自己写的部分，后面引用的原询价邮件（报价是空的）不看。
- 一个编号有多封报价邮件时取最新的一封；最新的写"待定"等 → 报价留空。
- 一张出库单对应一个报价：备注里的几个编号在报价表里共用同一个报价格子（同一格写多个编号，
  或报价格合并了几行）时只算一次。
- 表格读不出来的（没有表、一格里有多个数字…）交给 AI（开发者服务器 /quote），AI 的结果要能在原文里找到才用。
- 客户的"卡派出货通知"（主题含"出货通知"，表格有 No./费用情况/出库订单号/派送方式）= 客户确认的价格：
  先按出库单号、再按编号对上出库单，和我们的最新报价比，不一样要人工确认；没有我们的报价时用它的价格（标橙）。
- 只对账卡派（Dachser）单：有报价 / 出货通知写卡派 / 有询价 / 物流渠道是 Dachser 都算卡派；
  快递单和看不出来的单放进"未对账"表，不标红。
- 读过的邮件只把解析结果缓存在本机 feishu_mail_cache.json（不进 Git），下次不用重新下载。

用法:
  python reconcile.py 出库单.xlsx [--days 60] [-o 结果.xlsx]
  python reconcile.py 出库单.xlsx --eml 报价1.eml 报价2.eml     # 离线：用导出的邮件文件代替飞书
  python reconcile.py 出库单.xlsx --code "http://localhost:8080/callback?code=…"   # 在别的设备上授权后
"""
import argparse
import email
import email.utils
import gzip
import hashlib
import html
import json
import time
import re
import sys
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta
from email import policy
from html.parser import HTMLParser
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill

import app_ui
import license_check
from app_ui import APP_DIR

MR = re.compile(r"MR\d{8}(?!\d)")
TOKEN = re.compile(r"[A-Za-z0-9](?:[A-Za-z0-9\-_./]*[A-Za-z0-9])?")
CACHE_FILE = APP_DIR / "feishu_mail_cache.json"
RAW_DIR = APP_DIR / "feishu_mail_raw"
PARSER_VERSION = 4           # 解析规则变了就加 1：旧记录用本机保存的原文重新解析，不用重新下载
PENDING = re.compile(r"待定|待报|待确认|未定|再定|TBD|TBC|pending", re.I)
QUOTE_SUBJECT = "报价"
NOTICE_SUBJECT = "出货通知"
INQUIRY_SUBJECT = "询价"
EXPRESS = re.compile(r"快递|DHL|UPS|GLS|FedEx|DPD|TNT|Hermes|PostNL|Colissimo|Parcel", re.I)
TRUCK = re.compile(r"Dachser|卡派", re.I)
AI_TEXT_LIMIT = 6000
MAX_MESSAGES = 5000          # 一个文件夹最多看多少封，防止无限翻页
PARALLEL = 4

HEADERS = ["出库时间", "编号", "出库单号", "报价(欧元)", "状态", "依据", "客户确认价(欧元)", "成本价(欧元)",
           "报价邮件时间", "报价邮件主题", "出货通知时间", "说明"]
NOTE = HEADERS.index("说明")
SKIP_HEADERS = ["出库时间", "出库单号", "物流渠道", "判断", "原因", "备注"]
AI_MARK = "AI 读取："
NOTICE_MARK = "客户确认："
ONLY_NOTICE = "仅客户确认价"


# ------------------------------------------------------------------ 出库单

def read_outbound(path):
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
    it = ws.iter_rows(values_only=True)
    head = [str(h or "").strip() for h in next(it, ())]

    def col(prefix, required=True):
        i = next((i for i, h in enumerate(head) if h.startswith(prefix)), None)
        if i is None and required:
            if "编号" in head:
                raise SystemExit("这是询价表，不是领星导出的出库单。对账请选领星导出的出库单 Excel。")
            raise SystemExit(f"Excel 第 1 行找不到\"{prefix}\"列，看起来不是领星导出的出库单。")
        return i

    c_no, c_time, c_rm = col("Outbound Order No"), col("OutboundTime"), col("Remark")
    c_ship = col("Shipping service", False)       # 物流渠道：不一定准（常见"无需物流服务"），只作参考
    out = []
    for r in it:
        no = str(r[c_no] or "").strip() if c_no < len(r) else ""
        if not no:
            continue
        t = r[c_time] if c_time < len(r) else None
        remark = str(r[c_rm] or "") if c_rm < len(r) else ""
        ship = str(r[c_ship] or "").strip() if c_ship is not None and c_ship < len(r) else ""
        out.append({"no": no, "time": t, "remark": remark.strip(), "ship": ship, "codes": code_tokens(remark),
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


def _norm_cells(row):
    return [re.sub(r"\s+", "", c[0]) for c in row]


def notices_from_tables(h, subject):
    """客户的出货通知 → [[编号列表, 费用格子原文, 出库单号, 派送方式, 费用格子 id]]。
    表头要有 No.（或"编号"）和"费用情况"（或"运费报价"）；出库订单号、派送方式可以没有，
    没有派送方式时看主题里写的卡派 / 快递。"""
    p = _Tables()
    p.feed(h)
    default_way = "卡派" if "卡派" in subject else "快递" if "快递" in subject else ""
    out = []
    for t, rows in enumerate(p.tables):
        cols = None
        for row in rows:
            norm = _norm_cells(row)
            i_no = next((i for i, c in enumerate(norm) if c.lower().rstrip(".:：") in ("no", "编号")), None)
            i_pr = next((i for i, c in enumerate(norm) if c.startswith(("费用情况", "运费报价"))), None)
            if i_no is not None and i_pr is not None:
                cols = (i_no, i_pr,
                        next((i for i, c in enumerate(norm) if c.startswith(("出库订单号", "出库单号"))), None),
                        next((i for i, c in enumerate(norm) if c.startswith("派送方式")), None))
                continue
            if cols is None:
                continue
            get = lambda i: row[i][0].strip() if i is not None and i < len(row) else ""
            codes = code_tokens(get(cols[0]))
            obs = (TOKEN.findall(get(cols[2])) or [""])[0].upper()
            if not codes and not obs:
                continue
            cid = row[cols[1]][1] if cols[1] < len(row) else (-1, -1)
            out.append([codes, get(cols[1])[:100], obs, get(cols[3])[:30] or default_way, f"{t}:{cid[0]}:{cid[1]}"])
    return out


def inquiry_codes(subject, h, text):
    """询价邮件里的编号：主题和正文里的 MR 号 + 表格"编号"列里的编号。"""
    codes = MR.findall(subject.upper()) + MR.findall(text.upper())
    p = _Tables()
    p.feed(h or "")
    for rows in p.tables:
        c = None
        for row in rows:
            norm = _norm_cells(row)
            if c is None:
                c = norm.index("编号") if "编号" in norm else None
            elif c < len(row):
                codes += code_tokens(row[c][0])
    return list(dict.fromkeys(codes))


def parse_mail(subject, sender, date_ms, html_body, text_body):
    """一封邮件 → 缓存记录。q: {编号: [状态, 金额, 原文]}；ai: 要交给 AI 的编号；t: AI 用的正文；
    n: 出货通知的行（见 notices_from_tables）；iq: 询价邮件里的编号。"""
    rec = {"v": PARSER_VERSION, "d": date_ms, "s": subject, "f": sender, "q": {}}
    mine_html = own_html(html_body) if html_body else ""
    text = (html_to_text(mine_html) if mine_html else own_text(text_body))[:AI_TEXT_LIMIT]
    if NOTICE_SUBJECT in subject:         # 客户的出货通知：只当"客户确认价"，不当我们的报价
        rows = notices_from_tables(mine_html, subject) if mine_html else []
        if not rows:
            return {"v": PARSER_VERSION, "d": date_ms, "f": sender, "skip": 1}
        return {"v": PARSER_VERSION, "d": date_ms, "s": subject, "f": sender, "n": rows}
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
    if INQUIRY_SUBJECT in subject:
        iq = inquiry_codes(subject, mine_html, text)
        if iq:
            rec["iq"] = iq
    if not rec["q"] and not need and not rec.get("iq"):
        return {"v": PARSER_VERSION, "d": date_ms, "f": sender, "skip": 1}
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
    """解析结果缓存（含旧版本的记录，fetch_feishu 里用本机原文重新解析）。"""
    try:
        c = json.loads(CACHE_FILE.read_text(encoding="utf-8"))
        return {k: v for k, v in c.items() if isinstance(v, dict)}
    except Exception:
        return {}


def _raw_path(mid):
    return RAW_DIR / (hashlib.sha1(mid.encode("utf-8")).hexdigest() + ".json.gz")


def _save_raw(mid, m):
    """邮件原文压缩存在本机（feishu_mail_raw/，不进 Git、不打包）：解析规则改了不用重新下载。
    正文里内嵌的图片（data: 地址）去掉，省空间。"""
    strip = lambda s: re.sub(r"""(src=["']?)data:[^"'\s>]{200,}""", r"\1", s or "")
    rec = {"subject": m["subject"], "sender": m["sender"], "date": m["date"],
           "html": strip(m["html"]), "text": m["text"] or ""}
    RAW_DIR.mkdir(exist_ok=True)
    p = _raw_path(mid)
    tmp = p.with_suffix(".tmp")
    tmp.write_bytes(gzip.compress(json.dumps(rec, ensure_ascii=False).encode("utf-8")))
    tmp.replace(p)


def _load_raw(mid):
    try:
        return json.loads(gzip.decompress(_raw_path(mid).read_bytes()).decode("utf-8"))
    except Exception:
        return None


def _parse_raw(m):
    return parse_mail(m["subject"], m["sender"], m["date"], m["html"], m["text"])


def reparse_outdated(cache):
    """旧解析规则的记录：本机有原文就重新解析（保留已有的 AI 结果），没有原文的删掉、之后重新下载。"""
    old = [k for k, v in cache.items() if v.get("v") != PARSER_VERSION]
    done = 0
    for k in old:
        m = _load_raw(k)
        if m is None:
            del cache[k]
            continue
        rec = _parse_raw(m)
        ai_old = cache[k].get("aires") or {}
        keep = {mr: v for mr, v in ai_old.items() if mr in rec.get("ai", []) and rec.get("t") == cache[k].get("t")}
        if keep:
            rec["aires"] = keep
        cache[k] = rec
        done += 1
    if old:
        print(f"  解析规则更新了：{done} 封邮件用本机保存的原文重新解析"
              + (f"，{len(old) - done} 封本机没有原文，需要重新下载" if len(old) > done else "（不用重新下载）"))
    return bool(old)


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
    if reparse_outdated(cache):
        _save_cache(cache)
    cutoff = since.timestamp() * 1000
    new, own = 0, ""

    def fetch(mid):
        m = feishu_mail.get_message(token, mid)
        try:
            _save_raw(mid, m)
        except OSError as e:
            print(f"  邮件原文保存失败（不影响这次对账）：{e}")
        return mid, _parse_raw(m)

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
                # 没缓存的，或缓存了但本机没原文的（旧版本留下的）才下载
                todo = [i for i in ids if i not in cache or not _raw_path(i).exists()]
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


def _entries(k, mails):
    """这个编号在每封报价邮件里的报价：[(时间, [状态, 金额, 原文, …], 邮件记录, 报价格子, 是否 AI 读的)]。
    报价格子 = (邮件 id, 格子 id)，几个编号共用同一个格子时相同。
    表格里空着的、AI 说没有报价的都不算（可能只是转发、回复里提到）。"""
    found = []
    for mid, r in mails.items():
        v = next((val for c, val in r.get("q", {}).items() if code_key(c) == k), None)
        cell = (mid, v[3]) if v and len(v) > 3 else (mid, k)
        ai_code = next((c for c in r.get("ai", []) if code_key(c) == k), None)
        a = r.get("aires", {}).get(ai_code) if ai_code else None
        by_ai = bool(a and a[0] != "none")
        if by_ai:
            v = a
        elif a and a[0] == "none" and (not v or v[0] == "unclear"):
            v = ["unclear", None, v[2]] if v else None
        elif v is None and ai_code:
            v = ["unclear", None, ""]          # 没交给 AI（调用失败 / 次数用完）
        if v:
            found.append((r.get("d") or 0, v, r, cell, by_ai))
    return found


def lookup(k, mails):
    """这个编号最新的报价：(状态, 金额, 邮件记录, 原文, 报价格子, 是否 AI 读的)；没找到返回 None。"""
    found = _entries(k, mails)
    if not found:
        return None
    d, v, r, cell, by_ai = max(found, key=lambda x: x[0])
    return v[0], v[1], r, v[2], cell, by_ai


def fmt(n):
    return f"{n:g}" if isinstance(n, (int, float)) else str(n)


def day(ms):
    return datetime.fromtimestamp(ms / 1000).strftime("%Y-%m-%d %H:%M") if ms else ""


def quote_history(k, mails):
    """报价改过的编号：按时间排的 [(时间, 价格或"待定")]（相邻相同的合并）；只报过一个价返回 []。"""
    seq = []
    for d, v, *_ in sorted(_entries(k, mails), key=lambda x: x[0]):
        if v[0] in ("price", "pending"):
            lab = fmt(v[1]) if v[0] == "price" else "待定"
            if not seq or seq[-1][1] != lab:
                seq.append((d, lab))
    return seq if len(seq) > 1 else []


def eval_quote(o, by_key, by_mr, mails):
    """我们的报价 → {keys, codes, price, status, basis, src, notes, changed}。"""
    keys, why = match_order(o, by_key, by_mr)
    res = {"keys": keys, "codes": [sorted(by_key[k])[0] for k in keys], "price": None, "basis": "",
           "src": None, "notes": [], "changed": False, "cells": [], "qdate": 0}
    if not keys:
        res["status"] = "未找到报价邮件" if o["codes"] else "备注里没有编号"
        return res
    hist = {}                   # 报价改过（同一编号前后报了不同的价）要标出来；同样的改动只写一次
    for k in keys:
        h = quote_history(k, mails)
        if h:
            hist.setdefault(" → ".join(f"{day(d)[5:10]} {p}" for d, p in h), []).append(sorted(by_key[k])[0])
    for h, names in hist.items():
        res["changed"] = True
        res["notes"].append(f"{'+'.join(names)} 报价有更新：{h}")
    # 一张出库单一个报价：几个编号最新的报价落在同一个报价格子里（同一格写多个编号 / 合并格）只算一次
    cells, missing = {}, False
    for k in keys:
        hit = lookup(k, mails)
        if hit is None:
            missing = True
            continue
        st, price, rec, raw, cell, by_ai = hit
        raw = (AI_MARK + raw) if by_ai else raw
        cells.setdefault(cell, [st, price, rec, raw, []])[4].append(sorted(by_key[k])[0])
    res["cells"] = list(cells)
    res["qdate"] = max((c[2].get("d") or 0 for c in cells.values()), default=0)
    if why and len(cells) > 1:
        res.update(status="需人工确认")
        res["notes"].append(why)
        return res
    if len(cells) > 1:          # 一张出库单应该只有一个报价，分开报价说明有问题
        parts = "；".join(f"{'+'.join(c[4])}={c[3]}" for c in cells.values())
        res.update(status="需人工确认")
        res["notes"].append(f"多个编号分开报价，一张出库单应只有一个报价：{parts}")
        return res
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
    if len(keys) > 1 and ok:
        notes = [f"多个编号同一个报价：{notes[0]}"]
    res.update(status="已找到" if ok else next(s for s in states if s != "已找到"),
               price=round(sum(prices), 2) if ok else None,
               basis="；".join(c[3] for c in cells.values() if c[3]),
               src=max(srcs, key=lambda r: r.get("d") or 0) if srcs else None)
    if len(keys) > 1 or not ok:
        res["notes"] += notes
    return res


def index_notices(notices):
    """出货通知的行：{出库单号: [(邮件, 行)]}，{编号 key / MR 号: [(邮件, 行)]}。"""
    by_obs, by_key = {}, {}
    for r in notices:
        for row in r["n"]:
            if row[2]:
                by_obs.setdefault(row[2], []).append((r, row))
            for c in row[0]:
                for k in {code_key(c), *MR.findall(c)}:
                    by_key.setdefault(k, []).append((r, row))
    return by_obs, by_key


def find_notice(o, nidx):
    """这张出库单最新的出货通知：{st, amt, raw, way, codes, r}；先按出库单号找，再按备注里的编号找。"""
    by_obs, by_key = nidx
    items = by_obs.get(o["no"].upper())
    if not items:
        keys = {code_key(c) for c in o["codes"]} | set(o["mrs"])
        items = [it for k in keys for it in by_key.get(k, [])]
    if not items:
        return None
    latest = max(r.get("d") or 0 for r, _ in items)
    rows = list({(id(r), row[4]): (r, row) for r, row in items if (r.get("d") or 0) == latest}.values())
    raws = list(dict.fromkeys(row[1] for _, row in rows if row[1]))
    st = ("unclear", None) if len(raws) > 1 else parse_price(raws[0]) if raws else None
    return {"st": st[0] if st else "none", "amt": st[1] if st else None, "raw": " / ".join(raws),
            "way": "、".join(dict.fromkeys(row[3] for _, row in rows if row[3])),
            "codes": list(dict.fromkeys(c for _, row in rows for c in row[0])), "r": rows[0][0]}


def classify(o, q, n, has_inq):
    """(卡派 / 快递 / 不明, 原因, 是否改发快递)。
    有我们的报价 > 出货通知的派送方式 > 有询价 > 领星物流渠道；
    例外：有卡派报价，但客户之后发的出货通知写快递 → 改发快递了，不对账（标色提醒）。"""
    if q["keys"]:
        nd = (n["r"].get("d") or 0) if n else 0
        if n and n["way"] and EXPRESS.search(n["way"]) and nd > q["qdate"]:
            return "快递", (f"有卡派报价（{day(q['qdate'])[:10]}），但客户之后的出货通知（{day(nd)[:10]}）"
                            f"写的是{n['way']}：改成快递发货了"), True
        return "卡派", "", False
    return (*_classify_rest(o, n, has_inq), False)


def _classify_rest(o, n, has_inq):
    if n and n["way"] and EXPRESS.search(n["way"]):
        return "快递", f"出货通知写的是{n['way']}"
    if n and (TRUCK.search(n["way"]) or n["st"] == "price"):
        return "卡派", ""
    if has_inq:
        return "卡派", ""
    if TRUCK.search(o.get("ship", "")):
        return "卡派", ""
    if EXPRESS.search(o.get("ship", "")):
        return "快递", "物流渠道是快递"
    return "不明", "没有报价、出货通知和询价，物流渠道也看不出来（多半是快递）"


def reconcile(orders, quotes, notices=(), inquiries=(), costs=None):
    """返回 (对账行 [(值列表, {标记})], 未对账行)。costs = dachser_cost.costs() 的结果，None = 没查成本。"""
    by_key, by_mr = build_index(quotes)
    nidx = index_notices(notices)
    inq = {code_key(c) for r in inquiries for c in r["iq"]} | {m for r in inquiries for c in r["iq"] for m in MR.findall(c)}
    rows, skipped, used = [], [], []
    for o in orders:
        q = eval_quote(o, by_key, by_mr, quotes)
        n = find_notice(o, nidx)
        has_inq = any(code_key(c) in inq for c in o["codes"]) or any(m in inq for m in o["mrs"])
        kind, why, switched = classify(o, q, n, has_inq)
        if kind != "卡派":
            skipped.append(([time_text(o["time"]), o["no"], o.get("ship", ""), kind, why, o["remark"][:200]],
                            {"switched"} if switched else set()))
            continue
        price, status, notes, flags = q["price"], q["status"], list(q["notes"]), set()
        cust = n["amt"] if n and n["st"] == "price" else None
        basis = q["basis"]
        if n and n["raw"]:
            basis = "；".join(x for x in (basis, NOTICE_MARK + n["raw"]) if x)
        if n and n["way"] and EXPRESS.search(n["way"]):
            notes.append(f"出货通知写的是{n['way']}，但有我们的报价")
        if status == "已找到":
            if n is None:
                notes.append("没找到客户的出货通知")
            elif cust is None:
                notes.append(f"出货通知的费用看不懂：{n['raw']}" if n["raw"] else "出货通知里没写费用")
            elif abs(cust - price) > 0.005:
                notes.append(f"我们的最新报价 {fmt(price)} 和客户确认价 {fmt(cust)} 不一致")
                price, status = None, "需人工确认"
        elif cust is not None and status in ("未找到报价邮件", "备注里没有编号", "最新报价待定"):
            notes.append({"未找到报价邮件": "没找到我们的报价邮件", "备注里没有编号": "备注里没有编号",
                          "最新报价待定": "我们最新的报价写的是待定"}[status] + "，价格取自客户的出货通知")
            price, status = cust, ONLY_NOTICE
        elif cust is not None:
            notes.append(f"客户确认价 {fmt(cust)}")
        if has_inq and not q["keys"]:
            notes.append("有客户的询价邮件，但没找到我们的报价（可能是同事回复的）")
            if status == "未找到报价邮件":
                status = "有询价没找到报价"
        if status == "备注里没有编号" and not n:
            notes.append(o["remark"][:200])
        if q["changed"]:
            flags.add("changed")
        cost, cost_note = (costs or {}).get(o["no"], (None, ""))
        if cost_note:
            notes.append("成本：" + cost_note)
        shown = " ".join(q["codes"]) or " ".join(o["mrs"] or (n["codes"] if n else []) or o["codes"][:2])
        src = q["src"]
        rows.append(([time_text(o["time"]), shown, o["no"], price, status, basis, cust, cost,
                      day(src.get("d")) if src else "", src["s"] if src else "",
                      day(n["r"].get("d")) if n else "", "；".join(x for x in notes if x)], flags))
        used.append(q["cells"])
    # 同一个报价格子对上了好几张卡派出库单：可能会重复收钱，都要人工确认
    by_cell = {}
    for i, cells in enumerate(used):
        for c in cells:
            by_cell.setdefault(c, []).append(i)
    for idxs in by_cell.values():
        if len(idxs) < 2:
            continue
        for i in idxs:
            v, flags = rows[i]
            others = "、".join(rows[j][0][2] for j in idxs if j != i)
            v[3], v[4] = None, "需人工确认"
            v[NOTE] = "；".join(x for x in (v[NOTE], f"和出库单 {others} 用的是同一个报价，可能重复收费") if x)
            flags.add("shared")
    return rows, skipped


def write_xlsx(rows, skipped, out):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "对账"
    ws.append(HEADERS)
    for r, _ in rows:
        ws.append(r)
    for c, w in zip("ABCDEFGHIJKL", (20, 14, 18, 11, 16, 30, 13, 13, 17, 40, 17, 50)):
        ws.column_dimensions[c].width = w
    for cell in ws[1]:
        cell.font = Font(bold=True)
    ai_fill = PatternFill("solid", fgColor="FFF2CC")       # AI 读的：整行浅黄，要重点核对
    bad_fill = PatternFill("solid", fgColor="F8CBAD")      # 没填价格的状态：浅红
    cust_fill = PatternFill("solid", fgColor="F4B084")     # 价格只来自客户的出货通知：橙色
    chg_fill = PatternFill("solid", fgColor="BDD7EE")      # 报价改过：报价格浅蓝
    shared_fill = PatternFill("solid", fgColor="D9C3E9")   # 同一个报价对上多张出库单：出库单号浅紫
    switch_fill = PatternFill("solid", fgColor="FFD966")   # 有卡派报价但改发快递："未对账"整行金黄
    for (_, flags), row in zip(rows, ws.iter_rows(min_row=2)):
        if str(row[5].value or "").find(AI_MARK) >= 0:
            for c in row:
                c.fill = ai_fill
        if row[4].value == ONLY_NOTICE:
            row[4].fill = row[3].fill = cust_fill
        elif row[4].value != "已找到":
            row[4].fill = bad_fill
        if "changed" in flags:
            row[3].fill = chg_fill
        if "shared" in flags:
            row[2].fill = shared_fill
    for r in ws.iter_rows(min_row=2, min_col=4, max_col=8):
        r[0].number_format = r[3].number_format = r[4].number_format = "0.00"
    ws.freeze_panes = "A2"
    ws2 = wb.create_sheet("未对账")
    ws2.append(SKIP_HEADERS)
    for r, flags in skipped:
        ws2.append(r)
        if "switched" in flags:
            for c in ws2[ws2.max_row]:
                c.fill = switch_fill
    for c, w in zip("ABCDEF", (20, 18, 24, 8, 70, 40)):
        ws2.column_dimensions[c].width = w
    for cell in ws2[1]:
        cell.font = Font(bold=True)
    ws2.freeze_panes = "A2"
    while True:
        try:
            wb.save(out)
            return True
        except PermissionError:
            if app_ui.ask(f"保存失败：{out.name} 可能正在 Excel 里打开。关掉后点\"重试\"。",
                          [("r", "重试"), ("q", "不保存")]) == "q":
                return False


def run(args):
    t0 = time.monotonic()
    path = Path(args.xlsx)
    orders = read_outbound(path)
    if not orders:
        raise SystemExit("Excel 里没有出库单")
    times = [t for t in (parse_time(o["time"]) for o in orders) if t]
    since = (min(times) if times else datetime.now() - timedelta(days=31)) - timedelta(days=args.days)
    print(f"出库单 {len(orders)} 张；搜索 {since:%Y-%m-%d} 之后的邮件和 Dachser 网站询价")

    # 成本价：Dachser 网站询价的 Price（先做，需要登录时一开始就登）
    costs = None
    if getattr(args, "no_cost", False) or (args.eml and not getattr(args, "cost_html", None)):
        print("成本价：这次不查")
    else:
        import dachser_cost
        try:
            print("成本价：打开 Dachser 读网站询价（只读，不点任何按钮）…")
            drows, earliest = dachser_cost.fetch_rows(since, getattr(args, "cost_html", None))
            costs = dachser_cost.costs(orders, drows, earliest)
            print(f"成本价：Dachser 网站询价 {len(drows)} 条，对上 {sum(v[0] is not None for v in costs.values())} 张出库单")
        except Exception as e:
            print(f"⚠ 读 Dachser 网站询价失败，成本价不填：{str(e).splitlines()[0] if str(e) else e!r}")
    t_cost = time.monotonic()

    if args.eml:
        mails, own = {str(p): mail_from_eml(p) for p in args.eml}, ""
    else:
        mails, own = fetch_feishu(since, getattr(args, "code", None))
    t_mail = time.monotonic()
    live = {k: v for k, v in mails.items() if not v.get("skip")}
    quotes = {k: v for k, v in live.items() if v.get("q") or v.get("ai")}
    notices = [v for v in live.values() if v.get("n")]
    # 报价 = 我们发给客户的：只认和自己同一个邮箱域名的发件人（客户的邮件里也可能带报价表，可能抄的是旧价）
    # 询价 = 客户发来的；出货通知不限发件人（客户发的，我们转发的也一样）
    if own:
        other = sum(1 for v in quotes.values() if mail_domain(v.get("f")) != own)
        quotes = {k: v for k, v in quotes.items() if mail_domain(v.get("f")) == own}
        inquiries = [v for v in live.values() if v.get("iq") and mail_domain(v.get("f")) != own]
        print(f"找到 {len(quotes)} 封我们发出的报价邮件（@{own}；另有 {other} 封外部邮件带报价表，不算）")
    else:
        inquiries = [v for v in live.values() if v.get("iq") and not (v.get("q") or v.get("ai"))]
        print(f"找到 {len(quotes)} 封有报价的邮件" + ("" if args.eml else "（没读到自己的发件地址，所有发件人都算）"))
    print(f"找到 {len(notices)} 封客户的出货通知、{len(inquiries)} 封询价邮件")
    by_key, by_mr = build_index(quotes)
    wanted = {k for o in orders for k in match_order(o, by_key, by_mr)[0]}
    if run_ai(quotes, wanted) and not args.eml:
        _save_cache(mails)
    t_ai = time.monotonic()

    rows, skipped = reconcile(orders, quotes, notices, inquiries, costs)
    out = (Path(args.o) if args.o else path.with_name(f"{path.stem}_对账结果.xlsx")).resolve()
    saved = write_xlsx(rows, skipped, out)
    count = {}
    for r, _ in rows:
        count[r[4]] = count.get(r[4], 0) + 1
    kinds = {}
    for r, flags in skipped:
        k = "有卡派报价但改发快递" if "switched" in flags else "渠道不明" if r[3] == "不明" else r[3]
        kinds[k] = kinds.get(k, 0) + 1
    print(f"卡派 {len(rows)} 张对账：" + "，".join(f"{k} {v}" for k, v in count.items()))
    if skipped:
        print("不对账：" + "，".join(f"{k} {v}" for k, v in kinds.items()) + "（见\"未对账\"表）")
    for r, flags in skipped:
        if "switched" in flags:
            print(f"  [改发快递] {r[1]} {r[4]}")
    for r, _ in rows:
        if r[4] != "已找到":
            print(f"  [{r[4]}] {r[2]} {r[1]} {r[NOTE]}")
    dur, now = app_ui.duration, time.monotonic()
    print(f"用时：查成本 {dur(t_cost - t0)}，读邮件 {dur(t_mail - t_cost)}，AI {dur(t_ai - t_mail)}，对账和生成 Excel {dur(now - t_ai)}，共 {dur(now - t0)}")
    print(f"\n对账结果已保存到：{out}" if saved else "\n对账结果没有保存")


def main():
    ap = argparse.ArgumentParser(description="按出库单在飞书邮件里找报价，生成对账 Excel")
    ap.add_argument("xlsx", help="领星导出的出库单 Excel")
    ap.add_argument("--days", type=int, default=60, help="从最早的出库时间再往前找多少天的邮件（默认 60）")
    ap.add_argument("-o", help="输出文件，默认 <Excel名>_对账结果.xlsx")
    ap.add_argument("--eml", nargs="+", help="离线：用导出的 .eml 邮件代替飞书")
    ap.add_argument("--no-cost", action="store_true", help="不打开 Dachser 查成本价")
    ap.add_argument("--cost-html", help="离线：用保存的 My enquiries 列表页代替 Dachser 网站查成本价")
    ap.add_argument("--code", help="不在这台电脑上授权时：飞书授权后跳转的地址（或其中的 code），几分钟内有效")
    run(ap.parse_args())


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    main()
