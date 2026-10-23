"""第一步：读领星导出表 -> 取备注里的 MR 编号 -> 在 Dachser 报价列表里找对应行，并核对地址。

用法:
  直接运行 -> 默认读同文件夹里的 lingxing.xlsx 和 list.html
  读网页: python match_orders.py lingxing.xlsx live   (需先 pip install playwright; playwright install chromium)
  或: python match_orders.py 领星导出.xlsx Dachser列表.html [-o 结果.csv]

Dachser列表.html: 在 Dachser 报价历史页里 F12 -> Elements -> 右键 <html> -> Copy outerHTML，存成 .html
(之后接上 Playwright，这个文件就换成脚本直接读取的实时页面，匹配逻辑不变)
"""
import argparse, csv, re, sys, unicodedata
from html.parser import HTMLParser
import openpyxl

# 领星国家名 -> Dachser 国家代码
COUNTRY = {"austria":"A","belgium":"B","bulgaria":"BG","czech republic":"CZ","czechia":"CZ","denmark":"DK",
 "estonia":"EST","finland":"FIN","france":"F","germany":"D","great britain":"GB","united kingdom":"GB",
 "greece":"GR","hungary":"H","ireland":"IRL","italy":"I","latvia":"LV","lithuania":"LT","luxembourg":"L",
 "netherlands":"NL","norway":"N","poland":"PL","portugal":"P","romania":"RO","slovakia":"SK",
 "slovenia":"SLO","spain":"E","sweden":"S","switzerland":"CH"}
MR = re.compile(r"MR\d{8}")

def norm(s):
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]", "", s.lower())

# ---------- 领星 ----------
def read_lingxing(path):
    ws = openpyxl.load_workbook(path, data_only=True).worksheets[0]
    it = ws.iter_rows(values_only=True)
    head = [str(h or "").strip() for h in next(it)]
    def col(prefix):
        return next(i for i, h in enumerate(head) if h.startswith(prefix))
    c = dict(no=col("Outbound Order No"), rm=col("Remark"), rcp=col("Recipient/"), co=col("Company"),
             ctry=col("Country/Region"), city=col("City"), zip=col("Post code"), a1=col("Address1"),
             tel=col("Telephone"), email=col("Email/"), hno=col("House No"), a2=col("Address2"),
             prov=col("Province Name"), qty=col("Total Qty of SKU"))
    sku_cols = [i for i, h in enumerate(head) if re.fullmatch(r"SKU\s*\d+\s+SKU", h)]   # "SKU 1\nSKU", "SKU 2\nSKU"...
    out = []
    for r in it:
        if not r[c["no"]]:
            continue
        g = lambda k: str(r[c[k]] if r[c[k]] is not None else "").strip()
        out.append(dict(no=g("no"), remark=g("rm"), recipient=g("rcp"), company=g("co"),
                        country=g("ctry"), city=g("city"), zip=g("zip"), addr=g("a1"),
                        tel=g("tel"), email=g("email"), hno=g("hno"), addr2=g("a2"), province=g("prov"),
                        skus=[str(r[i]).strip() for i in sku_cols if r[i] not in (None, "")],
                        total_qty=re.sub(r"\.0+$", "", g("qty"))))
    return out

# ---------- Dachser 列表页 ----------
class ListParser(HTMLParser):
    """抓 table.specialTable 里每一行的: 单号、收货人(两行)、重量、价格"""
    def __init__(self):
        super().__init__(); self.rows=[]; self.cur=None; self.td=None; self.in_tbl=False
    def handle_starttag(self, tag, a):
        a = dict(a); cls = a.get("class") or ""
        if tag == "table" and "specialTable" in cls: self.in_tbl = True
        if not self.in_tbl: return
        if tag == "tr": self.cur = []
        if tag == "td" and self.cur is not None: self.td = ""
        if tag == "br" and self.td is not None: self.td += "\n"
    def handle_endtag(self, tag):
        if not self.in_tbl: return
        if tag == "td" and self.td is not None and self.cur is not None:
            self.cur.append(self.td.strip()); self.td = None
        if tag == "tr" and self.cur is not None:
            if len(self.cur) >= 6: self.rows.append(self.cur)
            self.cur = None
        if tag == "table": self.in_tbl = False
    def handle_data(self, d):
        if self.td is not None: self.td += d

def read_dachser(path):
    if path == "live":   # 直接读网页
        from fetch_dachser import fetch_list_html
        html = fetch_list_html()
    else:
        html = open(path, encoding="utf-8", errors="ignore").read()
    p = ListParser(); p.feed(html)
    out = []
    for r in p.rows:  # [勾选, 时间, 单号, 收货人, 重量, 价格, create...]
        order = re.sub(r"\s+", "", r[2])
        lines = [x.strip() for x in r[3].split("\n") if x.strip()]
        dest = lines[1] if len(lines) > 1 else ""
        m = re.match(r"(\S+)\s+(.*)", dest)
        code, rest = (m.group(1), m.group(2)) if m else ("", dest)
        out.append(dict(order=order, time=re.sub(r"\s+", " ", r[1]), consignee=lines[0] if lines else "",
                        dest=dest, code=code, rest=rest, weight=r[4], price=re.sub(r"\s+", " ", r[5])))
    return out

# ---------- 匹配 ----------
def check_address(lx, dz):
    """返回问题列表(空=一致)。国家/邮编硬比对；城市只提示。"""
    issues = []
    want = COUNTRY.get(lx["country"].lower())
    if want is None: issues.append(f"国家无法对应: {lx['country']}")
    elif want != dz["code"]: issues.append(f"国家不符: 领星{lx['country']}({want}) / Dachser {dz['code']}")
    z = norm(lx["zip"]).lstrip("0")
    if z and z not in norm(dz["rest"]).lstrip("0"): issues.append(f"邮编不符: 领星{lx['zip']} / Dachser {dz['rest']}")
    elif not z and dz["code"] not in ("IRL",): issues.append("领星无邮编")
    c = norm(lx["city"])
    if c and c not in norm(dz["rest"]) and norm(dz["rest"]) not in c: issues.append(f"城市可能不同: 领星{lx['city']} / Dachser {dz['rest']}")
    return issues

def match(lx_rows, dz_rows):
    by = {}
    for d in dz_rows: by.setdefault(d["order"], []).append(d)
    res = []
    for lx in lx_rows:
        m = MR.findall(lx["remark"])
        base = dict(lx=lx)
        if len(m) != 1:
            res.append({**base, "status": "备注无/多个MR编号", "note": repr(lx["remark"]), "dz": None}); continue
        mr = m[0]; hits = by.get(mr, [])   # 精确相等
        if not hits: res.append({**base, "mr": mr, "status": "Dachser未找到", "note": "跳过", "dz": None})
        elif len(hits) > 1: res.append({**base, "mr": mr, "status": "Dachser重复", "note": f"{len(hits)}条，需手动处理", "dz": None})
        else:
            iss = check_address(lx, hits[0])
            res.append({**base, "mr": mr, "dz": hits[0], "status": "可创建" if not iss else "地址有疑问", "note": "；".join(iss)})
    return res

def main():
    ap = argparse.ArgumentParser(); ap.add_argument("lingxing", nargs="?", default="lingxing.xlsx"); ap.add_argument("dachser", nargs="?", default="list.html"); ap.add_argument("-o", default="匹配结果.csv")
    a = ap.parse_args()
    lx, dz = read_lingxing(a.lingxing), read_dachser(a.dachser)
    print(f"领星 {len(lx)} 单，Dachser 列表 {len(dz)} 行")
    res = match(lx, dz)
    with open(a.o, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f); w.writerow(["领星出库单号","MR编号","状态","说明","领星收件人","领星公司","领星地址","Dachser收货地","Dachser重量","Dachser价格"])
        for r in res:
            l, d = r["lx"], r["dz"] or {}
            w.writerow([l["no"], r.get("mr",""), r["status"], r["note"], l["recipient"], l["company"],
                        f'{l["country"]} {l["zip"]} {l["city"]} {l["addr"]}', d.get("dest",""), d.get("weight",""), d.get("price","")])
            print(f'[{r["status"]}] {l["no"]} {r.get("mr","")} {r["note"]}')
    print("已写入", a.o)

if __name__ == "__main__":
    main()