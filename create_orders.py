#!/usr/bin/env python3
"""
领星出库单 -> 在 Dachser『My enquiries』报价列表里按备注找到报价 -> 点击 create

用法:
    python create_orders.py 领星导出.xlsx                 # 逐单搜索并点击 create
    python create_orders.py 领星导出.xlsx --dry           # 只搜索、匹配、报告，不点 create
    python create_orders.py 领星导出.xlsx --days 60       # 搜索最近 60 天的报价（默认 30）
    python create_orders.py 领星导出.xlsx --html eLogistics.html   # 离线：用保存的列表页做匹配预览

匹配规则（以备注为准）:
    备注里的 MR 编号（多个 MR 时全部都要出现）包含在 Dachser 单号里即算匹配，
    例如备注 MR12345678 能匹配 Dachser 单号 MR12345678-a+b。备注里没有 MR 编号时用整段备注匹配。
    已经 created 的报价不算。
    - 唯一匹配: 自动点击 create；地址和领星不一致只打印提醒，不拦截
    - 多个匹配: 列出来，由你在浏览器里自己点 create
    - 没有匹配 / 都已 created: 跳过
点击 create 之后脚本停下，等你在浏览器里处理建单页面，回车后继续下一单。
结果写入 建单结果.csv。
"""
import argparse
import csv
import datetime as dt
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import app_ui  # noqa: E402
from app_ui import APP_DIR, ask  # noqa: E402
from match_orders import MR, check_address, read_lingxing  # noqa: E402
from dachser_fill import (START_URL, Filler, _launch, _quotation_link, norm_country, norm_zip,  # noqa: E402
                          open_quotation, split_street)

# 在页面里读出报价表的每一行（含 create 按钮状态）
JS_ROWS = r"""() => {
    const tbl = document.querySelector('table.specialTable');
    if (!tbl) return null;
    const out = [];
    for (const tr of tbl.querySelectorAll('tr')) {
        if (tr.classList.contains('head')) continue;
        const tds = tr.querySelectorAll(':scope > td');
        if (tds.length < 7) continue;
        const btn = tr.querySelector('input.createOrder');
        out.push({
            time: tds[1].innerText.replace(/\s+/g, ' ').trim(),
            order: tds[2].innerText.trim(),
            consignee: tds[3].innerText.trim(),
            weight: tds[4].innerText.trim(),
            price: tds[5].innerText.replace(/\s+/g, ' ').trim(),
            btn_name: btn ? btn.getAttribute('name') : '',
            btn_value: btn ? btn.value : '',
            btn_disabled: btn ? btn.disabled : true,
        });
    }
    return out;
}"""


def to_dz(r):
    """把页面行整理成 match_orders.check_address 需要的格式。"""
    lines = [x.strip() for x in r["consignee"].split("\n") if x.strip()]
    dest = lines[1] if len(lines) > 1 else ""
    m = re.match(r"(\S+)\s+(.*)", dest)
    code, rest = (m.group(1), m.group(2).strip()) if m else ("", dest)
    return {**r, "order": re.sub(r"\s+", "", r["order"]), "name": lines[0] if lines else "",
            "dest": dest, "code": code, "rest": rest,
            "can_create": bool(r["btn_name"]) and not r["btn_disabled"]}


def remark_keys(remark):
    """备注 -> 匹配用的关键字列表。"""
    mrs = list(dict.fromkeys(MR.findall(remark)))
    if mrs:
        return mrs
    s = re.sub(r"\s+", "", remark)
    return [s] if s else []


def matches(order, keys):
    """Dachser 单号里包含全部关键字（MR 编号后面不能紧跟数字，避免 MR1234567 误配 MR12345678x）。"""
    return all(re.search(re.escape(k) + r"(?!\d)", order, re.I) for k in keys)


def classify(lx, rows):
    """返回 (状态, 可点的候选行列表, 说明)。"""
    keys = remark_keys(lx["remark"])
    if not keys:
        return "备注为空", [], ""
    hits = [r for r in rows if matches(r["order"], keys)]
    open_ = [r for r in hits if r["can_create"]]
    if not hits:
        return "Dachser未找到", [], ""
    if not open_:
        return "已创建过", [], f"{len(hits)} 条匹配，都已 created 或没有 create 按钮"
    if len(open_) > 1:
        return "多个匹配", open_, f"{len(open_)} 条可 create，需手动选择"
    return "唯一匹配", open_, "；".join(check_address(lx, open_[0]))


def fmt_row(r):
    return f"{r['time']}  {r['order']}  {r['name']} / {r['dest']}  {r['weight']}  {r['price']}"


# ---------------------------------------------------------------- 浏览器操作
def wait_ajax(page, ms=400):
    page.wait_for_timeout(ms)
    try:
        page.wait_for_load_state("domcontentloaded", timeout=15000)
        page.wait_for_function(
            "() => { const o = document.querySelector('.overlay');"
            " return !o || getComputedStyle(o).display === 'none'; }", timeout=15000)
    except Exception:
        pass
    page.wait_for_timeout(200)


def open_history(page):
    """打开 freightquotation -> My enquiries 列表页。"""
    if page.locator("form.filter-form input[name='orderNumber']").count():
        return
    tab = page.locator(".tab-row a", has_text=re.compile(r"My enquiries", re.I)).first
    if not tab.count():   # 不在 freightquotation 里（比如在建单页面），先点左侧菜单
        _quotation_link(page).click(timeout=10000)
        wait_ajax(page)
        tab = page.locator(".tab-row a", has_text=re.compile(r"My enquiries", re.I)).first
    tab.click(timeout=10000)
    page.wait_for_selector("form.filter-form input[name='orderNumber']", timeout=20000)
    wait_ajax(page)


def _submit_search(page, order_no, days):
    """填条件、点 Search，并等到结果表真的被换掉（不管 Search 是 ajax 还是整页刷新）。"""
    open_history(page)
    wait_ajax(page, 200)
    page.add_style_tag(content="#ui-datepicker-div { display: none !important; }")
    today = dt.date.today()
    page.evaluate("""([f, t]) => {
        const a = document.querySelector("form.filter-form input[name='fromDate']");
        const b = document.querySelector("form.filter-form input[name='toDate']");
        if (a) a.value = f; if (b) b.value = t;
        // 给当前结果表打标记，搜索完成后旧表会被替换掉，标记随之消失
        document.querySelectorAll('table.specialTable').forEach(t => t.setAttribute('data-old-result', '1'));
    }""", [(today - dt.timedelta(days=days)).strftime("%m/%d/%Y"), today.strftime("%m/%d/%Y")])
    page.fill("form.filter-form input[name='orderNumber']", order_no)
    page.click("form.filter-form input[name='filterButton']")
    try:
        page.wait_for_function("() => !document.querySelector('table[data-old-result]')", timeout=30000)
    except Exception:
        print("  [WARN] 30 秒内没等到新的搜索结果")
    page.wait_for_load_state("domcontentloaded")
    wait_ajax(page, 500)


def search(page, order_no, days):
    """在列表页按单号搜索（日期从 days 天前到今天），返回整理后的行。"""
    for attempt in range(2):
        _submit_search(page, order_no, days)
        try:
            cur = page.input_value("form.filter-form input[name='orderNumber']", timeout=10000).strip()
        except Exception:
            cur = None
        if cur == order_no:
            break
        print(f"  [WARN] 搜索框里是 {cur!r}，不是 {order_no!r}，重新搜索")
    rows = page.evaluate(JS_ROWS) or []
    print(f"  搜索 {order_no or '(全部)'}: {len(rows)} 行")
    return [to_dz(r) for r in rows]


def find_rows(page, keys, days):
    """先用单号搜；搜不到（Dachser 可能只做精确搜索，组合单号会漏）就清空单号、取整个日期范围再本地筛。"""
    rows = search(page, keys[0], days)
    if any(matches(r["order"], keys) for r in rows):
        return rows
    return search(page, "", days)


F_EDIT_SAVE = "form[action*='frmOrderDetail'] input[name='btnSave']"


def on_edit_page(page):
    """是否还停在建单编辑页（Process order）。停在这里时 Dachser 会锁住这张单。"""
    try:
        return page.locator(F_EDIT_SAVE).count() > 0
    except Exception:
        return False


def cancel_edit(page):
    """点编辑页的 Cancel 正常离开，Dachser 才会解除锁定。返回是否已离开编辑页。"""
    link = page.locator("form[action*='frmOrderDetail'] a.fw_formButtonLeft",
                        has_text=re.compile(r"^\s*Cancel\s*$")).first
    accept = lambda d: d.accept()      # 万一弹出"确定放弃吗"之类的确认框，选确定（Playwright 默认会点取消）
    page.on("dialog", accept)
    try:
        link.click(timeout=10000)
        page.wait_for_function("(s) => !document.querySelector(s)", arg=F_EDIT_SAVE, timeout=20000)
        wait_ajax(page, 500)
        return True
    except Exception as e:
        print(f"  [WARN] 自动点 Cancel 失败: {str(e).splitlines()[0]}")
        return False
    finally:
        page.remove_listener("dialog", accept)


def leave_edit_page(page, lx_no):
    """离开建单页前确认：没 Save 就让用户选择去保存还是放弃（放弃时点 Cancel 解锁）。返回 True=已离开。"""
    while on_edit_page(page):
        ans = ask(f"{lx_no}：页面还停在建单编辑页，看起来还没点 Save。\n"
                  "直接离开或关闭浏览器会让 Dachser 锁住这张单（之后手动打开会提示 being processed by someone else）。",
                  [("c", "我去点 Save，点完再点这里"), ("x", "不保存，放弃这单（点 Cancel）")])
        if ans == "x":
            if cancel_edit(page):
                print(f"  {lx_no}：已点 Cancel 放弃，锁定已解除")
                return True
            ask("自动点 Cancel 没成功，请在浏览器里手动点页面左下角的 Cancel。", [("c", "我点好了")])
    return True


def click_create(page, row):
    btn = page.locator(f"input[name='{row['btn_name']}']")
    btn.scroll_into_view_if_needed(timeout=5000)
    btn.click(timeout=10000)
    wait_ajax(page, 1000)


# ---------------------------------------------------------------- 建单页面填写
F_NAME1 = "input[name$='tabpanel:tfiName1Consignee']"
F_NAME2 = "input[name$='tabpanel:tfiName2Consignee']"
F_NAME3 = "input[name$='tabpanel:tfiName3Consignee']"
F_CONTACT = "input[name$='tfiContactPerson:contextSelect:input']"
F_TEL = "input[name$='ctnContactInformation:tfiTelNumber']"
F_MOBILE = "input[name$='ctnContactInformation:tfiMobileNumber']"
F_ORDERNO = "input[name$='listOfOrderNums:0:tfiOrderNumField1']"
F_COUNTERS = "input[name$='panel:counter']"
F_STREET = "input[name$='tabpanel:tfiStreetConsignee']"
F_COUNTRY = "select[name$='selCountriesConsignee']"
F_ZIP = "input[name$='zipContainer:tfiZipConsignee']"
F_CITY = "input[name$='tabpanel:tfiCityConsignee']"
F_PHONE_BOOKING = "input[name$='selDeliveryOptions'][value='AT']"


def norm_phone(s):
    """'+353 87 123 4567' -> '+353871234567'（Dachser 要求 +49891234567 这种格式）。"""
    s = re.sub(r"[^\d+]", "", s or "")
    return "+" + s[2:] if s.startswith("00") else s


def set_field(page, sel, value, label, warns):
    """填一个输入框并触发 change；超过 maxlength 截断并提醒。返回是否成功。"""
    loc = page.locator(sel).first
    try:
        loc.wait_for(state="attached", timeout=5000)
        ml = loc.get_attribute("maxlength")
        if ml and ml.isdigit() and len(value) > int(ml):
            warns.append(f"{label} 超长({len(value)}>{ml})已截断: {value!r}")
            value = value[:int(ml)]
        loc.scroll_into_view_if_needed(timeout=3000)
        loc.fill(value, timeout=5000)
        loc.dispatch_event("change")
        loc.evaluate("e => e.blur()")
        wait_ajax(page, 200)
        return True
    except Exception as e:
        warns.append(f"{label} 填写失败: {str(e).splitlines()[0]}")
        return False


# 货物行危险品弹窗
DG_UN_NO = "3480"                    # LITHIUM ION BATTERIES
DG_PACKAGING = "Fibreboard box"
F_DG_MASS = "input[name$='tfiRTDangerMass']"
F_DG_QTY = "input[name$='tfiRTQuantity']"
F_DG_PACK = "select[name$='selBoxRTDangerPack']"
F_DG_SAVE = "input[name$='frmDlgAddDGoodsOrder:lnkSave']"


def _visible_button(page, text_re):
    return page.locator("a.fw_lightButton, a.close", has_text=re.compile(text_re)).locator("visible=true").first


def _dg_list_has_un(page):
    """第一个弹窗（该货物行的危险品列表）里是否已经有这个 UN 号。"""
    return page.evaluate("""(un) => [...document.querySelectorAll('fieldset#hazmatKeys table.results tr')]
        .some(tr => tr.offsetParent !== null && !tr.querySelector('input[name*="tblDGoodsList"]')
                    && [...tr.querySelectorAll('td')].some(td => td.innerText.trim() === un))""", DG_UN_NO)


def add_dg_line(page, i, qty, warns):
    """第 i 行货物：点 DG 图标 -> Add dangerous goods -> 勾 UN3480 -> Add -> 填重量/件数/包装 -> Save -> Back。"""
    label = f"第 {i + 1} 行危险品"
    weight = page.locator(f"input[name$='listView:{i}:nested:4:panel:weight']").first.input_value().strip()
    if not weight:
        warns.append(f"{label}: 这一行没有重量，未添加")
        return
    if not qty:
        warns.append(f"{label}: 领星没有 Total Qty of SKU，未添加")
        return
    try:
        row = page.locator(f"tr:has(input[name$='listView:{i}:nested:1:panel:counter'])").first
        row.locator("a.dangerous-goods-link").click(timeout=10000)
        add_btn = _visible_button(page, r"^\s*Add dangerous goods\s*$")
        add_btn.wait_for(timeout=15000)
        wait_ajax(page, 300)
        if _dg_list_has_un(page):
            warns.append(f"{label}: 已经有 UN{DG_UN_NO}，没有重复添加")
        else:
            add_btn.click(timeout=10000)
            # 第二个弹窗：在 UN 号 = 3480 的那一行勾选
            name = None
            for _ in range(30):
                page.wait_for_timeout(500)
                name = page.evaluate("""(un) => { for (const tr of document.querySelectorAll('tr')) {
                        const cb = tr.querySelector('input[type=checkbox][name*="tblDGoodsList"]');
                        if (cb && cb.offsetParent !== null &&
                            [...tr.querySelectorAll('td')].some(td => td.innerText.trim() === un)) return cb.name; }
                        return null; }""", DG_UN_NO)
                if name:
                    break
            if not name:
                raise RuntimeError(f"危险品列表里找不到 UN{DG_UN_NO}")
            cb = page.locator(f"input[name='{name}']")
            cb.click(timeout=5000)        # 行本身也有点击事件，点完再确认一次是不是勾上了
            page.wait_for_timeout(200)
            if not cb.is_checked():
                cb.evaluate("e => { e.checked = true; e.dispatchEvent(new Event('change', {bubbles: true})); }")
            _visible_button(page, r"^\s*Add\s*$").click(timeout=10000)
            page.locator(F_DG_MASS).first.wait_for(state="visible", timeout=15000)
            wait_ajax(page, 300)
            set_field(page, F_DG_MASS, weight, f"{label} Dangerous goods mass", warns)
            set_field(page, F_DG_QTY, qty, f"{label} No. of packages", warns)
            Filler(page).pick_by_text(page.locator(F_DG_PACK).first, DG_PACKAGING)
            page.locator(F_DG_SAVE).first.click(timeout=10000)
            try:     # Save 成功后第二个弹窗自动关闭
                page.locator(F_DG_SAVE).first.wait_for(state="detached", timeout=15000)
            except Exception:
                fb = page.evaluate("() => [...document.querySelectorAll('.feedbackPanel, .feedbackPanelERROR, "
                                   "li.feedbackPanelERROR')].map(e => e.innerText.trim()).filter(Boolean).join(' | ')")
                warns.append(f"{label}: 点 Save 后弹窗没关闭{('：' + fb) if fb else ''}，已取消")
                _visible_button(page, r"^\s*Cancel\s*$").click(timeout=5000)
            wait_ajax(page, 500)
            if not _dg_list_has_un(page):
                warns.append(f"{label}: 列表里没看到 UN{DG_UN_NO}，请手动检查")
            else:
                print(f"    {label}: 已添加 UN{DG_UN_NO}，{weight} kg，{qty} 件，{DG_PACKAGING}")
        _visible_button(page, r"^\s*Back\s*$").click(timeout=10000)
        add_btn.wait_for(state="hidden", timeout=15000)
        wait_ajax(page, 300)
    except Exception as e:
        warns.append(f"{label} 添加失败: {str(e).splitlines()[0]}（请手动处理，必要时关掉弹窗）")


F_DG_UPLOAD = ".DGNButton input[type='file']"     # "Dangerous goods documents" 上传按钮背后的文件框
DG_DIR = APP_DIR / "危险品文件"


def find_dg_files(skus, dg_dir):
    """危险品文件名（不含扩展名）和 SKU 完全一致才算。返回 (找到的文件, 没找到的 SKU)。"""
    d = Path(dg_dir)
    files = [f for f in d.iterdir() if f.is_file()] if d.is_dir() else []
    found, missing = [], []
    for sku in dict.fromkeys(skus):
        hit = [f for f in files if f.stem == sku]
        found += hit
        if not hit:
            missing.append(sku)
    return found, missing


def upload_dg(page, lx, dg_dir, warns):
    """点 Dangerous goods documents 上传和 SKU 同名的危险品文件；找不到就不传。"""
    if not lx.get("skus"):
        warns.append("领星没有 SKU，未上传危险品文件")
        return
    if not Path(dg_dir).is_dir():
        warns.append(f"危险品文件夹不存在: {dg_dir}，未上传")
        return
    files, missing = find_dg_files(lx["skus"], dg_dir)
    if missing:
        warns.append(f"没找到危险品文件: {', '.join(missing)}（未上传）")
    for f in files:
        try:
            page.locator(F_DG_UPLOAD).first.set_input_files(str(f), timeout=10000)
            wait_ajax(page, 1000)
            try:     # 上传成功后文件名会出现在 Documents 区域
                page.wait_for_function(
                    "(n) => [...document.querySelectorAll('#upload, .uploaded-files-edit')]"
                    ".some(e => e.innerText.includes(n))", arg=f.name, timeout=30000)
                print(f"    已上传危险品文件: {f.name}")
            except Exception:
                warns.append(f"已选择上传 {f.name}，但 30 秒内页面上没看到这个文件，请确认")
        except Exception as e:
            warns.append(f"上传 {f.name} 失败: {str(e).splitlines()[0]}")


def fill_order(page, lx, dg_dir=DG_DIR):
    """在点击 create 后出现的 Transport Order 页面里填收件人、联系方式、单号、货物内容，上传危险品文件。不点 Save。"""
    warns = []
    try:
        page.wait_for_selector(F_NAME1, state="attached", timeout=30000)
    except Exception:
        return ["没等到建单页面（找不到 Name 1 输入框），未填写"]
    page.add_style_tag(content="#ui-datepicker-div, #ui-timepicker-div { display: none !important; }")
    # 地址区默认就在 Consignee 标签；万一停在 Shipper，先切过去
    if not page.locator(F_NAME1).first.is_visible():
        page.locator("#editAddr .tablinks a", has_text="Consignee").first.click(timeout=5000)
        wait_ajax(page)

    # 1) 先选 Phone delivery booking，Contact person 等字段才会出现
    #    （联系人表格上的 inactive 样式选完后也可能一直在，只影响外观，所以不看它；真填不进去 set_field 会提醒）
    try:
        page.locator(F_PHONE_BOOKING).first.click(timeout=5000)
        wait_ajax(page, 500)
    except Exception as e:
        warns.append(f"选 Phone delivery booking 失败: {str(e).splitlines()[0]}")
    try:
        page.wait_for_function(
            "(s) => { const e = document.querySelector(s); return !!e && !e.disabled && e.offsetParent !== null; }",
            arg=F_CONTACT, timeout=10000)
    except Exception:
        pass

    # 2) 收货地址以领星为准（先国家，国家变化会刷新邮编/城市区域）；County 不动，保持页面原值
    country = None
    try:
        country = norm_country(lx["country"])
        sel = page.locator(F_COUNTRY).first
        cur_c = sel.evaluate("e => ((e.options[e.selectedIndex] || {}).text || '').trim()")
        if cur_c.lower() != country.lower():
            Filler(page).pick_by_text(sel, country)
            warns.append(f"Country 已从 {cur_c} 改为 {country}")
    except Exception as e:
        warns.append(f"Country 设置失败: {str(e).splitlines()[0]}")
    zip_n = norm_zip(lx["zip"], country) if country else lx["zip"]
    street, num = split_street(lx["addr"], zip_n, lx["city"])
    num = lx["hno"] or num
    street_text = f"{street} {num}".strip() if not num or num not in street else street
    if lx["addr2"]:
        street_text = f"{street_text}, {lx['addr2']}"
    for sel, val, label in [(F_STREET, street_text, "Street"),
                            (F_ZIP, zip_n, "PCD"),
                            (F_CITY, lx["city"], "City")]:
        if not val:
            warns.append(f"领星没有 {label}，已清空")
        set_field(page, sel, val, label, warns)

    # 3) 收件人 / 联系方式
    phone = norm_phone(lx["tel"])
    name1 = lx["company"] or lx["recipient"]
    for sel, val, label in [(F_NAME1, name1, "Name 1"), (F_NAME2, lx["email"], "Name 2(邮箱)"),
                            (F_NAME3, phone, "Name 3(电话)"), (F_CONTACT, lx["recipient"], "Contact person"),
                            (F_TEL, phone, "Telephone"), (F_MOBILE, phone, "Mobile")]:
        if not val:
            warns.append(f"领星没有 {label} 对应的数据，留空")
        set_field(page, sel, val, label, warns)

    # 4) Order number 后面加 +出库单号；超长时截掉原单号前面的部分，出库单号保持完整
    try:
        loc = page.locator(F_ORDERNO).first
        cur = loc.input_value(timeout=5000).strip()
        ml = int(loc.get_attribute("maxlength") or 35)
        no = lx["no"]
        if no in cur:
            warns.append(f"Order number 里已有出库单号: {cur!r}")
        else:
            keep = ml - len(no) - 1
            if not cur or keep <= 0:
                new = no
            elif len(cur) > keep:
                new = f"{cur[-keep:].lstrip('+-/ ')}+{no}"
                warns.append(f"Order number 超长，原单号 {cur!r} 只保留后 {keep} 位")
            else:
                new = f"{cur}+{no}"
            set_field(page, F_ORDERNO, new, "Order number", warns)
    except Exception as e:
        warns.append(f"Order number 填写失败: {str(e).splitlines()[0]}")

    # 只处理填了数量的货物行：写 Content，再加危险品
    counters = page.locator(F_COUNTERS)
    lines = [i for i in range(counters.count()) if counters.nth(i).input_value().strip()]
    for i in lines:
        set_field(page, f"input[name$='listView:{i}:nested:3:panel:content:contextSelect:input']",
                  "battery", f"第 {i + 1} 行 Content", warns)
    if len(lines) > 1:
        warns.append(f"有 {len(lines)} 行货物，每行危险品 No. of packages 都填了总数量 {lx['total_qty']}，请按实际改")
    for i in lines:
        add_dg_line(page, i, lx["total_qty"], warns)

    # 5) 危险品文件（最后传：上传是带整张表单提交的 ajax，前面填的内容会一起带上）
    upload_dg(page, lx, dg_dir, warns)
    return warns


# ---------------------------------------------------------------- 主流程
def print_result(lx, status, cands, note):
    print(f"  结果: {status}" + (f"  ({note})" if note else ""))
    for i, r in enumerate(cands, 1):
        iss = check_address(lx, r)
        print(f"    [{i}] {fmt_row(r)}" + (f"\n        ⚠ {'；'.join(iss)}" if iss and len(cands) > 1 else ""))


def run_offline(lx_rows, html_path):
    from playwright.sync_api import sync_playwright
    html = Path(html_path).read_text(encoding="utf-8", errors="ignore")
    with sync_playwright() as pw:
        b = pw.chromium.launch()
        page = b.new_page()
        page.set_content(html, wait_until="domcontentloaded")
        page.route("**/*", lambda r: r.abort())
        rows = [to_dz(r) for r in (page.evaluate(JS_ROWS) or [])]
        b.close()
    print(f"列表页共 {len(rows)} 行，可 create {sum(r['can_create'] for r in rows)} 行")
    res = []
    for n, lx in enumerate(lx_rows, 1):
        print(f"\n[{n}/{len(lx_rows)}] {lx['no']}  备注 {lx['remark']!r}")
        status, cands, note = classify(lx, rows)
        print_result(lx, status, cands, note)
        res.append((lx, status, cands, note, "离线预览，未点击"))
    return res


def run_live(lx_rows, args):
    from playwright.sync_api import sync_playwright
    res = []
    with sync_playwright() as pw:
        ctx = _launch(pw, args)
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        open_quotation(page, args.url)
        for n, lx in enumerate(lx_rows, 1):
            if app_ui.stop_requested():
                print("\n已停止。")
                break
            print(f"\n===== [{n}/{len(lx_rows)}] {lx['no']}  备注 {lx['remark']!r} =====")
            keys = remark_keys(lx["remark"])
            if not keys:
                print("  结果: 备注为空，跳过")
                res.append((lx, "备注为空", [], "", "跳过"))
                continue
            try:
                rows = find_rows(page, keys, args.days)
            except Exception as e:
                msg = str(e).splitlines()[0]
                print(f"  [ERR] 搜索失败: {msg}")
                res.append((lx, "搜索失败", [], msg, "跳过"))
                if ask(f"{lx['no']} 搜索失败：{msg}", [("c", "跳过这单继续"), ("q", "结束")]) == "q":
                    break
                continue
            status, cands, note = classify(lx, rows)
            print_result(lx, status, cands, note)
            action = "跳过"
            if status == "唯一匹配":
                if args.dry:
                    action = "dry 模式，未点击"
                else:
                    if note:
                        print(f"  ⚠ 地址和领星不一致：{note}（照样点击 create，请在建单页面核对）")
                    try:
                        click_create(page, cands[0])
                        action = "已点击 create"
                        print("  已点击 create")
                    except Exception as e:
                        action = f"点击失败: {str(e).splitlines()[0]}"
                        print(f"  [ERR] {action}")
            elif status == "多个匹配" and not args.dry:
                print("  请在浏览器里自己选一条点击 create（上面列出的是候选）。")
                ans = ask(f"{lx['no']}（备注 {lx['remark'].strip()}）找到 {len(cands)} 条可 create 的报价：\n"
                          + "\n".join(f"  [{i}] {fmt_row(r)}" for i, r in enumerate(cands, 1))
                          + "\n\n请在浏览器里自己选一条点击 create。",
                          [("c", "我点好了"), ("s", "这单不建"), ("q", "结束")])
                if ans == "q":
                    res.append((lx, status, cands, note, "未处理（结束）"))
                    break
                action = "跳过（手动）" if ans == "s" else "已手动点击 create"
            elif status == "多个匹配":
                action = "dry 模式，需手动选择"
            if action.startswith("已"):
                fw = fill_order(page, lx, args.dg_dir)
                print("  已填写建单页面" + ("，以下需注意:" if fw else "，全部 [OK]"))
                for w in fw:
                    print(f"    ⚠ {w}")
                action += " + 已填写" + (f"（{len(fw)} 条提醒）" if fw else "")
                note = "；".join(filter(None, [note] + fw))
            res.append((lx, status, cands, note, action))
            if action.startswith("已"):
                tips = "\n".join(f"  ⚠ {w}" for w in fw)
                ans = ask(f"{lx['no']} 已填写完成。\n请在浏览器里检查，确认无误后点 Save，再点下面的按钮。"
                          + (f"\n\n注意：\n{tips}" if tips else ""),
                          [("c", "已保存，继续下一单" if n < len(lx_rows) else "已保存，完成"), ("q", "结束")])
                leave_edit_page(page, lx["no"])      # 没 Save 就离开会锁单
                if ans == "q":
                    break
        leave_edit_page(page, "当前这单")            # 关浏览器前再确认一次
        if not args.yes:
            ask("处理完毕。", [("c", "关闭浏览器")])
        ctx.close()
    return res


def write_report(res, path):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["领星出库单号", "备注", "匹配状态", "操作", "说明", "领星地址",
                    "Dachser单号", "Dachser收货地", "Dachser重量", "Dachser价格", "候选数"])
        for lx, status, cands, note, action in res:
            d = cands[0] if len(cands) == 1 else {}
            w.writerow([lx["no"], lx["remark"].strip(), status, action, note,
                        f'{lx["country"]} {lx["zip"]} {lx["city"]} {lx["addr"]}',
                        d.get("order", ""), d.get("dest", ""), d.get("weight", ""), d.get("price", ""),
                        len(cands)])


def default_report_path(xlsx):
    """结果表放在所选 Excel 旁边：领星导出_1023.xlsx -> 领星导出_1023_建单结果.csv"""
    p = Path(xlsx).resolve()
    return str(p.with_name(f"{p.stem}_建单结果.csv"))


def run(args):
    lx_rows = read_lingxing(args.lingxing)
    print(f"领星 {len(lx_rows)} 单")
    if not lx_rows:
        return
    res = run_offline(lx_rows, args.html) if args.html else run_live(lx_rows, args)
    out = args.o or default_report_path(args.lingxing)
    write_report(res, out)

    print("\n===== 汇总 =====")
    for lx, status, cands, note, action in res:
        print(f"  {lx['no']:<18} {lx['remark'].strip():<16} {status:<8} {action}" + (f"  | {note}" if note else ""))
    print("结果已写入", out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("lingxing", help="领星导出的出库单 Excel")
    ap.add_argument("--dry", action="store_true", help="只搜索匹配，不点击 create")
    ap.add_argument("--html", help="离线模式：用保存的 My enquiries 列表页 HTML 做匹配预览")
    ap.add_argument("--days", type=int, default=30, help="搜索最近多少天的报价（默认 30）")
    ap.add_argument("--dg-dir", default=str(DG_DIR), help="危险品文件所在文件夹，文件名(不含扩展名)=SKU（默认 ./危险品文件）")
    ap.add_argument("-o", default=None, help="结果 CSV（默认放在 Excel 旁边：<Excel名>_建单结果.csv）")
    ap.add_argument("--url", default=START_URL)
    ap.add_argument("--profile", default="dachser_profile")
    ap.add_argument("--chromium-path", default=None, help="指定浏览器可执行文件（测试用）")
    ap.add_argument("--yes", action="store_true", help="结束时不等回车直接关浏览器")
    return run(ap.parse_args())


if __name__ == "__main__":
    sys.exit(main())
