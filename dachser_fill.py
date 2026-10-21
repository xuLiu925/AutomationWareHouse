#!/usr/bin/env python3
"""
Dachser freightquotation 询价表单自动填写（第一步：只填写，不提交）

用法（Windows，命令行）:
    pip install playwright openpyxl
    python dachser_fill.py 询价表.xlsx            # 填第 1 条数据（Excel 第 2 行）
    python dachser_fill.py 询价表.xlsx --row 3    # 填 Excel 第 3 行
    python dachser_fill.py 询价表.xlsx --dry      # 只解析 Excel，不打开浏览器

首次运行会打开一个独立的 Chrome 窗口：请手动登录 Dachser，
进入 freightquotation -> Enquiry 页面后回到命令行按回车。
登录状态保存在 ./dachser_profile 目录里，下次不用重新登录。
脚本不会点击页面底部的 "Enquiry" 提交按钮，填完后由你检查并手动提交。
"""
import argparse
import re
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
from pallet_plan import plan_pallets  # noqa: E402

START_URL = "https://elogistics.dachser.com/"

# ---------------------------------------------------------------- 国家映射
# Dachser 下拉框里的国家名（英文）。key 为各种写法（小写），value 为下拉框文字。
_COUNTRIES = {
    "Germany": ["de", "deu", "germany", "deutschland", "德国"],
    "Belgium": ["be", "bel", "belgium", "belgie", "belgique", "比利时"],
    "Italy": ["it", "ita", "italy", "italia", "意大利"],
    "France": ["fr", "fra", "france", "法国"],
    "Austria": ["at", "aut", "austria", "osterreich", "österreich", "奥地利"],
    "Bulgaria": ["bg", "bulgaria", "保加利亚"],
    "Croatia": ["hr", "croatia", "hrvatska", "克罗地亚"],
    "Cyprus": ["cy", "cyprus", "塞浦路斯"],
    "Czech Republic": ["cz", "czech republic", "czechia", "česko", "cesko", "捷克"],
    "Denmark": ["dk", "denmark", "danmark", "丹麦"],
    "Estonia": ["ee", "est", "estonia", "eesti", "爱沙尼亚"],
    "Finland": ["fi", "fin", "finland", "suomi", "芬兰"],
    "Great Britain": ["gb", "uk", "great britain", "united kingdom", "england",
                      "scotland", "wales", "britain", "英国"],
    "Greece": ["gr", "greece", "希腊"],
    "Hungary": ["hu", "hungary", "magyarorszag", "匈牙利"],
    "Ireland": ["ie", "irl", "ireland", "爱尔兰"],
    "Latvia": ["lv", "latvia", "拉脱维亚"],
    "Lithuania": ["lt", "lithuania", "立陶宛"],
    "Luxembourg": ["lu", "lux", "luxembourg", "卢森堡"],
    "Malta": ["mt", "malta", "马耳他"],
    "Netherlands": ["nl", "nld", "netherlands", "the netherlands", "holland",
                    "nederland", "荷兰"],
    "Poland": ["pl", "poland", "polska", "波兰"],
    "Portugal": ["pt", "portugal", "葡萄牙"],
    "Romania": ["ro", "romania", "罗马尼亚"],
    "Slovakia": ["sk", "slovakia", "slovensko", "斯洛伐克"],
    "Slovenia": ["si", "slovenia", "slovenija", "斯洛文尼亚"],
    "Spain": ["es", "spain", "espana", "españa", "西班牙"],
    "Sweden": ["se", "sweden", "sverige", "瑞典"],
}
ALIAS = {a: name for name, al in _COUNTRIES.items() for a in al}

# 邮编只允许数字的国家 -> 位数。Excel 会吃掉前导 0，所以要补齐。
NUMERIC_ZIP_LEN = {
    "Germany": 5, "France": 5, "Italy": 5, "Spain": 5, "Austria": 4,
    "Belgium": 4, "Denmark": 4, "Finland": 5, "Hungary": 4, "Luxembourg": 4,
    "Netherlands": 4,  # 荷兰真实邮编是 4 位数字+2 字母，Dachser 不接受字母时只填数字
    "Poland": 5, "Portugal": 4, "Czech Republic": 5, "Slovakia": 5,
    "Sweden": 5, "Slovenia": 4, "Croatia": 5, "Bulgaria": 4, "Cyprus": 4,
    "Estonia": 5, "Latvia": 4, "Lithuania": 5, "Romania": 6, "Greece": 5,
}


def norm_country(raw):
    key = str(raw or "").strip().lower()
    key = re.sub(r"\s+", " ", key)
    if key in ALIAS:
        return ALIAS[key]
    # 容错：包含关系（如 "France (FR)"）
    for a, name in ALIAS.items():
        if len(a) > 3 and a in key:
            return name
    raise ValueError(f"无法识别国家: {raw!r}")


def norm_zip(raw, country):
    s = str(raw if raw is not None else "").strip()
    if isinstance(raw, float) and raw.is_integer():
        s = str(int(raw))
    if country in NUMERIC_ZIP_LEN:
        need = NUMERIC_ZIP_LEN[country]
        digits = re.sub(r"\D", "", s)
        if len(digits) > need:
            digits = digits[:need]
        return digits.zfill(need) if digits else ""
    # 英国/爱尔兰/马耳他等：保留字母数字，去掉 "GB-" 之类前缀
    s = re.sub(r"^[A-Za-z]{1,3}\s*[-–]\s*(?=\w)", "", s) if re.match(
        r"^[A-Za-z]{2}\s*-", s) else s
    return re.sub(r"\s+", " ", s).strip().upper()


# ---------------------------------------------------------------- 地址拆分
def split_street(address, zip_raw, city):
    """'51 allée de courbiat 87350, Panazol' -> ('allée de courbiat', '51')"""
    a = str(address or "").strip()
    a = re.sub(r"\s+", " ", a)
    if city:
        a = re.sub(r",?\s*" + re.escape(str(city)) + r"\s*$", "", a, flags=re.I)
    z = str(zip_raw or "").strip()
    cands = {z, re.sub(r"^[A-Za-z]{1,3}\s*-\s*", "", z), re.sub(r"\D", "", z)}
    for c in sorted((c for c in cands if c), key=len, reverse=True):
        a = re.sub(r",?\s*(?:[A-Za-z]{1,3}\s*-\s*)?" + re.escape(c) + r"\b", "", a)
    a = a.strip(" ,;")
    # 门牌号：开头（法/英/荷风格）或结尾（德风格）
    m = re.match(r"^(\d+\s?[A-Za-z]?(?:\s?[-/]\s?\d+[A-Za-z]?)?)[,\s]+(.+)$", a)
    if m:
        return m.group(2).strip(" ,"), m.group(1).replace(" ", "")
    m = re.match(r"^(.+?)[,\s]+(\d+\s?[A-Za-z]?(?:\s?[-/]\s?\d+[A-Za-z]?)?)$", a)
    if m:
        return m.group(1).strip(" ,"), m.group(2).replace(" ", "")
    return a, ""


# ---------------------------------------------------------------- 读 Excel
def read_row(path, row_no):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
    vals = [c.value for c in ws[row_no]]
    d = dict(zip(headers, vals))
    if not d.get("编号"):
        raise SystemExit(f"第 {row_no} 行没有编号，检查 --row 参数")

    country = norm_country(d.get("国家"))
    zip_ = norm_zip(d.get("邮编"), country)
    city = str(d.get("城市") or "").strip()
    street, number = split_street(d.get("地址"), d.get("邮编"), city)

    try:
        qty = float(d.get("数量") or 0)
    except ValueError:
        qty = 0
    dg_text = str(d.get("Dangerous Goods Type/危险品类型") or "")
    material = str(d.get("产品物料") or "")
    is_battery = bool(re.search(r"电池|batter|lithium|li-?ion", dg_text + material, re.I))

    def num(key):
        try:
            return float(d.get(key))
        except (TypeError, ValueError):
            return None

    unit = str(d.get("Unit/单位") or "")
    if unit and "cm" not in unit.lower():
        raise SystemExit(f"单位不是公制 (cm/kg)：{unit!r}，请先人工处理")
    try:
        pallet_lines, pallet_notes = plan_pallets(
            material, qty, num("单个物料（长）"), num("单个物料（宽）"),
            num("单个物料（高）"), num("单个物料（重量）"))
    except ValueError as e:
        pallet_lines, pallet_notes = [], [f"无法自动计算托盘: {e}"]

    return {
        "pallet_lines": pallet_lines,
        "pallet_notes": pallet_notes,
        "order_no": str(d["编号"]).strip(),
        "country": country,
        "zip": zip_,
        "city": city,
        "street": street,
        "number": number,
        "goods_value": "5000" if qty <= 4 else "10000",
        "dangerous": is_battery,
        "tail_lift": str(d.get("是否需要尾板") or "").strip() in ("是", "yes", "Yes", "Y", "y"),
        "_raw_address": d.get("地址"),
        "_qty": qty,
    }


# ---------------------------------------------------------------- 填表
class Filler:
    def __init__(self, page):
        self.p = page
        self.cons = page.locator("fieldset.addr.float-right")  # Consignee 区域
        # 询价阶段不需要选 Collection date：隐藏日期选择弹窗，避免它弹出来挡住别的输入框或被误点
        try:
            page.add_style_tag(content="#ui-datepicker-div { display: none !important; }")
        except Exception:
            pass
        self.date0 = self._date_value()    # 记下日期框原来的值，最后检查有没有被改动

    def _date_value(self):
        try:
            return self.p.locator("input[name$='containerDatePicker:datePicker']").input_value(timeout=3000)
        except Exception:
            return None

    def _leave_date_field(self):
        """Tab 可能把焦点带进 Collection date 输入框（会弹出日期选择），如果是就立刻移开焦点。"""
        try:
            self.p.evaluate("() => { const a = document.activeElement;"
                            " if (a && a.classList && a.classList.contains('hasDatepicker')) a.blur(); }")
        except Exception:
            pass

    def wait_ajax(self, ms=400):
        """等 Wicket 的 ajax 结束（遮罩层 .overlay 消失）。"""
        self.p.wait_for_timeout(ms)
        try:
            self.p.wait_for_function(
                "() => { const o = document.querySelector('.overlay');"
                " return !o || getComputedStyle(o).display === 'none'; }",
                timeout=8000)
        except Exception:
            pass
        self.p.wait_for_timeout(200)

    def fill_text(self, locator, value):
        locator.scroll_into_view_if_needed()
        locator.click()
        locator.fill("")
        locator.type(str(value), delay=15)
        locator.press("Tab")      # 触发 change / blur
        self._leave_date_field()
        self.wait_ajax()

    def pick_fine_select(self, wrapper_class, option_text=None, option_value=None):
        """Dachser 自定义下拉框 (fineSelect)：点输入框 -> 点 li。"""
        box = self.p.locator(f"div.fineSelect.{wrapper_class}")
        box.locator("input.fineSelect").click()
        ul = box.locator("ul.optlist")
        ul.wait_for(state="visible", timeout=3000)
        li = (ul.locator(f"li[val='{option_value}']") if option_value
              else ul.locator("li", has_text=re.compile(rf"^\s*{re.escape(option_text)}\s*$")))
        li.first.click()
        self.wait_ajax()

    # --- 各字段 ---
    def order_number(self, v):
        self.fill_text(self.p.locator("input[name$='ordernumber']"), v)

    def country(self, name):
        try:
            self.pick_fine_select("selCountries", option_text=name)
        except Exception:
            # 兜底：直接改隐藏的 <select> 并触发 change
            self.p.evaluate(
                """(name) => { const s = document.querySelector('#selCountries');
                   const o=[...s.options].find(o=>o.text.trim()===name);
                   if(!o) throw new Error('no option '+name);
                   s.value=o.value; s.dispatchEvent(new Event('change',{bubbles:true})); }""",
                name)
            self.wait_ajax()

    def zip(self, v):
        self.fill_text(self.cons.locator("input[name$='ctnZipCity:zip']"), v)

    def _consignee_inputs(self):
        """返回收货方区域里所有可见的输入控件 (name, tag)。"""
        return self.cons.evaluate(
            """el => [...el.querySelectorAll('input[type=text],select')]
               .filter(i => !i.closest('.addressMatchcode') && i.offsetParent !== null
                           || i.tagName==='SELECT' && !i.closest('.addressMatchcode'))
               .map(i => ({name:i.name||'', tag:i.tagName,
                           label:(i.closest('tr')?.querySelector('td')?.innerText||'').trim()}))""")

    def _find(self, name_re, label_re):
        """先按 name 找，再按所在行的标签文字找。返回 locator 列表。"""
        found = []
        inputs = self._consignee_inputs()
        for it in inputs:
            if it["name"] and (re.search(name_re, it["name"], re.I)
                               or re.search(label_re, it["label"], re.I)):
                if "ctnZipCity:zip" in it["name"] or it["name"].endswith(":name1"):
                    continue
                found.append(it)
        return [self.cons.locator(f"[name='{it['name']}']") for it in found]

    def city(self, v):
        locs = self._find(r"city|location|ort\b", r"^\s*(city|location|town)")
        if not locs:
            raise RuntimeError("找不到 City 输入框")
        loc = locs[0]
        if loc.evaluate("e => e.tagName") == "SELECT":
            opts = loc.evaluate("e => [...e.options].map(o => [o.value, o.text])")
            hit = next((o for o in opts if v.lower() in o[1].lower()), None)
            if hit is None:
                hit = next((o for o in opts if o[0]), None)
            loc.select_option(hit[0])
            self.wait_ajax()
        else:
            self.fill_text(loc, v)

    def street(self, street, number):
        locs = self._find(r"street|strasse|str\b|road", r"^\s*street")
        if not locs:
            raise RuntimeError("找不到 Street 输入框")
        if len(locs) >= 2 or self._find(r"house|hnr|nr\b|number", r"house|no\."):
            nums = self._find(r"house|hnr|nr\b|number", r"house|no\.")
            num_loc = nums[0] if nums else locs[1]
            self.fill_text(locs[0], street)
            if number:
                self.fill_text(num_loc, number)
        else:
            self.fill_text(locs[0], f"{street} {number}".strip())

    @staticmethod
    def _fine_box(sel):
        """隐藏的 <select> 后面紧跟的 fineSelect 自定义下拉框。"""
        return sel.locator("xpath=following-sibling::div[contains(@class,'fineSelect')][1]")

    def _fine_state(self, sel):
        """返回 (页面上可见框里显示的文字, 隐藏 <select> 当前选中的文字)。
        两个都对才算真的选上了（人看到的和提交给服务器的一致）。"""
        box = self._fine_box(sel)
        shown = box.locator("input.fineSelect").evaluate(
            "e => (e.value || e.getAttribute('selected_text') || '').trim()")
        real = sel.evaluate("e => ((e.options[e.selectedIndex] || {}).text || '').trim()")
        return shown, real

    def pick_by_text(self, sel, text):
        """在 fineSelect 下拉框里按显示文字选一项（只点选，不输入）。
        依次尝试：真实鼠标点击 -> 在页面里派发鼠标事件 -> 同步修改隐藏 <select> 和可见框。
        每一步之后都检查“可见框”和“隐藏 select”是否都已经是目标值。"""
        box = self._fine_box(sel)
        inp = box.locator("input.fineSelect")
        ul = box.locator("ul.optlist")
        li = ul.locator("li", has_text=re.compile(rf"^\s*{re.escape(text)}\s*$", re.I)).first
        want = text.strip().lower()

        def done():
            shown, real = self._fine_state(sel)
            return shown.lower() == want and real.lower() == want

        def close_list():
            try:
                if ul.is_visible():
                    self.p.keyboard.press("Escape")
                    self.p.wait_for_timeout(150)
                if ul.is_visible():      # 还开着就直接收起，免得盖住下面的输入框
                    ul.evaluate("e => { e.style.display = 'none'; }")
            except Exception:
                pass

        if done():
            return
        for _ in range(2):
            # 1) 真实鼠标点击：点输入框打开列表（若点一下反而把列表关了，再点一次），再点选项
            try:
                inp.scroll_into_view_if_needed()
                inp.click()
                self.p.wait_for_timeout(250)
                if not ul.is_visible():
                    inp.click()
                    self.p.wait_for_timeout(250)
                li.click(timeout=3000)
            except Exception:
                pass
            self.p.wait_for_timeout(300)
            if done():
                close_list()
                self.wait_ajax()
                return
            # 2) 真实点击没生效：直接在页面里对选项派发鼠标事件
            try:
                li.evaluate("""el => { for (const t of ['mouseover', 'mousedown', 'mouseup', 'click'])
                    el.dispatchEvent(new MouseEvent(t, {bubbles: true, cancelable: true, view: window})); }""")
            except Exception:
                pass
            self.p.wait_for_timeout(300)
            if done():
                close_list()
                self.wait_ajax()
                return
        # 3) 兜底：同步修改隐藏 <select>（提交给服务器的值）和可见框（人看到的）
        sel.evaluate(
            """(e, name) => {
                const o = [...e.options].find(o => o.text.trim().toLowerCase() === name.toLowerCase());
                if (!o) throw new Error('no option ' + name);
                e.value = o.value;
                e.dispatchEvent(new Event('change', {bubbles: true}));
                const box = e.parentElement.querySelector('div.fineSelect');
                if (box) {
                    const inp = box.querySelector('input.fineSelect');
                    if (inp) { inp.value = o.text.trim();
                               inp.setAttribute('selected_value', o.value);
                               inp.setAttribute('selected_text', o.text.trim()); }
                    box.querySelectorAll('ul.optlist li').forEach(
                        li => li.classList.toggle('selected', li.getAttribute('val') === o.value));
                }
            }""", text)
        close_list()
        self.wait_ajax()
        if not done():
            shown, real = self._fine_state(sel)
            raise RuntimeError(f"选不上 {text!r}：可见框显示 {shown!r}，隐藏 select 是 {real!r}")

    def order_lines(self, lines, goods_desc=None):
        """Order lines：第 i 行 = listView:i。页面默认只有 2 行。
        goods_desc：货物描述（如电池危险品填 battery），填到每一行的描述框。
        体积 (nested:6:panel:volume) 填完长宽高后页面自动生成，不用填。"""
        if not lines:
            raise RuntimeError("没有可填的托盘方案")
        for i, ln in enumerate(lines):
            pre = f"orderLinePanel:listView:{i}:nested"
            self.fill_text(self.p.locator(f"input[name='{pre}:1:panel:counter']"), ln["qty"])
            # 包装类型：按下拉框里显示的文字选（如 Disposable Pallet）
            self.pick_by_text(self.p.locator(f"select[name='{pre}:2:panel:packing']"),
                              ln["packaging"])
            if goods_desc:   # 货物描述（freeText，最长 20 个字符）
                self.fill_text(self.p.locator(
                    f"input[name='{pre}:3:panel:content:contextSelect:input']"), goods_desc)
            self.fill_text(self.p.locator(f"input[name='{pre}:4:panel:weight']"), ln["weight"])
            self.fill_text(self.p.locator(f"input[name='{pre}:5:panel:container:length']"), ln["L"])
            self.fill_text(self.p.locator(f"input[name='{pre}:5:panel:container:width']"), ln["W"])
            self.fill_text(self.p.locator(f"input[name='{pre}:5:panel:container:height']"), ln["H"])

    def goods_value(self, v):
        self.fill_text(self.p.locator("input[name='ctnOrderData:tfiGoodsValue']"), v)

    def check(self, selector, want):
        cb = self.p.locator(selector)
        cb.scroll_into_view_if_needed()
        if cb.is_checked() != want:
            cb.set_checked(want)
            self.wait_ajax()

    def show_full_view(self):
        """开始输入前，先点一下收货方地址区域的 displayFullView 链接（展开完整视图）。
        图标是 show 才点，已经展开就不再点，避免把它收回去。"""
        link = self.cons.locator("a.displayFullView")
        link.wait_for(state="attached", timeout=5000)
        src = link.locator("img").get_attribute("src") or ""
        if "show" in src.lower():
            link.scroll_into_view_if_needed()
            link.click()
            self.wait_ajax(800)

    def verify_packing(self, lines):
        """检查每个订单行的包装类型是否真的是想要的（比如 Disposable Pallet），
        不是就重选一次；还不对就报错，这样不会静默地留下默认的 Euro Pallet。"""
        for i, ln in enumerate(lines):
            sel = self.p.locator(f"select[name='orderLinePanel:listView:{i}:nested:2:panel:packing']")
            want = ln["packaging"]
            shown, real = self._fine_state(sel)
            if shown.lower() != want.lower() or real.lower() != want.lower():
                self.pick_by_text(sel, want)      # 被后面的操作改回去了就重选
                shown, real = self._fine_state(sel)
                if shown.lower() != want.lower() or real.lower() != want.lower():
                    raise RuntimeError(f"第 {i + 1} 行包装类型：可见框 {shown!r}，隐藏 select {real!r}，"
                                       f"不是 {want!r}")

    def check_date_untouched(self):
        """Collection date 询价阶段不用管；只检查脚本有没有碰它，改了就报出来。"""
        now = self._date_value()
        if self.date0 is not None and now != self.date0:
            raise RuntimeError(f"Collection date 被改动了：{self.date0!r} -> {now!r}，请手动检查")

    def dangerous_goods(self, want):
        """勾选 Dangerous goods；是危险品时再在后面的下拉框里选 Dangerous goods。"""
        self.check("input[name$='ctnDG:cbDG']", want)
        if want:
            sel = self.p.locator("select[name$='ctnDG:dangerousGoodsChoice']")
            sel.wait_for(state="attached", timeout=5000)
            self.pick_by_text(sel, "Dangerous goods")

    def phone_booking(self):
        r = self.p.locator("input[name$='ctnAvis:selAdvice'][value='AT']")
        r.scroll_into_view_if_needed()
        r.check()
        self.wait_ajax()


def fill_form(page, d):
    f = Filler(page)
    steps = [
        ("Show full view", f.show_full_view),
        ("Order number", lambda: f.order_number(d["order_no"])),
        ("Country", lambda: f.country(d["country"])),
        ("PCD", lambda: f.zip(d["zip"])),
        ("City", lambda: f.city(d["city"])),
        ("Street", lambda: f.street(d["street"], d["number"])),
        ("Goods value", lambda: f.goods_value(d["goods_value"])),
        # 先处理危险品（勾选 + 下拉），再填订单行，避免页面刷新把订单行内容清掉
        ("Dangerous goods", lambda: f.dangerous_goods(d["dangerous"])),
        ("Order lines", lambda: f.order_lines(
            d["pallet_lines"], "battery" if d["dangerous"] else None)),
        ("Packing type check", lambda: f.verify_packing(d["pallet_lines"])),
        ("Collection date untouched", f.check_date_untouched),
        ("Tail lift delivery",
         lambda: f.check("input[name$='specialContainer:cbAutoLift']", d["tail_lift"])),
        ("Phone delivery booking", f.phone_booking),
    ]
    results = []
    for name, fn in steps:
        try:
            fn()
            results.append((name, True, ""))
            print(f"  [OK]   {name}")
        except Exception as e:  # 某一项失败不影响其他项
            results.append((name, False, str(e).splitlines()[0]))
            print(f"  [FAIL] {name}: {str(e).splitlines()[0]}")
            if name == "PCD":   # 邮编失败会导致后面的 City/Street 也找不到
                pass
    if not all(ok for _, ok, _ in results):
        print("\n收货方区域当前可见的输入框（用于排查）:")
        try:
            for it in f._consignee_inputs():
                print("   ", it)
        except Exception:
            pass
    return results


# ---------------------------------------------------------------- 主程序
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("--row", type=int, default=2, help="Excel 行号（默认 2，即第一条数据）")
    ap.add_argument("--url", default=START_URL)
    ap.add_argument("--yes", action="store_true", help="跳过托盘方案确认")
    ap.add_argument("--dry", action="store_true", help="只解析 Excel，不开浏览器")
    ap.add_argument("--profile", default="dachser_profile")
    ap.add_argument("--chromium-path", default=None, help="指定浏览器可执行文件（测试用）")
    ap.add_argument("--skip-login-wait", action="store_true", help="测试用")
    args = ap.parse_args()

    d = read_row(args.xlsx, args.row)
    print("解析结果:")
    for k, v in d.items():
        if not k.startswith("_"):
            print(f"  {k:12} = {v!r}")
    print(f"  (原始地址: {d['_raw_address']!r}, 数量: {d['_qty']})")
    print("\n托盘方案:")
    for ln in d["pallet_lines"]:
        print(f"  - {ln['desc']}  ->  数量 {ln['qty']} / {ln['packaging']} / "
              f"{ln['weight']}kg / {ln['L']}x{ln['W']}x{ln['H']}")
    for n in d["pallet_notes"]:
        print(f"  * {n}")
    if not args.dry and not args.yes:
        if input("\n托盘方案是否确认？回车=确认并开始填写，输入 n 退出: ").strip().lower() == "n":
            return
    if args.dry:
        return

    from playwright.sync_api import sync_playwright
    with sync_playwright() as pw:
        kw = dict(user_data_dir=str(Path(args.profile).resolve()), headless=False,
                  viewport=None, args=["--start-maximized"])
        if args.chromium_path:
            kw["executable_path"] = args.chromium_path
        else:
            kw["channel"] = "chrome"
        ctx = pw.chromium.launch_persistent_context(**kw)
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        if args.url:
            page.goto(args.url)
        if not args.skip_login_wait:
            input("\n请在浏览器里登录并打开 freightquotation -> Enquiry 页面，"
                  "看到 New enquiry 表单后按回车继续...")
        page.wait_for_selector("input[name$='ordernumber']", timeout=15000)
        print("\n开始填写:")
        fill_form(page, d)
        print("\n填写完成。请在浏览器里检查，确认无误后自行点击页面底部的 Enquiry。")
        if not args.skip_login_wait:
            input("检查完毕后按回车关闭浏览器...")
        ctx.close()


if __name__ == "__main__":
    sys.exit(main())