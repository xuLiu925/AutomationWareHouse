#!/usr/bin/env python3
"""
Dachser freightquotation 询价表单自动填写（第一步：只填写，不提交）

用法（Windows，命令行）:
    pip install playwright openpyxl
    python dachser_fill.py 询价表.xlsx            # 填第 1 条数据（Excel 第 2 行所在分组）
    python dachser_fill.py 询价表.xlsx --row 3    # 填 Excel 第 3 行所在的订单分组
    python dachser_fill.py 询价表.xlsx --dry      # 只解析 Excel，不打开浏览器
    python dachser_fill.py 询价表.xlsx --all      # 批量：依次填写所有订单（按运费报价列合并情况分组）
    python dachser_fill.py 询价表.xlsx --all --dry   # 批量预览：只列出每条订单的解析结果，不开浏览器

订单分组规则：
    "运费报价(欧元)"列合并的行属于同一订单（地址相同），脚本合并处理后只填一张表单。
    未合并的行各自独立。

首次运行会打开一个独立的 Chrome 窗口：请手动登录 Dachser，
进入 freightquotation -> Enquiry 页面后回到命令行按回车。
登录状态保存在 ./dachser_profile 目录里，下次不用重新登录。
脚本不会点击页面底部的 "Enquiry" 提交按钮，填完后由你检查并手动提交。
"""
import argparse
import math
import re
import sys
import time
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
import app_ui  # noqa: E402
from app_ui import ask  # noqa: E402
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


# ---------------------------------------------------------------- 订单分组检测
def detect_order_groups(path):
    """
    按"运费报价(欧元)"列的合并单元格识别订单分组。
    同一合并范围内的有编号行属于同一订单（地址相同）；未合并的行各自独立。
    返回: [[行号, ...], ...]  每个子列表是一个订单的所有 Excel 行号（按原顺序）
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]

    id_col = next((i + 1 for i, h in enumerate(headers) if h == "编号"), None)
    fare_col = next((i + 1 for i, h in enumerate(headers) if h == "运费报价(欧元)"), None)
    if id_col is None:
        raise SystemExit('Excel 第 1 行找不到“编号”列')

    data_rows_set = set()
    for r in range(2, ws.max_row + 1):
        v = ws.cell(row=r, column=id_col).value
        if v is not None and str(v).strip():
            data_rows_set.add(r)
    if not data_rows_set:
        return []

    # 运费报价列的合并范围 -> 将每行映射到组首行
    row_to_group = {}
    if fare_col:
        for merge in ws.merged_cells.ranges:
            if merge.min_col <= fare_col <= merge.max_col:
                for r in range(merge.min_row, merge.max_row + 1):
                    row_to_group[r] = merge.min_row

    groups: dict[int, list[int]] = {}
    for r in sorted(data_rows_set):
        key = row_to_group.get(r, r)
        groups.setdefault(key, []).append(r)

    return [rows for _, rows in sorted(groups.items())]


# ---------------------------------------------------------------- 读 Excel（支持多行订单）
def read_order_group(path, row_nums):
    """
    读取一个订单分组（一行或多行），返回可供 fill_form() 使用的字典。
    多行订单：地址字段取第一行（合并订单地址相同），所有产品的托盘方案合并为一个列表。

    托盘估算失败时：用原始箱体尺寸占位（qty=件数，packaging=Disposable Pallet），
    并在 pallet_notes 中标注 ⚠ 供人工核实。
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]

    def rowdict(r):
        return dict(zip(headers, [c.value for c in ws[r]]))

    first = rowdict(row_nums[0])
    if not first.get("编号"):
        raise SystemExit(f"第 {row_nums[0]} 行没有编号，检查输入")

    country = norm_country(first.get("国家"))
    zip_ = norm_zip(first.get("邮编"), country)
    city = str(first.get("城市") or "").strip()
    street, number = split_street(first.get("地址"), first.get("邮编"), city)
    tail_lift = str(first.get("是否需要尾板") or "").strip() in ("是", "yes", "Yes", "Y", "y")

    pallet_lines: list[dict] = []
    pallet_notes: list[str] = []
    is_dangerous = False
    all_order_nos: list[str] = []
    total_qty = 0.0

    for rn in row_nums:
        d = rowdict(rn)
        ono = str(d.get("编号") or "").strip()
        if ono:
            all_order_nos.append(ono)

        try:
            qty = float(d.get("数量") or 0)
        except (TypeError, ValueError):
            qty = 0.0
        if not qty:
            pallet_notes.append(f"{ono}: 数量为零，跳过")
            continue
        total_qty += qty

        dg_text = str(d.get("Dangerous Goods Type/危险品类型") or "")
        material = str(d.get("产品物料") or "")
        if re.search(r"电池|batter|lithium|li-?ion", dg_text + material, re.I):
            is_dangerous = True

        def num(key, _d=d):
            try:
                return float(_d.get(key))
            except (TypeError, ValueError):
                return None

        unit = str(d.get("Unit/单位") or "")
        if unit and "cm" not in unit.lower():
            pallet_notes.append(f"{ono}: 单位 {unit!r} 非公制，请人工处理此行")
            continue

        try:
            pl, pn = plan_pallets(
                material, qty,
                num("单个物料（长）"), num("单个物料（宽）"),
                num("单个物料（高）"), num("单个物料（重量）"))
            pallet_lines.extend(pl)
            pallet_notes.extend(pn)
        except (ValueError, Exception) as e:
            # 降级：用原始箱体尺寸占位，qty=件数，标注人工核实
            L_, W_, H_, Wk = (num("单个物料（长）"), num("单个物料（宽）"),
                               num("单个物料（高）"), num("单个物料（重量）"))
            if all(v is not None for v in (L_, W_, H_, Wk)):
                pallet_lines.append({
                    "desc": f"{material} 原始尺寸（需人工核实）",
                    "qty": int(math.ceil(qty)),
                    "packaging": "Disposable Pallet",
                    "weight": math.ceil(Wk * qty),
                    "L": math.ceil(L_), "W": math.ceil(W_), "H": math.ceil(H_),
                })
                pallet_notes.append(
                    f"⚠ {ono}/{material}: 托盘估算失败（{e}），"
                    "已用原始箱体尺寸 + 件数占位，请人工核实包装方式和托盘数")
            else:
                pallet_notes.append(
                    f"⚠ {ono}/{material}: 托盘估算失败（{e}），且缺少尺寸数据，请完全人工处理")

    return {
        "pallet_lines": pallet_lines,
        "pallet_notes": pallet_notes,
        "order_no": all_order_nos[0] if all_order_nos else "",
        "all_order_nos": all_order_nos,
        "country": country,
        "zip": zip_,
        "city": city,
        "street": street,
        "number": number,
        "goods_value": "5000" if total_qty <= 4 else "10000",
        "dangerous": is_dangerous,
        "tail_lift": tail_lift,
        "_raw_address": first.get("地址"),
        "_qty": total_qty,
        "_row_nums": row_nums,
    }


def read_row(path, row_no):
    """找到 row_no 所属的订单分组后整组读取（--row 模式入口）。"""
    for group in detect_order_groups(path):
        if row_no in group:
            return read_order_group(path, group)
    return read_order_group(path, [row_no])


# ---------------------------------------------------------------- 填表
class Filler:
    def __init__(self, page):
        self.p = page
        self.cons = page.locator("fieldset.addr.float-right")  # Consignee 区域
        self._last_filled_lines: list[dict] = []
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
        每一步之后都检查"可见框"和"隐藏 select"是否都已经是目标值。"""
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

    def _add_one_order_line(self):
        """
        点击"添加订单行"按钮，让表单多出一行。
        尝试常见的 Wicket 按钮选择器，成功返回 True，找不到按钮返回 False。
        """
        for sel in [
            "a[id$=':addOrderLine']",
            "a[id$=':addLine']",
            "a[id*='addLine']",
            "a:text-is('Add line')",
            "a:text-is('Add another line')",
            "input[type='submit'][value*='Add line']",
            ".addLine a",
            "a.add-line",
        ]:
            try:
                btn = self.p.locator(sel).first
                if btn.is_visible(timeout=500):
                    btn.scroll_into_view_if_needed()
                    btn.click()
                    self.wait_ajax()
                    return True
            except Exception:
                pass
        return False

    def order_lines(self, lines, goods_desc=None):
        """Order lines：第 i 行 = listView:i。页面默认只有 2 行。
        lines 超过 2 行时先尝试点添加按钮；无法自动添加则截断并警告。
        goods_desc：货物描述（如电池危险品填 battery），填到每一行的描述框。
        体积 (nested:6:panel:volume) 填完长宽高后页面自动生成，不用填。"""
        if not lines:
            raise RuntimeError("没有可填的托盘方案")

        # 超出默认 2 行时尝试扩展表单
        fillable = list(lines)
        for i in range(2, len(fillable)):
            if not self._add_one_order_line():
                print(f"  [WARN] 表单目前有 {i} 行，共需 {len(fillable)} 行；"
                      "未能自动添加更多行，后续行需人工填写")
                fillable = fillable[:i]
                break

        self._last_filled_lines = fillable
        for i, ln in enumerate(fillable):
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

    def product(self, name):
        """Transport information 里的 Product 下拉框（必填），统一选 targoflex。"""
        sel = self.p.locator("select[name$='ctnOrderData:selProducts']")
        sel.wait_for(state="attached", timeout=5000)
        self.pick_by_text(sel, name)

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

    def verify_packing(self, _lines=None):
        """检查每个订单行的包装类型是否真的是想要的（比如 Disposable Pallet），
        不是就重选一次；还不对就报错。使用 order_lines() 实际填写的行（截断后的）。"""
        lines = self._last_filled_lines or (_lines or [])
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
        ("Product", lambda: f.product("targoflex")),
        ("Country", lambda: f.country(d["country"])),
        ("PCD", lambda: f.zip(d["zip"])),
        ("City", lambda: f.city(d["city"])),
        ("Street", lambda: f.street(d["street"], d["number"])),
        ("Goods value", lambda: f.goods_value(d["goods_value"])),
        # 先处理危险品（勾选 + 下拉），再填订单行，避免页面刷新把订单行内容清掉
        ("Dangerous goods", lambda: f.dangerous_goods(d["dangerous"])),
        ("Order lines", lambda: f.order_lines(
            d["pallet_lines"], "battery" if d["dangerous"] else None)),
        ("Packing type check", f.verify_packing),
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


# ---------------------------------------------------------------- 批量模式
def parse_all(path):
    """按订单分组解析所有行 -> [{"rows", "order_no", "data", "error"}]。
    解析失败、或托盘方案全部算不出来的订单，error 里记原因，批量填写时会跳过。"""
    items = []
    for row_nums in detect_order_groups(path):
        it: dict = {"rows": row_nums, "order_no": None, "data": None, "error": None}
        try:
            d = read_order_group(path, row_nums)
            it["order_no"] = d["order_no"]
            it["data"] = d
            if not d["pallet_lines"]:
                it["error"] = "; ".join(d["pallet_notes"]) or "没有可用的托盘方案"
        except (Exception, SystemExit) as e:
            it["error"] = str(e) or repr(e)
        items.append(it)
    return items


def _plan_text(d):
    lines = d["pallet_lines"]
    if not lines:
        return "（无托盘）"
    parts = [f"{ln['qty']}托/{ln['L']}×{ln['W']}×{ln['H']}/{ln['weight']}kg"
             for ln in lines]
    return " + ".join(parts)


def _rows_label(rows):
    if len(rows) == 1:
        return f"第{rows[0]:>3}行"
    return f"第{rows[0]}-{rows[-1]}行({len(rows)}件)"


def print_batch_table(items):
    total_rows = sum(len(it["rows"]) for it in items)
    print(f"\nExcel 共 {len(items)} 个订单（合计 {total_rows} 行）：")
    for it in items:
        head = f"  {_rows_label(it['rows']):<16}  {(it['order_no'] or '?'):<22}"
        if it["error"]:
            print(f"{head} 【跳过】{it['error']}")
        else:
            d = it["data"]
            warn = "  ⚠有警告" if any("⚠" in n for n in d["pallet_notes"]) else ""
            print(f"{head} {d['country']:<12} 总数量 {d['_qty']:g}  {len(d['pallet_lines'])} 行订单"
                  + ("  [电池/危险品]" if d["dangerous"] else "") + warn)
    dups = {}
    for it in items:
        if it["order_no"]:
            dups[it["order_no"]] = dups.get(it["order_no"], 0) + 1
    dup_keys = [k for k, v in dups.items() if v > 1]
    if dup_keys:
        print(f"\n注意: 编号重复: {', '.join(dup_keys)}")


def _launch(pw, args):
    # chromium_sandbox=True：不让 Playwright 加 --no-sandbox，Chrome 顶部就不会出现"不受支持的命令行标记"提示
    kw = dict(user_data_dir=str(app_ui.app_path(args.profile)), headless=False,
              viewport=None, args=["--start-maximized"], chromium_sandbox=True)
    if args.chromium_path:
        kw["executable_path"] = args.chromium_path
    else:
        kw["channel"] = "chrome"
    try:
        return pw.chromium.launch_persistent_context(**kw)
    except Exception as e:
        print(f"  [WARN] 开启沙箱启动浏览器失败（{str(e).splitlines()[0]}），改用 --no-sandbox 启动")
        kw["chromium_sandbox"] = False
        return pw.chromium.launch_persistent_context(**kw)


LOGIN_USER = "input[name='user:unit:textfield']"
LOGIN_PASS = "input[name='password:password']"
LOGIN_BTN = "form:has(input[name='password:password']) input[type='submit']"
AUTOFILL_WINDOW_S = 6     # 登录页出现后这么多秒内检测到账号密码才自动点 Login（之后有值多半是你在手动输入）
LOGIN_TIMEOUT_S = 600


def _quotation_link(page):
    return page.locator("a.application-name-link", has=page.locator(
        "span.blue-application-name", has_text=re.compile(r"^\s*quotation\s*$", re.I))).first


def _login_filled(page):
    """账号、密码框都有值（包括 Chrome 自动填充：自动填充的密码在用户操作前 JS 读不到 value，要看 :-webkit-autofill）。"""
    return page.evaluate("""([u, p]) => [u, p].every(s => {
        const e = document.querySelector(s);
        if (!e) return false;
        try { if (e.matches(':-webkit-autofill')) return true; } catch (_) {}
        return e.value.trim() !== '';
    })""", [LOGIN_USER, LOGIN_PASS])


def open_quotation(page, url):
    """打开 Dachser -> 登录 -> 点左侧 freightquotation 进入询价页面。
    已登录：直接进；登录页账号密码已自动填好：自动点 Login；否则等你在浏览器里登录。"""
    page.goto(url)
    login_seen = None        # 第一次看到登录页的时间
    auto_clicked = told = False
    deadline = time.time() + LOGIN_TIMEOUT_S
    while True:
        page.wait_for_timeout(500)
        try:
            if _quotation_link(page).count():
                break
            if page.locator(LOGIN_PASS).count():
                login_seen = login_seen or time.time()
                if (not auto_clicked and time.time() - login_seen < AUTOFILL_WINDOW_S
                        and _login_filled(page)):
                    auto_clicked = True
                    print("检测到浏览器已保存的账号密码，自动点击 Login ...")
                    page.locator(LOGIN_BTN).first.click(timeout=5000)
                    page.wait_for_timeout(1500)
                    continue
                if not told:
                    print("请在浏览器里输入账号、密码并点击 Login，登录成功后脚本会自动继续 ...")
                    told = True
        except Exception:
            pass             # 页面跳转中，下一轮再看
        if time.time() > deadline:
            ask("一直没检测到登录成功。请在浏览器里登录 Dachser。", [("c", "继续等待")])
            deadline = time.time() + LOGIN_TIMEOUT_S
    print("已登录，进入 freightquotation ...")
    _goto_quotation(page)


def _wait_blank_form(page, args):
    """确认页面是一个空白的新询价表单（订单号框是空的），避免把上一条已填的内容覆盖掉。
    返回 True=可以填；False=用户选择结束。"""
    while True:
        cur, why = None, ""
        try:
            page.wait_for_selector("input[name$='ordernumber']", timeout=15000)
            cur = page.locator("input[name$='ordernumber']").input_value(timeout=3000).strip()
        except Exception as e:
            why = str(e).splitlines()[0]
        if cur == "":
            return True
        if args.skip_login_wait:        # 测试模式：不弹提示
            return False
        if cur is None:
            msg = f"找不到订单号输入框（{why}），请确认当前是 freightquotation -> Enquiry 的新表单页面。"
        else:
            msg = f"页面订单号框里已有内容 {cur!r}，看起来不是空白新表单。为避免覆盖上一条，请先打开新表单。"
        print("\n" + msg)
        if ask(msg + "\n处理好后重试。", [("c", "重试"), ("q", "结束")]) == "q":
            return False


RESULT_WAIT_MS = 600000      # 等你点 Enquiry 提交的最长时间（毫秒），超时会问你怎么办


def _read_result_info(page):
    """从询价结果页读出询价号和总价，读不到就返回空字符串（只用于汇总，不影响流程）。"""
    try:
        v = page.evaluate("""() => {
            const q = document.querySelector('.main-content h1 span.float-left');
            const t = document.querySelector('td.border b span');
            const h = [...document.querySelectorAll('th span')].map(e => e.innerText).find(s => /Amount in/i.test(s));
            return {title: q ? q.innerText : '', total: t ? t.innerText.trim() : '', head: h || ''};
        }""")
        qn = re.search(r"query number:\s*(\d+)", v["title"], re.I)
        cur = re.search(r"Amount in\s+(\w+)", v["head"], re.I)
        parts = []
        if qn:
            parts.append(f"询价号 {qn.group(1)}")
        if v["total"]:
            parts.append(f"总价 {v['total']} {cur.group(1) if cur else ''}".strip())
        return "，".join(parts)
    except Exception:
        return ""


def _wait_result_page(page):
    """等你点击 Enquiry 提交后出现的询价结果页（页面底部有 New enquiry 按钮）。
    返回 ("submitted", 摘要) / ("skipped", None) / ("quit", None)。"""
    print("等待你点击 Enquiry 提交……")
    while True:
        try:
            page.wait_for_selector("input[name='newSearch']", timeout=RESULT_WAIT_MS)
            return "submitted", _read_result_info(page)
        except Exception:
            ans = ask("一直没有检测到询价结果页。",
                      [("c", "继续等待"), ("s", "这条不提交，换下一条"), ("q", "结束")])
            if ans == "q":
                return "quit", None
            if ans == "s":
                return "skipped", None


def _goto_quotation(page):
    """点左侧菜单里的 freightquotation（蓝色的 quotation），回到空白的询价表单。"""
    link = _quotation_link(page)
    try:
        link.click(timeout=8000)
    except Exception:
        try:
            link.evaluate("e => e.click()")
        except Exception as e:
            print(f"  [WARN] 自动点击 quotation 失败：{str(e).splitlines()[0]}。请手动打开新的询价表单。")
            return
    try:    # 等到出现订单号框为空的新表单（旧表单还在时订单号非空，所以不会误判）
        page.wait_for_function(
            "() => { const i = document.querySelector(\"input[name$='ordernumber']\");"
            " return !!i && i.value.trim() === ''; }", timeout=20000)
    except Exception:
        print("  [WARN] 点击 quotation 后没有等到空白表单，请检查页面。")


def main_batch(args):
    items = parse_all(args.xlsx)
    if not items:
        raise SystemExit("Excel 里没有找到带编号的数据行")
    print_batch_table(items)

    # 打印所有警告备注
    has_warn = False
    for it in items:
        if it["data"]:
            warns = [n for n in it["data"]["pallet_notes"] if "⚠" in n]
            if warns:
                if not has_warn:
                    print("\n⚠ 以下托盘方案需人工核实：")
                    has_warn = True
                label = _rows_label(it["rows"])
                for w in warns:
                    print(f"  {label}: {w}")

    todo = [it for it in items if not it["error"]]
    print(f"\n将填写 {len(todo)} 个订单，跳过 {len(items) - len(todo)} 个（跳过的请手动处理）。")
    if not todo or args.dry:
        return
    if not args.yes and ask(f"将填写 {len(todo)} 个订单，跳过 {len(items) - len(todo)} 个。"
                            "托盘方案见运行记录，确认开始？", [("c", "开始"), ("n", "退出")]) == "n":
        return

    status = {}      # rows[0] -> 结果说明
    from playwright.sync_api import sync_playwright
    with sync_playwright() as pw:
        ctx = _launch(pw, args)
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        if args.skip_login_wait:
            if args.url:
                page.goto(args.url)
        else:
            open_quotation(page, args.url)
        for n, it in enumerate(todo, 1):
            if app_ui.stop_requested():
                print("\n已停止。")
                break
            d = it["data"]
            rows_label = _rows_label(it["rows"])
            print(f"\n===== [{n}/{len(todo)}] Excel {rows_label}  编号 {d['order_no']} =====")
            if len(it["rows"]) > 1:
                print(f"  此订单含 {len(it['rows'])} 行产品: {', '.join(d['all_order_nos'])}")
            print(f"  托盘方案: {_plan_text(d)}  ({len(d['pallet_lines'])} 行)")
            for note in d["pallet_notes"]:
                print(f"  * {note}")
            if not _wait_blank_form(page, args):
                break
            results = fill_form(page, d)
            failed = [f"{name}: {msg}" for name, ok, msg in results if not ok]
            base = ("已填写，全部 [OK]" if not failed
                    else "已填写，但有失败项: " + " | ".join(failed))
            status[it["rows"][0]] = base
            if args.skip_login_wait:
                continue
            print("\n请在浏览器里检查，确认无误后点击页面底部的 Enquiry 提交。"
                  "提交后脚本会自动点击 quotation 回到空白表单，继续下一条。")
            outcome, info = _wait_result_page(page)
            if outcome == "quit":
                break
            if outcome == "submitted":
                status[it["rows"][0]] = base + " | 已提交" + (f"：{info}" if info else "")
                print(f"  已检测到询价结果页{('：' + info) if info else ''}")
            else:
                status[it["rows"][0]] = base + " | 未提交（已跳过）"
            if n < len(todo):
                _goto_quotation(page)
            else:
                ask("全部处理完毕。", [("c", "关闭浏览器")])
        ctx.close()

    print("\n===== 汇总 =====")
    for it in items:
        rows_label = _rows_label(it["rows"])
        if it["error"]:
            text = f"跳过: {it['error']}"
        else:
            text = status.get(it["rows"][0], "未处理")
        print(f"  {rows_label:<18} {(it['order_no'] or '?'):<22} {text}")
    print('\n提示: “已填写”只表示脚本已把内容填进表单，是否已提交以你在页面上的操作为准。')


# ---------------------------------------------------------------- 主程序
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("--all", action="store_true",
                    help="批量模式：依次填写 Excel 里所有订单（按运费报价列合并情况分组）")
    ap.add_argument("--row", type=int, default=2,
                    help="Excel 行号（默认 2），自动识别所属订单分组后整组处理")
    ap.add_argument("--url", default=START_URL)
    ap.add_argument("--yes", action="store_true", help="跳过托盘方案确认")
    ap.add_argument("--dry", action="store_true", help="只解析 Excel，不开浏览器")
    ap.add_argument("--profile", default="dachser_profile")
    ap.add_argument("--chromium-path", default=None, help="指定浏览器可执行文件（测试用）")
    ap.add_argument("--skip-login-wait", action="store_true", help="测试用")
    args = ap.parse_args()
    if args.all:
        return main_batch(args)

    d = read_row(args.xlsx, args.row)
    print("解析结果:")
    for k, v in d.items():
        if not k.startswith("_"):
            print(f"  {k:16} = {v!r}")
    print(f"  (原始地址: {d['_raw_address']!r}, 总数量: {d['_qty']}, 行号: {d['_row_nums']})")
    print(f"\n  此订单包含 {len(d['all_order_nos'])} 个编号: {', '.join(d['all_order_nos'])}")
    print(f"\n托盘方案（{len(d['pallet_lines'])} 行）:")
    for ln in d["pallet_lines"]:
        print(f"  - {ln['desc']}  ->  qty={ln['qty']} / {ln['packaging']} / "
              f"{ln['weight']}kg / {ln['L']}×{ln['W']}×{ln['H']}")
    for n in d["pallet_notes"]:
        print(f"  * {n}")
    if not args.dry and not args.yes:
        if ask("\n托盘方案是否确认？", [("c", "确认并开始填写"), ("n", "退出")]) == "n":
            return
    if args.dry:
        return

    from playwright.sync_api import sync_playwright
    with sync_playwright() as pw:
        ctx = _launch(pw, args)
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        if args.skip_login_wait:
            if args.url:
                page.goto(args.url)
        else:
            open_quotation(page, args.url)
        page.wait_for_selector("input[name$='ordernumber']", timeout=15000)
        print("\n开始填写:")
        fill_form(page, d)
        print("\n填写完成。请在浏览器里检查，确认无误后自行点击页面底部的 Enquiry。")
        if not args.skip_login_wait:
            ask("检查完毕后关闭浏览器。", [("c", "关闭浏览器")])
        ctx.close()


if __name__ == "__main__":
    sys.exit(main())
