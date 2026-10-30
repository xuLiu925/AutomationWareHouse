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

编号格里混写的备注：取开头的字母数字当编号，其余当备注；认不准时会问。
"备注"列和编号格里的备注脚本不处理，只列出来提示人工查看。
"是否需要尾板"列（前缀匹配）：先看否定词（不/无/否/no…），再看肯定词；空或认不出时会问。

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
import ai_clean  # noqa: E402
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


# ---------------------------------------------------------------- 地址整理
NUMBER_FIRST = {"France", "Ireland", "United Kingdom", "Luxembourg"}      # 这些国家门牌号写在街道前面
_HOUSE_NO = r"\d+\s?[A-Za-z]?(?:\s?[-/]\s?\d+[A-Za-z]?)?"
_STREET_WORD = (r"(stra(ß|ss)e|str\.|weg|gasse|platz|allee|ring|damm|ufer|steig|straat|laan|dreef|"
                r"plein|rue|avenue|chemin|via|calle|road|street)")
_ALL_COUNTRY_NAMES = sorted({a for al in _COUNTRIES.values() for a in al if len(a) > 3}
                            | {n.lower() for n in _COUNTRIES}, key=len, reverse=True)


def _zip_text(zip_raw):
    if isinstance(zip_raw, float) and zip_raw.is_integer():
        zip_raw = int(zip_raw)
    return str(zip_raw if zip_raw is not None else "").strip()


def _zip_pattern(zip_raw):
    """地址里邮编的写法：允许 'LT-21134'、'619 00'/'61900'，纯数字邮编允许 Excel 吃掉的前导 0。"""
    z = _zip_text(zip_raw)
    if not z:
        return ""
    if re.fullmatch(r"(?:[A-Za-z]{1,3}\s*-\s*)?[\d\s]+", z) and len(re.sub(r"\D", "", z)) >= 4:
        core = r"\s*".join(re.sub(r"\D", "", z).lstrip("0"))
        return r"(?<![\w-])(?:[A-Za-z]{1,3}\s*-\s*)?0*" + core + r"(?!\w)"
    return r"(?<!\w)" + r"\s*".join(re.escape(c) for c in z.replace(" ", "")) + r"(?!\w)"


def _strip_tail(s, word):
    """去掉末尾的城市/国家名。前面要有逗号或门牌号，或者粘在小写字母后（'RiedstadtGermany'），
    否则可能是街名本身（'Rue de Paris'、'Rue de France'）。"""
    m = re.search(r"([,\s]*)" + re.escape(word) + r"\s*$", s, re.I)
    if not m or not m.start():
        return s
    before, sep = s[:m.start()], m.group(1)
    glued = not sep and before[-1].islower() and s[m.end(1)].isupper()
    if "," in sep or glued or re.search(r"\d[A-Za-z]?$", before):
        return before
    return s


def clean_street(address, zip_raw, city, names=()):
    """Street 栏只填街道 + 门牌号（一个格子，保留原顺序，法国等门牌号在前也不动）。
    去掉地址里的括号备注、国家、邮编及其后的内容、城市、z.H./c/o 前缀、开头的收件人/公司名（names）。
    返回 (street, 括号备注列表)。"""
    s = str(address or "").translate(str.maketrans("，、；（）：", ",,;():"))
    notes = [n.strip() for n in re.findall(r"\(([^)]*)\)", s) if n.strip()]
    s = re.sub(r"\([^)]*\)", " ", s)
    city = str(city or "").strip()
    for n in _ALL_COUNTRY_NAMES:
        s = _strip_tail(s.rstrip(), n)
        city = _strip_tail(city, n)
    s = s.strip(" ,;")
    pat = _zip_pattern(zip_raw)
    m = re.search(pat, s) if pat else None
    if m:                                 # 邮编后面只会是城市/国家；前面没内容时才保留后面
        s = s[:m.start()] if s[:m.start()].strip(" ,;") else s[m.end():]
    if city:
        s = _strip_tail(s.strip(" ,;"), city)
    if re.match(r"^\s*(z\.?\s*h(d|änden)?\.?|c/o|attn\.?:?)\s+", s, re.I):
        w = re.search(r"[^\s,]*" + _STREET_WORD, s, re.I)
        if w:
            s = s[w.start():]
    for n in names:
        n = re.sub(r"\s+", " ", str(n or "")).strip()
        if len(n) >= 3:
            s = re.sub(r"^\s*" + re.escape(n) + r"[\s,]+(?=\S)", "", s, flags=re.I)
    # 逗号分段时取第一段同时有字母和数字的（公司/仓库名在前、区名在后都不要）
    parts = [p.strip() for p in s.split(",") if p.strip()]
    hit = [p for p in parts if re.search(r"\d", p) and re.search(r"[^\W\d_]", p)]
    if len(parts) > 1 and hit:
        s = hit[0]
    return re.sub(r"\s+", " ", s).strip(" ,;"), notes


def street_warnings(street):
    out = []
    if not re.search(r"\d", street):
        out.append(f"⚠ Street {street!r} 没有门牌号，请核对")
    m = re.search(r"(?<![\w.-])\d{4,6}(?![\w.-])", street)
    if m:
        out.append(f"⚠ Street {street!r} 里还有像邮编的数字 {m.group(0)}，可能和邮编列不一致")
    return out


def street_input(address, zip_raw, city, country, recipient, company, hno="", addr2=""):
    """询价和建单的 Street 用同一份输入（也是发给 AI 的字段）；询价表没有 House No / Address2，留空。"""
    return {"地址": address, "House No": re.sub(r"\.0$", "", str(hno or "").strip()),
            "Address2": str(addr2 or "").strip(), "邮编": zip_raw, "城市": city, "国家": country,
            "收件人": recipient, "公司名称": company}


def rule_street(inp):
    """规则版 Street（询价、建单共用）：整理地址后补 House No，保留原顺序。Address2 只有纯门牌号时才用。
    返回 (street, 括号备注列表, 提醒列表)。"""
    notes = []
    street, addr_notes = clean_street(inp["地址"], inp["邮编"], inp["城市"], (inp["收件人"], inp["公司名称"]))
    hno, addr2 = inp["House No"], inp["Address2"]
    if addr2 and not hno and not re.search(r"\d", street) and re.fullmatch(_HOUSE_NO, addr2):
        hno, addr2 = addr2, ""
    if addr2 and addr2 != hno:
        notes.append(f"⚠ Address2 没有填进 Street，请确认是否需要: {addr2!r}")
    if not hno or re.search(rf"(?<![\w/-]){re.escape(hno)}(?![\w/-])", street, re.I):
        return street, addr_notes, notes
    try:
        front = norm_country(inp["国家"]) in NUMBER_FIRST
    except ValueError:
        front = False
    if re.match(r"\d", street) if front else re.search(_HOUSE_NO + "$", street):
        notes.append(f"⚠ House No {hno!r} 和地址里的门牌号不一致，Street 保留地址原文，请核对")
        return street, addr_notes, notes
    return (f"{hno} {street}" if front else f"{street} {hno}").strip(), addr_notes, notes


def street_ok(street, address, zip_raw, city="", names=(), rule="", hno=""):
    """校验 AI 给的街道，不通过返回原因，通过返回 ""（防止 AI 编造、漏门牌号、没删干净）。
    address 是原地址全文；rule 是规则结果，用来判断原地址有没有门牌号；hno 是 House No 列。"""
    if not street:
        return "是空的"
    if not re.search(r"[^\W\d_]", street):
        return "没有街道名"
    raw = str(address or "").lower()
    words = re.findall(r"\w+", street.lower())
    if not all(re.search(rf"(?<!\d){t}(?!\d)" if t.isdigit() else re.escape(t), raw) for t in words):
        return "有原地址里没有的内容"
    pat = _zip_pattern(zip_raw)
    if pat and re.search(pat, street):
        return "还含有邮编"
    if re.search(r"(?<!\d)\d{1,3}(?!\d)", rule) and not re.search(r"\d", street):
        return "漏了门牌号"
    if any(t and _strip_tail(street, t) != street for t in [city] + _ALL_COUNTRY_NAMES):
        return "还含有城市或国家"
    for n in names:
        n = re.sub(r"\s+", " ", str(n or "")).strip().lower()
        if len(n) >= 3 and street.lower().startswith(n) and len(street) > len(n):
            return "还含有收件人或公司名"
    if hno and not re.search(rf"(?<!\w){re.escape(hno)}(?!\w)", street, re.I):
        return "没有用 House No"
    return ""


def merge_ai_street(rec, inp, a):
    """把 AI 结果合并进 rec（street / street_src / street_alt / street_notes），询价、建单共用。
    通过 street_ok 就用 AI；不通过且和规则不同，留到 ask_street 让人选。"""
    ai = re.sub(r"\s+", " ", (a.get("street") or "")).strip(" ,;")
    raw = " ".join(str(inp[k] or "") for k in ("地址", "House No", "Address2"))
    bad = street_ok(ai, raw, inp["邮编"], inp["城市"], (inp["收件人"], inp["公司名称"]), rec["street"], inp["House No"])
    if not bad:
        rec["street"], rec["street_src"] = ai, "AI"
        rec["street_notes"] = [n for n in rec["street_notes"] if not n.startswith("⚠ House No")]
    elif ai and ai != rec["street"]:
        rec["street_alt"] = ai
        rec["street_notes"].append(f"⚠ AI 给的 Street {ai!r} {bad}，规则结果 {rec['street']!r}，需要确认")
    if (a.get("issues") or "").strip():
        rec["street_notes"].append("⚠ AI 提示: " + a["issues"].strip())


def street_all_notes(rec):
    return rec["street_notes"] + street_warnings(rec["street"])


def ask_street(rec, label, raw, can_skip=True):
    """AI 结果没通过校验时让人选。返回 False = 用户选择跳过这单。"""
    if not rec.get("street_alt"):
        return True
    ans = ask(f"{label} 的 Street 认不准：\n原文：{raw}",
              [("r", f"用规则结果 {rec['street']}"), ("a", f"用 AI 结果 {rec['street_alt']}")]
              + ([("s", "跳过这单，手动处理")] if can_skip else []))
    if ans == "s":
        return False
    if ans == "a":
        rec["street"] = rec["street_alt"]
    rec["street_alt"] = ""
    return True


# ---------------------------------------------------------------- 编号 / 尾板 识别
ORDER_NO_MAX = 35      # Dachser Order number 输入框的最大长度


def parse_order_no(raw):
    """客户常把备注写进编号格：'MR12345678，客户收货时间…'。
    取开头连续的字母/数字/-_/. 作为编号，其余当备注。
    返回 (编号, 备注, 是否有把握)；没把握时由调用方问用户。"""
    if isinstance(raw, float) and raw.is_integer():
        raw = int(raw)
    s = str(raw if raw is not None else "").strip()
    m = re.match(r"[A-Za-z0-9][A-Za-z0-9\-_/.]*", s)
    if not m:
        return "", s, False                     # 开头就是中文/符号
    no = m.group(0).rstrip("-_/.")
    rest = s[m.end():]
    remark = rest.strip(" \t\r\n,，;；:：、")
    sure = 3 <= len(no) <= ORDER_NO_MAX
    # 只隔一个空格就接着字母数字（'AB 123'）：可能是编号的一部分，也可能是备注
    if re.match(r"[ \t]+[A-Za-z0-9]", rest):
        sure = False
    # 紧跟着 # * 之类的符号（'PO#4711'）：分不清
    if rest and not re.match(r"[\s,，;；:：、(（一-鿿]", rest):
        sure = False
    return no, remark, sure


_TAIL_NO = re.compile(r"不|无|否|没|\bno\b|\bnein\b|\bnicht\b|\bkein|\bfalse\b|^n$", re.I)
_TAIL_YES = re.compile(r"需要|要|是|有|尾板|\byes\b|\bja\b|\btrue\b|^y$|tail|hebeb", re.I)


def parse_tail_lift(raw):
    """是否需要尾板 -> True / False / None（空或认不出，交给用户选）。先判否定词。"""
    s = str(raw if raw is not None else "").strip()
    if not s:
        return None
    if _TAIL_NO.search(s):
        return False
    if _TAIL_YES.search(s):
        return True
    return None


def _col(headers, prefix):
    """按前缀找列名（客户可能写成 '是否需要尾板车辆' 之类）。"""
    return next((h for h in headers if h.startswith(prefix)), None)


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
        if any(h.startswith(("Outbound Order No", "Remark")) for h in headers):
            raise SystemExit("这是领星导出的出库单，不是询价表。请把功能改成\"建单\"再开始。")
        shown = "、".join(h for h in headers if h)[:300] or "（第 1 行是空的）"
        raise SystemExit(f"Excel 第 1 行找不到\"编号\"列，看起来不是询价表。\n第 1 行读到的是：{shown}")

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
    addr_input = street_input(first.get("地址"), first.get("邮编"), city, first.get("国家"),
                              first.get("收件人"), first.get("公司名称"))
    street, addr_notes, street_notes = rule_street(addr_input)
    tail_col = _col(headers, "是否需要尾板")
    tail_raw = first.get(tail_col) if tail_col else None
    tail_lift = parse_tail_lift(tail_raw)

    # 编号格里的备注 + "备注"列（可能拆在合并区域的下面几行）都收集起来，提示人工处理
    remarks: list[str] = []
    note_col = _col(headers, "备注")
    note_rows = set(row_nums)
    id_idx = headers.index("编号") + 1
    for rn in row_nums:
        for mr in ws.merged_cells.ranges:
            if mr.min_col <= id_idx <= mr.max_col and mr.min_row == rn:
                note_rows.update(range(mr.min_row, mr.max_row + 1))

    pallet_lines: list[dict] = []
    pallet_notes: list[str] = []
    is_dangerous = False
    all_order_nos: list[str] = []
    order_no_raw = ""
    order_no_sure = True
    id_remark = ""
    total_qty = 0.0

    for rn in row_nums:
        d = rowdict(rn)
        raw_no = d.get("编号")
        ono, extra, sure = parse_order_no(raw_no)
        if not order_no_raw:
            order_no_raw = str(raw_no or "").strip()
            order_no_sure = sure
        if extra:
            remarks.append(extra)
            if rn == row_nums[0]:
                id_remark = extra
        ono = ono or str(raw_no or "").strip()
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

    for rn in sorted(note_rows):
        v = rowdict(rn).get(note_col) if note_col else None
        if v is not None and str(v).strip():
            remarks.append(str(v).strip())
    remarks += [f"地址括号里: {n}" for n in addr_notes]
    # 去重：被另一条更长备注包含的就不再单独列（备注列常把编号格里的内容拆开再写一遍）
    flat = [re.sub(r"\s+", " ", r) for r in remarks]
    remarks = [r for i, r in enumerate(remarks)
               if not any(j != i and flat[i] in flat[j] and (len(flat[j]) > len(flat[i]) or j < i)
                          for j in range(len(flat)))]

    return {
        "pallet_lines": pallet_lines,
        "pallet_notes": pallet_notes,
        "order_no": all_order_nos[0] if all_order_nos else "",
        "order_no_raw": order_no_raw,
        "order_no_sure": order_no_sure,
        "remarks": remarks,
        "_id_remark": id_remark,
        "all_order_nos": all_order_nos,
        "country": country,
        "zip": zip_,
        "city": city,
        "street": street,
        "number": "",                      # 询价单 Street 栏只有一个格子，门牌号已在 street 里
        "street_src": "规则",
        "street_alt": "",
        "street_notes": street_notes,
        "clean_notes": [],
        "_ai_input": {"row": row_nums[0], "编号": first.get("编号"), "是否需要尾板": tail_raw, **addr_input},
        "goods_value": "5000" if total_qty <= 4 else "10000",
        "dangerous": is_dangerous,
        "tail_lift": tail_lift,
        "tail_raw": "" if tail_raw is None else str(tail_raw).strip(),
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
         lambda: f.check("input[name$='specialContainer:cbAutoLift']", bool(d["tail_lift"]))),
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


def apply_ai(datas):
    """用 AI 整理结果补充/修正规则识别（编号、尾板、街道），结果写回每个 d。
    规则有把握且与 AI 一致 -> 照用；规则没把握 -> 用 AI（需通过校验）；两者冲突 -> 标为待确认。
    返回说明文字（打印用）。"""
    datas = [d for d in datas if d]
    if not datas:
        return ""
    if not ai_clean.available():
        return "AI 整理：未填写授权码，使用规则识别"
    print(f"AI 整理中（{len(datas)} 单，约需几秒）…", flush=True)
    res = ai_clean.clean_rows([d["_ai_input"] for d in datas])
    if res is None:
        return "AI 整理：调用失败，本次使用规则识别"
    for d in datas:
        inp = d["_ai_input"]
        a = res.get(inp["row"])
        if not a:
            d["clean_notes"].append("AI 没有返回这一单的结果，使用规则识别")
            continue
        # 编号：必须原样出自编号格
        raw_no = d["order_no_raw"]
        ai_no = (a.get("order_no") or "").strip()
        if ai_no and ai_no in raw_no:
            if not d["order_no_sure"]:
                d["order_no"], d["order_no_sure"] = ai_no, True
                d["clean_notes"].append(f"编号由 AI 识别: {ai_no}")
                # 编号格剩下的备注也改用 AI 拆出来的
                if d["_id_remark"] in d["remarks"]:
                    d["remarks"].remove(d["_id_remark"])
                ai_remark = (a.get("remark") or "").strip()
                if ai_remark and ai_remark not in d["remarks"]:
                    d["remarks"].insert(0, ai_remark)
            elif ai_no != d["order_no"]:
                d["order_no_alt"] = ai_no
                d["order_no_sure"] = False
                d["clean_notes"].append(f"编号: 规则识别为 {d['order_no']}，AI 识别为 {ai_no}，开始前确认")
        # 尾板
        ai_tail = {"yes": True, "no": False}.get(a.get("tail_lift"))
        if d["tail_lift"] is None and ai_tail is not None:
            d["tail_lift"] = ai_tail
            d["clean_notes"].append(f"尾板由 AI 判断: {'需要' if ai_tail else '不需要'}（原文 {d['tail_raw']!r}）")
        elif d["tail_lift"] is not None and ai_tail is not None and ai_tail != d["tail_lift"]:
            d["tail_lift"] = None
            d["clean_notes"].append("尾板: 规则和 AI 判断不一致，开始前确认")
        # 街道：AI 结果要通过校验才用
        merge_ai_street(d, inp, a)
    return f"AI 整理：已启用（{ai_clean.last_model or 'Gemini'}）"


def print_streets(items):
    """Street 和原地址不同的（被整理过的）列出来，方便核对。"""
    rows = [it for it in items if it["data"] and it["data"]["street"] != str(it["data"]["_raw_address"] or "").strip()]
    if rows:
        print("\n🏠 以下地址已整理成 Street（请核对）：")
        for it in rows:
            d = it["data"]
            print(f"  {_rows_label(it['rows'])} {d['street']!r:34} ← 原文 {d['_raw_address']!r}  [{d['street_src']}]")
    notes = [(it, n) for it in items if it["data"]
             for n in it["data"]["clean_notes"] + street_all_notes(it["data"])]
    if notes:
        print("\n🔎 整理说明（编号 / 尾板 / 地址）：")
        for it, n in notes:
            print(f"  {_rows_label(it['rows'])} {n}")


def resolve_uncertain(d, label):
    """编号没把握、尾板认不出、Street 的 AI 结果没通过校验时问用户，结果写回 d。返回 False = 用户选择跳过这单。"""
    if not d["order_no_sure"]:
        raw = d["order_no_raw"]
        alt = d.get("order_no_alt")
        opts = ([("c", f"用 {d['order_no']}")] if d["order_no"] and d["order_no"] != raw else []) + (
            [("a", f"用 AI 识别的 {alt}")] if alt and alt != d["order_no"] else []) + [
            ("f", f"用整格原文（超过 {ORDER_NO_MAX} 字会被截断）"), ("s", "跳过这单，手动处理")]
        ans = ask(f"{label} 的编号认不准：\n{raw}\n\n脚本识别为：{d['order_no'] or '（无）'}", opts)
        if ans == "s":
            return False
        if ans == "f":
            d["order_no"] = raw
        elif ans == "a":
            d["order_no"] = alt
        d["order_no_sure"] = True
    if d["tail_lift"] is None:
        ans = ask(f"{label}（编号 {d['order_no']}）的\"是否需要尾板\""
                  f"{'是空的' if not d['tail_raw'] else '写的是：' + d['tail_raw']}，需要尾板吗？",
                  [("y", "需要尾板"), ("n", "不需要尾板")])
        d["tail_lift"] = ans == "y"
    return ask_street(d, f"{label}（编号 {d['order_no']}）", d["_raw_address"])


def _tail_text(d):
    return {True: "要尾板", False: "不要尾板", None: "尾板?"}[d["tail_lift"]]


def print_remarks(items):
    """编号格里的备注和"备注"列内容：脚本不处理，列出来提示人工处理。"""
    rows = [it for it in items if it["data"] and it["data"]["remarks"]]
    if not rows:
        return
    print("\n📝 以下订单有备注（收货时间、后备联系人等），脚本不处理，请人工查看：")
    for it in rows:
        print(f"  {_rows_label(it['rows'])} {it['order_no']}:")
        for r in it["data"]["remarks"]:
            print("      " + r.replace("\n", " / "))


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
            if not d["order_no_sure"]:
                warn += "  ⚠编号待确认"
            print(f"{head} {d['country']:<12} 总数量 {d['_qty']:g}  {len(d['pallet_lines'])} 行订单  "
                  f"{_tail_text(d)}" + ("  [电池/危险品]" if d["dangerous"] else "") + warn)
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
    print("\n" + apply_ai([it["data"] for it in items]))
    for it in items:
        if it["data"]:
            it["order_no"] = it["data"]["order_no"]
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
    print_streets(items)
    print_remarks(items)

    todo = [it for it in items if not it["error"]]
    unsure = [it for it in todo if not it["data"]["order_no_sure"] or it["data"]["tail_lift"] is None
              or it["data"].get("street_alt")]
    if unsure:
        print(f"\n有 {len(unsure)} 个订单的编号、尾板或 Street 需要确认，开始前会逐个询问。")
    print(f"\n将填写 {len(todo)} 个订单，跳过 {len(items) - len(todo)} 个（跳过的请手动处理）。")
    if not todo or args.dry:
        return
    for it in unsure:
        if not resolve_uncertain(it["data"], _rows_label(it["rows"])):
            it["error"] = "用户选择跳过（信息认不准）"
        it["order_no"] = it["data"]["order_no"]
    todo = [it for it in todo if not it["error"]]
    if not todo:
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
            for r in d["remarks"]:
                print(f"  📝 备注（请人工处理）: {r.replace(chr(10), ' / ')}")
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
    print(apply_ai([d]))
    for n in d["clean_notes"] + street_all_notes(d):
        print(f"  🔎 {n}")
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
    for r in d["remarks"]:
        print(f"  📝 备注（请人工处理）: {r.replace(chr(10), ' / ')}")
    if not d["order_no_sure"]:
        print(f"  ⚠ 编号认不准，原文: {d['order_no_raw']!r}")
    if not args.dry and not resolve_uncertain(d, f"第 {d['_row_nums'][0]} 行"):
        return
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
