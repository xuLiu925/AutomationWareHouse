"""
托盘方案计算：按产品型号去 pallet_params.json 查参数，再算整托盘 + 零散托盘。

对外接口（给 dachser_fill.py 用）:
    plan_pallets(material, qty, L, W, H, weight) -> (lines, notes)

    lines: [{"desc", "qty", "packaging", "weight", "L", "W", "H"}, ...]
           最多 2 行：整托盘一行（qty=整托盘数）+ 零散托盘一行（qty=1）
    notes: 给人看的提示/警告
    查不到产品或参数没填全时抛 ValueError（dachser_fill 会捕获并提示人工处理）

参数表单位：尺寸 cm，重量 kg。
"""
import json
import math
import re
from pathlib import Path

PARAMS_FILE = Path(__file__).with_name("pallet_params.json")

# Dachser 订单行里的 weight 填什么，需要第一次对照页面确认：
#   "total"      = 这一行的总重量（整托盘行 = 单托盘重量 × 托盘数）
#   "per_pallet" = 每个托盘的重量
LINE_WEIGHT_MODE = "total"

REQUIRED = [
    "packaging_code",
    "carton_height_cm", "carton_weight_kg",
    "pallet_length_cm", "pallet_width_cm", "pallet_height_cm", "pallet_weight_kg",
    "per_layer", "max_layers",
]

_PARAMS = None


def _norm(s):
    return re.sub(r"\s+", "", str(s or "")).upper()


def _load():
    global _PARAMS
    if _PARAMS is None:
        with open(PARAMS_FILE, encoding="utf-8") as f:
            raw = json.load(f)
        _PARAMS = {k: v for k, v in raw.items() if not k.startswith("_")}
    return _PARAMS


def get_params(material):
    """按型号精确匹配（忽略大小写和空格），也支持 aliases。查不到就报错。"""
    want = _norm(material)
    for model, p in _load().items():
        names = [model] + list(p.get("aliases") or [])
        if want in {_norm(n) for n in names}:
            missing = [k for k in REQUIRED if p.get(k) is None]
            if missing:
                raise ValueError(f"{model} 的打托参数没填全，缺: {', '.join(missing)}")
            return model, p
    raise ValueError(f"产品 {material!r} 未配置打托参数，请先在 pallet_params.json 添加")


def _build_pallet(p, n):
    """一个托盘上放 n 个箱子 -> (高 cm, 重 kg)"""
    layers = math.ceil(n / p["per_layer"])
    height = layers * p["carton_height_cm"] + p["pallet_height_cm"]
    weight = n * p["carton_weight_kg"] + p["pallet_weight_kg"]
    return layers, height, weight


def plan_pallets(material, qty, L=None, W=None, H=None, weight=None):
    model, p = get_params(material)

    if qty is None or qty <= 0 or abs(qty - round(qty)) > 1e-9:
        raise ValueError(f"数量必须是正整数，当前: {qty!r}")
    qty = int(round(qty))

    notes = []

    # 用 Excel 里的单箱数据核对参数表，不一致只提示，不改结果
    def _check(label, excel, table, tol):
        if excel is not None and table is not None and abs(excel - table) > tol:
            notes.append(f"注意: Excel 单箱{label} {excel} 与参数表 {table} 不一致，请核对")
    _check("长", L, p.get("carton_length_cm"), 0.5)
    _check("宽", W, p.get("carton_width_cm"), 0.5)
    _check("高", H, p.get("carton_height_cm"), 0.5)
    _check("重量", weight, p.get("carton_weight_kg"), 0.5)

    per_pallet = p["per_layer"] * p["max_layers"]
    full, rest = divmod(qty, per_pallet)
    pl, pw = p["pallet_length_cm"], p["pallet_width_cm"]
    pack = p["packaging_code"]

    def _line(desc, count, height, weight_one):
        w = weight_one * count if LINE_WEIGHT_MODE == "total" else weight_one
        return {"desc": desc, "qty": count, "packaging": pack,
                "weight": math.ceil(w), "L": math.ceil(pl), "W": math.ceil(pw),
                "H": math.ceil(height)}

    lines = []
    if full:
        _, h, w = _build_pallet(p, per_pallet)
        if p.get("full_pallet_height_cm"):      # 有实测值就用实测值
            h = p["full_pallet_height_cm"]
        if p.get("full_pallet_weight_kg"):
            w = p["full_pallet_weight_kg"]
        lines.append(_line(f"{model} 整托盘 {per_pallet} 箱/托", full, h, w))
    if rest:
        layers, h, w = _build_pallet(p, rest)
        lines.append(_line(f"{model} 零散托盘 {rest} 箱 ({layers} 层)", 1, h, w))

    notes.append(f"{model}: {qty} 箱 = {full} 个整托盘({per_pallet} 箱/托)"
                 + (f" + 1 个零散托盘({rest} 箱)" if rest else ""))
    notes.append("托盘由参数表系统计算，发出前请确认")
    return lines, notes


if __name__ == "__main__":
    import sys
    m, q = (sys.argv[1], float(sys.argv[2])) if len(sys.argv) >= 3 else ("US2000", 100)
    ls, ns = plan_pallets(m, q)
    for x in ls:
        print(x)
    for n in ns:
        print("*", n)