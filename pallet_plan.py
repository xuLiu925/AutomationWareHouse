"""
托盘方案计算：按产品型号去 pallet_params.json 查本地参数（单箱尺寸重量、每层个数、
最大层数、托盘长宽、包装代码），托盘高度和自重统一按 DEFAULT_PALLET，
再算整托盘 + 零散托盘。本地参数为准；Excel 里的单个物料数据只用来核对，
不一致时只提示，不影响计算结果。

对外接口（给 dachser_fill.py 用）:
    plan_pallets(material, qty, L, W, H, weight) -> (lines, notes)

    L/W/H/weight: Excel 里"单个物料"的长宽高(cm)和重量(kg)，仅用于核对，可为 None
    lines: [{"desc", "qty", "packaging", "weight", "L", "W", "H"}, ...]
           最多 2 行：整托盘一行（qty=整托盘数）+ 零散托盘一行（qty=1）
    notes: 给人看的提示/警告
    查不到产品或参数没填全时抛 ValueError（dachser_fill 会捕获并提示人工处理）
"""
import json
import math
import re
from pathlib import Path

PARAMS_FILE = Path(__file__).with_name("pallet_params.json")

# 所有产品统一的托盘高度(cm)和自重(kg)。某个产品在 JSON 里写了同名字段则以 JSON 为准。
DEFAULT_PALLET = {
    "pallet_height_cm": 15,
    "pallet_weight_kg": 25,
}

# Dachser 订单行里的 weight 填什么，需要第一次对照页面确认：
#   "total"      = 这一行的总重量（整托盘行 = 单托盘重量 × 托盘数）
#   "per_pallet" = 每个托盘的重量
LINE_WEIGHT_MODE = "total"

# 每个产品必须在 JSON 里配置的字段
REQUIRED = [
    "packaging",
    "carton_length_cm", "carton_width_cm", "carton_height_cm", "carton_weight_kg",
    "pallet_length_cm", "pallet_width_cm",
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
            merged = dict(DEFAULT_PALLET)
            merged.update({k: v for k, v in p.items() if v is not None})
            return model, merged
    raise ValueError(f"产品 {material!r} 未配置打托参数，请先在 pallet_params.json 添加")


def _build_pallet(p, n):
    """一个托盘上放 n 个箱子 -> (层数, 高 cm, 重 kg)

    摆法优先级：
    1. 产品的 layers_by_qty 表里有这个箱数，就直接用表里的层数（{"7": 3} = 7 个放 3 层）；
    2. 否则箱数 >= per_layer（满一层）时每层放 per_layer 个，
       不满一层时每层放 small_per_layer 个（没配置就同 per_layer）。
    """
    override = (p.get("layers_by_qty") or {}).get(str(n))
    if override is not None:
        if (not isinstance(override, int) or isinstance(override, bool)
                or override < 1 or override > p["max_layers"]):
            raise ValueError(f"layers_by_qty[{n}] = {override!r} 无效，应为 1~{p['max_layers']} 的整数")
        if override * p["per_layer"] < n:
            raise ValueError(f"layers_by_qty[{n}] = {override} 层放不下 {n} 个（每层最多 {p['per_layer']} 个）")
        layers = override
    else:
        per_layer = p["per_layer"] if n >= p["per_layer"] else p.get("small_per_layer", p["per_layer"])
        layers = math.ceil(n / per_layer)
    height = layers * p["carton_height_cm"] + p["pallet_height_cm"]
    weight = n * p["carton_weight_kg"] + p["pallet_weight_kg"]
    return layers, height, weight


def plan_pallets(material, qty, L=None, W=None, H=None, weight=None):
    model, p = get_params(material)

    if qty is None or qty <= 0 or abs(qty - round(qty)) > 1e-9:
        raise ValueError(f"数量必须是正整数，当前: {qty!r}")
    qty = int(round(qty))

    notes = []

    # Excel 的单个物料数据只核对：不一致提示，结果始终按本地参数算
    for label, excel, key in (("长", L, "carton_length_cm"), ("宽", W, "carton_width_cm"),
                              ("高", H, "carton_height_cm"), ("重量", weight, "carton_weight_kg")):
        if excel is not None and abs(excel - p[key]) > 0.05:
            notes.append(f"注意: Excel 单个物料{label} {excel} 与本地 {p[key]} 不一致，已按本地计算")

    per_pallet = p["per_layer"] * p["max_layers"]
    full, rest = divmod(qty, per_pallet)
    pl, pw = p["pallet_length_cm"], p["pallet_width_cm"]
    pack = p["packaging"]          # Dachser 包装类型下拉框里显示的文字，如 Disposable Pallet

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
    notes.append("托盘由系统按本地参数计算，发出前请确认")
    return lines, notes


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 3:
        sys.exit("用法: python pallet_plan.py 型号 数量")
    ls, ns = plan_pallets(sys.argv[1], float(sys.argv[2]))
    for x in ls:
        print(x)
    for n in ns:
        print("*", n)