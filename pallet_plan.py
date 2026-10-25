"""
托盘方案计算：按产品物料编码（领星 SKU）去 pallet_params.json 查本地参数（单箱尺寸重量、每层个数、
最大层数、托盘长宽、包装代码），再算整托盘 + 零散托盘。

查到产品时：本地参数为准；Excel 里的单个物料数据只用来核对，不一致时只提示。
查不到产品时：用 Excel 单箱长宽高重量 + _defaults 兜底估算托盘方案，并在 notes 里
              标出"⚠ 未知产品，按通用托盘估算"，不会静默填写。

对外接口（给 dachser_fill.py 用）:
    plan_pallets(material, qty, L, W, H, weight) -> (lines, notes)

    L/W/H/weight: Excel 里"单个物料"的长宽高(cm)和重量(kg)；已知产品仅核对，
                  未知产品必须提供，否则报错
    lines: [{"desc", "qty", "packaging", "weight", "L", "W", "H"}, ...]
           最多 2 行：整托盘一行（qty=整托盘数）+ 零散托盘一行（qty=1）
    notes: 给人看的提示/警告
    参数没填全 / 箱子放不下 / 单托超重时抛 ValueError（dachser_fill 会捕获并提示人工处理）
"""
import json
import os
import math
import re
import shutil
from app_ui import APP_DIR

PARAMS_FILE = APP_DIR / "pallet_params.json"     # 放在程序外面，打包成 exe 后也能直接改

# 硬编码兜底值——_defaults 里没有对应键时才用这里
_HARDCODED_DEFAULTS = {
    "pallet_length_cm":      120,
    "pallet_width_cm":        80,
    "pallet_total_height_cm": 220,   # 含托盘底板的总高度上限
    "max_pallet_weight_kg":  1500,   # 含托盘自重的单托最大重量（所有产品通用）
    "pallet_height_cm":       15,
    "pallet_weight_kg":       25,
    "packaging":              "Disposable Pallet",
}

# Dachser 订单行里的 weight 填什么：
#   "total"      = 这一行的总重量（整托盘行 = 单托盘重量 × 托盘数）
#   "per_pallet" = 每个托盘的重量
LINE_WEIGHT_MODE = "total"

# 每个产品必须在 JSON 里配置的字段
REQUIRED = [
    "packaging",
    "carton_length_cm", "carton_width_cm", "carton_height_cm", "carton_weight_kg",
    "pallet_length_cm", "pallet_width_cm",
    "per_layer", "max_layers",
    "full_pallet_height_cm", "full_pallet_weight_kg",
]

_PARAMS = None
_DEFAULTS = None


def _norm(s):
    return re.sub(r"\s+", "", str(s or "")).upper()


def _load():
    global _PARAMS, _DEFAULTS
    if _PARAMS is None:
        with open(PARAMS_FILE, encoding="utf-8") as f:
            raw = json.load(f)
        _PARAMS = {k: v for k, v in raw.items() if not k.startswith("_")}
        raw_defaults = {k: v for k, v in (raw.get("_defaults") or {}).items()
                        if not k.startswith("_")}
        _DEFAULTS = {**_HARDCODED_DEFAULTS, **raw_defaults}
    return _PARAMS, _DEFAULTS


def reload():
    """参数文件被编辑器改过后调用，下次计算重新读取。"""
    global _PARAMS, _DEFAULTS
    _PARAMS = _DEFAULTS = None


def load_raw():
    """原样读取整个 JSON（含 _说明、_defaults），给编辑器用。"""
    with open(PARAMS_FILE, encoding="utf-8") as f:
        return json.load(f)


def save_raw(raw):
    """写回 JSON：先备份旧文件为 .bak，再写临时文件替换，写一半出错也不会把原文件弄坏。"""
    if PARAMS_FILE.exists():
        shutil.copy2(PARAMS_FILE, PARAMS_FILE.with_name(PARAMS_FILE.name + ".bak"))
    tmp = PARAMS_FILE.with_name(PARAMS_FILE.name + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(raw, f, ensure_ascii=False, indent=2)
        f.write("\n")
    os.replace(tmp, PARAMS_FILE)
    reload()


def _is_int(v):
    return isinstance(v, int) and not isinstance(v, bool)


def _is_pos(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and v > 0


DEFAULTS_REQUIRED = list(_HARDCODED_DEFAULTS)


def validate_defaults(d):
    """_defaults 校验 -> 错误列表"""
    errors = []
    for k in DEFAULTS_REQUIRED:
        v = d.get(k)
        if k == "packaging":
            if not str(v or "").strip():
                errors.append("packaging 不能为空")
        elif not _is_pos(v):
            errors.append(f"{k} 必须是大于 0 的数字")
    if not errors and d["pallet_height_cm"] >= d["pallet_total_height_cm"]:
        errors.append("托盘底板高度必须小于总高度上限")
    if not errors and d["pallet_weight_kg"] >= d["max_pallet_weight_kg"]:
        errors.append("托盘自重必须小于单托重量上限")
    return errors


def validate_product(sku, p, defaults):
    """单个产品校验 -> (errors, warnings)。errors 不允许保存，warnings 需要用户确认。

    defaults 是合并后的通用参数（托盘底板高/重、总高/重量上限）。
    """
    errors, warnings = [], []
    if not str(sku or "").strip():
        errors.append("产品物料编码不能为空")
    elif str(sku).strip().startswith("_"):
        errors.append("产品物料编码不能以 _ 开头（_ 开头的键会被程序忽略）")
    if not str(p.get("packaging") or "").strip():
        errors.append("packaging（包装类型）不能为空")
    for k in REQUIRED:
        if k == "packaging":
            continue
        v = p.get(k)
        if not _is_pos(v):
            errors.append(f"{k} 必须是大于 0 的数字")
        elif k in ("per_layer", "max_layers") and not _is_int(v):
            errors.append(f"{k} 必须是整数")
    if errors:
        return errors, warnings

    spl = p.get("small_per_layer")
    if spl is not None and (not _is_int(spl) or not 1 <= spl <= p["per_layer"]):
        errors.append(f"small_per_layer 必须是 1~{p['per_layer']} 的整数（每层个数）")
    per_pallet = p["per_layer"] * p["max_layers"]
    for q, layers in (p.get("layers_by_qty") or {}).items():
        if not str(q).isdigit() or not 1 <= int(q) < per_pallet:
            errors.append(f"layers_by_qty 箱数 {q!r} 必须是 1~{per_pallet - 1} 的整数（不满一整托的箱数）")
        elif not _is_int(layers) or not 1 <= layers <= p["max_layers"]:
            errors.append(f"layers_by_qty {q} 箱的层数 {layers!r} 必须是 1~{p['max_layers']} 的整数")
        elif layers * p["per_layer"] < int(q):
            errors.append(f"layers_by_qty {q} 箱放 {layers} 层放不下（每层最多 {p['per_layer']} 个）")

    total_h, max_w = defaults["pallet_total_height_cm"], defaults["max_pallet_weight_kg"]
    calc_h = p["max_layers"] * p["carton_height_cm"] + defaults["pallet_height_cm"]
    calc_w = per_pallet * p["carton_weight_kg"] + defaults["pallet_weight_kg"]
    if calc_h > total_h:
        errors.append(f"码满 {p['max_layers']} 层高 {calc_h:g} cm（{p['max_layers']}×{p['carton_height_cm']:g}"
                      f" + 托盘 {defaults['pallet_height_cm']:g}），超过总高上限 {total_h:g} cm")
    if calc_w > max_w:
        errors.append(f"整托 {per_pallet} 箱重 {calc_w:g} kg（{per_pallet}×{p['carton_weight_kg']:g}"
                      f" + 托盘 {defaults['pallet_weight_kg']:g}），超过单托上限 {max_w:g} kg")
    if p["full_pallet_height_cm"] > total_h:
        errors.append(f"整托实测高度 {p['full_pallet_height_cm']:g} cm 超过总高上限 {total_h:g} cm")
    if p["full_pallet_weight_kg"] > max_w:
        errors.append(f"整托实测重量 {p['full_pallet_weight_kg']:g} kg 超过单托上限 {max_w:g} kg")
    if errors:
        return errors, warnings

    if p["full_pallet_height_cm"] < calc_h:
        warnings.append(f"整托实测高度 {p['full_pallet_height_cm']:g} cm 比按箱子算的 {calc_h:g} cm 还低")
    if p["full_pallet_weight_kg"] < calc_w:
        warnings.append(f"整托实测重量 {p['full_pallet_weight_kg']:g} kg 比按箱子算的 {calc_w:g} kg 还轻")
    if (p["carton_length_cm"] > p["pallet_length_cm"] or p["carton_width_cm"] > p["pallet_width_cm"]) and \
            (p["carton_length_cm"] > p["pallet_width_cm"] or p["carton_width_cm"] > p["pallet_length_cm"]):
        warnings.append("单箱长宽比托盘还大")
    return errors, warnings


def get_params(material):
    """按产品物料编码（JSON 键名）精确匹配，忽略大小写和空格。

    查不到产品 -> 抛 LookupError（调用方捕获后走兜底）
    找到但参数没填全 -> 抛 ValueError（配置错误，不走兜底）
    """
    want = _norm(material)
    params, defaults = _load()
    for model, p in params.items():
        if want == _norm(model):
            missing = [k for k in REQUIRED if p.get(k) is None]
            if missing:
                raise ValueError(f"{model} 的打托参数没填全，缺: {', '.join(missing)}")
            merged = {
                "pallet_height_cm": defaults["pallet_height_cm"],
                "pallet_weight_kg": defaults["pallet_weight_kg"],
            }
            merged.update({k: v for k, v in p.items() if v is not None})
            return model, merged
    raise LookupError(f"产品 {material!r} 未配置打托参数")


def _fallback_params(material, L, W, H, weight):
    """查不到产品时，用 Excel 单箱数据 + _defaults 推算托盘参数。

    摆放规则：不旋转（长对应托盘长边 120，宽对应短边 80），只取整层。
    层数同时受高度和重量上限约束，取两者中较小值（保守报价原则）。
    """
    _, defaults = _load()

    missing_fields = [(label, v) for label, v in
                      [("长", L), ("宽", W), ("高", H), ("重量", weight)] if v is None]
    if missing_fields:
        names = "、".join(lbl for lbl, _ in missing_fields)
        raise ValueError(
            f"产品 {material!r} 未配置打托参数，且 Excel 缺少单箱数据（{names}），无法估算托盘"
        )

    pl   = defaults["pallet_length_cm"]
    pw   = defaults["pallet_width_cm"]
    ph   = defaults["pallet_height_cm"]
    pw_  = defaults["pallet_weight_kg"]
    th   = defaults["pallet_total_height_cm"]
    mw   = defaults["max_pallet_weight_kg"]
    pack = defaults["packaging"]

    # 每层箱数：不旋转，长对 120，宽对 80
    per_layer = math.floor(pl / L) * math.floor(pw / W)
    if per_layer <= 0:
        raise ValueError(
            f"产品 {material!r} 单箱长宽 {L}×{W} cm 超出托盘 {pl}×{pw} cm，无法自动估算"
        )

    # 高度限制：可码货高度 = 总高 - 托盘底板高
    cargo_h = th - ph
    max_layers_h = math.floor(cargo_h / H)
    if max_layers_h <= 0:
        raise ValueError(
            f"产品 {material!r} 单箱高 {H} cm 超出可码货高度 {cargo_h} cm（总高 {th} - 托盘 {ph}），无法自动估算"
        )

    # 重量限制：只取整层（保守报价），每层重 = per_layer × 单箱重
    layer_weight = per_layer * weight
    max_layers_w = math.floor((mw - pw_) / layer_weight)
    if max_layers_w <= 0:
        raise ValueError(
            f"产品 {material!r} 单层重量 {layer_weight:.1f} kg 已超出可载重 {mw - pw_:.1f} kg"
            f"（上限 {mw} - 托盘 {pw_}），无法自动估算"
        )

    max_layers = min(max_layers_h, max_layers_w)

    return {
        "packaging":          pack,
        "carton_length_cm":   L,
        "carton_width_cm":    W,
        "carton_height_cm":   H,
        "carton_weight_kg":   weight,
        "pallet_length_cm":   pl,
        "pallet_width_cm":    pw,
        "pallet_height_cm":   ph,
        "pallet_weight_kg":   pw_,
        "per_layer":          per_layer,
        "max_layers":         max_layers,
    }


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


def _check_weight(model, label, actual_weight):
    """检查单托重量是否超过全局上限，超了就抛 ValueError（转人工处理）。"""
    _, defaults = _load()
    limit = defaults["max_pallet_weight_kg"]
    if actual_weight > limit:
        raise ValueError(
            f"{model} {label} 单托重量 {actual_weight:.1f} kg 超出上限 {limit} kg，请人工处理"
        )


def plan_pallets(material, qty, L=None, W=None, H=None, weight=None):
    _load()  # 确保缓存已加载

    if qty is None or qty <= 0 or abs(qty - round(qty)) > 1e-9:
        raise ValueError(f"数量必须是正整数，当前: {qty!r}")
    qty = int(round(qty))

    notes = []
    is_unknown = False

    try:
        model, p = get_params(material)
    except LookupError:
        # 查不到产品 -> 用 Excel 尺寸兜底，标注为未知产品
        is_unknown = True
        model = str(material)
        p = _fallback_params(material, L, W, H, weight)
        notes.append(
            f"⚠ 产品 {material!r} 未配置打托参数，已按 Excel 单箱尺寸 + 通用托盘估算，"
            f"发出前请人工核实（每托最多 {p['per_layer']} 箱/层 × {p['max_layers']} 层 = "
            f"{p['per_layer'] * p['max_layers']} 箱）"
        )

    if not is_unknown:
        # 已知产品：核对 Excel 数据，不一致只提示，结果按本地参数
        for label, excel, key in (("长", L, "carton_length_cm"), ("宽", W, "carton_width_cm"),
                                  ("高", H, "carton_height_cm"), ("重量", weight, "carton_weight_kg")):
            if excel is not None and abs(excel - p[key]) > 0.05:
                notes.append(f"注意: Excel 单个物料{label} {excel} 与本地 {p[key]} 不一致，已按本地计算")

    per_pallet = p["per_layer"] * p["max_layers"]
    full, rest = divmod(qty, per_pallet)
    pl, pw = p["pallet_length_cm"], p["pallet_width_cm"]
    pack = p["packaging"]

    def _line(desc, count, height, weight_one):
        w = weight_one * count if LINE_WEIGHT_MODE == "total" else weight_one
        return {"desc": desc, "qty": count, "packaging": pack,
                "weight": math.ceil(w), "L": math.ceil(pl), "W": math.ceil(pw),
                "H": math.ceil(height)}

    lines = []
    if full:
        _, h, w = _build_pallet(p, per_pallet)
        if not is_unknown:                      # 已知产品整托用实测值（必填）
            h = p["full_pallet_height_cm"]
            w = p["full_pallet_weight_kg"]
        _check_weight(model, f"整托盘({per_pallet} 箱)", w)
        lines.append(_line(f"{model} 整托盘 {per_pallet} 箱/托", full, h, w))
    if rest:
        layers, h, w = _build_pallet(p, rest)
        _check_weight(model, f"零散托盘({rest} 箱)", w)
        lines.append(_line(f"{model} 零散托盘 {rest} 箱 ({layers} 层)", 1, h, w))

    notes.append(f"{model}: {qty} 箱 = {full} 个整托盘({per_pallet} 箱/托)"
                 + (f" + 1 个零散托盘({rest} 箱)" if rest else ""))
    notes.append("托盘由系统按本地参数计算，发出前请确认")
    return lines, notes


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 3:
        sys.exit("用法: python pallet_plan.py 产品物料编码 数量 [长 宽 高 重量]")
    args = sys.argv[1:]
    material_ = args[0]
    qty_ = float(args[1])
    dims = [float(x) for x in args[2:6]] if len(args) >= 6 else [None, None, None, None]
    ls, ns = plan_pallets(material_, qty_, *dims)
    for x in ls:
        print(x)
    for n in ns:
        print("*", n)
