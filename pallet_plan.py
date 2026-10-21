"""
托盘方案计算。所有业务规则都集中在本文件顶部，后续新增/修改型号只改这里。

单位：厘米 / 公斤。
"""
import math
import re

# ------------------------------------------------------------ 规则来自 pallet_rules.json
import json
from pathlib import Path

_CFG = json.loads((Path(__file__).resolve().parent / "pallet_rules.json").read_text(encoding="utf-8"))
_G = _CFG["global"]
MAX_PALLET = tuple(_G["max_pallet"])          # 通用货物托盘最大长宽高（高度含托盘）
PALLET_BASE_H = _G["pallet_base_h"]           # 托盘自身高度
PALLET_WEIGHT = _G["pallet_weight"]           # 托盘自重
MAX_PALLET_WEIGHT = _G["max_pallet_weight"]   # 单托最大总重
ALLOW_TIPPING = _G["allow_tipping"]           # 是否允许侧放/倒放
WEIGHT_IS_LINE_TOTAL = _G["weight_is_line_total"]  # 页面重量是整行总重？首次使用请核对
PRODUCT_RULES = _CFG["products"]


def find_rule(material):
    for r in PRODUCT_RULES:
        if re.search(r["match"], str(material or ""), re.I):
            return r
    return None


# ------------------------------------------------------------ 通用：一层能放几个
def _grid(L, W, a, b):
    """把 a×b 的矩形按整齐网格放进 L×W，返回个数。"""
    if a <= 0 or b <= 0:
        return 0
    return int(L // a) * int(W // b)


def per_layer(L, W, a, b):
    """
    一层最多放几个 a×b 的箱子（允许旋转 90°）。
    考虑：整齐网格（两种朝向）+ 两种朝向分区混放（简单切割）。
    """
    best = max(_grid(L, W, a, b), _grid(L, W, b, a))
    for (LL, WW) in ((L, W), (W, L)):
        for (x, y) in ((a, b), (b, a)):
            # 先沿 LL 方向放 k 排 (x×y)，剩余长度用旋转后的 (y×x) 填
            for k in range(1, int(LL // x) + 1):
                n1 = k * int(WW // y)
                n2 = _grid(LL - k * x, WW, y, x)
                best = max(best, n1 + n2)
    return best


def generic_capacity(l, w, h, unit_weight):
    """通用货物：一托最多几个。返回 (容量, 说明)。"""
    L, W, H = MAX_PALLET
    usable_h = H - PALLET_BASE_H
    orientations = [(l, w, h)]
    if ALLOW_TIPPING:
        orientations = [(l, w, h), (l, h, w), (w, h, l)]
    best, best_note = 0, ""
    for (a, b, hh) in orientations:
        layers = int(usable_h // hh) if hh > 0 else 0
        n = per_layer(L, W, a, b) * layers
        if n > best:
            best, best_note = n, f"每层 {per_layer(L, W, a, b)} 个 × {layers} 层"
    if unit_weight and unit_weight > 0:
        wcap = int((MAX_PALLET_WEIGHT - PALLET_WEIGHT) // unit_weight)
        if wcap < best:
            best, best_note = wcap, f"受 {MAX_PALLET_WEIGHT}kg 限重，最多 {wcap} 个"
    return best, best_note


# ------------------------------------------------------------ 主入口
def plan_pallets(material, qty, l, w, h, unit_weight):
    """
    返回 (lines, notes)。
    lines: 最多 2 行 [{"qty","packaging","weight","L","W","H","desc"}]
           第 1 行 = 满托，第 2 行 = 不满的最后一托。
    notes: 需要人工留意的提示（list[str]）。
    """
    notes = []
    qty = int(math.ceil(qty or 0))
    if qty <= 0:
        raise ValueError("数量为 0 或缺失")
    if not unit_weight or unit_weight <= 0:
        raise ValueError("单件重量缺失")

    rule = find_rule(material)
    if rule:
        cap = rule["per_pallet"]
        pL, pW, pH = rule["dims"]
        packaging = rule["packaging"]
        # 按重量再校验一遍
        wcap = int((MAX_PALLET_WEIGHT - PALLET_WEIGHT) // unit_weight)
        if wcap < cap:
            notes.append(f"{rule['name']} 满托 {cap} 箱按单重 {unit_weight}kg 超过 "
                         f"{MAX_PALLET_WEIGHT}kg，已降为 {wcap} 箱/托")
            cap = wcap
        src = f"已知型号 {rule['name']}"
        item_min = None
    else:
        if not (l and w and h):
            raise ValueError("未知型号且长/宽/高缺失，无法计算托盘")
        cap, how = generic_capacity(l, w, h, unit_weight)
        if cap < 1:
            raise ValueError(f"单件 {l}×{w}×{h} 放不进 {MAX_PALLET} 托盘，需要人工处理")
        pL, pW, pH = MAX_PALLET
        packaging = "EU"
        src = f"通用算法（{how}）"
        item_min = min(l, w, h)
        notes.append("型号不在已知列表，按表中尺寸通用计算，请核对")

    full, rest = divmod(qty, cap)
    lines = []

    def full_weight(n):
        return n * unit_weight + PALLET_WEIGHT

    fixed = rule.get("full_weight") if rule else None
    if full:
        # 有实测满托重量且满托箱数没被限重改过 -> 直接用实测值，否则按单箱重量计算
        use_fixed = fixed is not None and cap == rule["per_pallet"]
        wt = fixed if use_fixed else full_weight(cap)
        lines.append({"qty": full, "packaging": packaging, "L": pL, "W": pW, "H": pH,
                      "weight": wt * (full if WEIGHT_IS_LINE_TOTAL else 1),
                      "desc": f"{full} 托满托，每托 {cap} 箱，单托 {wt:.1f}kg" + ("（实测）" if use_fixed else "")})
    if rest:
        # 不满的托盘：高度按装载比例估算，避免按 220 高虚报体积
        ratio_h = PALLET_BASE_H + (pH - PALLET_BASE_H) * rest / cap
        min_h = PALLET_BASE_H + (item_min or 0) + (0 if item_min else 15)
        rest_h = int(min(pH, max(ratio_h, min_h)))
        rest_h = int(math.ceil(rest_h / 5.0) * 5)
        rest_h = min(rest_h, pH)
        wt = full_weight(rest)
        lines.append({"qty": 1, "packaging": packaging, "L": pL, "W": pW, "H": rest_h,
                      "weight": wt,
                      "desc": f"1 托散托，{rest} 箱，高度估算 {rest_h}cm，{wt:.1f}kg"})
        notes.append(f"散托高度 {rest_h}cm 为按装载比例估算，请确认")

    if len(lines) > 2:
        raise ValueError("超过页面预置的行数")
    for ln in lines:
        ln["weight"] = int(math.ceil(ln["weight"]))
        if ln["weight"] > 99999 or ln["qty"] > 999:
            raise ValueError("数量/重量超出页面输入框长度限制")
        if ln["weight"] / (ln["qty"] if WEIGHT_IS_LINE_TOTAL else 1) > MAX_PALLET_WEIGHT + 1:
            raise ValueError("单托超重")
    return lines, [f"算法来源: {src}"] + notes