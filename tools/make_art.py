#!/usr/bin/env python3
"""画软件图标和运行记录欢迎画面的插画，输出到 assets/（开发时运行一次，结果提交进仓库、随程序打包）。

    pip install pillow        # 只有画图时需要，程序本身不用
    python tools/make_art.py

assets/xlfp.ico           exe / 窗口图标：货车 + 速度线（16、24 像素去掉速度线）
assets/scene_<时段>.png   欢迎画面插画：morning 早上 / day 白天 / night 晚上
用到 Windows 自带字体（Segoe UI）。
"""
import math
import random
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter, ImageFont

OUT = Path(__file__).resolve().parent.parent / "assets"
F = "C:/Windows/Fonts/"


def lerp(a, b, t):
    return tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(3))


def hexc(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def gradient(w, h, c1, c2, diagonal=True):
    g = Image.new("RGB", (w, h))
    px = g.load()
    for y in range(h):
        for x in range(w):
            t = (x + y) / (w + h) if diagonal else y / h
            px[x, y] = lerp(c1, c2, t)
    return g


# ------------------------------------------------------------------ 图标
def icon(size=1024, small=False):
    """深蓝圆角底 + 白色货车侧影 + 橙色速度线；small（16/24 像素）去掉速度线、货车放大。"""
    S = size
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    mask = Image.new("L", (S, S), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, S - 1, S - 1], radius=int(S * 0.22), fill=255)
    bg = gradient(S, S, hexc("#1d4ed8"), hexc("#0b1f4d")).convert("RGBA")
    img.paste(bg, (0, 0), mask)
    d = ImageDraw.Draw(img)
    white, navy, dark, orange = (255, 255, 255, 255), hexc("#1e40af") + (255,), hexc("#0b1f4d") + (255,),         hexc("#fb923c") + (255,)
    u = S / 100
    dx, sc = (-8, 1.16) if small else (0, 1.0)         # 小图：货车放大、往左挪（没有速度线）

    def X(v):
        return (50 + (v - 50) * sc + dx) * u

    def Y(v):
        return (50 + (v - 50) * sc) * u

    d.rounded_rectangle([X(26), Y(30), X(62), Y(64)], radius=3 * u * sc, fill=white)            # 车厢
    d.rounded_rectangle([X(64), Y(41), X(81), Y(64)], radius=4.5 * u * sc, fill=white)          # 车头
    d.rounded_rectangle([X(68), Y(45), X(78), Y(53)], radius=1.5 * u * sc, fill=navy)           # 车窗
    for wx in (36, 52, 73):                                                                     # 轮子
        r1, r2 = 6.5, 3.6
        d.ellipse([X(wx - r1), Y(66 - r1), X(wx + r1), Y(66 + r1)], fill=dark)
        d.ellipse([X(wx - r2), Y(66 - r2), X(wx + r2), Y(66 + r2)], fill=white)
    if not small:
        for y, x0 in ((38, 9), (47, 13), (56, 9)):                                             # 速度线
            d.rounded_rectangle([X(x0), Y(y - 2), X(21), Y(y + 2)], radius=2 * u, fill=orange)
    return img


# ------------------------------------------------------------------ 插画
def scene(kind, W=560, H=210, k=3):
    w, h = W * k, H * k
    pal = {
        "morning": ("#fde68a", "#fb7185", "#f97316"),
        "day": ("#bae6fd", "#60a5fa", "#fde047"),
        "night": ("#1e1b4b", "#4c1d95", "#fef3c7"),
    }[kind]
    sky = gradient(w, h, hexc(pal[0]), hexc(pal[1]), diagonal=False).convert("RGBA")
    mask = Image.new("L", (w, h), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, w - 1, h - 1], radius=28 * k, fill=255)
    img = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    img.paste(sky, (0, 0), mask)
    d = ImageDraw.Draw(img)
    # 太阳 / 月亮 + 光晕
    cx, cy, r = w * 0.80, h * 0.30, 30 * k
    glow = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    gd = ImageDraw.Draw(glow)
    gd.ellipse([cx - r * 2.2, cy - r * 2.2, cx + r * 2.2, cy + r * 2.2], fill=hexc(pal[2]) + (70,))
    glow = glow.filter(ImageFilter.GaussianBlur(18 * k))
    glow.putalpha(Image.composite(glow.getchannel("A"), Image.new("L", (w, h), 0), mask))
    img.alpha_composite(glow)
    d = ImageDraw.Draw(img)
    if kind == "night":
        m = Image.new("L", (w, h), 0)
        md = ImageDraw.Draw(m)
        md.ellipse([cx - r, cy - r, cx + r, cy + r], fill=255)
        md.ellipse([cx - r * 0.35, cy - r * 1.15, cx + r * 1.45, cy + r * 0.7], fill=0)
        img.paste(Image.new("RGBA", (w, h), hexc(pal[2]) + (255,)), (0, 0), m)
        d = ImageDraw.Draw(img)
    else:
        d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=hexc(pal[2]) + (255,))
    if kind == "night":
        random.seed(3)
        for _ in range(40):
            sx, sy = random.uniform(0, w * 0.7), random.uniform(0, h * 0.5)
            rr = random.uniform(1, 2.6) * k
            d.ellipse([sx - rr, sy - rr, sx + rr, sy + rr], fill=(255, 255, 255, random.randint(120, 255)))
    # 远山
    hill = {"morning": hexc("#fdba74"), "day": hexc("#86efac"), "night": hexc("#312e81")}[kind]
    for i, (amp, base, alpha) in enumerate([(26, 0.62, 110), (18, 0.70, 200)]):
        pts = [(0, h)]
        for xx in range(0, w + 10, 10):
            pts.append((xx, h * base - amp * k * math.sin(xx / w * math.pi * (2.3 + i) + i)))
        pts.append((w, h))
        d.polygon(pts, fill=hill + (alpha,))
    # 路
    road_y = h * 0.80
    d.rectangle([0, road_y, w, h], fill=(55, 65, 81, 255))
    for xx in range(-20 * k, w, 60 * k):
        d.rounded_rectangle([xx, road_y + 22 * k, xx + 32 * k, road_y + 27 * k], radius=3 * k, fill=(250, 204, 21, 255))
    img.putalpha(Image.composite(img.getchannel("A"), Image.new("L", (w, h), 0), mask))
    d = ImageDraw.Draw(img)
    # 货车
    tx, ty = w * 0.22, road_y
    trailer = [tx, ty - 92 * k, tx + 210 * k, ty - 14 * k]
    d.rounded_rectangle(trailer, radius=10 * k, fill=(255, 255, 255, 255))
    d.rounded_rectangle([trailer[0], trailer[3] - 14 * k, trailer[2], trailer[3]], radius=6 * k, fill=(203, 213, 225, 255))
    # 车厢上的字
    f = ImageFont.truetype(F + "seguibl.ttf", int(24 * k))
    d.text((tx + 34 * k, ty - 82 * k), "FreightPilot", font=f, fill=(30, 58, 138, 255))
    f2 = ImageFont.truetype(F + "seguisb.ttf", int(14 * k))
    d.text((tx + 36 * k, ty - 48 * k), "on time · every time", font=f2, fill=(100, 116, 139, 255))
    # 车头
    cab = [tx + 214 * k, ty - 68 * k, tx + 278 * k, ty - 14 * k]
    d.rounded_rectangle(cab, radius=12 * k, fill=(249, 115, 22, 255))
    d.rounded_rectangle([cab[0] + 24 * k, cab[1] + 9 * k, cab[2] - 8 * k, cab[1] + 30 * k], radius=5 * k, fill=(186, 230, 253, 255))
    d.rectangle([tx + 208 * k, ty - 30 * k, tx + 216 * k, ty - 16 * k], fill=(71, 85, 105, 255))
    # 车灯
    d.ellipse([cab[2] - 8 * k, ty - 30 * k, cab[2] + 2 * k, ty - 22 * k], fill=(254, 240, 138, 255))
    # 轮子
    for wx in (tx + 40 * k, tx + 70 * k, tx + 175 * k, tx + 252 * k):
        rr = 14 * k
        d.ellipse([wx - rr, ty - 14 * k - rr + 4 * k, wx + rr, ty - 14 * k + rr + 4 * k], fill=(30, 41, 59, 255))
        rr2 = 6 * k
        d.ellipse([wx - rr2, ty - 10 * k - rr2, wx + rr2, ty - 10 * k + rr2], fill=(203, 213, 225, 255))
    # 速度线
    for i, (yy, ln) in enumerate([(-80, 70), (-58, 110), (-36, 80)]):
        x1 = tx - 20 * k
        d.rounded_rectangle([x1 - ln * k, ty + yy * k, x1, ty + (yy + 5) * k], radius=3 * k, fill=(255, 255, 255, 170))
    return img.resize((W, H), Image.LANCZOS)


def main():
    OUT.mkdir(exist_ok=True)
    big, small = icon(), icon(small=True)
    frames = [(small if s <= 24 else big).resize((s, s), Image.LANCZOS) for s in (256, 128, 64, 48, 32, 24, 16)]
    frames[0].save(OUT / "xlfp.ico", sizes=[f.size for f in frames], append_images=frames[1:])
    for kind in ("morning", "day", "night"):
        scene(kind).save(OUT / f"scene_{kind}.png", optimize=True)
    print(f"已生成到 {OUT}")


if __name__ == "__main__":
    main()
