"""运行记录的欢迎画面：按时段的插画 + 问候 + 每日一句（服务器上 AI 每天写一句，所有人当天一样）。

只在窗口启动和"清空记录"后显示在运行记录第一条，不弹窗。连不上服务器时用下面自带的句子。
"""
import random
from datetime import datetime

import license_check
from app_ui import app_path

# 服务器取不到每日一句时随机用一句（两行，和 AI 的格式一样）
FALLBACK = [
    "今天的托盘都很守规矩，没有一个想偷偷长过 220 厘米。\n愿你的报价和咖啡一样：一次就对，还提神。",
    "据可靠消息：今天的邮编都不缺前导零。\n消息来源：一个很乐观的程序。",
    "纸箱不会说话，但它们都希望被温柔地码上托盘。\n今天也请对它们好一点。",
    "货车有尾板，生活有退路。\n实在不行，先去倒杯咖啡。",
    "周一的报价像德国的天气，说变就变。\n还好你比天气预报靠谱。",
    "今天的地址都写了门牌号——这是一个美好的愿望。\n程序会帮你盯着的。",
    "托盘码得整齐，心情自然整齐。\n今天也是井井有条的一天。",
    "Dachser 的卡车在路上，你的咖啡在手上。\n一切都在按计划进行。",
    "如果邮编会说话，它一定想说：别再把我的 0 弄丢了。\n今天程序替它看着。",
    "一单一单地建，一杯一杯地喝。\n下班时间会自己走过来的。",
    "危险品要贴标签，坏心情也一样：贴上就扔掉。\n今天轻装上阵。",
    "货到了，报价对了，客户笑了。\n这就是物流人的小确幸。",
    "今天的尾板们精神饱满，随时准备上岗。\n你也一样，加油。",
    "Excel 的格子再多，也装不下你的认真。\n今天也辛苦了。",
    "生活就像拼托盘：多试几种摆法，总能放下。\n别急，慢慢来。",
    "据统计，喝了咖啡的人建单速度提升 100%。\n统计人：我自己。",
]
FACES = ["(๑•̀ㅂ•́)و✧", "ヾ(≧▽≦*)o", "( •̀ω•́ )✧", "(〃'▽'〃)", "(ง •̀_•́)ง", "o(*￣▽￣*)ブ", "(｡･∀･)ﾉﾞ"]
GENERIC_USERS = {"admin", "administrator", "user", "users", "owner", "pc", "guest", "default"}
WEEKDAYS = "一二三四五六日"


def user_name():
    """Windows 用户名；admin / user 这类通用名不拿来称呼。"""
    name = license_check.win_user().strip()
    return "" if name.lower() in GENERIC_USERS else name


def hello(now=None):
    now = now or datetime.now()
    name, h, face = user_name(), now.hour, random.choice(FACES)
    if h >= 21 or h < 5:
        return f"都 {now:%H:%M} 了，{name}还在忙 {face}"
    word = "早上好" if h < 11 else "中午好" if h < 13 else "下午好" if h < 18 else "晚上好"
    return f"{word}，{name}！ {face}" if name else f"{word}！ {face}"


def scene_kind(now=None):
    h = (now or datetime.now()).hour
    return "morning" if 5 <= h < 11 else "day" if 11 <= h < 18 else "night"


def scene_file(now=None):
    p = app_path(f"assets/scene_{scene_kind(now)}.png")
    return p if p.exists() else None


# 横幅上的字色：(问候, 每日一句, 底色兜底)；日期在深色路面上，统一浅色
BANNER_COLORS = {
    "morning": ("#7c2d12", "#431407", "#fde7c0"),
    "day": ("#1e3a8a", "#1e293b", "#c7e2fb"),
    "night": ("#fef3c7", "#e0e7ff", "#2e1d6b"),
}
DATE_COLOR = "#e2e8f0"


def date_line(now=None):
    now = now or datetime.now()
    return f"{now.month}月{now.day}日 星期{WEEKDAYS[now.weekday()]}"


def fallback_line():
    return random.choice(FALLBACK)


def daily_line(timeout=8):
    """服务器上的每日一句；取不到返回 None。"""
    try:
        status, data = license_check.post("/greeting", {}, timeout)
    except Exception:
        return None
    text = data.get("text") if status == 200 and isinstance(data, dict) else None
    return text.strip() if isinstance(text, str) and text.strip() else None
