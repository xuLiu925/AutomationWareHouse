"""
AI 整理询价表里客户乱写的格子（编号混备注、地址里带邮编城市国家、收件人前缀等）。

- 通过开发者的服务器（worker/worker.js，Cloudflare Worker）调用 Google Gemini；
  Gemini Key 和提示词都只在服务器上，这里只发数据、收结果。用的是授权码（license_check）。
- 每一单都发给 AI。只发送整理需要的字段（编号、地址、邮编、城市、国家、尾板、收件人、公司），不发电话和邮箱。
- 没有授权码或调用失败时返回 None，调用方改用规则版。
- AI 结果由调用方用规则校验（编号必须出自原文、街道不能含邮编等），不直接信任。
"""
import license_check

TIMEOUT_S = 120

FIELDS = ("编号", "地址", "邮编", "城市", "国家", "是否需要尾板", "收件人", "公司名称")

last_model = ""     # 服务器返回的模型名，打印用


def available():
    return bool(license_check.get_token())


def _cell(v):
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return "" if v is None else str(v).strip()


def clean_rows(rows):
    """rows: [{"row": Excel 行号, "编号": ..., "地址": ..., ...}]
    返回 {行号: {"order_no", "remark", "street", "tail_lift", "issues"}}；不可用或失败返回 None。"""
    global last_model
    if not available() or not rows:
        return None
    payload = [{"row": r["row"], **{k: _cell(r.get(k)) for k in FIELDS}} for r in rows]
    try:
        status, data = license_check.post("/clean", {"rows": payload}, TIMEOUT_S)
    except Exception as e:
        print(f"  [AI] 连不上服务器（{str(e).splitlines()[0] if str(e) else repr(e)}），改用规则识别")
        return None
    if status != 200 or not isinstance(data, dict) or not isinstance(data.get("results"), list):
        why = (data.get("error") or data.get("message")) if isinstance(data, dict) else ""
        print(f"  [AI] 调用失败（HTTP {status}{'：' + why if why else ''}），改用规则识别")
        return None
    last_model = data.get("model", "")
    out = {}
    for it in data["results"]:
        try:
            out[int(it["row"])] = it
        except (KeyError, TypeError, ValueError):
            pass
    return out
