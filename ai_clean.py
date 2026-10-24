"""
AI 整理询价表里客户乱写的格子（编号混备注、地址里带邮编城市国家、收件人前缀等）。

- 服务商：Google Gemini（REST，不需要额外安装包）。Key 从环境变量 GEMINI_API_KEY 读，
  模型可用 XLFP_AI_MODEL 覆盖。没有 Key 或调用失败时返回 None，调用方改用规则版。
- 只发送整理需要的字段（编号、地址、邮编、城市、国家、尾板、收件人、公司），不发电话和邮箱。
- AI 结果由调用方用规则校验（编号必须出自原文、街道不能含邮编等），不直接信任。
"""
import json
import os
import urllib.error
import urllib.request

DEFAULT_MODEL = "gemini-3.5-flash-lite"
API_URL = "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
TIMEOUT_S = 60
CHUNK = 40          # 一次请求最多多少行

FIELDS = ("编号", "地址", "邮编", "城市", "国家", "是否需要尾板", "收件人", "公司名称")

PROMPT = """你在帮物流公司整理客户填写的询价表。每一行是一个发往欧洲的订单，客户经常把不同信息混写在一个格子里。
对每一行输出：
- order_no：客户订单编号。通常是"编号"格开头的一串字母/数字（如 MR26101527），后面常跟着中文或德文备注。
  编号前面也可能有说明文字（如"客户编号 MR12345678 …"），这时取最像订单号的那串。
  只取编号本身，必须原样出自"编号"格，不要改写；分辨不出就留空，并在 issues 里说明。
- remark："编号"格里除编号以外的备注原文（收货时间、后备联系人等），没有就留空。
- street：只含街道名和门牌号，用于 Dachser 询价单的 Street 栏。
  从"地址"中去掉邮编、城市、国家、括号里的备注、"z.H./c/o/Attn"及其后的收件人或公司名。
  地址开头如果是收件人或公司名（与"收件人""公司名称"列相同或相近），即使没有 z.H. 也要去掉。
  保留原文的拼写和顺序（法国、比利时等门牌号在前也保持原样），不要翻译、不要补全、不要编造。
  地址里没有街道（只有地名）就原样保留地名。
- tail_lift："是否需要尾板"格的意思：yes / no，空或看不懂就填 unknown。
- issues：你不确定或发现矛盾的地方（如地址里的邮编和"邮编"列不一致、街道没有门牌号），用中文简短说明；没有就留空。
按输入的 row 原样返回 row。"""

SCHEMA = {
    "type": "ARRAY",
    "items": {
        "type": "OBJECT",
        "properties": {
            "row": {"type": "INTEGER"},
            "order_no": {"type": "STRING"},
            "remark": {"type": "STRING"},
            "street": {"type": "STRING"},
            "tail_lift": {"type": "STRING", "enum": ["yes", "no", "unknown"]},
            "issues": {"type": "STRING"},
        },
        "required": ["row", "order_no", "remark", "street", "tail_lift", "issues"],
    },
}


def _key():
    return os.environ.get("GEMINI_API_KEY", "").strip()


def model_name():
    return os.environ.get("XLFP_AI_MODEL", "").strip() or DEFAULT_MODEL


def available():
    return bool(_key())


def _cell(v):
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return "" if v is None else str(v).strip()


def _call(rows):
    body = {
        "systemInstruction": {"parts": [{"text": PROMPT}]},
        "contents": [{"role": "user", "parts": [{"text": json.dumps(rows, ensure_ascii=False)}]}],
        "generationConfig": {
            "temperature": 0,
            "responseMimeType": "application/json",
            "responseSchema": SCHEMA,
        },
    }
    req = urllib.request.Request(
        API_URL.format(model=model_name()),
        data=json.dumps(body, ensure_ascii=False).encode("utf-8"),
        headers={"Content-Type": "application/json", "x-goog-api-key": _key()},
    )
    with urllib.request.urlopen(req, timeout=TIMEOUT_S) as r:
        resp = json.load(r)
    cands = resp.get("candidates") or []
    if not cands:
        raise RuntimeError(f"AI 没有返回结果：{str(resp)[:200]}")
    text = "".join(p.get("text", "") for p in cands[0].get("content", {}).get("parts", []))
    return json.loads(text)


def clean_rows(rows):
    """rows: [{"row": Excel 行号, "编号": ..., "地址": ..., ...}]
    返回 {行号: {"order_no", "remark", "street", "tail_lift", "issues"}}；不可用或失败返回 None。"""
    if not available() or not rows:
        return None
    payload = [{"row": r["row"], **{k: _cell(r.get(k)) for k in FIELDS}} for r in rows]
    out = {}
    try:
        for i in range(0, len(payload), CHUNK):
            for it in _call(payload[i:i + CHUNK]):
                out[int(it["row"])] = it
    except urllib.error.HTTPError as e:
        print(f"  [AI] 调用失败（HTTP {e.code}）：{e.read()[:200].decode('utf-8', 'replace')}，改用规则识别")
        return None
    except Exception as e:
        print(f"  [AI] 调用失败（{str(e).splitlines()[0] if str(e) else repr(e)}），改用规则识别")
        return None
    return out
