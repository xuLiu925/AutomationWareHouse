/**
 * XL FreightPilot 服务端（Cloudflare Worker）：授权开关 + AI 整理中转。
 *
 * Gemini Key 和提示词只放在这里，exe 里没有；exe 用每个同事自己的授权码调用。
 *
 * 绑定（Worker → Settings → Variables / Bindings）：
 *   GEMINI_API_KEY  密钥（Secret）
 *   XLFP            KV 命名空间，里面的键：
 *     config        {"enabled": true, "message": ""}            总开关，缺省为开启
 *     user:<授权码>  {"name": "张三", "enabled": true, "daily": 100}  每个同事一条；daily = 每天最多整理几张表，0 或不写 = 不限
 *     usage:<日期>:<授权码>  程序自动写入的当天调用次数（2 天后自动过期）
 *   MODEL           可选变量，覆盖默认模型
 *
 * 接口（都要带 Authorization: Bearer <授权码>）：
 *   POST /check  → {"enabled": true/false, "message": "...", "name": "张三"}
 *   POST /clean  {"rows": [...]} → {"model": "...", "results": [...]}
 */

const DEFAULT_MODEL = "gemini-3.5-flash-lite";
const API_URL = "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent";
const CHUNK = 5;             // 一次请求 Gemini 最多多少行（拆小了并行发，整张表更快）
const PARALLEL = 10;         // 同时最多几个 Gemini 请求
const MAX_ROWS = 400;        // 一次最多接收多少行，防止滥用
const DISABLED_MESSAGE = "该软件已由开发者关闭使用，如需使用请联系开发者：xuliu00925@gmail.com";

const FIELDS = ["编号", "地址", "邮编", "城市", "国家", "是否需要尾板", "收件人", "公司名称"];

const PROMPT = `你在帮物流公司整理客户填写的询价表。每一行是一个发往欧洲的订单，客户经常把不同信息混写在一个格子里。
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
按输入的 row 原样返回 row。`;

const SCHEMA = {
  type: "ARRAY",
  items: {
    type: "OBJECT",
    properties: {
      row: { type: "INTEGER" },
      order_no: { type: "STRING" },
      remark: { type: "STRING" },
      street: { type: "STRING" },
      tail_lift: { type: "STRING", enum: ["yes", "no", "unknown"] },
      issues: { type: "STRING" },
    },
    required: ["row", "order_no", "remark", "street", "tail_lift", "issues"],
  },
};

const json = (data, status = 200) =>
  new Response(JSON.stringify(data), {
    status,
    headers: { "Content-Type": "application/json; charset=utf-8" },
  });

async function getJSON(env, key) {
  const v = await env.XLFP.get(key);
  if (!v) return null;
  try {
    return JSON.parse(v);
  } catch {
    return null;
  }
}

/** 校验授权码和开关。返回 {user, token} 或一个错误 Response。 */
async function authorize(request, env) {
  const m = (request.headers.get("Authorization") || "").match(/^Bearer\s+(\S+)$/);
  const token = m ? m[1] : "";
  const user = token ? await getJSON(env, `user:${token}`) : null;
  if (!user) return json({ enabled: false, message: "授权码无效，请联系开发者获取：xuliu00925@gmail.com" }, 401);
  const config = (await getJSON(env, "config")) || {};
  if (config.enabled === false) return json({ enabled: false, message: config.message || DISABLED_MESSAGE }, 403);
  if (user.enabled === false) return json({ enabled: false, message: user.message || DISABLED_MESSAGE }, 403);
  return { user, token };
}

const cell = (v) => (v === null || v === undefined ? "" : String(v).trim().slice(0, 500));

async function callGemini(env, model, rows) {
  const res = await fetch(API_URL.replace("{model}", model), {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-goog-api-key": env.GEMINI_API_KEY },
    body: JSON.stringify({
      systemInstruction: { parts: [{ text: PROMPT }] },
      contents: [{ role: "user", parts: [{ text: JSON.stringify(rows) }] }],
      generationConfig: { temperature: 0, responseMimeType: "application/json", responseSchema: SCHEMA },
    }),
  });
  if (!res.ok) throw new Error(`Gemini HTTP ${res.status}: ${(await res.text()).slice(0, 200)}`);
  const data = await res.json();
  const parts = data.candidates?.[0]?.content?.parts || [];
  if (!parts.length) throw new Error(`Gemini 没有返回结果: ${JSON.stringify(data).slice(0, 200)}`);
  return JSON.parse(parts.map((p) => p.text || "").join(""));
}

async function clean(request, env, { user, token }) {
  let body;
  try {
    body = await request.json();
  } catch {
    return json({ error: "请求格式不对" }, 400);
  }
  const rows = Array.isArray(body?.rows) ? body.rows : [];
  if (!rows.length || rows.length > MAX_ROWS) return json({ error: `行数必须在 1 到 ${MAX_ROWS} 之间` }, 400);

  // 每日次数限制（按表计）
  const usageKey = `usage:${new Date().toISOString().slice(0, 10)}:${token}`;
  const used = parseInt((await env.XLFP.get(usageKey)) || "0", 10);
  if (user.daily > 0 && used >= user.daily) {
    return json({ error: `今天的 AI 整理次数已用完（${user.daily} 次），明天再试或联系开发者` }, 429);
  }
  await env.XLFP.put(usageKey, String(used + 1), { expirationTtl: 2 * 86400 });

  const payload = rows.map((r) => ({ row: parseInt(r.row, 10), ...Object.fromEntries(FIELDS.map((k) => [k, cell(r[k])])) }));
  const model = env.MODEL || DEFAULT_MODEL;
  const chunks = [];
  for (let i = 0; i < payload.length; i += CHUNK) chunks.push(payload.slice(i, i + CHUNK));
  const parts = new Array(chunks.length);
  let next = 0;
  const worker = async () => {
    while (next < chunks.length) {
      const i = next++;
      parts[i] = await callGemini(env, model, chunks[i]);
    }
  };
  try {
    await Promise.all(Array.from({ length: Math.min(PARALLEL, chunks.length) }, worker));
  } catch (e) {
    return json({ error: String(e.message || e) }, 502);
  }
  return json({ model, results: parts.flat() });
}

export default {
  async fetch(request, env) {
    const path = new URL(request.url).pathname;
    if (request.method !== "POST" || !["/check", "/clean"].includes(path)) return json({ error: "not found" }, 404);
    const auth = await authorize(request, env);
    if (auth instanceof Response) return auth;
    if (path === "/check") return json({ enabled: true, message: "", name: auth.user.name || "" });
    return clean(request, env, auth);
  },
};
