/**
 * XL FreightPilot 服务端（Cloudflare Worker）：授权开关 + AI 整理中转 + 产品托盘参数（带审核）。
 *
 * Gemini Key 和提示词只放在这里，exe 里没有；exe 用每个同事自己的授权码调用。
 *
 * 绑定（Worker → Settings → Variables / Bindings）：
 *   GEMINI_API_KEY  密钥（Secret）
 *   XLFP            KV 命名空间，里面的键：
 *     config        {"enabled": true, "message": ""}            总开关，缺省为开启
 *     user:<授权码>  {"name": "张三", "enabled": true, "daily": 100}  每个同事一条；daily = 每天最多整理几张表，0 或不写 = 不限
 *                   加 "role": "admin" = 管理员（产品参数直接保存、审核别人的申请）
 *     usage:<日期>:<授权码>  程序自动写入的当天调用次数（2 天后自动过期）
 *     products      当前产品参数 {"version", "updated", "by", "note", "data": <pallet_params.json 的内容>}
 *     products:v:<版本号>  每个版本的快照（回滚用）
 *     proposal:<id> 修改申请（审完 90 天后自动删除）
 *   MODEL           可选变量，覆盖默认模型
 *
 * 接口（都要带 Authorization: Bearer <授权码>）：
 *   POST /check  → {"enabled": true/false, "message": "...", "name": "张三"}
 *   POST /clean  {"rows": [...]} → {"model": "...", "results": [...], "failed": 0}  询价整理；重试后仍失败的分块不在 results 里
 *   POST /street {"rows": [...]} → 同上                                              建单 Street 整理，Street 规则和询价完全相同
 *   产品参数（校验规则在程序里 pallet_plan.validate_product，这里只做格式检查、权限和版本）：
 *   POST /products/get       → {"version", "updated", "by", "data", "role", "pending"}  data 为 null = 还没上传过
 *   POST /products/save      {"base_version", "data", "note"}  管理员直接保存；版本对不上返回 409
 *   POST /products/history   → {"versions": [{"version", "updated", "by", "note"}]}  管理员
 *   POST /products/rollback  {"base_version", "version"}  管理员；把旧版本另存为新版本
 *   POST /proposals/submit   {"base_version", "action": "upsert"|"delete"|"defaults", "sku", "old_sku", "data", "note"}
 *   POST /proposals/list     → {"items": [...]}  管理员看全部，其他人只看自己的
 *   POST /proposals/approve  {"id"}  管理员 → {"version", "data"}
 *   POST /proposals/reject   {"id", "reason"}  管理员
 *   POST /proposals/withdraw {"id"}  提交人撤回自己待审核的申请
 */

const DEFAULT_MODEL = "gemini-3.5-flash-lite";
const API_URL = "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent";
const CHUNK = 5;             // 一次请求 Gemini 最多多少行（拆小了并行发，整张表更快）
const PARALLEL = 10;         // 同时最多几个 Gemini 请求
const MAX_ROWS = 400;        // 一次最多接收多少行，防止滥用
const RETRIES = 2;           // 一块失败后再试几次
const DISABLED_MESSAGE = "该软件已由开发者关闭使用，如需使用请联系开发者：xuliu00925@gmail.com";

const STREET_FIELDS = ["地址", "House No", "Address2", "邮编", "城市", "国家", "收件人", "公司名称"];
const FIELDS = ["编号", "是否需要尾板", ...STREET_FIELDS];

// 询价和建单共用同一份 Street 规则
const STREET_SPEC = `- street：只含街道名和门牌号，用于 Dachser 的 Street 栏（询价单和运输单规则相同）。
  街道信息主要在"地址"。"House No" 有值时就是门牌号，必须写进 street（已经在地址里就不要重复），放在该国习惯的位置，地址里已有的顺序保持不变。
  "Address2" 通常是楼栋、单元、公司或备注，不要写进 street；只有地址里没有街道或门牌号、而 Address2 里有时才从中取街道或门牌号。这两列为空就忽略。
  去掉邮编及其后的城市、国家，去掉括号里的备注、"z.H./c/o/Attn"及其后的收件人或公司名。
  街道前面的收件人、公司名、仓库名（与"收件人""公司名称"列相同或相近，或含 GmbH、Ltd、s.r.o.、Warehouse 等）即使没有 z.H. 也要去掉；
  街道后面的区名、村名、楼栋单元也要去掉。
  保留原文的拼写和顺序（法国、比利时等门牌号在前也保持原样，12 bis 这种写法完整保留），不要翻译、不要补全、不要编造。
  街名本身的数字和地名要保留（如 Rue du 8 Mai 1945、Rue de Paris）。地址里没有街道（只有地名）就原样保留地名。`;

const ISSUE_RULES = `只写需要人工核对的问题，用中文简短说明：地址里的邮编和"邮编"列确实不同（前导 0 不同不算）、街道没有门牌号、House No 和地址里的门牌号矛盾、分不清哪段是街道。
  地址里没有邮编、城市、国家都很正常，不是问题；提取顺利就留空，不要描述你做了什么。`;

const PROMPT = `你在帮物流公司整理客户填写的询价表。每一行是一个发往欧洲的订单，客户经常把不同信息混写在一个格子里。
对每一行输出：
- order_no：客户订单编号。通常是"编号"格开头的一串字母/数字（如 MR12345678），后面常跟着中文或德文备注。
  编号前面也可能有说明文字（如"客户编号 MR87654321 …"），这时取最像订单号的那串。
  只取编号本身，必须原样出自"编号"格，不要改写；分辨不出就留空，并在 issues 里说明。
- remark："编号"格里除编号以外的备注原文（收货时间、后备联系人等），没有就留空。
${STREET_SPEC}
- tail_lift："是否需要尾板"格的意思：yes / no，空或看不懂就填 unknown。
- issues：${ISSUE_RULES}
按输入的 row 原样返回 row。`;

const STREET_PROMPT = `你在帮物流公司整理收货地址，每一行是一个发往欧洲的订单。输出：
${STREET_SPEC}
- issues：${ISSUE_RULES}
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

const STREET_SCHEMA = {
  type: "ARRAY",
  items: {
    type: "OBJECT",
    properties: { row: { type: "INTEGER" }, street: { type: "STRING" }, issues: { type: "STRING" } },
    required: ["row", "street", "issues"],
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

async function callGemini(env, model, rows, prompt, schema) {
  const res = await fetch(API_URL.replace("{model}", model), {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-goog-api-key": env.GEMINI_API_KEY },
    body: JSON.stringify({
      systemInstruction: { parts: [{ text: prompt }] },
      contents: [{ role: "user", parts: [{ text: JSON.stringify(rows) }] }],
      generationConfig: { temperature: 0, responseMimeType: "application/json", responseSchema: schema },
    }),
  });
  if (!res.ok) throw new Error(`Gemini HTTP ${res.status}: ${(await res.text()).slice(0, 200)}`);
  const data = await res.json();
  const parts = data.candidates?.[0]?.content?.parts || [];
  if (!parts.length) throw new Error(`Gemini 没有返回结果: ${JSON.stringify(data).slice(0, 200)}`);
  return JSON.parse(parts.map((p) => p.text || "").join(""));
}

/** 一块失败时重试；还是失败返回 null，只影响这几行（程序对这几行改用规则）。 */
async function callWithRetry(env, model, rows, prompt, schema) {
  for (let attempt = 0; ; attempt++) {
    try {
      const out = await callGemini(env, model, rows, prompt, schema);
      if (!Array.isArray(out)) throw new Error("Gemini 返回的不是数组");
      return out;
    } catch (e) {
      if (attempt >= RETRIES) {
        console.error(String(e.message || e));
        return null;
      }
      await new Promise((r) => setTimeout(r, 1000 * (attempt + 1)));
    }
  }
}

async function clean(request, env, { user, token }, { fields, prompt, schema }) {
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

  const payload = rows.map((r) => ({ row: parseInt(r.row, 10), ...Object.fromEntries(fields.map((k) => [k, cell(r[k])])) }));
  const model = env.MODEL || DEFAULT_MODEL;
  const chunks = [];
  for (let i = 0; i < payload.length; i += CHUNK) chunks.push(payload.slice(i, i + CHUNK));
  const parts = new Array(chunks.length);
  let next = 0;
  const worker = async () => {
    while (next < chunks.length) {
      const i = next++;
      parts[i] = await callWithRetry(env, model, chunks[i], prompt, schema);
    }
  };
  await Promise.all(Array.from({ length: Math.min(PARALLEL, chunks.length) }, worker));
  const ok = parts.filter(Boolean);
  if (!ok.length) return json({ error: "Gemini 调用失败，稍后再试" }, 502);
  return json({ model, results: ok.flat(), failed: parts.length - ok.length });
}

// ------------------------------------------------------------------ 产品参数 + 审核

const MAX_DATA = 200_000;        // 整个参数文件最大字节数
const MAX_PRODUCT = 5_000;       // 单个产品最大字节数
const DONE_TTL = 90 * 86400;     // 审完的申请保留 90 天
const isAdmin = (user) => user.role === "admin";
const isObj = (v) => v !== null && typeof v === "object" && !Array.isArray(v);
const now = () => new Date().toISOString();
const who = (user) => user.name || "未命名";

async function readBody(request) {
  try {
    const b = await request.json();
    return isObj(b) ? b : {};
  } catch {
    return {};
  }
}

const getProducts = async (env) => (await getJSON(env, "products")) || { version: 0, data: null };

async function putProducts(env, cur, data, user, note) {
  const rec = { version: (cur.version || 0) + 1, updated: now(), by: who(user), note: String(note || "").slice(0, 300), data };
  const text = JSON.stringify(rec);
  await env.XLFP.put(`products:v:${rec.version}`, text);
  await env.XLFP.put("products", text);
  return rec;
}

const badSku = (sku) => typeof sku !== "string" || !sku.trim() || sku.trim().startsWith("_") || sku.length > 100;

/** 把申请应用到参数上，返回新的 data；冲突时抛 Error（文字给人看）。 */
function applyProposal(data, p) {
  const out = {};
  if (p.action === "defaults") {
    const old = isObj(data._defaults) ? data._defaults : {};
    const keep = Object.fromEntries(Object.entries(old).filter(([k]) => k.startsWith("_")));   // 保留 _说明
    for (const [k, v] of Object.entries(data)) out[k] = k === "_defaults" ? { ...keep, ...p.data } : v;
    if (!("_defaults" in out)) out._defaults = { ...keep, ...p.data };
    return out;
  }
  if (p.action === "delete") {
    if (!(p.sku in data)) throw new Error(`${p.sku} 已经不存在了`);
    for (const [k, v] of Object.entries(data)) if (k !== p.sku) out[k] = v;
    return out;
  }
  // upsert：改名时保持原来的位置
  const from = p.old_sku || null;
  if (from && !(from in data)) throw new Error(`原产品 ${from} 已经不存在了（可能被删除或改名）`);
  if (!from && p.sku in data) throw new Error(`${p.sku} 已经存在（申请提交后被别人加上了）`);
  if (from && from !== p.sku && p.sku in data) throw new Error(`${p.sku} 已经存在，不能改成这个名字`);
  for (const [k, v] of Object.entries(data)) out[k === from ? p.sku : k] = k === from ? p.data : v;
  if (!from) out[p.sku] = p.data;
  return out;
}

async function listProposalKeys(env) {
  const keys = [];
  let cursor;
  do {
    const r = await env.XLFP.list({ prefix: "proposal:", cursor });
    keys.push(...r.keys);
    cursor = r.list_complete ? undefined : r.cursor;
  } while (cursor);
  return keys;
}

const publicProposal = ({ token, ...rest }) => rest;     // 不把别人的授权码发出去

async function products(path, request, env, { user, token }) {
  const b = await readBody(request);
  const admin = isAdmin(user);
  const cur = await getProducts(env);

  if (path === "/products/get") {
    const pending = admin ? (await listProposalKeys(env)).filter((k) => k.metadata?.status === "pending").length : 0;
    return json({ version: cur.version, updated: cur.updated || "", by: cur.by || "", data: cur.data,
                  role: admin ? "admin" : "user", pending });
  }

  if (path === "/products/save" || path === "/products/rollback") {
    if (!admin) return json({ error: "只有管理员可以直接保存" }, 403);
    if (b.base_version !== cur.version)
      return json({ error: `服务器上的参数已经更新到第 ${cur.version} 版，请刷新后再改`, version: cur.version }, 409);
    let data = b.data, note = b.note;
    if (path === "/products/rollback") {
      const old = await getJSON(env, `products:v:${parseInt(b.version, 10)}`);
      if (!old) return json({ error: "找不到这个版本" }, 404);
      data = old.data;
      note = `回滚到第 ${old.version} 版`;
    }
    if (!isObj(data) || JSON.stringify(data).length > MAX_DATA) return json({ error: "参数格式不对" }, 400);
    const rec = await putProducts(env, cur, data, user, note);
    return json({ version: rec.version, data: rec.data });
  }

  if (path === "/products/history") {
    if (!admin) return json({ error: "只有管理员可以查看" }, 403);
    const nums = [];
    for (let v = cur.version; v >= 1 && nums.length < 50; v--) nums.push(v);
    const recs = await Promise.all(nums.map((v) => getJSON(env, `products:v:${v}`)));
    return json({ versions: recs.filter(Boolean).map((r) => ({ version: r.version, updated: r.updated, by: r.by, note: r.note })) });
  }

  if (path === "/proposals/submit") {
    if (cur.data === null) return json({ error: "服务器上还没有产品参数，请等管理员先上传" }, 409);
    const action = b.action;
    if (!["upsert", "delete", "defaults"].includes(action)) return json({ error: "申请类型不对" }, 400);
    if (action !== "defaults" && badSku(b.sku)) return json({ error: "产品物料编码不对" }, 400);
    if (action !== "delete" && (!isObj(b.data) || JSON.stringify(b.data).length > MAX_PRODUCT))
      return json({ error: "参数格式不对" }, 400);
    if (b.old_sku != null && badSku(b.old_sku)) return json({ error: "原产品编码不对" }, 400);
    const created = now();
    const id = `${Date.now().toString(36).padStart(9, "0")}-${crypto.randomUUID().slice(0, 8)}`;
    const p = {
      id, created, token, name: who(user), action, status: "pending",
      sku: action === "defaults" ? "" : b.sku.trim(), old_sku: b.old_sku || null,
      data: action === "delete" ? null : b.data, base_version: b.base_version ?? null,
      note: String(b.note || "").slice(0, 300),
    };
    await env.XLFP.put(`proposal:${id}`, JSON.stringify(p), { metadata: { status: "pending", token, created } });
    return json({ id });
  }

  if (path === "/proposals/list") {
    const keys = (await listProposalKeys(env)).filter((k) => admin || k.metadata?.token === token);
    const pending = keys.filter((k) => k.metadata?.status === "pending");
    const done = keys.filter((k) => k.metadata?.status !== "pending")
      .sort((x, y) => (x.name < y.name ? 1 : -1)).slice(0, 50);          // id 以时间开头，按名字倒序 = 最新在前
    const items = (await Promise.all([...pending, ...done].map((k) => getJSON(env, k.name)))).filter(Boolean);
    items.sort((x, y) => (y.status === "pending") - (x.status === "pending") || (x.created < y.created ? 1 : -1));
    return json({ items: items.map(publicProposal) });
  }

  // approve / reject / withdraw
  const key = `proposal:${String(b.id || "")}`;
  const p = await getJSON(env, key);
  if (!p) return json({ error: "找不到这条申请" }, 404);
  if (p.status !== "pending") return json({ error: "这条申请已经处理过了" }, 409);
  const finish = async (status, extra = {}) => {
    Object.assign(p, { status, reviewed: now(), reviewer: who(user) }, extra);
    await env.XLFP.put(key, JSON.stringify(p),
                       { metadata: { status, token: p.token, created: p.created }, expirationTtl: DONE_TTL });
  };

  if (path === "/proposals/withdraw") {
    if (p.token !== token) return json({ error: "只能撤回自己的申请" }, 403);
    await finish("withdrawn");
    return json({ ok: true });
  }
  if (!admin) return json({ error: "只有管理员可以审核" }, 403);
  if (path === "/proposals/reject") {
    await finish("rejected", { reason: String(b.reason || "").slice(0, 300) });
    return json({ ok: true });
  }
  // approve
  let data;
  try {
    data = applyProposal(cur.data || {}, p);
  } catch (e) {
    return json({ error: String(e.message || e) }, 409);
  }
  const label = p.action === "defaults" ? "通用参数" : p.sku;
  const verb = { upsert: p.old_sku ? "修改" : "新增", delete: "删除", defaults: "修改" }[p.action];
  const rec = await putProducts(env, cur, data, user, `审核通过 ${p.name} 的申请：${verb} ${label}`);
  await finish("approved", { applied_version: rec.version });
  return json({ version: rec.version, data: rec.data });
}

const PRODUCT_ROUTES = ["/products/get", "/products/save", "/products/history", "/products/rollback",
                        "/proposals/submit", "/proposals/list", "/proposals/approve", "/proposals/reject",
                        "/proposals/withdraw"];

export default {
  async fetch(request, env) {
    const path = new URL(request.url).pathname;
    if (request.method !== "POST" || !["/check", "/clean", "/street", ...PRODUCT_ROUTES].includes(path))
      return json({ error: "not found" }, 404);
    const auth = await authorize(request, env);
    if (auth instanceof Response) return auth;
    if (path === "/check") return json({ enabled: true, message: "", name: auth.user.name || "" });
    if (path === "/clean") return clean(request, env, auth, { fields: FIELDS, prompt: PROMPT, schema: SCHEMA });
    if (path === "/street") return clean(request, env, auth, { fields: STREET_FIELDS, prompt: STREET_PROMPT, schema: STREET_SCHEMA });
    return products(path, request, env, auth);
  },
};
