/**
 * XL FreightPilot 服务端（Cloudflare Worker）：授权开关 + AI 整理中转 + 产品托盘参数（带审核）。
 *
 * Gemini Key 和提示词只放在这里，exe 里没有。同事不用授权码：程序每次运行前问 /check 总开关，
 * 关掉或连不上都不能用；请求头 X-XLFP-User 带 Windows 用户名（URL 编码），只用来显示谁提交了参数修改。
 *
 * 绑定（Worker → Settings → Variables / Bindings）：
 *   GEMINI_API_KEY  密钥（Secret）
 *   XLFP            KV 命名空间，里面的键：
 *     config        {"enabled": true, "message": "", "ai_daily": 500}  总开关（缺省开启，关掉后所有接口都拒绝）；
 *                   ai_daily = 全软件每天最多调用几次 AI，0 或不写 = 不限
 *     user:<授权码>  {"name": "刘旭", "role": "admin"}  可选：带这个码的程序是管理员（产品参数直接保存、审核申请）；
 *                   也可写 "enabled": false 停用、"daily": 100 单独限次。没带码 = 普通用户
 *     usage:<日期>:all / usage:<日期>:<码或 win:用户名>  程序自动写入的当天 AI 调用次数（2 天后自动过期）
 *     products      当前产品参数 {"version", "updated", "by", "note", "data": <pallet_params.json 的内容>}
 *     products:v:<版本号>  每个版本的快照（回滚用）
 *     proposal:<id> 修改申请（审完 90 天后自动删除）
 *   MODEL           可选变量，覆盖默认模型
 *   FEISHU_APP_ID   飞书应用 App ID（变量）        对账功能读飞书邮件用；不放进公开仓库和 exe
 *   FEISHU_APP_SECRET  飞书应用 App Secret（Secret）
 *   FEISHU_REDIRECT_URI  可选变量，默认 http://localhost:8080/callback（要和飞书后台「安全设置」一致）
 *   RELEASES        R2 存储桶（可选）：在线更新用，放 latest.json 和压缩包（见 README 在线更新）
 *
 * 接口（管理员带 Authorization: Bearer <授权码>，其他人不用带）：
 *   POST /check  → {"enabled": true/false, "message": "...", "name": "张三"}
 *   POST /clean  {"rows": [...]} → {"model": "...", "results": [...], "failed": 0}  询价整理；重试后仍失败的分块不在 results 里
 *   POST /street {"rows": [...]} → 同上                                              建单 Street 整理，Street 规则和询价完全相同
 *   POST /quote  {"rows": [{"row", "编号", "邮件主题", "邮件正文"}]} → {"model", "results": [{"row", "price", "evidence"}], "failed"}
 *                对账：规则从报价邮件表格里读不出价格时，让 AI 读邮件找这个编号的报价
 *   POST /aftersales/read  {"rows": [{"row", "邮件主题", "发件人", "最新内容", "全文"}]}
 *                → results: [{"row", "is_aftersales", "needs_customer", "issue", "summary", "my_action", "customer_msg", "german_terms"}]
 *   POST /aftersales/reply {"rows": [{"row", "Dachser 邮件主题", "Dachser 最新邮件", "客户回复", "我的补充"}]}
 *                → results: [{"row", "german", "chinese", "missing"}]
 *                售后助手（单独的项目，不在 exe 里）：只起草，不发邮件；每次调用计入 daily 次数
 *   飞书邮件授权（App Secret 只在这里；邮件内容不经过这里，程序直接用 user_access_token 读飞书）：
 *   POST /feishu/config  → {"app_id", "redirect_uri"}
 *   POST /feishu/token   {"code"} 或 {"refresh_token"} → 飞书 /authen/v2/oauth/token 的返回（access_token、expires_in …）
 *   在线更新（没绑定 RELEASES 时 /update/latest 返回 {"version": null}）：
 *   POST /update/latest   → latest.json 的内容 {"version", "file", "sha256", "size", "notes"}
 *   POST /update/download → latest.json 指向的压缩包（程序按 sha256 校验）
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
const DISABLED_MESSAGE = "服务暂停使用，请联系刘旭（xuliu00925@gmail.com）";

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

// 对账：从报价邮件里找某个编号的报价（我们报给客户的运费，欧元）
const QUOTE_FIELDS = ["编号", "邮件主题", "邮件正文"];
const QUOTE_LIMITS = { 邮件正文: 6000 };
const QUOTE_PROMPT = `你在帮物流公司对账。每一行给出一个订单编号（如 MR12345678）和一封报价邮件（主题和正文）。
邮件是我们发给客户的运费报价，正文里通常有一张表，"编号"列是订单编号，"运费报价(欧元)"列是报价（如"320欧"）。
正文后面可能引用了客户原来的询价邮件，那里的报价是空的，不要用引用部分。
对每一行输出：
- price：这封邮件里给这个编号的报价金额（欧元，只填数字，如 320 或 320.5）。
  写的是"待定"、"TBD"、空白、或者这封邮件里根本没有这个编号的报价，就填 null。
  一个编号对应多个金额且分不清哪个是总报价，也填 null。不要计算、不要猜。
- evidence：price 出自正文的原文片段（包含金额的那几个字，如"320欧"），price 为 null 时留空。
按输入的 row 原样返回 row。`;
const QUOTE_SCHEMA = {
  type: "ARRAY",
  items: {
    type: "OBJECT",
    properties: { row: { type: "INTEGER" }, price: { type: "NUMBER", nullable: true }, evidence: { type: "STRING" } },
    required: ["row", "price", "evidence"],
  },
};

// 售后助手（D:\aftersales-agent，跑在开发者自己的手机上）：读 Dachser 售后邮件，起草给客户的中文消息；
// 把客户的回复起草成给 Dachser 的德语邮件。只起草，发不发由人决定。
const AS_READ_FIELDS = ["邮件主题", "发件人", "最新内容", "全文"];
const AS_READ_LIMITS = { 最新内容: 6000, 全文: 12000 };
const AS_READ_PROMPT = `你在帮一家德国海外仓（我们）处理物流商 Dachser 发来的邮件。我们用 Dachser 给客户（中国卖家）发货，
Dachser 发现问题时会发 "Active Report"（如地址错误、收件人联系不上、拒收、货物损坏、危险品标记问题、额外费用），之后会在同一封邮件里来回回复。
每一行是一封新收到的邮件："最新内容"是这封邮件本身，"全文"还包含下面引用的历史邮件（只作背景）。只根据邮件内容回答，不要编造。
对每一行输出（除 german_terms 外都用中文）：
- is_aftersales：这是不是关于某票货的售后问题（Active Report 或它的回复、货物异常、需要我们或客户处理的事）。报价、账单汇总、广告、系统通知等填 false。
- needs_customer：要不要联系客户（需要客户提供信息、确认、做决定，或者客户应该知道的送货问题，如联系不上收件人、地址不对、改约送货、拒收、货损）。
  只涉及我们和 Dachser 之间的事（我们自己的标记或申报错误、对我们收费、Dachser 只是回复我们的问题）填 false。
- issue：问题类型，一句话（如"收件人联系不上""危险品标记不正确，收费 75 欧"）。
- summary：3 到 6 行：哪票货（客户单号、Dachser 运单号）、发生了什么、费用、Dachser 要我们做什么、期限。最新内容和历史不一致时以最新内容为准，并说明 Dachser 最新的态度。
- my_action：我们这边下一步要做什么，一两句话。
- customer_msg：needs_customer 为 true 时，写一条发给客户微信群的消息：开头写客户单号，说清楚问题、需要客户做什么、什么时候之前，语气礼貌简洁，像同事在群里说话，不要用"尊敬的"这类客套。
  不要提 Dachser 内部的流程和条款，不要写邮件里没有的信息（没有期限就不写期限）。needs_customer 为 false 时留空。
- german_terms：邮件里关键的德语原句（最多 3 句，原文照抄），方便核对；没有就留空。
按输入的 row 原样返回 row。`;
const AS_READ_SCHEMA = {
  type: "ARRAY",
  items: {
    type: "OBJECT",
    properties: {
      row: { type: "INTEGER" }, is_aftersales: { type: "BOOLEAN" }, needs_customer: { type: "BOOLEAN" },
      issue: { type: "STRING" }, summary: { type: "STRING" }, my_action: { type: "STRING" },
      customer_msg: { type: "STRING" }, german_terms: { type: "STRING" },
    },
    required: ["row", "is_aftersales", "needs_customer", "issue", "summary", "my_action", "customer_msg", "german_terms"],
  },
};
const AS_REPLY_FIELDS = ["Dachser 邮件主题", "Dachser 最新邮件", "客户回复", "我的补充"];
const AS_REPLY_LIMITS = { "Dachser 最新邮件": 8000, 客户回复: 4000, 我的补充: 2000 };
const AS_REPLY_PROMPT = `你在帮一家德国海外仓（我们）回复物流商 Dachser 的售后邮件。"Dachser 最新邮件"是 Dachser 发来的邮件（可能含引用的历史），
"客户回复"是我们的客户（中国卖家）在微信群里的回复（中文，可能很口语），"我的补充"是我们自己要加的话（可能为空）。
对每一行输出：
- german：给 Dachser 的德语回复邮件正文，以我们公司的口吻（不是客户的口吻），用 "Sehr geehrte Damen und Herren," 开头，用 "Mit freundlichen Grüßen" 结尾，不写签名和名字。
  只回答 Dachser 问的事，写进客户回复和我的补充里的信息（地址、电话、日期、决定），专有名词、号码、地址原样照抄，不要翻译成别的写法。
  不要编造客户没说的信息、不要替客户做承诺；Dachser 要的信息客户没给，就不要在德语里假装给了。语气专业、简短。
- chinese：german 的逐段中文对照。
- missing：Dachser 要求的信息里，客户回复和我的补充还没提供的（中文，一句一条）；都齐了就留空。
按输入的 row 原样返回 row。`;
const AS_REPLY_SCHEMA = {
  type: "ARRAY",
  items: {
    type: "OBJECT",
    properties: { row: { type: "INTEGER" }, german: { type: "STRING" }, chinese: { type: "STRING" }, missing: { type: "STRING" } },
    required: ["row", "german", "chinese", "missing"],
  },
};

const FEISHU_TOKEN_URL = "https://open.feishu.cn/open-apis/authen/v2/oauth/token";
const DEFAULT_REDIRECT = "http://localhost:8080/callback";

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

/** 程序发来的 Windows 用户名（URL 编码，头 X-XLFP-User），用来显示谁提交了参数修改。 */
function winUser(request) {
  let name = request.headers.get("X-XLFP-User") || "";
  try {
    name = decodeURIComponent(name);
  } catch {
    /* 原样用 */
  }
  return name.replace(/[\u0000-\u001f]/g, "").trim().slice(0, 40);
}

/** 开关 + 身份。返回 {user, token, config} 或一个错误 Response。
 *  同事不用授权码：没带码（或码不在 KV 里）= 普通用户，身份 = Windows 用户名。
 *  带了 KV 里有的码 = 那条 user 记录（"role": "admin" 才是管理员）。总开关关掉时所有接口都拒绝。 */
async function authorize(request, env) {
  const config = (await getJSON(env, "config")) || {};
  if (config.enabled === false) return json({ enabled: false, message: config.message || DISABLED_MESSAGE }, 403);
  const m = (request.headers.get("Authorization") || "").match(/^Bearer\s+(\S+)$/);
  const known = m ? await getJSON(env, `user:${m[1]}`) : null;
  if (known) {
    if (known.enabled === false) return json({ enabled: false, message: known.message || DISABLED_MESSAGE }, 403);
    return { user: known, token: m[1], config };
  }
  const name = winUser(request) || "同事";
  return { user: { name, role: "user" }, token: `win:${name}`, config };
}

const cell = (v, max = 500) => (v === null || v === undefined ? "" : String(v).trim().slice(0, max));

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

async function clean(request, env, { user, token, config }, { fields, prompt, schema, limits = {} }) {
  let body;
  try {
    body = await request.json();
  } catch {
    return json({ error: "请求格式不对" }, 400);
  }
  const rows = Array.isArray(body?.rows) ? body.rows : [];
  if (!rows.length || rows.length > MAX_ROWS) return json({ error: `行数必须在 1 到 ${MAX_ROWS} 之间` }, 400);

  // 每日次数限制（按请求计；一张表 / 一次对账通常是一次）：config.ai_daily = 全软件每天总共多少次，
  // user.daily = 某个授权码自己的上限；0 或不写 = 不限。超了返回 429，程序改用规则。
  const today = new Date().toISOString().slice(0, 10);
  const allKey = `usage:${today}:all`, userKey = `usage:${today}:${token}`;
  const [allUsed, userUsed] = (await Promise.all([env.XLFP.get(allKey), env.XLFP.get(userKey)]))
    .map((v) => parseInt(v || "0", 10));
  if (config.ai_daily > 0 && allUsed >= config.ai_daily)
    return json({ error: `今天的 AI 次数已用完（全软件每天 ${config.ai_daily} 次），这次改用规则处理` }, 429);
  if (user.daily > 0 && userUsed >= user.daily)
    return json({ error: `今天的 AI 次数已用完（${user.daily} 次），这次改用规则处理` }, 429);
  await Promise.all([env.XLFP.put(allKey, String(allUsed + 1), { expirationTtl: 2 * 86400 }),
                     env.XLFP.put(userKey, String(userUsed + 1), { expirationTtl: 2 * 86400 })]);

  const payload = rows.map((r) => ({ row: parseInt(r.row, 10),
                                     ...Object.fromEntries(fields.map((k) => [k, cell(r[k], limits[k])])) }));
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

// ------------------------------------------------------------------ 飞书邮件授权（只换 token，不经手邮件内容）

async function feishu(path, request, env) {
  if (!env.FEISHU_APP_ID || !env.FEISHU_APP_SECRET) return json({ error: "服务器还没有配置飞书应用" }, 503);
  const redirect_uri = env.FEISHU_REDIRECT_URI || DEFAULT_REDIRECT;
  if (path === "/feishu/config") return json({ app_id: env.FEISHU_APP_ID, redirect_uri });
  const b = await readBody(request);
  const base = { client_id: env.FEISHU_APP_ID, client_secret: env.FEISHU_APP_SECRET };
  let body;
  if (typeof b.code === "string" && b.code.trim()) body = { ...base, grant_type: "authorization_code", code: b.code.trim(), redirect_uri };
  else if (typeof b.refresh_token === "string" && b.refresh_token.trim())
    body = { ...base, grant_type: "refresh_token", refresh_token: b.refresh_token.trim() };
  else return json({ error: "缺少 code 或 refresh_token" }, 400);
  const res = await fetch(FEISHU_TOKEN_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json; charset=utf-8" },
    body: JSON.stringify(body),
  });
  let data;
  try {
    data = await res.json();
  } catch {
    return json({ error: `飞书返回 HTTP ${res.status}` }, 502);
  }
  return json(data, res.ok && data.access_token ? 200 : 400);
}

// ------------------------------------------------------------------ 在线更新
// R2 存储桶（绑定名 RELEASES）里放：latest.json（build.py 生成）和它指向的压缩包。
// 总开关开着时程序才能查到和下载，地址固定，不怕共享服务器的临时链接变化。
async function update(path, env) {
  if (!env.RELEASES) return json({ version: null, message: "服务器还没有配置在线更新" });
  const obj = await env.RELEASES.get("latest.json");
  if (!obj) return json({ version: null, message: "服务器上还没有发布版本" });
  let latest;
  try {
    latest = JSON.parse(await obj.text());
  } catch {
    return json({ error: "latest.json 格式不对" }, 500);
  }
  if (path === "/update/latest") return json(latest);
  const file = latest.file ? await env.RELEASES.get(latest.file) : null;
  if (!file) return json({ error: `服务器上找不到更新包 ${latest.file || ""}` }, 404);
  return new Response(file.body, {
    headers: { "Content-Type": "application/zip", "X-Version": String(latest.version || "") },
  });
}

const UPDATE_ROUTES = ["/update/latest", "/update/download"];
const PRODUCT_ROUTES = ["/products/get", "/products/save", "/products/history", "/products/rollback",
                        "/proposals/submit", "/proposals/list", "/proposals/approve", "/proposals/reject",
                        "/proposals/withdraw"];
const FEISHU_ROUTES = ["/feishu/config", "/feishu/token"];

export default {
  async fetch(request, env) {
    const path = new URL(request.url).pathname;
    if (request.method !== "POST" ||
        !["/check", "/clean", "/street", "/quote", "/aftersales/read", "/aftersales/reply", ...FEISHU_ROUTES, ...UPDATE_ROUTES, ...PRODUCT_ROUTES].includes(path))
      return json({ error: "not found" }, 404);
    const auth = await authorize(request, env);
    if (auth instanceof Response) return auth;
    if (path === "/check") return json({ enabled: true, message: "", name: auth.user.name || "" });
    if (path === "/clean") return clean(request, env, auth, { fields: FIELDS, prompt: PROMPT, schema: SCHEMA });
    if (path === "/street") return clean(request, env, auth, { fields: STREET_FIELDS, prompt: STREET_PROMPT, schema: STREET_SCHEMA });
    if (path === "/quote")
      return clean(request, env, auth, { fields: QUOTE_FIELDS, prompt: QUOTE_PROMPT, schema: QUOTE_SCHEMA, limits: QUOTE_LIMITS });
    if (path === "/aftersales/read")
      return clean(request, env, auth, { fields: AS_READ_FIELDS, prompt: AS_READ_PROMPT, schema: AS_READ_SCHEMA, limits: AS_READ_LIMITS });
    if (path === "/aftersales/reply")
      return clean(request, env, auth, { fields: AS_REPLY_FIELDS, prompt: AS_REPLY_PROMPT, schema: AS_REPLY_SCHEMA, limits: AS_REPLY_LIMITS });
    if (FEISHU_ROUTES.includes(path)) return feishu(path, request, env);
    if (UPDATE_ROUTES.includes(path)) return update(path, env);
    return products(path, request, env, auth);
  },
};
