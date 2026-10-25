// ==UserScript==
// @name         XL FreightPilot · WMS 序列号助手
// @namespace    https://xlfp.xuliu00925.workers.dev/
// @version      1.0
// @description  领星 WMS 复核/验货：在批量序列号框里扫二维码时，把逗号自动换成换行，并去掉重复的序列号。Developed by Xu Liu（刘旭）· xuliu00925@gmail.com
// @author       Xu Liu（刘旭）
// @match        https://wms.xlwms.com/*
// @grant        GM_xmlhttpRequest
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_registerMenuCommand
// @connect      xlfp.xuliu00925.workers.dev
// @run-at       document-idle
// ==/UserScript==

/*
 * 授权：和 XL FreightPilot 用同一个授权码（Tampermonkey 图标 → "设置授权码"），
 * 用之前联网问开发者的服务器能不能用；授权码无效、被停用或连不上时不做转换，
 * WMS 本身照常可以手动使用。脚本只改序列号输入框里的文字，不调用 WMS 的接口。
 */
(function () {
  "use strict";

  const SERVER = "https://xlfp.xuliu00925.workers.dev";
  const CONTACT = "xuliu00925@gmail.com";
  const RECHECK_MS = 60 * 60 * 1000;   // 授权通过后 1 小时内不再重复验证
  const TEXTAREA = ".xl-input-batch-popover textarea";   // 勾选"支持批量复核"后序列号的多行框
  const SEP = /[,，]/;

  const DISABLED = `该软件已由开发者关闭使用，如需使用请联系开发者：${CONTACT}`;
  const NO_TOKEN = `还没有填写授权码：点浏览器右上角 Tampermonkey 图标 →"设置授权码"。没有授权码请联系开发者：${CONTACT}`;
  const OFFLINE = `无法连接授权服务器，暂时不能自动转换逗号，请检查网络。如一直无法使用，请联系开发者：${CONTACT}`;

  let auth = { ok: false, msg: "", at: 0 };
  let checking = null;
  let warned = null;   // 当前这个框已经提示过授权问题，不再每个字都提示

  function check() {
    const token = String(GM_getValue("token", "")).trim();
    if (!token) return Promise.resolve({ ok: false, msg: NO_TOKEN });
    return new Promise((resolve) => {
      const fail = () => resolve({ ok: false, msg: OFFLINE });
      GM_xmlhttpRequest({
        method: "POST",
        url: SERVER + "/check",
        headers: { "Content-Type": "application/json", Authorization: "Bearer " + token },
        data: "{}",
        timeout: 10000,
        onload(r) {
          let d = {};
          try { d = JSON.parse(r.responseText) || {}; } catch (e) { /* 非 JSON 当作连不上 */ }
          if (r.status === 200 && d.enabled === true) resolve({ ok: true, msg: "" });
          else if (r.status === 401 || r.status === 403) resolve({ ok: false, msg: d.message || DISABLED });
          else fail();
        },
        onerror: fail,
        ontimeout: fail,
      });
    });
  }

  function ensureAuth() {
    if (auth.ok && Date.now() - auth.at < RECHECK_MS) return Promise.resolve(auth);
    if (!checking) {
      checking = check().then((a) => {
        auth = { ...a, at: Date.now() };
        checking = null;
        return auth;
      });
    }
    return checking;
  }

  function toast(msg, error) {
    const el = document.createElement("div");
    el.textContent = msg;
    el.style.cssText = "position:fixed;top:16px;left:50%;transform:translateX(-50%);z-index:2147483647;" +
      "max-width:600px;padding:10px 16px;border-radius:4px;font-size:14px;line-height:1.5;" +
      "box-shadow:0 2px 12px rgba(0,0,0,.2);color:#fff;background:" + (error ? "#e6492d" : "#3f8cff");
    document.body.appendChild(el);
    setTimeout(() => el.remove(), error ? 8000 : 4000);
  }

  /** 逗号换成换行、去掉空格和空行、已扫完的序列号去重；最后一个还在扫的不动。没变化返回 null。 */
  function normalize(text) {
    const parts = text.split(/[,，\r\n]+/);
    const tail = parts.pop().trim();          // 结尾是分隔符时为空串
    const seen = new Set(), out = [], dups = [];
    for (const p of parts.map((s) => s.trim()).filter(Boolean)) {
      if (seen.has(p)) dups.push(p);
      else { seen.add(p); out.push(p); }
    }
    const value = out.map((s) => s + "\n").join("") + tail;
    return value === text ? null : { value, dups };
  }

  const setValue = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, "value").set;
  let busy = false;

  function apply(ta) {
    const text = ta.value;
    const atEnd = ta.selectionStart === text.length;
    // 有逗号就整理；没逗号时只在末尾刚扫完一行时整理（去重），不打扰在中间手动改
    if (!SEP.test(text) && !(atEnd && text.endsWith("\n"))) return;
    const r = normalize(text);
    if (!r) return;
    busy = true;
    try {
      setValue.call(ta, r.value);
      ta.selectionStart = ta.selectionEnd = r.value.length;
      ta.dispatchEvent(new Event("input", { bubbles: true }));   // 让 WMS 页面收到新内容
    } finally {
      busy = false;
    }
    if (r.dups.length) toast(`已去掉重复的序列号：${r.dups.join("、")}`);
  }

  function onInput(ev) {
    const ta = ev.target;
    if (busy || !(ta instanceof HTMLTextAreaElement) || !ta.matches(TEXTAREA)) return;
    if (auth.ok && Date.now() - auth.at < RECHECK_MS) return apply(ta);
    if (!SEP.test(ta.value)) return;
    ensureAuth().then((a) => {
      if (a.ok) apply(ta);   // 验证期间扫进来的内容也一起整理
      else if (warned !== ta) { warned = ta; toast(a.msg, true); }
    });
  }

  document.addEventListener("input", onInput, true);
  document.addEventListener("focusin", (ev) => {
    if (ev.target instanceof HTMLTextAreaElement && ev.target.matches(TEXTAREA)) {
      warned = null;
      ensureAuth();   // 打开多行框时提前验证，扫码时不用等
    }
  }, true);

  GM_registerMenuCommand("设置授权码", () => {
    const token = prompt("请输入 XL FreightPilot 授权码：", GM_getValue("token", ""));
    if (token === null) return;
    GM_setValue("token", token.trim());
    auth = { ok: false, msg: "", at: 0 };
    ensureAuth().then((a) => toast(a.ok ? "授权码有效，可以使用。" : a.msg, !a.ok));
  });

  ensureAuth();
})();
