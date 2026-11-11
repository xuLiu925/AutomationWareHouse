# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**XL FreightPilot** (author: Xu Liu / 刘旭) — Python + Playwright automation for the Dachser eLogistics portal (`elogistics.dachser.com`), used by the logistics team:

- **询价 (quotation)**: Excel enquiry sheet → pallet plan → fills the freightquotation Enquiry form (`dachser_fill.py`)
- **建单 (transport order)**: Lingxing (领星) outbound export → finds the quotation by MR number in "My enquiries" → clicks `create` → fills the transport order page (`create_orders.py`)
- **对账 (reconcile)**: Lingxing outbound export → MR numbers from the remark → quote emails in the user's own Feishu mailbox → `<Excel名>_对账结果.xlsx` (`reconcile.py`, `feishu_mail.py`)
- Shipped to colleagues as a Nuitka-compiled Windows GUI (`gui.py`, `build.py`). The repo is **public** (portfolio, the user's own project). Never add business data (customer orders, addresses, phones, real Excel exports) to tracked files; use made-up examples in comments and docs. `eLogistics.html` predates this and the user chose to leave it and the history as they are.

The scripts fill but **never submit** (no Enquiry / Save click) — the user reviews and submits.

Docs: `README.md` (developer). There is no colleague-facing manual in the repo or the build (user's choice — small company, colleagues ask the user directly); `使用说明.txt` exists only on the user's machine (gitignored), don't recreate or bundle it. Keep README in sync with behavior changes. Communicate with the user in Chinese.

## Running

```powershell
pip install playwright openpyxl
python gui.py                                            # GUI (same as the exe)
python dachser_fill.py 询价表.xlsx --all [--dry]           # batch quotation
python dachser_fill.py 询价表.xlsx [--row 3] [--dry]       # single order group
python create_orders.py 领星.xlsx [--dry] [--days 30] [--dg-dir DIR] [-o out.csv]
python create_orders.py 领星.xlsx --html eLogistics.html   # offline match preview (no browser login)
python match_orders.py 领星.xlsx list.html|live            # standalone matching CSV
python pallet_plan.py 11US2000ZZD-20001 44
python build.py                                          # package (see below)
```

There is no test suite. Verify changes by: `--dry` / `--html eLogistics.html` runs, `--all --dry` on `test_inqery.xlsx`, and Playwright against small hand-made HTML fixtures (put them in the scratchpad, not the repo) that mimic the real Dachser markup the user pasted. Things that need the live site must be stated as unverified.

## Architecture

```
gui.py            tkinter window; runs dachser_fill.main_batch / create_orders.run in a worker thread
app_ui.py         APP_DIR, app_path(), ask(), stop flag — shared by CLI and GUI
license_check.py  service switch check() via Worker before each GUI run, post()/headers() (Windows user name, optional admin token)
dachser_fill.py   Excel parsing, Filler (page wrapper), _launch(), open_quotation() login, quotation batch
ai_clean.py       AI cleanup of messy quotation cells via the Worker (error / daily limit → rules)
worker/           Cloudflare Worker (JS): auth + Gemini relay + pallet params/review; holds key + prompt; not bundled
pallet_plan.py    plan_pallets(): params from pallet_params.json -> up to 2 order lines + notes; load_raw/save_raw/validate_product/validate_defaults
params_editor.py  tkinter editor for pallet_params.json (GUI button "产品参数"): submit for review / review / history
params_sync.py    server sync of pallet params (Worker /products/*, /proposals/*), ServerError
create_orders.py  search/match/create, fill_order(), add_dg_line(), upload_dg(), leave_edit_page()
match_orders.py   read_lingxing(), check_address(), COUNTRY map, MR regex
dachser_cost.py   对账 cost column: reads My enquiries (create_orders.search, read-only) and matches Dachser prices to outbound orders
reconcile.py      对账: read_outbound(), parse_mail() (own part only, tables 编号/运费报价), run_ai() via Worker /quote, lookup() latest quote, xlsx output; --eml offline mode
feishu_mail.py    Feishu OAuth (local callback server on the redirect port, code → token via Worker), feishu_token.json, list/get messages
updater.py        online update: version.txt check via Worker /update/latest, download + sha256, _update/apply.ps1 swaps files after exit and restarts
fetch_dachser.py  thin fetch helper for match_orders live mode (browser_profile/)
build.py          Nuitka build
xlfp-wms-sn.user.js  Tampermonkey script for Lingxing WMS (wms.xlwms.com) review page: commas → newlines in the batch serial textarea; copied to dist via EXTRA_FILES
script.py, script1.py   PyCharm sample files, unused
```

### Cross-cutting rules
- **Never call `input()`** in the scripts — use `app_ui.ask(msg, [(key, label), ...])` (first option = default). CLI → `input()`, GUI → dialog via `set_ask_handler`. Loops check `app_ui.stop_requested()` at the start of each order.
- **Paths**: resolve data files against `app_ui.APP_DIR` (script dir, or exe dir when compiled — `"__compiled__" in globals()`), never CWD or `Path(__file__)`. Applies to `pallet_params.json`, `危险品文件/`, `dachser_profile/` (`_launch` uses `app_path(args.profile)`).
- **Report output** of `create_orders` defaults to `<Excel名>_建单结果.csv` next to the input Excel.
- New modules must be added to `build.py` `SOURCES`.

### Browser / login
- `_launch()`: persistent context in `dachser_profile/`, `channel="chrome"` (system Chrome, browsers are not bundled), `chromium_sandbox=True` (removes the `--no-sandbox` warning bar; falls back if launch fails).
- `open_quotation()`: goes to START_URL; already logged in → continue; login form autofilled by Chrome within ~6 s (`:-webkit-autofill`, values unreadable before a user gesture) → clicks Login; else polls until the user logs in; then clicks the freightquotation menu link (`_quotation_link`).

### Dachser form quirks
- Wicket AJAX: wait for `.overlay` to hide after interactions (`Filler.wait_ajax` / `create_orders.wait_ajax`). Element ids are regenerated — select by `name$=` suffixes or text, never by `id`.
- `fineSelect` widgets (hidden `<select>` + visible div): use `Filler.pick_by_text(select_locator, text)` (real click → JS events → DOM sync fallback).
- Collection date / time pickers are hidden via CSS to avoid popups.
- My enquiries search (`create_orders.search`): mark the old result table (`data-old-result`) before clicking Search and wait until it is replaced — Search may be a full reload, `domcontentloaded` alone returns stale results.
- Transport order edit page **locks the order** until Save or Cancel. `leave_edit_page()` must run before navigating away or closing the browser (asks: saved / discard → clicks Cancel, accepting the confirm dialog). Otherwise reopening shows "The order is being processed by someone else".
- Loading point (mandatory fineSelect `select[name$='loadingParty:select']`, added to the transport order page in Nov 2026): `pick_loading_point` picks the option whose text starts with `Loading point A` (fallback value `-1`), first thing in `fill_order`; skipped if the page has no such select.
- Collection date (`pnlTransportInfo:orderDatePicker`, page default = today): `set_collection_date` leaves it before 11:00 local time, from 11:00 on sets tomorrow (plain next day, weekends not skipped — user's rule), format taken from the current value (`mm/dd/yyyy` in English).
- Contact fields: select Phone delivery booking (`selDeliveryOptions` value `AT`) first; the contact table keeps class `inactive` — ignore it.
- DG dialog per order line: DG icon → "Add dangerous goods" → tick the row whose UN no. is `3480` (row clicks toggle the checkbox too, so re-check state) → Add → mass = line weight, No. of packages = Lingxing `Total Qty of SKU`, packaging `Fibreboard box` → Save → Back. Skips if 3480 already listed.

### 建单 fill rules (`fill_order`)
Street (`prepare_streets` before the run; same logic as 询价, see Excel contracts), Country (`norm_country`, only if different), PCD (`norm_zip`, re-pads lost leading zeros), City from Lingxing; **County untouched**. Name1 = company or recipient, Name2 = email, Name3/Telephone/Mobile = phone (`norm_phone`: digits and `+` only, `00` → `+`), Contact person = recipient. Order number += `+出库单号`; if over maxlength (35) the original is cut from the front and leading separators stripped, 出库单号 kept whole. Content `battery` only on lines with qty. Over-maxlength values are truncated with a warning. Upload DG files last (multipart AJAX submits the whole form).

Matching (`classify` / `find_hits`): keys = all MR numbers in the remark (or the whole remark if none); a Dachser order matches if it contains every key (`MR\d{8}` not followed by a digit). No forward hit → reverse match (`in_remark`): the Dachser order number appears verbatim in the remark (no adjacent letter/digit or `-1`-style continuation, ≥3 chars) — 询价 only fills the leading 编号 token, the remark may carry extra text or a second MR. Search term (`search_term`): first MR, else `parse_order_no` of the remark. No AI here (user's choice). `created`/buttonless rows excluded. 0 → not found / already created; 1 → auto create; >1 → user chooses.

### Excel contracts
- 询价 row-1 headers: `编号`, `国家`, `邮编`, `城市`, `地址`, `数量`, `Unit/单位`, `单个物料（长/宽/高/重量）`, `Dangerous Goods Type/危险品类型`, `产品物料`, `是否需要尾板` (prefix match); optional `运费报价(欧元)` — merged cells group rows into one order; optional `备注…`.
  - Rule given to colleagues: each piece of info in ONE cell (merge or in-cell newlines), no continuation rows.
  - `编号` may carry free text: `parse_order_no` takes the leading `[A-Za-z0-9-_/.]` run; unsure cases (no token, <3 or >35 chars, space+alnum or a symbol right after) are asked via `resolve_uncertain` before the batch starts (not in `--dry`).
  - Tail lift: `parse_tail_lift` → True/False/None (negation words first); None is asked before start. Remarks (编号 leftover + 备注 column, incl. rows under the 编号 merge) are only listed for manual handling, never filled.
  - Street: **询价 and 建单 share one logic — keep it that way, never fork it.** Dachser has ONE street field: street + house number only, original order kept, never split. Shared in `dachser_fill.py`: `street_input` (same input fields), `rule_street`/`clean_street` (rules), `merge_ai_street`/`street_ok` (AI + validation), `street_all_notes`, `ask_street`; in `worker.js` one `STREET_SPEC`. House No is added; Address2 is never appended (pure house number only), otherwise just warned. `street_warnings` flags leftover zip-like numbers / no house number.
  - AI (`ai_clean.py` → Worker `POST /clean` → Gemini; every row is sent, by user's choice; 建单 uses `POST /street` with Lingxing address fields, also user's choice; a failed chunk is retried, then only its rows fall back to rules): `apply_ai` merges into the rule result — AI order no must be a substring of the cell, AI street must pass `street_ok` (words from the source; no zip/city/country/name; house number kept), conflicts become "ask before start". Any error (incl. 429 daily limit) → rules only. Only the needed fields are sent (no phone/email). Prompt, schema, model and Gemini key live only in `worker/worker.js` — never move them back into the exe or repo config. EEA distribution requires the paid Gemini tier.
- 领星 (prefix match): `Outbound Order No`, `Remark`, `Recipient/`, `Company`, `Country/Region`, `City`, `Post code`, `House No`, `Address1`, `Address2`, `Province Name`, `Telephone`, `Email/`, `Total Qty of SKU`, `SKU n\nSKU`.
- Wrong sheet type for the mode raises `SystemExit` with a hint to switch mode.

### pallet_params.json
Gitignored (local cache of the server version; must exist locally before `build.py`, which bundles it). Keyed by the product material code (Lingxing SKU, as in the 询价 `产品物料` column; exact match ignoring case/spaces, no aliases). Per product: `packaging` (exact Dachser dropdown text), `carton_length_cm`, `carton_width_cm`, `carton_height_cm`, `carton_weight_kg`, `pallet_length_cm`, `pallet_width_cm`, `per_layer`, `max_layers`, `full_pallet_height_cm`, `full_pallet_weight_kg` (measured full pallet, always used for full pallets). Optional: `small_per_layer`, `layers_by_qty`. `_defaults` holds shared limits (pallet height/weight, max total height 220, max weight 1500). Lives next to the exe; editing it needs no rebuild. Edited via `params_editor.py`: save-time validation in `pallet_plan.validate_product` (errors block: required > 0, ints, max_layers×carton height + pallet height ≤ total height, full pallet computed/measured weight ≤ max, layers_by_qty sane, duplicate SKU; warnings need confirm: measured lower/lighter than computed, carton bigger than pallet). Runtime has no separate height check by design. `save_raw` backs up to `.bak`, writes atomically and calls `reload()`.

**Server is the source of truth** (Worker KV `products`, `products:v:<n>` history, `proposal:<id>`). `gui._run` calls `params_sync.sync_local()` before a quotation run → writes the local file as a cache with `_version` (failure → warn, use cache). Editor: the admin is the machine whose `xlfp_config.json` holds a token with `role: "admin"` in KV `user:<token>`; it saves directly (`/products/save` with `base_version`, 409 on stale) and reviews; everyone else submits proposals (`upsert` with `old_sku` for rename / `delete` / `defaults`) that only the admin approves — user's choice: review required, admin is the only reviewer, new + change + delete all reviewed. Approve re-validates the proposal against current data in the client; the Worker applies it (`applyProposal`) and only checks shape/permissions. Everyone else has no token: identity = Windows user name (header `X-XLFP-User`, URL-encoded; proposals owned by `win:<name>`). Offline / disabled → editor is read-only. CLI runs use the local file without syncing.

## GUI (`gui.py`)
- Worker thread + queue; `_poll` (every 100 ms) must never stop — errors are logged, a failed dialog falls back to `messagebox`, otherwise the worker waits forever.
- Dialogs: not modal (no `grab_set`), not `transient`, topmost, centered on the current foreground window (`_foreground_rect`, outer frame measured from the root window); clicking the main window lifts a pending dialog.
- Footer shows `license_check.NOTICE`; title `XL FreightPilot — Freight Quotation & Order Automation`.

## License + AI server (`license_check.py`, `worker/`)
Cloudflare Worker (`worker/worker.js`, deploy/admin steps in `worker/README.md`), URL in `license_check.SERVER_URL` (override with env `XLFP_SERVER`). Colleagues have **no token** (user's choice, Nov 2026: no 授权码 UI, no per-user codes). Every GUI run first asks `POST /check`; global switch `config.enabled === false` → all routes 403 with `config.message` (default "服务暂停使用，请联系刘旭"); unreachable server also blocks use (user's choice — the switch must be authoritative). `config.ai_daily` = whole-software AI calls per day (`usage:<date>:all`; set online by the user, 0/absent = unlimited; over → 429 → client falls back to rules). Requests carry `X-XLFP-User` (Windows user name). Only the admin's own machine has `xlfp_config.json` `{"token"}`; a token found in KV `user:<token>` gives that record (`role`, optional `enabled`/`daily`); unknown/no token = plain user, never rejected. Disabled only stops automation — never delete or alter files. Keep the behavior disclosed (GUI footer); don't add hidden mechanisms. Test locally with a Node server wrapping `worker.js` (in-memory KV) in the scratchpad.
`/aftersales/read` and `/aftersales/reply` serve the user's separate after-sales agent (`D:\aftersales-agent`, not in this repo, runs on their phone in Termux): drafts only, it mails nothing but the user's own mailbox.

## 对账 (`reconcile.py`, `feishu_mail.py`)
Quote email = our reply to the customer's "出库询价" mail (subject contains `报价`); its first table has `编号` and `运费报价(欧元)` ("320欧"). Rules (user's choice): only mails whose sender has the user's own mail domain count as quotes (domain taken from a SENT-folder sender at runtime, never hardcoded; `--eml` / no SENT mail → all senders); per MR the **latest** quote email wins; a code whose quotes changed over time gets a light-blue price cell + history note.
Customer's **卡派出货通知** (subject contains `出货通知`, any sender; table `No.`/`编号`, `费用情况`/`运费报价`, `出库订单号` = Lingxing Outbound Order No, `派送方式`) = customer-confirmed price, never a quote: matched by outbound no first, then codes; latest notice wins. Quote found + notice price differs → 需人工确认; no quote (or none in remark / latest 待定) + notice price → price from notice, status `仅客户确认价`, orange. Customer `询价` mails (external sender, `iq` codes) only feed status `有询价没找到报价`.
Only 卡派 (Dachser) orders are reconciled: quote found > notice 派送方式 (快递/DHL… → express, 卡派 → truck) > inquiry exists > `Shipping service` contains Dachser; else express/unknown → sheet 未对账, not red. `Shipping service` alone is unreliable (often 无需物流服务) — never let it drop an order that has a quote/notice/inquiry. Exception (user's choice): quote exists but the customer's notice saying 快递 is newer than the latest quote → 改发快递, sheet 未对账, row filled gold. One quote cell matched by several 卡派 orders → all of them 需人工确认 with no price (double-billing guard), outbound no cell light purple (user's choice).
Quote rules continued: latest says 待定 → price blank (status 最新报价待定); blank price = the software found nothing; one outbound order = one price (user's rule): codes whose latest quote is the same price cell (several codes in one 编号 cell, or a rowspan-merged price cell — `_Tables` tracks cell ids, `q` entries carry the cell id as 4th item) count once; codes quoted separately (different price cells) → price blank, 需人工确认 listing each price — never summed (user's choice). The result path is printed as the last log line. The output is used to bill customers: column `依据` holds the raw price cell or `AI 读取：<evidence>` (AI rows filled light yellow so they get checked); non-已找到 status cells light red. Only the mail's own part is parsed (cut at `history-quote-wrapper` / blockquote / … — the quoted inquiry has empty prices). Empty price cells don't count as a quote. Unparseable cells (no number, several numbers) and MRs mentioned in a 报价 mail's own text without a table entry go to the AI (`POST /quote`); the AI price must appear in its evidence and the evidence in the mail text, else 需人工确认. Each user reads **their own** mailbox (all folders except DRAFT/SPAM/TRASH via `list_folders`; needs `mail:user_mailbox.folder:read`, falls back to INBOX + SENT); App ID / Secret live only in the Worker (`FEISHU_APP_ID`, `FEISHU_APP_SECRET` — keep them out of the repo and exe); mail content never goes through the Worker except the AI fallback. `feishu_token.json`, `feishu_mail_cache.json` (parsed results, `PARSER_VERSION`) and `feishu_mail_raw/` (one gzip JSON per mail, file name = sha1 of the message id, inline `data:` images stripped) are local and gitignored. Bumping `PARSER_VERSION` re-parses old records from the raw files (`reparse_outdated`, keeps AI results when the text is unchanged) — no re-download; mails cached without a raw file are downloaded once to fill it in. Cost (`dachser_cost.py`, user's choice: one column `成本价(欧元)`, every run): Dachser My enquiries Price = our cost (account shared by everyone). `reconcile.run` opens Dachser first (`_launch` + `open_quotation`, `create_orders.search(page, "", days)` from earliest outbound − `--days`), read-only. Match: Dachser order number contains the outbound no (建单 appends `+出库单号`) → that row; else `create_orders.find_hits` on the remark (shared rule, never fork), minus rows carrying another outbound no of this batch; prefer created rows; several prices / none → blank + `成本：…` in 说明. Email-quoted costs (price in a PDF attachment) are not read (user's choice). Fetch failure → column blank, reconcile continues. `--no-cost`, `--cost-html FILE` (offline list page); `--eml` without `--cost-html` skips cost. Unverified live: whether a long date range is capped/paged in My enquiries. Timing: `reconcile.run` prints per-phase durations before the result-path line; `gui._run` prints the total for every mode (`app_ui.duration`). No `offline_access` yet → re-authorize every ~2 h. Paging stops when a whole page is older than (earliest OutboundTime − `--days`, default 60) — assumes newest-first listing. The live Feishu API (list order, folder ids, body encoding) is unverified until the user runs it; test with `--eml` and a mock Worker.

## WMS serial helper (`xlfp-wms-sn.user.js`)
Lingxing WMS 按单复核 with "支持批量复核" ticked turns the serial input into `XlInputBatch` (popover textarea, one SN per line, whole list submitted at once). QR codes yield comma-separated SNs; the script rewrites the textarea on `input` (commas → newlines, trim, dedupe finished lines, last partial line untouched) via the native value setter + dispatched `input` event. No token (same service switch as the exe): `POST /check` via `GM_xmlhttpRequest` (`@connect` the Worker host), cached 1 h when OK; not OK → no conversion, one toast per textarea focus. Plain JS, so the check is bypassable — disclosed. Never call WMS APIs from it. Test with a Playwright fixture stubbing `GM_*`.

## Online update (`updater.py`)
Release files live in Cloudflare R2 (Worker binding `RELEASES`, bucket `xlfp-releases`; user's choice over GitHub Releases: private, served only while the switch is on). `build.py [notes]` writes `version.txt` (`YYYY.MM.DD.HHMM`, string-compared) into the dist and `dist/release/` = `XL-FreightPilot-<ver>.zip` + `latest.json` {version, file, sha256, size, notes}; the user uploads zip first, then latest.json. Worker `POST /update/latest` returns latest.json (`{version: null}` without binding), `/update/download` streams that zip; both are blocked when the switch is off. GUI checks 1.5 s after start (silent on errors; skipped without version.txt = running .py) and via the 检查更新 button; never while a task runs. Install: extract to `_update/new` (zip-slip checked) → `_update/apply.ps1` (UTF-8 BOM, started with CREATE_NO_WINDOW — DETACHED_PROCESS makes PowerShell exit; tries CREATE_BREAKAWAY_FROM_JOB first) waits for the exe pid, robocopies over APP_DIR excluding an existing `pallet_params.json` and `version.txt`, copies `version.txt` last (so a failed copy keeps the old version and re-offers the update), writes `_update/failed.txt` on robocopy ≥ 8 (shown at next start), restarts the exe. User data files are not in the package. Tests: mock Worker with a fake R2 object + a fake app dir; the Bash/PowerShell tool job kills child processes, so launch the run phase via `Win32_Process.Create` to test the swap.

## Build (`build.py`)
Nuitka `--standalone`, `--enable-plugin=tk-inter`, `--windows-console-mode=disable`, `--include-package=playwright` (the Nuitka playwright plugin bundles the driver; adding it manually conflicts). Nuitka's DLL tool breaks on non-ASCII paths: when the project root is ASCII (`D:\XLFreightPilot`) it compiles in place with the current Python into `build/` (`IN_PLACE`); otherwise sources are copied to `D:\liuxu_build` (`LIUXU_BUILD_DIR`) and built in a dedicated venv there. `NUITKA_CACHE_DIR` = `<stage>/cache` (compiler auto-downloaded, Zig). Output `dist/XL FreightPilot/` + zip, then runs `"XL FreightPilot.exe" --selftest`. Fails with a clear message if the old exe is still running. The user packages manually — don't run `build.py` unless asked.

## Git
Remote `origin` = public `xuLiu925/AutomationWareHouse`, branch `main`. Commit/push only when asked; everything pushed is public. Never commit `feishu_token.json`, `feishu_mail_cache.json`, `feishu_mail_raw/`, `pallet_params.json`, `使用说明.txt`, `危险品文件/`, `dachser_profile/`, `browser_profile/`, `*.xlsx`, `*.csv`, `license.json`, `xlfp_config.json`, `build/`, `dist/`.
