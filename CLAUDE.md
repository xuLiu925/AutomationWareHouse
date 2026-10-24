# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**XL FreightPilot** (author: Xu Liu / 刘旭) — Python + Playwright automation for the Dachser eLogistics portal (`elogistics.dachser.com`), used by the logistics team:

- **询价 (quotation)**: Excel enquiry sheet → pallet plan → fills the freightquotation Enquiry form (`dachser_fill.py`)
- **建单 (transport order)**: Lingxing (领星) outbound export → finds the quotation by MR number in "My enquiries" → clicks `create` → fills the transport order page (`create_orders.py`)
- Shipped to colleagues as a Nuitka-compiled Windows GUI (`gui.py`, `build.py`). The repo is **private**; never make source public (license switch would be bypassable, `eLogistics.html` contains business data).

The scripts fill but **never submit** (no Enquiry / Save click) — the user reviews and submits.

User-facing docs: `README.md` (developer), `使用说明.txt` (colleagues, bundled into the build). Keep them in sync with behavior changes. Communicate with the user in Chinese.

## Running

```powershell
pip install playwright openpyxl
python gui.py                                            # GUI (same as the exe)
python dachser_fill.py 询价表.xlsx --all [--dry]           # batch quotation
python dachser_fill.py 询价表.xlsx [--row 3] [--dry]       # single order group
python create_orders.py 领星.xlsx [--dry] [--days 30] [--dg-dir DIR] [-o out.csv]
python create_orders.py 领星.xlsx --html eLogistics.html   # offline match preview (no browser login)
python match_orders.py 领星.xlsx list.html|live            # standalone matching CSV
python pallet_plan.py US2000 44
python build.py                                          # package (see below)
```

There is no test suite. Verify changes by: `--dry` / `--html eLogistics.html` runs, `--all --dry` on `test_inqery.xlsx`, and Playwright against small hand-made HTML fixtures (put them in the scratchpad, not the repo) that mimic the real Dachser markup the user pasted. Things that need the live site must be stated as unverified.

## Architecture

```
gui.py            tkinter window; runs dachser_fill.main_batch / create_orders.run in a worker thread
app_ui.py         APP_DIR, app_path(), ask(), stop flag — shared by CLI and GUI
license_check.py  remote on/off switch, checked by gui.py before each run
dachser_fill.py   Excel parsing, Filler (page wrapper), _launch(), open_quotation() login, quotation batch
pallet_plan.py    plan_pallets(): params from pallet_params.json -> up to 2 order lines + notes
create_orders.py  search/match/create, fill_order(), add_dg_line(), upload_dg(), leave_edit_page()
match_orders.py   read_lingxing(), check_address(), COUNTRY map, MR regex
fetch_dachser.py  thin fetch helper for match_orders live mode (browser_profile/)
build.py          Nuitka build
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
- Contact fields: select Phone delivery booking (`selDeliveryOptions` value `AT`) first; the contact table keeps class `inactive` — ignore it.
- DG dialog per order line: DG icon → "Add dangerous goods" → tick the row whose UN no. is `3480` (row clicks toggle the checkbox too, so re-check state) → Add → mass = line weight, No. of packages = Lingxing `Total Qty of SKU`, packaging `Fibreboard box` → Save → Back. Skips if 3480 already listed.

### 建单 fill rules (`fill_order`)
Street (Address1 minus trailing city/zip, + house no / Address2), Country (`norm_country`, only if different), PCD (`norm_zip`, re-pads lost leading zeros), City from Lingxing; **County untouched**. Name1 = company or recipient, Name2 = email, Name3/Telephone/Mobile = phone (`norm_phone`: digits and `+` only, `00` → `+`), Contact person = recipient. Order number += `+出库单号`; if over maxlength (35) the original is cut from the front and leading separators stripped, 出库单号 kept whole. Content `battery` only on lines with qty. Over-maxlength values are truncated with a warning. Upload DG files last (multipart AJAX submits the whole form).

Matching (`classify`): keys = all MR numbers in the remark (or the whole remark if none); a Dachser order matches if it contains every key (`MR\d{8}` not followed by a digit); `created`/buttonless rows excluded. 0 → not found / already created; 1 → auto create; >1 → user chooses.

### Excel contracts
- 询价 row-1 headers: `编号`, `国家`, `邮编`, `城市`, `地址`, `数量`, `Unit/单位`, `单个物料（长/宽/高/重量）`, `Dangerous Goods Type/危险品类型`, `产品物料`, `是否需要尾板`; optional `运费报价(欧元)` — merged cells group rows into one order.
- 领星 (prefix match): `Outbound Order No`, `Remark`, `Recipient/`, `Company`, `Country/Region`, `City`, `Post code`, `House No`, `Address1`, `Address2`, `Province Name`, `Telephone`, `Email/`, `Total Qty of SKU`, `SKU n\nSKU`.
- Wrong sheet type for the mode raises `SystemExit` with a hint to switch mode.

### pallet_params.json
Per product: `packaging` (exact Dachser dropdown text), `carton_length_cm`, `carton_width_cm`, `carton_height_cm`, `carton_weight_kg`, `pallet_length_cm`, `pallet_width_cm`, `per_layer`, `max_layers`. Optional: `aliases`, `small_per_layer`, `layers_by_qty`, `full_pallet_height_cm`, `full_pallet_weight_kg`. `_defaults` holds shared limits (pallet height/weight, max total height 220, max weight 1500). Lives next to the exe; editing it needs no rebuild.

## GUI (`gui.py`)
- Worker thread + queue; `_poll` (every 100 ms) must never stop — errors are logged, a failed dialog falls back to `messagebox`, otherwise the worker waits forever.
- Dialogs: not modal (no `grab_set`), not `transient`, topmost, centered on the current foreground window (`_foreground_rect`, outer frame measured from the root window); clicking the main window lifts a pending dialog.
- Footer shows `license_check.NOTICE`; title `XL FreightPilot — Freight Quotation & Order Automation`.

## License switch (`license_check.py`)
Fetches `https://raw.githubusercontent.com/xuLiu925/liuxu-license/main/license.json` (`{"enabled": true, "message": ""}`) before each GUI run. Only `enabled is True` passes; network/format errors block use (user's choice). Disabled only stops automation — never delete or alter files. Keep the behavior disclosed (footer, `使用说明.txt`); don't add hidden mechanisms.

## Build (`build.py`)
Nuitka `--standalone`, `--enable-plugin=tk-inter`, `--windows-console-mode=disable`, `--include-package=playwright` (the Nuitka playwright plugin bundles the driver; adding it manually conflicts). Nuitka's DLL tool breaks on non-ASCII paths: when the project root is ASCII (`D:\XLFreightPilot`) it compiles in place with the current Python into `build/` (`IN_PLACE`); otherwise sources are copied to `D:\liuxu_build` (`LIUXU_BUILD_DIR`) and built in a dedicated venv there. `NUITKA_CACHE_DIR` = `<stage>/cache` (compiler auto-downloaded, Zig). Output `dist/XL FreightPilot/` + zip, then runs `"XL FreightPilot.exe" --selftest`. Fails with a clear message if the old exe is still running. The user packages manually — don't run `build.py` unless asked.

## Git
Remote `origin` = private `xuLiu925/AutomationWareHouse`, branch `main`. Commit/push only when asked; confirm the repo is still private before pushing. Never commit `危险品文件/`, `dachser_profile/`, `browser_profile/`, `*.xlsx`, `*.csv`, `license.json`, `build/`, `dist/`.
