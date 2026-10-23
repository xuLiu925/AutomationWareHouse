# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Python automation toolset for Dachser freight quotation workflows. It reads shipment data from Excel files and uses Playwright to automatically fill out the Dachser eLogistics web portal (`elogistics.dachser.com`).

## Setup

```powershell
pip install playwright openpyxl
playwright install chromium
```

## Running the Scripts

### Main entry point: `dachser_fill.py`

Fill a single row from Excel (row 2 by default):
```powershell
python dachser_fill.py 询价表.xlsx
python dachser_fill.py 询价表.xlsx --row 3
```

Parse Excel without opening browser:
```powershell
python dachser_fill.py 询价表.xlsx --dry
```

Batch mode — fill all rows with order numbers sequentially:
```powershell
python dachser_fill.py 询价表.xlsx --all
python dachser_fill.py 询价表.xlsx --all --dry
```

### Match orders: `match_orders.py`

Matches Lingxing (领星) export to Dachser quotation history by MR number:
```powershell
python match_orders.py lingxing.xlsx list.html
python match_orders.py lingxing.xlsx live   # live fetch via Playwright
```

### Create orders: `create_orders.py`

Reads Lingxing export, searches each `Remark/备注` MR number in Dachser "My enquiries", clicks `create`:
```powershell
python create_orders.py lingxing.xlsx                       # live: search + click create
python create_orders.py lingxing.xlsx --dry                 # live search, no click
python create_orders.py lingxing.xlsx --html eLogistics.html  # offline match preview
```
Rules: match = Dachser order number contains all MR numbers from the remark; `created` rows ignored. Single match → auto-click (address mismatch only warned); multiple matches → user clicks manually. After create, `fill_order()` fills the Transport Order page: selects Phone delivery booking first (activates contact fields); consignee Street/Country/PCD/City overwritten from Lingxing (County left untouched; the contact table's `inactive` class may persist and is ignored); Name1=company (or recipient), Name2=email, Name3/Telephone/Mobile=phone (spaces stripped), Contact person=recipient, Order number += `+出库单号` (if over maxlength 35, original number is cut from the front, 出库单号 kept whole), Content=`battery` on lines with qty; for each such line `add_dg_line()` clicks the DG icon → Add dangerous goods → checks UN 3480 → Add → mass = line weight, No. of packages = Lingxing `Total Qty of SKU`, packaging = Fibreboard box → Save → Back (skips if UN3480 already listed); finally uploads dangerous goods documents via the "Dangerous goods documents" button — files in `危险品文件/` (or `--dg-dir`) whose name without extension equals a Lingxing `SKU n\nSKU` value exactly; missing files are skipped with a warning. Does **not** click Save; pauses for the user. Writes `建单结果.csv`.

### GUI + packaging for colleagues ("XL FreightPilot", author Xu Liu)
```powershell
python gui.py        # tkinter window: 询价 (dachser_fill.main_batch) / 建单 (create_orders.run)
python build.py      # Nuitka standalone build -> dist/XL FreightPilot/ + dist/XL FreightPilot.zip, runs `"XL FreightPilot.exe" --selftest`
```
- `app_ui.py`: `APP_DIR` (script dir, or exe dir when compiled) — `pallet_params.json`, `危险品文件/`, `dachser_profile/` resolve against it; `ask(msg, options)` replaces every `input()` (CLI = input, GUI = modal dialog via `set_ask_handler`); stop flag checked at the start of each order.
- Never call `input()` directly in the scripts — use `app_ui.ask()` so the GUI keeps working.
- GUI dialogs: not modal (no `grab_set`) and not `transient`, topmost, centered on the foreground window; `_poll` must never stop (errors are logged, failed dialogs fall back to `messagebox`), otherwise the worker thread waits forever.
- `create_orders` report defaults to `<Excel名>_建单结果.csv` next to the Excel.
- The build never includes `dachser_profile/` (login data). Colleagues need only Chrome installed.

### Pallet calculator standalone:
```powershell
python pallet_plan.py US2000 44
```

## Architecture

### Data flow
1. **Excel input** → `dachser_fill.py:read_row()` parses one row into a normalized dict
2. **Pallet planning** → `pallet_plan.py:plan_pallets()` looks up product params in `pallet_params.json` and calculates pallet layout (full pallets + remainder)
3. **Browser automation** → `dachser_fill.py:Filler` class drives Playwright to fill each field in sequence

### Key files
- `dachser_fill.py` — main script; contains `Filler` class (Playwright page wrapper), `read_row()` (Excel parser), `fill_form()` (orchestrates all fill steps), `main_batch()` (batch mode)
- `pallet_plan.py` — pallet calculation logic; reads `pallet_params.json`; `plan_pallets()` returns `(lines, notes)` for up to 2 order lines (full pallets + partial)
- `pallet_params.json` — product specs: carton dimensions/weight, pallet dimensions, per-layer count, max layers, Dachser packaging type label
- `pallet_rules.json` — older/alternate product rules format (currently used for reference)
- `fetch_dachser.py` — thin Playwright helper used by `match_orders.py` in `live` mode; saves browser session in `browser_profile/`
- `match_orders.py` — reads Lingxing export + Dachser HTML list → matches by MR number → outputs `匹配结果.csv`

### Browser session persistence
- `dachser_fill.py` saves login state to `dachser_profile/` (gitignored)
- `fetch_dachser.py` saves login state to `browser_profile/` (gitignored)
- `dachser_fill.open_quotation()` (used by `dachser_fill.py` and `create_orders.py`): opens the site; if already logged in goes straight on; if the login form is autofilled by Chrome within ~6s it clicks Login; otherwise waits (polling, no Enter needed) until the user logs in; then clicks freightquotation
- `_launch()` passes `chromium_sandbox=True` so Chrome doesn't show the `--no-sandbox` warning bar (falls back to no-sandbox if launch fails)

### Excel column contract (`dachser_fill.py`)
Required columns in row 1 (header): `编号`, `国家`, `邮编`, `城市`, `地址`, `数量`, `Unit/单位`, `单个物料（长）`, `单个物料（宽）`, `单个物料（高）`, `单个物料（重量）`, `Dangerous Goods Type/危险品类型`, `产品物料`, `是否需要尾板`

### Adding a new product to pallet_params.json
Each entry needs: `packaging` (exact text from Dachser dropdown), `carton_length_cm`, `carton_width_cm`, `carton_height_cm`, `carton_weight_kg`, `pallet_length_cm`, `pallet_width_cm`, `per_layer`, `max_layers`. Optional: `aliases` (list), `small_per_layer`, `layers_by_qty` (dict of qty→layers overrides), `full_pallet_height_cm`, `full_pallet_weight_kg`.

### Dachser form quirks
- The form uses Wicket AJAX; `Filler.wait_ajax()` waits for `.overlay` to disappear after each field interaction
- Country and packaging type use a custom `fineSelect` widget (hidden `<select>` + visible `div`); `pick_by_text()` tries real click → JS mouse events → direct DOM sync as fallback
- Collection date is intentionally suppressed via CSS to prevent calendar popups
- The script fills but does **not** submit — user must manually click "Enquiry"
- Transport order edit page (Process order) **locks the order** until Save or Cancel. Never navigate away / close the browser while on it: `create_orders.leave_edit_page()` asks the user (save, or discard → clicks Cancel) before moving on or closing. Otherwise reopening shows "The order is being processed by someone else".
