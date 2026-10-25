"""产品托盘参数编辑窗口（从 gui.py 的"产品参数"按钮打开）。

参数以服务器为准（params_sync.py ↔ worker/worker.js）：
- 管理员：保存 / 删除 / 改通用参数直接生效；"审核"页处理别人的申请；"历史版本"页可回滚
- 其他人：保存变成"提交审核"，管理员通过后生效；"我的申请"页看进度、撤回
- 连不上服务器 / 没有授权码：只读，显示本机缓存
校验规则在 pallet_plan.py（validate_product / validate_defaults），这里只管界面。
"""
import tkinter as tk
from tkinter import messagebox, simpledialog, ttk

import pallet_plan
import params_sync
from params_sync import ServerError

TITLE = "产品托盘参数"

# 产品表单：(键, 标签, 类型, 说明)  类型 text / num / int / lbq；键为 None = 分组标题
PRODUCT_FIELDS = [
    (None, "基本", None, None),
    ("packaging", "包装类型", "text", "Dachser 包装类型下拉框里显示的文字"),
    (None, "单箱", None, None),
    ("carton_length_cm", "单箱长 cm", "num", ""),
    ("carton_width_cm", "单箱宽 cm", "num", ""),
    ("carton_height_cm", "单箱高 cm", "num", ""),
    ("carton_weight_kg", "单箱重 kg", "num", ""),
    (None, "托盘与摆放", None, None),
    ("pallet_length_cm", "托盘长 cm", "num", ""),
    ("pallet_width_cm", "托盘宽 cm", "num", ""),
    ("per_layer", "每层箱数", "int", "满一层时每层放几箱"),
    ("max_layers", "最多层数", "int", "整托盘码几层"),
    (None, "整托实测", None, None),
    ("full_pallet_height_cm", "整托高度 cm", "num", "实际打好的整托盘总高（含托盘）"),
    ("full_pallet_weight_kg", "整托重量 kg", "num", "实际打好的整托盘总重（含托盘、膜）"),
    (None, "高级（选填）", None, None),
    ("small_per_layer", "不满一层时每层", "int", "箱数不满一层时每层放几箱，空 = 同每层箱数"),
    ("layers_by_qty", "指定层数", "lbq", "特殊箱数指定层数，写成 箱数=层数，逗号隔开，如 4=2, 7=3"),
]
OPTIONAL = {"small_per_layer", "layers_by_qty"}
LABELS = {k: lbl for k, lbl, _, _ in PRODUCT_FIELDS if k}

DEFAULT_FIELDS = [
    ("packaging", "未知产品包装类型", "text", "查不到产品、按 Excel 尺寸估算时用"),
    ("pallet_length_cm", "托盘长 cm", "num", "未知产品估算用"),
    ("pallet_width_cm", "托盘宽 cm", "num", "未知产品估算用"),
    ("pallet_height_cm", "托盘底板高 cm", "num", "零散托盘高度 = 层数×箱高 + 这个"),
    ("pallet_weight_kg", "托盘自重 kg", "num", "零散托盘重量 = 箱数×箱重 + 这个"),
    ("pallet_total_height_cm", "总高度上限 cm", "num", "所有产品通用，含托盘"),
    ("max_pallet_weight_kg", "单托重量上限 kg", "num", "所有产品通用，含托盘"),
]
DLABELS = {k: lbl for k, lbl, _, _ in DEFAULT_FIELDS}


def _fmt(v):
    if v is None:
        return ""
    if isinstance(v, float):
        return f"{v:g}"
    if isinstance(v, dict):
        return _fmt_lbq(v)
    return str(v)


def _parse_num(s, kind, label):
    """文本 -> 数字；空返回 None；格式不对抛 ValueError（带中文说明）。"""
    s = s.strip().replace("，", ".").replace(",", ".")
    if not s:
        return None
    try:
        v = float(s)
    except ValueError:
        raise ValueError(f"{label}：'{s}' 不是数字")
    if v.is_integer():
        return int(v)
    if kind == "int":
        raise ValueError(f"{label}：必须是整数")
    return v


def _fmt_lbq(d):
    return ", ".join(f"{q}={n}" for q, n in (d or {}).items())


def _parse_lbq(s):
    out = {}
    for part in s.replace("，", ",").replace("；", ",").replace(";", ",").split(","):
        part = part.strip()
        if not part:
            continue
        q, sep, n = part.replace("：", "=").replace(":", "=").partition("=")
        if not sep or not q.strip().isdigit() or not n.strip().isdigit():
            raise ValueError(f"指定层数：'{part}' 格式不对，应写成 箱数=层数，例如 4=2")
        out[str(int(q))] = int(n)
    return out or None


def _norm(s):
    return pallet_plan._norm(s)


def _merged_defaults(raw):
    d = {k: v for k, v in (raw.get("_defaults") or {}).items() if not k.startswith("_")}
    return {**pallet_plan._HARDCODED_DEFAULTS, **d}


class ParamsEditor:
    _instance = None

    @classmethod
    def open(cls, root):
        """同一时间只开一个编辑窗口；已经开着就拉到前面。"""
        if cls._instance is not None:
            try:
                if cls._instance.win.winfo_exists():
                    cls._instance.win.deiconify()
                    cls._instance.win.lift()
                    cls._instance.win.focus_force()
                    return cls._instance
            except tk.TclError:
                pass
        cls._instance = cls(root)
        return cls._instance

    def __init__(self, root):
        self.root = root
        self.raw = {}
        self.version = None
        self.role = None             # "admin" / "user"；None = 只读（离线）
        self.offline_reason = ""
        self.current = None          # 正在编辑的产品键；None = 新产品
        self.dirty = False
        self._loading = False
        self.proposals = []
        self.entries = []            # 只读模式下要禁用的输入框

        win = self.win = tk.Toplevel(root)
        win.title(TITLE)
        win.geometry("940x660")
        win.minsize(800, 580)
        win.protocol("WM_DELETE_WINDOW", self._close)
        win.bind("<Control-s>", lambda e: self._save_current())

        head = ttk.Frame(win, padding=(8, 8, 8, 0))
        head.pack(fill="x")
        self.banner = ttk.Label(head, wraplength=760, justify="left")
        self.banner.pack(side="left", fill="x", expand=True)
        ttk.Button(head, text="刷新", command=self._refresh_from_server).pack(side="right")

        self.nb = ttk.Notebook(win)
        self.nb.pack(fill="both", expand=True, padx=8, pady=8)
        self._build_products(self.nb)
        self._build_defaults(self.nb)
        self.review_page = self._build_review(self.nb)
        self.history_page = self._build_history(self.nb)
        self.nb.bind("<<NotebookTabChanged>>", self._on_tab)

        self._load()

    # ------------------------------------------------------------ 加载 / 模式
    def _busy(self, on):
        try:
            self.win.configure(cursor="watch" if on else "")
            self.win.update_idletasks()
        except tk.TclError:
            pass

    def _load(self):
        """从服务器读参数；连不上就只读显示本机缓存。"""
        self._busy(True)
        try:
            r = params_sync.fetch()
        except ServerError as e:
            r, self.offline_reason = None, str(e)
        finally:
            self._busy(False)

        self.role = None
        if r is not None and r.get("data") is None:
            if r.get("role") == "admin" and messagebox.askyesno(
                    TITLE, "服务器上还没有产品参数。\n把本机的 pallet_params.json 上传为第 1 版吗？", parent=self.win):
                try:
                    r = {**r, **params_sync.save(r.get("version", 0), self._local_raw() or {}, "首次上传")}
                except ServerError as e:
                    messagebox.showerror(TITLE, f"上传失败：{e}", parent=self.win)
                    r, self.offline_reason = None, f"上传失败：{e}"
            else:
                r, self.offline_reason = None, "服务器上还没有产品参数（等管理员上传）"
        if r is not None:
            self.role = r.get("role") or "user"
            self._set_data(r["data"], r["version"])
            self.pending = r.get("pending") or 0
        else:
            self.raw = params_sync.strip_local(self._local_raw() or {"_defaults": dict(pallet_plan._HARDCODED_DEFAULTS)})
            self.version = None
            self.pending = 0
        self._apply_mode()
        self._refresh_list()
        self._load_defaults_form()
        keys = self._product_keys()
        if self.current in self.raw:
            self._select(self.current)
        elif keys:
            self._select(keys[0])
        else:
            self._new(force=True)
        if self.role:
            self._load_proposals()

    def _local_raw(self):
        try:
            return pallet_plan.load_raw() if pallet_plan.PARAMS_FILE.exists() else None
        except Exception as e:
            messagebox.showerror(TITLE, f"本机的 {pallet_plan.PARAMS_FILE.name} 读取失败：{e}", parent=self.win)
            return None

    def _set_data(self, data, version):
        """服务器返回的新数据：更新界面用的数据，同时写本机缓存。"""
        self.raw, self.version = params_sync.strip_local(data), version
        try:
            if params_sync.local_version() != version:
                params_sync.write_cache(data, version)
        except OSError as e:
            messagebox.showwarning(TITLE, f"写本机缓存失败（不影响服务器上的数据）：{e}", parent=self.win)

    def _apply_mode(self):
        admin, online = self.role == "admin", self.role is not None
        if admin:
            text, color = f"管理员：保存后直接生效。当前第 {self.version} 版。", "#1a7f37"
        elif online:
            text, color = f"修改需要提交给管理员审核，通过后才生效。当前第 {self.version} 版。", "#0b5cad"
        else:
            text, color = f"只读：{self.offline_reason}\n显示的是本机缓存的参数。", "#c62828"
        self.banner.configure(text=text, foreground=color)
        self.save_btn.configure(text="保存 (Ctrl+S)" if admin else "提交审核 (Ctrl+S)")
        self.del_btn.configure(text="删除" if admin else "申请删除")
        self.dsave_btn.configure(text="保存通用参数" if admin else "提交审核")
        st = "normal" if online else "disabled"
        for w in [self.save_btn, self.del_btn, self.new_btn, self.copy_btn, self.revert_btn, self.dsave_btn,
                  *self.entries]:
            w.configure(state=st)
        if online:
            self.nb.tab(self.review_page, state="normal")
            self._set_review_title()
        else:
            self.nb.tab(self.review_page, state="hidden")
        self.nb.tab(self.history_page, state="normal" if admin else "hidden")
        self.approve_btn.pack_forget()
        self.reject_btn.pack_forget()
        self.withdraw_btn.pack_forget()
        if admin:
            self.approve_btn.pack(side="left")
            self.reject_btn.pack(side="left", padx=6)
        else:
            self.withdraw_btn.pack(side="left")

    def _set_review_title(self):
        n = sum(1 for p in self.proposals if p.get("status") == "pending") if self.proposals else self.pending
        name = "审核" if self.role == "admin" else "我的申请"
        self.nb.tab(self.review_page, text=f"{name} ({n})" if n else name)

    def _refresh_from_server(self):
        if not self._confirm_leave():
            return
        self._load()

    def _on_tab(self, _e=None):
        try:
            sel = self.nb.select()
        except tk.TclError:
            return
        if sel == str(self.history_page) and self.role == "admin":
            self._load_history()

    # ------------------------------------------------------------ 数据
    def _product_keys(self):
        return [k for k in self.raw if not k.startswith("_")]

    def _defaults(self):
        return _merged_defaults(self.raw)

    def _commit_admin(self, new_raw, note):
        """管理员直接保存到服务器。成功返回 True。"""
        try:
            r = params_sync.save(self.version, new_raw, note)
        except ServerError as e:
            if e.status == 409 and messagebox.askyesno(TITLE, f"{e}\n\n现在刷新吗？（你这次的修改会丢掉）",
                                                       parent=self.win):
                self.dirty = False
                self._load()
            elif e.status != 409:
                messagebox.showerror(TITLE, f"保存失败：{e}", parent=self.win)
            return False
        self._set_data(r["data"], r["version"])
        self._apply_mode()
        return True

    def _submit(self, action, sku="", old_sku=None, data=None):
        """普通用户提交申请。成功返回 True。"""
        note = simpledialog.askstring(TITLE, "给管理员的说明（选填，比如数据来源）：", parent=self.win)
        if note is None:
            return False
        try:
            params_sync.submit(self.version, action, sku, old_sku, data, note.strip())
        except ServerError as e:
            messagebox.showerror(TITLE, f"提交失败：{e}", parent=self.win)
            return False
        messagebox.showinfo(TITLE, "已提交，管理员审核通过后生效。\n进度可以在\"我的申请\"页查看。", parent=self.win)
        self._load_proposals()
        return True

    # ------------------------------------------------------------ 产品页
    def _build_products(self, nb):
        page = ttk.Frame(nb, padding=8)
        nb.add(page, text="产品")

        left = ttk.Frame(page)
        left.pack(side="left", fill="y")
        self.search = tk.StringVar()
        self.search.trace_add("write", lambda *a: self._refresh_list())
        ttk.Label(left, text="搜索：").pack(anchor="w")
        ttk.Entry(left, textvariable=self.search, width=28).pack(fill="x")
        box = ttk.Frame(left)
        box.pack(fill="both", expand=True, pady=6)
        self.tree = ttk.Treeview(box, columns=("sku",), show="headings", selectmode="browse", height=18)
        self.tree.heading("sku", text="产品物料编码")
        self.tree.column("sku", width=200)
        sb = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        bar = ttk.Frame(left)
        bar.pack(fill="x")
        self.new_btn = ttk.Button(bar, text="新增", command=self._new, width=7)
        self.new_btn.pack(side="left")
        self.copy_btn = ttk.Button(bar, text="复制", command=self._copy, width=7)
        self.copy_btn.pack(side="left", padx=4)
        self.del_btn = ttk.Button(bar, text="删除", command=self._delete, width=8)
        self.del_btn.pack(side="left")

        right = ttk.Frame(page, padding=(16, 0, 0, 0))
        right.pack(side="left", fill="both", expand=True)
        form = ttk.Frame(right)
        form.pack(fill="x")
        self.vars = {}
        r = 0
        ttk.Label(form, text="产品物料编码 *").grid(row=r, column=0, sticky="w", pady=2)
        self.sku = tk.StringVar()
        e = ttk.Entry(form, textvariable=self.sku, width=30)
        e.grid(row=r, column=1, sticky="we", padx=6)
        self.entries.append(e)
        ttk.Label(form, text="和询价表“产品物料”列一致（领星 SKU），忽略大小写和空格",
                  foreground="#666").grid(row=r, column=2, sticky="w")
        self.sku.trace_add("write", self._changed)
        r += 1
        for key, label, kind, hint in PRODUCT_FIELDS:
            if key is None:
                ttk.Label(form, text=label, font=("Microsoft YaHei UI", 9, "bold")).grid(
                    row=r, column=0, columnspan=3, sticky="w", pady=(8, 2))
                r += 1
                continue
            v = tk.StringVar()
            v.trace_add("write", self._changed)
            self.vars[key] = v
            ttk.Label(form, text=label + ("" if key in OPTIONAL else " *")).grid(row=r, column=0, sticky="w", pady=2)
            if key == "packaging":
                e = self.pack_combo = ttk.Combobox(form, textvariable=v, width=28)
            else:
                e = ttk.Entry(form, textvariable=v, width=30)
            e.grid(row=r, column=1, sticky="we", padx=6)
            self.entries.append(e)
            ttk.Label(form, text=hint, foreground="#666").grid(row=r, column=2, sticky="w")
            r += 1
        form.columnconfigure(2, weight=1)

        self.check = self._text(right, 6)
        self.check.pack(fill="both", expand=True, pady=(10, 4))

        bar = ttk.Frame(right)
        bar.pack(fill="x")
        self.save_btn = ttk.Button(bar, text="保存 (Ctrl+S)", command=self._save_current)
        self.save_btn.pack(side="left")
        self.revert_btn = ttk.Button(bar, text="撤销修改", command=self._revert)
        self.revert_btn.pack(side="left", padx=6)
        self.status = ttk.Label(bar, foreground="#666")
        self.status.pack(side="left", padx=8)

    def _text(self, parent, height):
        t = tk.Text(parent, height=height, wrap="word", relief="flat", background=self.win.cget("background"),
                    font=("Microsoft YaHei UI", 9), state="disabled")
        t.tag_configure("ok", foreground="#1a7f37")
        t.tag_configure("err", foreground="#c62828")
        t.tag_configure("warn", foreground="#b26a00")
        t.tag_configure("info", foreground="#333")
        t.tag_configure("head", font=("Microsoft YaHei UI", 9, "bold"))
        return t

    @staticmethod
    def _write_text(t, parts):
        """parts: [(文字, tag), ...]"""
        t.configure(state="normal")
        t.delete("1.0", "end")
        for s, tag in parts:
            t.insert("end", s, tag)
        t.configure(state="disabled")

    def _refresh_list(self):
        want = _norm(self.search.get())
        self.tree.delete(*self.tree.get_children())
        for k in self._product_keys():
            if want in _norm(k):
                self.tree.insert("", "end", iid=k, values=(k,))
        packs = {str(p.get("packaging")) for k, p in self.raw.items()
                 if not k.startswith("_") and isinstance(p, dict) and p.get("packaging")}
        packs.add(str(self._defaults()["packaging"]))
        self.pack_combo.configure(values=sorted(packs))
        if self.current and self.tree.exists(self.current):
            self._loading = True       # 只是同步高亮，不触发切换
            self.tree.selection_set(self.current)
            self.tree.see(self.current)
            self._loading = False

    def _fill(self, sku, p):
        self._loading = True
        try:
            self.sku.set(sku)
            for key, _, kind, _ in PRODUCT_FIELDS:
                if key is not None:
                    self.vars[key].set(_fmt(p.get(key)))
        finally:
            self._loading = False
        self.dirty = False
        self._validate_live()

    def _read_form(self):
        """表单 -> (sku, 产品 dict, 格式错误列表)。保留原产品里表单没有的其它键。"""
        errors = []
        base = dict(self.raw.get(self.current) or {}) if self.current else {}
        base.pop("aliases", None)
        p = {}
        for key, label, kind, _ in PRODUCT_FIELDS:
            if key is None:
                continue
            s = self.vars[key].get()
            try:
                if kind == "text":
                    val = s.strip() or None
                elif kind == "lbq":
                    val = _parse_lbq(s)
                else:
                    val = _parse_num(s, kind, label)
            except ValueError as e:
                errors.append(str(e))
                val = None
            p[key] = val
        for k, v in base.items():       # 表单之外手工加的键原样保留
            if k not in p:
                p[k] = v
        p = {k: v for k, v in p.items() if not (k in OPTIONAL and v is None)}
        return self.sku.get().strip(), p, errors

    def _validate(self, sku, p, raw, old_sku):
        """校验一个产品放进 raw 后是否合格 -> (errors, warnings)"""
        errs, warns = pallet_plan.validate_product(sku, p, _merged_defaults(raw))
        dup = next((k for k in raw if not k.startswith("_") and k != old_sku and _norm(k) == _norm(sku)), None)
        if sku and dup:
            errs.insert(0, f"产品物料编码和已有的 {dup} 重复")
        return errs, warns

    def _check(self):
        """-> (sku, p, errors, warnings)"""
        sku, p, errors = self._read_form()
        if errors:
            return sku, p, errors, []
        errs, warns = self._validate(sku, p, self.raw, self.current)
        return sku, p, errs, warns

    def _changed(self, *_):
        if self._loading:
            return
        self.dirty = True
        self.status.configure(text="有未保存的修改")
        self._validate_live()

    def _summary(self, p, d):
        n = p["per_layer"] * p["max_layers"]
        calc_h = p["max_layers"] * p["carton_height_cm"] + d["pallet_height_cm"]
        calc_w = n * p["carton_weight_kg"] + d["pallet_weight_kg"]
        return (f"每托 {p['per_layer']} 箱/层 × {p['max_layers']} 层 = {n} 箱；"
                f"按箱子算整托 {calc_h:g} cm / {calc_w:g} kg，"
                f"实测 {p['full_pallet_height_cm']:g} cm / {p['full_pallet_weight_kg']:g} kg"
                f"（上限 {d['pallet_total_height_cm']:g} cm / {d['max_pallet_weight_kg']:g} kg）\n")

    def _check_parts(self, errors, warnings, ok_text="✔ 校验通过\n"):
        parts = [(f"✘ {e}\n", "err") for e in errors] + [(f"⚠ {w}（需确认）\n", "warn") for w in warnings]
        if not errors and not warnings:
            parts.append((ok_text, "ok"))
        return parts

    def _validate_live(self):
        sku, p, errors, warnings = self._check()
        parts = [] if errors else [(self._summary(p, self._defaults()), "info")]
        self._write_text(self.check, parts + self._check_parts(errors, warnings))

    def _confirm_leave(self):
        """切换/关闭前处理未保存修改。返回 True = 可以离开。"""
        if not self.dirty or self.role is None:
            self.dirty = False
            return True
        name = self.current or "新产品"
        verb = "保存" if self.role == "admin" else "提交审核"
        ans = messagebox.askyesnocancel(TITLE, f"{name} 有未保存的修改，要{verb}吗？", parent=self.win)
        if ans is None:
            return False
        if ans:
            return self._save_current()
        self.dirty = False
        return True

    def _on_tree_select(self, _e=None):
        if self._loading:
            return
        sel = self.tree.selection()
        if not sel or sel[0] == self.current:
            return
        if not self._confirm_leave():
            self._refresh_list()        # 恢复原来的高亮
            return
        self._select(sel[0])

    def _select(self, key):
        self.current = key
        self._fill(key, self.raw[key])
        self._refresh_list()
        self.status.configure(text="")

    def _new(self, force=False):
        if not force and not self._confirm_leave():
            return
        d = self._defaults()
        self.current = None
        self.tree.selection_remove(*self.tree.selection())
        self._fill("", {"packaging": d["packaging"], "pallet_length_cm": d["pallet_length_cm"],
                        "pallet_width_cm": d["pallet_width_cm"]})
        self.status.configure(text="新产品：填好后保存" if self.role == "admin" else "新产品：填好后提交审核")

    def _copy(self):
        if not self.current:
            messagebox.showinfo(TITLE, "先在左边选一个要复制的产品", parent=self.win)
            return
        if not self._confirm_leave():
            return
        src = self.raw[self.current]
        self.current = None
        self.tree.selection_remove(*self.tree.selection())
        self._fill("", src)
        self.dirty = True
        self.status.configure(text="已复制，填上新的产品物料编码")

    def _delete(self):
        if not self.current:
            return
        sku = self.current
        if self.role == "admin":
            if not messagebox.askyesno(TITLE, f"确定删除 {sku} 吗？（可以在\"历史版本\"里回滚）", parent=self.win):
                return
            new_raw = {k: v for k, v in self.raw.items() if k != sku}
            if not self._commit_admin(new_raw, f"删除 {sku}"):
                return
            self.dirty = False
            self.current = None
            self._refresh_list()
            keys = self._product_keys()
            if keys:
                self._select(keys[0])
            else:
                self._new(force=True)
            self.status.configure(text=f"已删除 {sku}")
        else:
            if not messagebox.askyesno(TITLE, f"申请删除 {sku}？管理员审核通过后才会删除。", parent=self.win):
                return
            if self._submit("delete", sku):
                self.status.configure(text=f"已提交删除 {sku} 的申请")

    def _revert(self):
        if self.current:
            self._fill(self.current, self.raw[self.current])
        else:
            self._new(force=True)
        self.status.configure(text="已撤销修改")

    def _save_current(self):
        """保存（管理员）或提交审核（其他人）当前产品。成功返回 True。"""
        if self.role is None:
            return False
        sku, p, errors, warnings = self._check()
        if errors:
            messagebox.showerror(TITLE, "不能保存，请先改正：\n\n" + "\n".join(f"• {e}" for e in errors),
                                 parent=self.win)
            return False
        if warnings and not messagebox.askyesno(
                TITLE, "\n".join(f"• {w}" for w in warnings) + "\n\n确认数据没错，继续吗？", parent=self.win):
            return False
        if self.role != "admin":
            if self.current and self.current == sku and p == self.raw.get(sku):
                messagebox.showinfo(TITLE, "没有修改，不需要提交", parent=self.win)
                return False
            if not self._submit("upsert", sku, self.current, p):
                return False
            self.dirty = False
            if self.current:
                self._fill(self.current, self.raw[self.current])     # 表单回到当前生效的值
            else:
                self._new(force=True)
            self.status.configure(text=f"已提交 {sku}，等待审核")
            return True
        new_raw = {}
        for k, v in self.raw.items():            # 改名时保持原来的位置
            new_raw[sku if k == self.current else k] = p if k == self.current else v
        if self.current is None:
            new_raw[sku] = p
        verb = "新增" if self.current is None else ("修改" if self.current == sku else f"{self.current} 改名为")
        if not self._commit_admin(new_raw, f"{verb} {sku}"):
            return False
        self.current = sku
        self.dirty = False
        self._refresh_list()
        self._validate_live()
        self.status.configure(text=f"已保存 {sku}（第 {self.version} 版）")
        return True

    # ------------------------------------------------------------ 通用参数页
    def _build_defaults(self, nb):
        page = ttk.Frame(nb, padding=12)
        nb.add(page, text="通用参数")
        ttk.Label(page, text="所有产品共用的托盘底板、高度/重量上限，以及查不到产品时的估算参数。",
                  foreground="#666").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))
        self.dvars = {}
        for i, (key, label, kind, hint) in enumerate(DEFAULT_FIELDS, start=1):
            v = tk.StringVar()
            self.dvars[key] = v
            ttk.Label(page, text=label + " *").grid(row=i, column=0, sticky="w", pady=2)
            e = ttk.Entry(page, textvariable=v, width=24)
            e.grid(row=i, column=1, sticky="w", padx=6)
            self.entries.append(e)
            ttk.Label(page, text=hint, foreground="#666").grid(row=i, column=2, sticky="w")
        bar = ttk.Frame(page)
        bar.grid(row=len(DEFAULT_FIELDS) + 1, column=0, columnspan=3, sticky="w", pady=(12, 0))
        self.dsave_btn = ttk.Button(bar, text="保存通用参数", command=self._save_defaults)
        self.dsave_btn.pack(side="left")
        self.dstatus = ttk.Label(bar, foreground="#666")
        self.dstatus.pack(side="left", padx=8)

    def _load_defaults_form(self):
        d = self._defaults()
        for key, v in self.dvars.items():
            v.set(_fmt(d.get(key)))

    def _defaults_errors(self, d, raw):
        """通用参数 d 用到 raw 里所有产品上是否合格 -> errors"""
        errors = pallet_plan.validate_defaults(d)
        if not errors:      # 上限改小后，已有产品可能不再合格
            for k in (k for k in raw if not k.startswith("_")):
                errs, _ = pallet_plan.validate_product(k, raw[k], {**pallet_plan._HARDCODED_DEFAULTS, **d})
                errors += [f"{k}：{e}" for e in errs]
        return errors

    def _save_defaults(self):
        errors, d = [], {}
        for key, label, kind, _ in DEFAULT_FIELDS:
            s = self.dvars[key].get()
            try:
                d[key] = (s.strip() or None) if kind == "text" else _parse_num(s, kind, label)
            except ValueError as e:
                errors.append(str(e))
        errors = errors or self._defaults_errors(d, self.raw)
        if errors:
            messagebox.showerror(TITLE, "不能保存，请先改正：\n\n" + "\n".join(f"• {e}" for e in errors[:15]),
                                 parent=self.win)
            return
        if all(d[k] == self._defaults().get(k) for k in d):
            messagebox.showinfo(TITLE, "没有修改", parent=self.win)
            return
        if self.role != "admin":
            if self._submit("defaults", data=d):
                self._load_defaults_form()
                self.dstatus.configure(text="已提交，等待审核")
            return
        old = self.raw.get("_defaults") or {}
        new_raw = dict(self.raw)
        new_raw["_defaults"] = {**{k: v for k, v in old.items() if k.startswith("_")}, **d}     # 保留 _说明
        if self._commit_admin(new_raw, "修改通用参数"):
            self.dstatus.configure(text=f"已保存（第 {self.version} 版）")
            self._validate_live()

    # ------------------------------------------------------------ 审核 / 我的申请
    def _build_review(self, nb):
        page = ttk.Frame(nb, padding=8)
        nb.add(page, text="审核")
        box = ttk.Frame(page)
        box.pack(fill="x")
        cols = ("time", "who", "action", "sku", "status")
        self.ptree = ttk.Treeview(box, columns=cols, show="headings", selectmode="browse", height=9)
        for c, t, w in zip(cols, ("提交时间", "提交人", "类型", "产品", "状态"), (130, 90, 80, 220, 80)):
            self.ptree.heading(c, text=t)
            self.ptree.column(c, width=w, stretch=c == "sku")
        sb = ttk.Scrollbar(box, orient="vertical", command=self.ptree.yview)
        self.ptree.configure(yscrollcommand=sb.set)
        self.ptree.pack(side="left", fill="x", expand=True)
        sb.pack(side="left", fill="y")
        self.ptree.tag_configure("pending", foreground="#b26a00")
        self.ptree.bind("<<TreeviewSelect>>", lambda e: self._show_proposal())
        self.pdetail = self._text(page, 14)
        self.pdetail.configure(relief="solid", borderwidth=1)
        self.pdetail.pack(fill="both", expand=True, pady=8)
        bar = ttk.Frame(page)
        bar.pack(fill="x")
        self.approve_btn = ttk.Button(bar, text="通过", command=self._approve)
        self.reject_btn = ttk.Button(bar, text="驳回", command=self._reject)
        self.withdraw_btn = ttk.Button(bar, text="撤回申请", command=self._withdraw)
        ttk.Button(bar, text="刷新列表", command=self._load_proposals).pack(side="right")
        return page

    def _load_proposals(self):
        try:
            self.proposals = params_sync.list_proposals()
        except ServerError as e:
            self._write_text(self.pdetail, [(f"读取申请列表失败：{e}", "err")])
            return
        self.ptree.delete(*self.ptree.get_children())
        for p in self.proposals:
            sku = "通用参数" if p["action"] == "defaults" else p.get("sku", "")
            if p.get("old_sku") and p["old_sku"] != p.get("sku"):
                sku = f"{p['old_sku']} → {sku}"
            action = "新增" if p["action"] == "upsert" and not p.get("old_sku") else params_sync.ACTION[p["action"]]
            self.ptree.insert("", "end", iid=p["id"], tags=(p["status"],), values=(
                params_sync.fmt_time(p.get("created")), p.get("name", ""), action, sku,
                params_sync.STATUS.get(p["status"], p["status"])))
        self._set_review_title()
        self._show_proposal()

    def _selected_proposal(self):
        sel = self.ptree.selection()
        return next((p for p in self.proposals if p["id"] == sel[0]), None) if sel else None

    def _proposal_check(self, p):
        """审核前在当前数据上校验申请 -> (errors, warnings, 对比行)"""
        raw = self.raw
        diff, errors, warnings = [], [], []

        def changes(old, new, labels):
            keys = list(dict.fromkeys([*labels, *old, *new]))
            for k in keys:
                if k.startswith("_") or k == "aliases":
                    continue
                o, n = old.get(k), new.get(k)
                if o != n:
                    diff.append((f"  {labels.get(k, k)}：", "info"))
                    diff.append((f"{_fmt(o) or '（空）'}", "err"))
                    diff.append(("  →  ", "info"))
                    diff.append((f"{_fmt(n) or '（空）'}\n", "ok"))
            if not diff:
                diff.append(("  没有任何变化\n", "warn"))

        if p["action"] == "defaults":
            changes(self._defaults(), {**self._defaults(), **(p.get("data") or {})}, DLABELS)
            errors = self._defaults_errors({**self._defaults(), **(p.get("data") or {})}, raw)
        elif p["action"] == "delete":
            if p["sku"] not in raw:
                errors.append(f"{p['sku']} 已经不存在了")
            else:
                diff.append((f"  删除整个产品 {p['sku']}\n", "err"))
        else:
            old_sku, sku, data = p.get("old_sku"), p["sku"], p.get("data") or {}
            if old_sku and old_sku not in raw:
                errors.append(f"原产品 {old_sku} 已经不存在了（可能被删除或改名）")
            elif not old_sku and sku in raw:
                errors.append(f"{sku} 已经存在（申请提交后被加上了）")
            else:
                if old_sku and old_sku != sku:
                    diff.append((f"  产品物料编码：{old_sku}  →  {sku}\n", "warn"))
                changes(raw.get(old_sku) or {}, data, LABELS)
                e, w = self._validate(sku, data, raw, old_sku)
                errors += e
                warnings += w
        return errors, warnings, diff

    def _show_proposal(self):
        p = self._selected_proposal()
        if not p:
            self._write_text(self.pdetail, [("在上面选一条申请查看详情", "info")] if self.proposals
                             else [("没有申请记录", "info")])
            return
        st = params_sync.STATUS.get(p["status"], p["status"])
        parts = [(f"{p.get('name', '')} 于 {params_sync.fmt_time(p.get('created'))} 提交 · {st}\n", "head")]
        if p.get("note"):
            parts.append((f"说明：{p['note']}\n", "info"))
        if p["status"] != "pending":
            parts.append((f"{p.get('reviewer', '')} 于 {params_sync.fmt_time(p.get('reviewed'))} 处理", "info"))
            if p.get("applied_version"):
                parts.append((f"，生效为第 {p['applied_version']} 版", "info"))
            parts.append(("\n", "info"))
            if p.get("reason"):
                parts.append((f"驳回原因：{p['reason']}\n", "err"))
        if p["status"] == "pending":
            if p.get("base_version") != self.version:
                parts.append((f"提交时基于第 {p.get('base_version')} 版，现在是第 {self.version} 版，"
                              "下面按当前版本对比\n", "warn"))
            errors, warnings, diff = self._proposal_check(p)
            parts.append(("\n修改内容（当前 → 申请）：\n", "head"))
            parts += diff
            parts.append(("\n校验：\n", "head"))
            parts += self._check_parts(errors, warnings)
        else:
            data = p.get("data") or {}
            if data:
                parts.append(("\n申请的内容：\n", "head"))
                labels = DLABELS if p["action"] == "defaults" else LABELS
                parts += [(f"  {labels.get(k, k)}：{_fmt(v)}\n", "info") for k, v in data.items()]
        self._write_text(self.pdetail, parts)
        pending = p["status"] == "pending"
        for b in (self.approve_btn, self.reject_btn):
            b.configure(state="normal" if pending else "disabled")
        self.withdraw_btn.configure(state="normal" if pending else "disabled")

    def _approve(self):
        p = self._selected_proposal()
        if not p or p["status"] != "pending":
            return
        errors, warnings, _ = self._proposal_check(p)
        if errors:
            messagebox.showerror(TITLE, "这条申请在当前参数上不合格，不能通过（可以驳回）：\n\n"
                                 + "\n".join(f"• {e}" for e in errors[:15]), parent=self.win)
            return
        msg = "确定通过这条申请吗？通过后立即对所有人生效。"
        if warnings:
            msg = "\n".join(f"• {w}" for w in warnings) + "\n\n" + msg
        if not self._confirm_leave() or not messagebox.askyesno(TITLE, msg, parent=self.win):
            return
        try:
            r = params_sync.approve(p["id"])
        except ServerError as e:
            messagebox.showerror(TITLE, f"操作失败：{e}", parent=self.win)
            self._load_proposals()
            return
        self._set_data(r["data"], r["version"])
        self._apply_mode()
        self._refresh_list()
        self._load_defaults_form()
        if self.current in self.raw:
            self._select(self.current)
        self._load_proposals()

    def _reject(self):
        p = self._selected_proposal()
        if not p or p["status"] != "pending":
            return
        reason = simpledialog.askstring(TITLE, "驳回原因（会显示给提交人）：", parent=self.win)
        if reason is None:
            return
        try:
            params_sync.reject(p["id"], reason.strip())
        except ServerError as e:
            messagebox.showerror(TITLE, f"操作失败：{e}", parent=self.win)
        self._load_proposals()

    def _withdraw(self):
        p = self._selected_proposal()
        if not p or p["status"] != "pending":
            return
        if not messagebox.askyesno(TITLE, "确定撤回这条申请吗？", parent=self.win):
            return
        try:
            params_sync.withdraw(p["id"])
        except ServerError as e:
            messagebox.showerror(TITLE, f"操作失败：{e}", parent=self.win)
        self._load_proposals()

    # ------------------------------------------------------------ 历史版本（管理员）
    def _build_history(self, nb):
        page = ttk.Frame(nb, padding=8)
        nb.add(page, text="历史版本")
        cols = ("ver", "time", "who", "note")
        self.htree = ttk.Treeview(page, columns=cols, show="headings", selectmode="browse")
        for c, t, w in zip(cols, ("版本", "时间", "修改人", "说明"), (60, 130, 90, 420)):
            self.htree.heading(c, text=t)
            self.htree.column(c, width=w, stretch=c == "note")
        self.htree.pack(fill="both", expand=True)
        bar = ttk.Frame(page)
        bar.pack(fill="x", pady=(8, 0))
        ttk.Button(bar, text="回滚到选中的版本", command=self._rollback).pack(side="left")
        ttk.Label(bar, text="回滚会把选中版本另存为一个新版本，之后的记录都保留", foreground="#666").pack(
            side="left", padx=8)
        ttk.Button(bar, text="刷新列表", command=self._load_history).pack(side="right")
        return page

    def _load_history(self):
        try:
            versions = params_sync.history()
        except ServerError as e:
            messagebox.showerror(TITLE, f"读取历史版本失败：{e}", parent=self.win)
            return
        self.htree.delete(*self.htree.get_children())
        for v in versions:
            cur = "（当前）" if v["version"] == self.version else ""
            self.htree.insert("", "end", iid=str(v["version"]), values=(
                f"{v['version']}{cur}", params_sync.fmt_time(v.get("updated")), v.get("by", ""), v.get("note", "")))

    def _rollback(self):
        sel = self.htree.selection()
        if not sel:
            return
        v = int(sel[0])
        if v == self.version:
            messagebox.showinfo(TITLE, "选中的就是当前版本", parent=self.win)
            return
        if not self._confirm_leave() or not messagebox.askyesno(
                TITLE, f"把参数恢复成第 {v} 版吗？会立即对所有人生效。", parent=self.win):
            return
        try:
            r = params_sync.rollback(self.version, v)
        except ServerError as e:
            messagebox.showerror(TITLE, f"回滚失败：{e}", parent=self.win)
            return
        self._set_data(r["data"], r["version"])
        self._apply_mode()
        self.current = self.current if self.current in self.raw else None
        self._refresh_list()
        self._load_defaults_form()
        if self.current:
            self._select(self.current)
        self._load_history()

    # ------------------------------------------------------------ 关闭
    def _close(self):
        if not self._confirm_leave():
            return
        ParamsEditor._instance = None
        self.win.destroy()


def open_editor(root):
    return ParamsEditor.open(root)


if __name__ == "__main__":      # 单独调试：python params_editor.py
    _root = tk.Tk()
    _root.withdraw()
    if open_editor(_root):
        def _quit_when_closed():
            if ParamsEditor._instance is None:
                _root.destroy()
            else:
                _root.after(200, _quit_when_closed)
        _quit_when_closed()
        _root.mainloop()
