#!/usr/bin/env python3
"""XL FreightPilot —— Dachser 询价 / 建单的窗口界面（作者：Xu Liu / 刘旭）。

浏览器自动化在后台线程里跑（不卡窗口）；脚本里所有 print 显示到"运行记录"，
所有 ask() 变成弹窗。打包成 exe 的入口就是这个文件。
"""
import argparse
import os
import queue
import sys
import threading
import tkinter as tk
import traceback
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk
from tkinter.scrolledtext import ScrolledText

sys.path.insert(0, str(Path(__file__).resolve().parent))
import app_ui  # noqa: E402
import license_check  # noqa: E402
from app_ui import APP_DIR  # noqa: E402

APP_NAME = "XL FreightPilot"
APP_TITLE = f"{APP_NAME} — Freight Quotation & Order Automation"
HELP_FILE = APP_DIR / "使用说明.txt"


def _foreground_rect():
    """Windows 当前最前面窗口的 (x, y, 宽, 高)；取不到（或已最小化）返回 None。"""
    try:
        import ctypes
        from ctypes import wintypes
        u = ctypes.windll.user32
        hwnd = u.GetForegroundWindow()
        r = wintypes.RECT()
        if hwnd and u.GetWindowRect(hwnd, ctypes.byref(r)):
            w, h = r.right - r.left, r.bottom - r.top
            if w > 200 and h > 150 and r.left > -10000:      # 最小化的窗口坐标是 -32000
                return r.left, r.top, w, h
    except Exception:
        pass
    return None


def _virtual_screen():
    """所有显示器拼起来的范围 (x, y, 宽, 高)，用来防止弹窗跑到屏幕外。"""
    try:
        import ctypes
        m = ctypes.windll.user32.GetSystemMetrics
        return m(76), m(77), m(78), m(79)      # SM_XVIRTUALSCREEN / Y / CX / CY
    except Exception:
        return None


class QueueWriter:
    """把 print 的内容转给窗口（后台线程不能直接操作 tkinter）。"""
    def __init__(self, q):
        self.q = q

    def write(self, s):
        if s:
            self.q.put(("log", s))

    def flush(self):
        pass


class App:
    def __init__(self, root):
        self.root = root
        self.q = queue.Queue()
        self.worker = None
        root.title(APP_TITLE)
        root.geometry("820x640")
        root.minsize(640, 480)

        top = ttk.Frame(root, padding=10)
        top.pack(fill="x")

        self.mode = tk.StringVar(value="inquiry")
        r = ttk.Frame(top)
        r.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))
        ttk.Label(r, text="功能：").pack(side="left")
        ttk.Radiobutton(r, text="询价（按询价表填 Enquiry）", value="inquiry", variable=self.mode,
                        command=self._mode_changed).pack(side="left", padx=(0, 16))
        ttk.Radiobutton(r, text="建单（按领星出库单找报价并 create）", value="create", variable=self.mode,
                        command=self._mode_changed).pack(side="left")

        self.xlsx = tk.StringVar()
        ttk.Label(top, text="Excel：").grid(row=1, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.xlsx).grid(row=1, column=1, sticky="we", padx=4)
        ttk.Button(top, text="选择…", command=self._pick_xlsx).grid(row=1, column=2)

        self.dg_dir = tk.StringVar(value=str(APP_DIR / "危险品文件"))
        self.dg_label = ttk.Label(top, text="危险品文件夹：")
        self.dg_label.grid(row=2, column=0, sticky="w", pady=4)
        self.dg_entry = ttk.Entry(top, textvariable=self.dg_dir)
        self.dg_entry.grid(row=2, column=1, sticky="we", padx=4)
        self.dg_btn = ttk.Button(top, text="选择…", command=self._pick_dg)
        self.dg_btn.grid(row=2, column=2)

        opt = ttk.Frame(top)
        opt.grid(row=3, column=0, columnspan=3, sticky="w", pady=(2, 0))
        self.dry = tk.BooleanVar(value=False)
        self.dry_cb = ttk.Checkbutton(opt, variable=self.dry)
        self.dry_cb.pack(side="left")
        self.days_label = ttk.Label(opt, text="    搜索最近")
        self.days_label.pack(side="left")
        self.days = tk.IntVar(value=30)
        self.days_spin = ttk.Spinbox(opt, from_=1, to=365, width=5, textvariable=self.days)
        self.days_spin.pack(side="left")
        self.days_label2 = ttk.Label(opt, text="天的报价")
        self.days_label2.pack(side="left")
        top.columnconfigure(1, weight=1)

        btns = ttk.Frame(root, padding=(10, 0))
        btns.pack(fill="x")
        self.start_btn = ttk.Button(btns, text="开始", command=self._start)
        self.start_btn.pack(side="left")
        self.stop_btn = ttk.Button(btns, text="停止", command=self._stop, state="disabled")
        self.stop_btn.pack(side="left", padx=6)
        ttk.Button(btns, text="使用说明", command=self._help).pack(side="right")
        ttk.Button(btns, text="授权码", command=self._ask_token).pack(side="right", padx=(6, 0))
        ttk.Button(btns, text="清空记录", command=lambda: self.log.delete("1.0", "end")).pack(side="right", padx=6)

        ttk.Label(root, text="运行记录：", padding=(10, 8, 10, 0)).pack(anchor="w")
        ttk.Label(root, text=license_check.NOTICE, foreground="#666", padding=(10, 0, 10, 6)).pack(
            side="bottom", anchor="w")
        self.log = ScrolledText(root, height=20, font=("Microsoft YaHei UI", 9))
        self.log.pack(fill="both", expand=True, padx=10, pady=(0, 6))

        self._mode_changed()
        app_ui.set_ask_handler(self._ask_from_worker)
        sys.stdout = sys.stderr = QueueWriter(self.q)
        root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.pending_dlg = None
        root.bind("<Button-1>", self._raise_pending, add="+")
        root.bind("<FocusIn>", self._raise_pending, add="+")
        root.after(100, self._poll)

    # ------------------------------------------------------------ 界面
    def _mode_changed(self):
        create = self.mode.get() == "create"
        st = "normal" if create else "disabled"
        for w in (self.dg_entry, self.dg_btn, self.days_spin):
            w.configure(state=st)
        self.dry_cb.configure(text="只预览（只搜索匹配，不点 create）" if create
                              else "只预览（只解析 Excel、看托盘方案，不打开网页）")

    def _pick_xlsx(self):
        f = filedialog.askopenfilename(title="选择 Excel", filetypes=[("Excel", "*.xlsx *.xlsm"), ("所有文件", "*.*")])
        if f:
            self.xlsx.set(f)

    def _pick_dg(self):
        d = filedialog.askdirectory(title="选择危险品文件夹", initialdir=self.dg_dir.get() or str(APP_DIR))
        if d:
            self.dg_dir.set(d)

    def _help(self):
        if HELP_FILE.exists():
            os.startfile(HELP_FILE)
        else:
            messagebox.showinfo(APP_NAME, f"找不到 {HELP_FILE.name}")

    def _poll(self):
        """每 100ms 取一次后台线程发来的日志/弹窗。任何一步出错都不能让这个循环停掉，否则界面会卡死。"""
        try:
            while True:
                try:
                    kind, *rest = self.q.get_nowait()
                except queue.Empty:
                    break
                try:
                    if kind == "log":
                        self.log.insert("end", rest[0])
                        self.log.see("end")
                    elif kind == "ask":
                        try:
                            self._show_ask(*rest)
                        except Exception:
                            self._fallback_ask(*rest)
                    elif kind == "done":
                        self.start_btn.configure(state="normal")
                        self.stop_btn.configure(state="disabled")
                except Exception:
                    self.log.insert("end", f"\n[界面出错] {traceback.format_exc()}\n")
        finally:
            self.root.after(100, self._poll)

    def _fallback_ask(self, msg, options, ev, box):
        """自定义弹窗显示失败时，用系统自带的提示框兜底（保证一定能回答，后台不会一直等）。"""
        keys = [k for k, _ in options]
        try:
            if len(options) == 1:
                messagebox.showinfo(APP_NAME, msg)
                box["key"] = keys[0]
            else:
                ok = messagebox.askyesno(APP_NAME, f"{msg}\n\n是 = {options[0][1]}\n否 = {options[1][1]}")
                box["key"] = keys[0] if ok else keys[1]
        except Exception:
            box["key"] = keys[0]
        ev.set()

    def _raise_pending(self, _event=None):
        """点主窗口时，把还没回答的弹窗拉回最前面（防止它被别的窗口挡住找不到）。"""
        dlg = getattr(self, "pending_dlg", None)
        if dlg is not None:
            try:
                if dlg.winfo_exists():
                    dlg.deiconify()
                    dlg.lift()               # 只拉到前面，不抢焦点，主窗口里照样能滚动看记录
            except tk.TclError:
                pass

    # ------------------------------------------------------------ 弹窗（替代命令行的"按回车"）
    def _ask_from_worker(self, msg, options):
        """后台线程调用：把问题交给主线程弹窗，等用户点按钮。"""
        ev, box = threading.Event(), {}
        self.q.put(("ask", msg, options, ev, box))
        ev.wait()
        return box["key"]

    def _show_ask(self, msg, options, ev, box):
        target = _foreground_rect()          # 先记下当前在前面的窗口（多半是浏览器），弹窗放在它正中间
        dlg = tk.Toplevel(self.root)
        dlg.withdraw()                       # 摆好位置再显示，避免先在角落闪一下
        dlg.title(f"{APP_NAME} - 请确认")
        # 不设 transient：主窗口最小化或被挡住时弹窗不会跟着藏起来，任务栏上也有它自己的图标
        dlg.resizable(False, False)
        ttk.Label(dlg, text=msg, padding=16, wraplength=560, justify="left").pack()
        bar = ttk.Frame(dlg, padding=(16, 0, 16, 16))
        bar.pack()

        def choose(k):
            box["key"] = k
            self.pending_dlg = None
            try:
                dlg.destroy()
            finally:
                ev.set()

        for i, (k, label) in enumerate(options):
            b = ttk.Button(bar, text=label, command=lambda k=k: choose(k))
            b.pack(side="left", padx=4)
            if i == 0:
                b.focus_set()
                dlg.bind("<Return>", lambda e, k=k: choose(k))
        # 点右上角 X：有"结束/退出"就当结束，否则当默认项
        keys = [k for k, _ in options]
        dlg.protocol("WM_DELETE_WINDOW", lambda: choose(next((k for k in ("q", "n") if k in keys), keys[0])))
        dlg.update_idletasks()
        dw, dh = dlg.winfo_reqwidth(), dlg.winfo_reqheight()
        if target is None:
            target = (self.root.winfo_rootx(), self.root.winfo_rooty(), self.root.winfo_width(), self.root.winfo_height())
        tx, ty, tw, th = target
        # 边框宽、标题栏高：弹窗还没显示量不到，用一直显示着的主窗口量（样式一样）
        fx, fy = self.root.winfo_rootx() - self.root.winfo_x(), self.root.winfo_rooty() - self.root.winfo_y()
        if not (0 <= fx < 50 and 0 <= fy < 80):
            fx, fy = 8, 31
        ow, oh = dw + 2 * fx, dh + fy + fx                    # 整个弹窗（含标题栏）的大小
        x, y = tx + (tw - ow) // 2, ty + (th - oh) // 2
        vs = _virtual_screen()
        if vs:
            vx, vy, vw, vh = vs
            x = min(max(x, vx), vx + vw - ow)
            y = min(max(y, vy), vy + vh - oh)
        dlg.geometry(f"+{x}+{y}")
        self.pending_dlg = dlg
        dlg.deiconify()
        dlg.attributes("-topmost", True)    # 浏览器在前台时也能看到
        dlg.lift()
        dlg.focus_force()
        # 不用 grab_set（模态）：弹窗一旦被挡住，主窗口会怎么点都没反应

    def _ask_token(self):
        """填写/修改授权码，保存到 xlfp_config.json。返回是否填了。"""
        t = simpledialog.askstring(APP_NAME, f"请输入授权码（向开发者获取：{license_check.CONTACT}）：",
                                   initialvalue=license_check.get_token(), parent=self.root)
        if not t or not t.strip():
            return False
        try:
            license_check.save_token(t)
        except OSError as e:
            messagebox.showerror(APP_NAME, f"授权码保存失败：{e}")
            return False
        return True

    # ------------------------------------------------------------ 运行
    def _start(self):
        xlsx = self.xlsx.get().strip().strip('"')
        if not xlsx or not Path(xlsx).is_file():
            messagebox.showwarning(APP_NAME, "请先选择 Excel 文件")
            return
        if not license_check.get_token() and not self._ask_token():
            return
        mode = self.mode.get()
        try:
            days = int(self.days.get())
        except (tk.TclError, ValueError):
            days = 30
        app_ui.clear_stop()
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.log.insert("end", f"\n{'=' * 20} 开始{'询价' if mode == 'inquiry' else '建单'}：{xlsx} {'=' * 20}\n")
        self.worker = threading.Thread(target=self._run, args=(mode, xlsx, self.dry.get(), days,
                                                               self.dg_dir.get().strip()), daemon=True)
        self.worker.start()

    def _run(self, mode, xlsx, dry, days, dg_dir):
        try:
            print("正在验证授权 ...")
            ok, why = license_check.check()
            if not ok:
                print(why + "\n")
                app_ui.ask(why, [("c", "确定")])
                return
            if mode == "inquiry":
                import dachser_fill
                dachser_fill.main_batch(argparse.Namespace(
                    xlsx=xlsx, dry=dry, yes=False, url=dachser_fill.START_URL, profile="dachser_profile",
                    chromium_path=None, skip_login_wait=False))
            else:
                import create_orders
                create_orders.run(argparse.Namespace(
                    lingxing=xlsx, dry=dry, html=None, days=days, dg_dir=dg_dir or str(create_orders.DG_DIR),
                    o=None, url=create_orders.START_URL, profile="dachser_profile", chromium_path=None,
                    yes=False))
            print("\n完成。\n")
        except SystemExit as e:
            print(f"\n{e}\n")
        except BaseException as e:
            print(f"\n[出错] {e}\n{traceback.format_exc()}")
            s = str(e).lower()
            if "executable doesn't exist" in s or ("chrome" in s and "not found" in s):
                print("看起来电脑上没有安装 Chrome，请先安装 Google Chrome。\n")
        finally:
            self.q.put(("done",))

    def _stop(self):
        app_ui.request_stop()
        self.stop_btn.configure(state="disabled")
        print("\n已请求停止：当前这一单处理完后停止（如果正在等你操作，请在弹窗里点\"结束\"）。\n")

    def _on_close(self):
        if self.worker and self.worker.is_alive():
            if not messagebox.askyesno(APP_NAME, "还在运行，确定要退出吗？浏览器也会一起关闭。\n\n"
                                       "如果浏览器里还有没点 Save 的建单页，请先在浏览器里点 Save 或 Cancel，"
                                       "否则这张单会被 Dachser 锁住一段时间。"):
                return
        self.root.destroy()
        os._exit(0)     # 后台线程可能还在等浏览器，直接结束整个程序


def selftest():
    """打包后检查：能否找到 Chrome、启动浏览器（不登录）。用法: "XL FreightPilot.exe" --selftest"""
    out = APP_DIR / "selftest.txt"
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as pw:
            b = pw.chromium.launch(channel="chrome", headless=True, chromium_sandbox=True)
            p = b.new_page()
            p.set_content("<p id=x>ok</p>")
            r = p.inner_text("#x")
            b.close()
        import openpyxl  # noqa: F401
        import pallet_plan
        pallet_plan.plan_pallets("US2000", 44)
        lic = license_check.check()      # 只报告，不影响自检结果（授权文件可能还没放好）
        out.write_text(f"OK browser={r} app_dir={APP_DIR} license={lic}\n", encoding="utf-8")
    except BaseException:
        out.write_text("FAIL\n" + traceback.format_exc(), encoding="utf-8")


def main():
    if "--selftest" in sys.argv:
        return selftest()
    root = tk.Tk()
    try:
        ttk.Style().theme_use("vista")
    except tk.TclError:
        pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
