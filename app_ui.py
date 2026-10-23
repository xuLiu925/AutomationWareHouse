"""命令行和窗口界面共用的东西：
- APP_DIR: 程序所在目录。直接运行 .py 时是脚本目录；打包成 exe 后是 exe 所在目录
  （pallet_params.json、危险品文件夹、登录信息都放这里）
- ask(): 需要人确认/选择时调用。命令行下用 input()；窗口界面运行时由 gui.py 换成弹窗
- 停止标记: 窗口里点"停止"后，批量循环在下一单开始前退出
"""
import sys
import threading
from pathlib import Path

if "__compiled__" in globals() or getattr(sys, "frozen", False):
    APP_DIR = Path(sys.argv[0]).resolve().parent
else:
    APP_DIR = Path(__file__).resolve().parent


def app_path(p):
    """相对路径按程序所在目录解析（双击 exe 时当前目录不一定是 exe 所在目录）。"""
    p = Path(p)
    return p if p.is_absolute() else APP_DIR / p


def _cli_ask(msg, options):
    keys = [k for k, _ in options]
    hint = "，".join(f"{'回车' if i == 0 else k}={label}" for i, (k, label) in enumerate(options))
    while True:
        a = input(f"{msg}\n[{hint}]: ").strip().lower()
        if not a:
            return keys[0]
        if a in keys:
            return a


_handler = _cli_ask
_stop = threading.Event()


def ask(msg, options):
    """options: [(key, 说明), ...]，第一个是默认项（命令行直接回车）。返回选中的 key。"""
    return _handler(msg, options)


def set_ask_handler(f):
    global _handler
    _handler = f


def request_stop():
    _stop.set()


def clear_stop():
    _stop.clear()


def stop_requested():
    return _stop.is_set()
