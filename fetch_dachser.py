"""用 Playwright 打开 Dachser，让你登录并停在报价历史页，然后读取页面内容。
登录状态保存在 browser_profile 文件夹里，下次通常不用重新登录。"""
from pathlib import Path
from playwright.sync_api import sync_playwright

START_URL = "https://elogistics.dachser.com/"
PROFILE = str(Path(__file__).with_name("browser_profile"))

def fetch_list_html(start_url=START_URL, headless=False, wait=None):
    """返回当前页面 HTML。wait: 可传入函数代替 input()（测试用）。"""
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(PROFILE, headless=headless, viewport={"width": 1400, "height": 900})
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.goto(start_url)
        (wait or (lambda: input("\n请在弹出的浏览器里登录 Dachser，并打开『报价历史』列表页，\n确认能看到带 create 按钮的表格后，回到这里按回车... ")))()
        page.wait_for_selector("table.specialTable", timeout=15000)
        html = page.content()
        ctx.close()
        return html