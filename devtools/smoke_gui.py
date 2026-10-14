"""离屏冒烟测试：验证界面构建、布局、四步流水线（桩化紫鸟，零网络）。

    QT_QPA_PLATFORM=offscreen python devtools/smoke_gui.py

覆盖：
    [1] 主窗口构建 + 布局溢出检查 + 截图
    [2] 店铺列表 → 表格列自动推导 → 选中
    [3] 流水线（导出订单 + 抓客户信息）
        ③订单：桩化「导出订单」面板（面板文案/导出历史/下载按钮都由桩给出），
          但**真实跑** tiktok_export 的整套流程与解析（含跳过字段说明行）；
        ④客户：真实走 BrowserTransport 的 JS 构造 / 响应解析
    [4] 页面来源不对 → 自动换来源重试（客户阶段）
    [5] 仅抓客户信息（从订单 Excel 回读 Order ID 列）
    [6] 所有区域站都看不到这批订单 → 预检后即中止（整批 SKIPPED，
        绝不把 618 条都重复请求一遍）
    [7] 真机场景（2026-10-14）：通用站回 98001008 / eu 站 "No matching route"
        / seller.us.* 同源通 → 自动换到 us 站并**把组合锁定**，之后每条只发 1 次
    [8] 同源写法打不通 → 自动换「绝对 host」写法（并记住，之后单发）
"""

from __future__ import annotations

import io
import json
import os
import re
import sys
import tempfile
import time
from pathlib import Path
from typing import Any

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import openpyxl  # noqa: E402
from PyQt6.QtWidgets import QApplication, QWidget  # noqa: E402

import customer_gui as cg  # noqa: E402
import tiktok_client as tc  # noqa: E402
import tiktok_export as te  # noqa: E402
import ziniao_bridge as zb  # noqa: E402

SHOTS = ROOT / "docs" / "screenshots"
SHOTS.mkdir(parents=True, exist_ok=True)

# 让导出流程在测试里不用真等（生产是 6s / 0.5s）
te.PAGE_SETTLE = 0.0
te.PANEL_SETTLE = 0.01

ORDER_IDS = ["576966020592278290", "576966018927597605", "576965865881115353"]

# 导出文件的表头（真实导出是 63 列，这里取带客户字段的一小撮）
EXPORT_COLUMNS = [
    "Order ID", "Order Status", "Buyer Nickname", "Recipient", "Phone #",
    "Country", "State", "City", "Zipcode", "Address Line 1",
]

STORE = {
    "browserName": "UK店-TikTok",
    "browserOauth": "9a8c00004e204e200000199c4e22000001a133b0a91900000136d6d7ef",
    "siteId": 1,
    "platform_name": "TikTok Shop",
    "store_username": "hanwei_uk",
    "browserIp": "45.12.34.56",
    "proxyType": "静态住宅",
    "isExpired": False,
}


# ----------------------------------------------------------------------------
# 导出文件（真实形状：第 1 行表头、第 2 行字段说明、第 3 行起数据）
# ----------------------------------------------------------------------------
def build_export_xlsx() -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "OrderSKUList"
    ws.append(EXPORT_COLUMNS)
    ws.append(["Platform unique order ID.", "Current order status.",
               "Buyer's nickname on the platform.", None, None, None, None, None,
               None, None])  # 字段说明行 —— 解析时应被跳过
    for i, oid in enumerate(ORDER_IDS):
        ws.append([oid, "AWAITING_SHIPMENT", f"buyer_{oid[-4:]}",
                   f"G*** R***{i}", "(+44)751*****87", "英国", "北爱尔兰",
                   "贝尔法斯特", "BT14 8GL", f"*** street {i}"])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _order_get_payload(oid: str) -> dict:
    return {
        "code": 0,
        "message": "success",
        "data": {"main_order": [{
            "main_order_id": oid,
            "buyer_info_module": {
                "buyer_nickname": f"buyer_{oid[-4:]}",
                "shipping_address": {
                    "items": [
                        {"key": "name", "value": "G*** R***"},
                        {"key": "first_name", "value": "Geordie"},
                        {"key": "last_name", "value": "Reid"},
                        {"key": "phone", "value": "(+44)751*****87"},
                        {"key": "email", "value": "g***@***"},
                        {"key": "zipcode", "value": "BT14 8GL"},
                        {"key": "address", "value": "*** street"},
                        {"key": "address_detail", "value": ""},
                    ],
                    "region": {"name": "United Kingdom", "code": "GB"},
                    "districts": [{"name": "Northern Ireland"}, {"name": "Belfast"}],
                    "address_id": "7512584504721785878",
                },
            },
        }]},
    }


_GET_ID = re.compile(r'main_order_id\\?"?\s*:\s*\[\\?"(\d+)\\?"\]')
# fetch URL 里的 seller_id 查询参数（oec_seller_id= 也会被匹配到，值相同）
_QS_SELLER = re.compile(r"seller_id=(\d{6,})")
_STR = re.compile(r'"((?:[^"\\]|\\.)*)"')


def _js_literal(script: str) -> str:
    """取 JS 里第一个字符串字面量的值（用于解析「点哪个标签/哪一行」）。"""
    for m in _STR.finditer(script):
        try:
            return json.loads('"' + m.group(1) + '"')
        except Exception:
            continue
    return ""


class FakeOrderPage:
    """假的后台订单页：导出面板 + 导出历史 + 下载按钮。"""

    def __init__(self):
        self.url = "about:blank"
        self.origin = ""
        self.panel_open = False
        self.scope = "(未选)"
        self.fmt = "(未选)"
        self.records: list[dict[str, Any]] = []
        self.files: dict[str, bytes] = {}
        self.seq = 0
        self.export_clicks = 0
        self.download_clicks = 0
        self.get_calls = 0
        self.blocked_origins: set[str] = set()
        self.scripts: list[str] = []
        self.permission_error = False
        # 页面自己请求里出现的 seller_id（新的「店铺身份」可信来源）
        self.seller_counts: dict[str, int] = {}
        # cookie 里读到的店铺 id（可能因为多店铺同名 cookie 串味而与上面不同）
        self.cookie_seller_id = "7494876782645708274"
        # 这些登录态看不到订单（服务端语义：按会话判定可见店铺）
        self.denied_sellers: set[str] = set()
        # 每个站自己的 cookie seller_id（cookie 是**按站隔离**的）
        self.host_sellers: dict[str, str] = {}
        # 只允许这些 host 的请求成功（其余回 10000）—— 用来测「换写法」
        self.host_ok_only: set[str] = set()
        # ---- 区域站行为（模拟 2026-10-14 真机日志）----
        #   host_denied  → 98001008（这个站的登录态看不到这批订单）
        #   host_noroute → HTTP 200 "No matching route"（这个站没这个路由）
        #   host_jsfail  → 页面 JS 报错（api16-* 那种跨域/hook 报错）
        #   host_ok      → 只有这些 host 能通（非空时生效）
        self.host_denied: set[str] = set()
        self.host_noroute: set[str] = set()
        self.host_jsfail: set[str] = set()
        self.host_ok: set[str] = set()
        # 打开某站会被重定向到另一站（实测：eu 地址 → 页面落在 us）
        self.redirect: dict[str, str] = {}

    # -- 面板 ---------------------------------------------------------------
    def click_entry(self) -> str:
        self.panel_open = not self.panel_open
        return "clicked"

    def click_label(self, text: str) -> str:
        if text.startswith("筛选出的订单"):
            self.scope = "筛选出的订单"
        elif text.startswith("当前标签页的所有订单"):
            self.scope = "当前标签页的所有订单"
        elif text in ("Excel", "CSV"):
            self.fmt = text
        return "clicked"

    def click_export(self) -> str:
        self.export_clicks += 1
        self.seq += 1
        key = f"exportfilekey{self.seq:04d}"
        name = f"全部 笔订单-2026-10-13-08_0{self.seq}.xlsx"
        self.records.insert(0, {
            "file_name": _url_quote(name), "download": 2, "file_key": key,
            "export_time": 1791896000 + self.seq,
            "export_task_id": f"81322168074921085{self.seq:02d}",
        })
        self.files[name] = build_export_xlsx()
        return "clicked-1"

    def export_records_payload(self) -> str:
        return json.dumps({"code": 0, "data": {"export_records": list(self.records)},
                           "message": "success"})

    def download(self, name: str) -> bool:
        if name not in self.files:
            return False
        self.download_clicks += 1
        return True

    # -- 运输层（order/get）-----------------------------------------------
    def transport(self, script: str) -> str:
        if self.origin in self.blocked_origins:
            return json.dumps({"status": 200,
                               "text": json.dumps({"code": 10000, "message": None})})
        hm = re.search(r"https://([^/\"']+)", script)
        host = hm.group(1) if hm else ""
        # ---- 区域站行为：us 通 / eu 无路由 / 通用站无权限（真机日志）----
        if host and host in self.host_jsfail:
            self.get_calls += 1
            raise RuntimeError("页面 JS 报错：Cannot read properties of "
                               "undefined (reading 'headers')")
        if host and host in self.host_noroute:
            self.get_calls += 1
            return json.dumps({"status": 200, "text": "No matching route"})
        if host and (host in self.host_denied
                     or (self.host_ok and host not in self.host_ok)):
            self.get_calls += 1
            return json.dumps({"status": 200, "text": json.dumps(
                {"code": 98001008, "data": None, "message": "需要权限"})})
        if self.host_ok_only:
            if host not in self.host_ok_only:
                # 这个写法不对（host/前缀不对）→ 页面来源是对的，但接口打不通
                self.get_calls += 1
                return json.dumps({"status": 200,
                                   "text": json.dumps({"code": 10000, "message": None})})
        m = _GET_ID.search(script)
        oid = m.group(1) if m else ORDER_IDS[0]
        sid_m = _QS_SELLER.search(script)
        sid = sid_m.group(1) if sid_m else ""
        self.get_calls += 1
        if self.permission_error or (sid and sid in self.denied_sellers):
            # 模拟「cookie 是别的店铺的」：登录态没权限看这批订单
            return json.dumps({"status": 200, "text": json.dumps(
                {"code": 98001008, "data": None, "message": "需要权限"})})
        return json.dumps({"status": 200, "text": json.dumps(_order_get_payload(oid))})


def _url_quote(text: str) -> str:
    import urllib.parse

    return urllib.parse.quote(text)


class FakeSession:
    """模拟 ziniao_bridge.BrowserSession（原生 CDP 版）的对外接口。

    真实接口里被用到的是：navigate / current_url / add_init_script / eval /
    run_js / read_json / intercept_url / fetch_bytes / cookies / close。
    """

    def __init__(self, *a, **kw):
        self.page = FakeOrderPage()
        self.closed = False
        self.visited: list[str] = []
        self.init_scripts: list[str] = []

    # -- 基础 ---------------------------------------------------------------
    def navigate(self, url, retry=5, interval=2.0, wait=60.0):  # noqa: ARG002
        host = zb._host_of(url)
        target = self.page.redirect.get(host, host)   # 区域站之间会互相重定向
        if target and target != host:
            url = url.replace(host, target, 1)
        self.page.url = url
        self.page.origin = zb.origin_of(url)
        self.page.panel_open = False
        if "/order?" in url:
            self.page.records.clear()      # 新文档 = hook 重新开始
        self.visited.append(url)
        return self

    def ensure_on(self, url):
        host = zb.origin_of(url)
        if host and host in self.page.url:
            return self
        return self.navigate(url)

    def current_url(self):
        return self.page.url

    def add_init_script(self, source):
        self.init_scripts.append(source)
        return "stub-init"

    def eval(self, script, await_promise=True, timeout=120):  # noqa: ARG002
        return self._dispatch(script)

    def run_js(self, script, timeout=120, retry=2):  # noqa: ARG002
        return self._dispatch(script)

    def _dispatch(self, script: str):
        self.page.scripts.append(script)
        s = script
        # --- 导出面板相关 ---
        if "export_entry" in s or "XPathResult" in s and "indexOf" not in s:
            return self.page.click_entry()
        if te.PANEL_MARKER in s and "offsetParent" in s:
            return "1" if self.page.panel_open else ""
        if "label" in s and "indexOf(want) === 0" in s:
            return self.page.click_label(_js_literal(s))
        if "=== '导出'" in s:
            return self.page.click_export()
        if "'checked'" in s:
            return json.dumps([self.page.scope, self.page.fmt])
        if "=== '下载'" in s:
            return json.dumps({"ok": self.page.download(_js_literal(s))})
        # --- 运输层（BrowserTransport 的 fetch JS）---
        if "/api/fulfillment" in s or "fetch(" in s:
            return self.page.transport(s)
        return ""

    def read_json(self, expr, default=None):
        if "searchParams" in expr:          # tiktok_export.PAGE_SELLER_JS
            return dict(self.page.seller_counts)
        if "__tk_exp" in expr:
            return [self.page.export_records_payload()]
        return default

    # -- 下载 ---------------------------------------------------------------
    def intercept_url(self, url_pattern, click_js, timeout=25.0, marker=None):  # noqa: ARG002
        name = _js_literal(click_js)
        if self.page.download(name):
            return [f"https://seller.us.tiktokglobalshop.com/easesafe/"
                    f"oec_trade_order_export/object/{name}?use-disposable=1"]
        return []

    def fetch_bytes(self, url, timeout=180.0, chunk=120_000, tag="__ziniao_dl"):  # noqa: ARG002
        name = url.split("/object/", 1)[-1].split("?", 1)[0]
        data = self.page.files.get(name)
        if data is None:
            raise zb.ZiniaoBridgeError(f"桩里没有这个文件：{name}")
        return data

    # -- 其余 ---------------------------------------------------------------
    def cookies(self, domain_keyword=None, url=None):  # noqa: ARG002
        sid = self.page.cookie_seller_id
        if url:                      # 按站取：每个站带自己的那份 seller_id
            host = url.split("//")[-1].split("/")[0]
            sid = self.page.host_sellers.get(host, sid)
        return {"oec_seller_id_unified_seller_env": sid,
                "app_id_unified_seller_env": "6556",
                "sessionid": "stub-session"}

    def store_context(self, url, expect="", timeout=25.0, interval=1.5):  # noqa: ARG002
        sid = self.page.cookie_seller_id
        return {"cookies": dict(self.cookies()),
                "seller_id": sid, "stable": True, "seen": [sid],
                "candidates": [{"name": "oec_seller_id_unified_seller_env",
                                "value": sid,
                                "domain": "seller.eu.tiktokshopglobalselling.com",
                                "path": "/"}]}

    def reconnect(self):
        return True

    def alive(self):
        return True

    def tabs(self):
        return [{"id": "stub", "type": "page", "url": self.page.url}]

    def close(self):
        self.closed = True


class FakeClient:
    """模拟 ZiniaoClient 中流水线用到的那几个方法。"""

    def __init__(self, *a, **kw):
        self.opened: list[Any] = []
        self.closed: list[str] = []
        self.ip_checked = False
        self.session = FakeSession()

    def open_store(self, store, **kw):
        self.opened.append(store)
        return {
            "browserOauth": (store.get("browserOauth") if isinstance(store, dict) else store),
            "debuggingPort": 9222,
            "ipDetectionPage": "https://ip.ziniao.example/check",
            "launcherPage": "https://seller.tiktokshopglobalselling.com",
        }

    def open_ip_check(self, browser, url, timeout=60):
        self.ip_checked = True
        return True

    def close_store(self, key):
        self.closed.append(str(key))

    def get_exit(self):
        pass


# 所有被测代码共用的假客户端实例
FAKE = FakeClient()


def install_stubs() -> None:
    zc = zb.load_ziniao()["client"]
    zb.make_client = lambda cfg, log=None, root=None: FAKE
    zc.ZiniaoClient = lambda *a, **kw: FAKE  # 以防有别的路径直接实例化

    zb.open_session = lambda port, log=None, home_urls=None, retry=20, interval=1.0: FAKE.session
    zb.connect_browser = zb.open_session
    zb.close_store = lambda client, key, log=None: client.close_store(key)


def pump(app: QApplication, predicate, timeout: float = 30.0) -> bool:
    """在离屏环境里跑事件循环，直到 predicate 成立或超时。"""
    end = time.time() + timeout
    while time.time() < end:
        app.processEvents()
        if predicate():
            return True
        time.sleep(0.02)
    return predicate()


def check_layout(win: cg.MainWindow) -> bool:
    app = QApplication.instance()
    app.processEvents()
    ok = True
    width = win.centralWidget().width()
    for w in win.centralWidget().findChildren(QWidget):
        if not w.isVisible():
            continue
        g = w.geometry()
        if g.x() + g.width() > width + 2 and w.parent() is win.centralWidget():
            print(f"  [!] 溢出: {type(w).__name__} right={g.x() + g.width()} > {width}")
            ok = False
    return ok


def main() -> int:
    app = QApplication(sys.argv)
    install_stubs()

    print("[1] 构建主窗口")
    win = cg.MainWindow()
    win.resize(1360, 900)
    win.show()
    app.processEvents()
    print(f"    紫鸟项目: {win.ed_root.text().strip()}")
    print(f"    账号: {win.lbl_account.text()}")
    print(f"    ③导出选项: 范围={cg.DEFAULT_EXPORT_SCOPE}｜格式={cg.DEFAULT_EXPORT_FORMAT}"
          "（界面已固定走默认，无对应控件）")

    print("[2] 布局溢出检查")
    print("    无溢出" if check_layout(win) else "    布局存在问题")
    shot = SHOTS / "main_window.png"
    win.grab().save(str(shot))
    print(f"    截图: {shot}")

    print("[3] 店铺列表 → 表格列推导")
    win._on_stores(True, [STORE, dict(STORE, browserName="US店", siteId=2)], "共 2 个店铺")
    app.processEvents()
    cols = [win.tbl_stores.horizontalHeaderItem(i).text()
            for i in range(win.tbl_stores.columnCount())]
    assert win.tbl_stores.rowCount() == 2, win.tbl_stores.rowCount()
    print(f"    列: {cols}")
    win.ed_filter.setText("US")
    app.processEvents()
    assert win.tbl_stores.rowCount() == 1, "过滤后应为 1 行"
    win.ed_filter.clear()
    app.processEvents()
    win.tbl_stores.selectRow(0)
    app.processEvents()
    print(f"    {win.lbl_selected.text()}")

    print("[4] 流水线（导出订单 + 抓客户信息）")
    with tempfile.TemporaryDirectory() as td:
        order_out = Path(td) / "订单.xlsx"
        cust_out = Path(td) / "客户.xlsx"
        win.ed_order_out.setText(str(order_out))
        win.ed_cust_out.setText(str(cust_out))
        win.ed_cust_in.setText(str(order_out))
        win.sp_interval.setRange(0.0, 60.0)  # 冒烟加速：临时放开 UI 的 1.5s 下限
        win.sp_interval.setValue(0.0)
        win.chk_skip_ip.setChecked(True)

        win._start_pipeline(do_orders=True, do_customers=True)
        ok = pump(app, lambda: win.worker is None, timeout=60)
        assert ok, "流水线超时未结束"
        app.processEvents()

        assert order_out.exists(), "订单 Excel 未生成"
        assert cust_out.exists(), "客户 Excel 未生成"
        assert win.tbl_orders.rowCount() == len(ORDER_IDS), win.tbl_orders.rowCount()
        # 订单表列 = 导出文件自己的表头
        tcols = [win.tbl_orders.horizontalHeaderItem(i).text()
                 for i in range(win.tbl_orders.columnCount())]
        assert tcols == EXPORT_COLUMNS, tcols
        assert win.tbl_customers.rowCount() == len(ORDER_IDS), win.tbl_customers.rowCount()

        back = tc.read_order_ids(order_out, "订单", "Order ID")
        assert back == ORDER_IDS, f"回读订单不一致: {back}"

        # 解析必须跳过「字段说明行」
        header, rows = te.parse_export(order_out)
        assert header == EXPORT_COLUMNS, header
        assert len(rows) == len(ORDER_IDS), f"字段说明行没被跳过：{len(rows)} 行"
        assert rows[0][0] == ORDER_IDS[0], rows[0][0]

        ws = openpyxl.load_workbook(cust_out).active
        crows = list(ws.iter_rows(values_only=True))
        assert [str(c) for c in crows[0]] == tc.CUSTOMER_COLUMNS, "客户表头不一致"
        assert crows[1][1] == "buyer_8290", crows[1][1]
        assert crows[3][16] == "OK", crows[3][16]
        assert FAKE.opened and FAKE.closed, "店铺未正确开关"

        # ★ 抓客户前必须把页面停到**订单页**（2026-10-14 用户真机定案的
        #   98001008 第一前提：站对、页不对照样报需要权限）★
        logs = win.log_box.toPlainText()
        assert "订单页" in logs, "抓客户前没有把页面导航到订单页"
        # 日志必须写明身份从哪来 —— 否则用户会怀疑「是不是读了本地 cookie 文件」
        assert "非本地文件" in logs, "日志没写清登录态来源"
        # 日志必须**同时落盘**：exe 是窗口程序，Windows 用户只能靠这个文件把现场发出来
        logf = getattr(win, "_log_path", None)
        assert logf is not None and Path(logf).is_file(), f"日志没落盘：{logf}"
        assert "已加载紫鸟项目" in Path(logf).read_text(encoding="utf-8",
                                                   errors="replace"), \
            "落盘日志与界面日志不一致"

        print(f"    导出点击 {FAKE.session.page.export_clicks} 次，"
              f"下载点击 {FAKE.session.page.download_clicks} 次；"
              f"订单 {len(rows)} 行 / 客户 {win.tbl_customers.rowCount()} 行")
        print(f"    订单表列数 {len(tcols)}（= 导出文件表头）；店铺已关闭={bool(FAKE.closed)}")
        assert FAKE.session.closed, "会话未断开"

        print("[4b] 客户阶段：页面来源不对 → 自动换来源重试")
        blocked = zb.origin_of("https://seller.tiktokshopglobalselling.com")
        FAKE.session.page.blocked_origins = {blocked}
        FAKE.session.visited.clear()
        FAKE.session.page.scripts.clear()
        win.chk_skip_ip.setChecked(False)   # 顺带跑一遍 IP 检测（桩里直接返回已检测到）
        win.tbl_customers.setRowCount(0)
        win._start_pipeline(do_orders=False, do_customers=True)
        ok = pump(app, lambda: win.worker is None, timeout=60)
        assert ok, "换来源重试超时"
        app.processEvents()
        assert win.tbl_customers.rowCount() == len(ORDER_IDS), win.tbl_customers.rowCount()
        # 第一个来源返回 code=10000，应换到下一个来源并成功
        assert blocked in FAKE.session.visited, FAKE.session.visited
        last_origin = zb.origin_of(FAKE.session.visited[-1])
        assert last_origin != blocked, f"没有换来源：{FAKE.session.visited}"
        print(f"    先在 {blocked} 被回 code=10000，换到 {last_origin} 后成功 → OK")
        FAKE.session.page.blocked_origins = set()

        print("[5] 仅抓客户信息（从订单 Excel 回读）")
        FAKE.session.page.scripts.clear()
        win.tbl_customers.setRowCount(0)
        win._start_pipeline(do_orders=False, do_customers=True)
        ok = pump(app, lambda: win.worker is None, timeout=60)
        assert ok, "第二段流水线超时"
        app.processEvents()
        assert win.tbl_customers.rowCount() == len(ORDER_IDS), win.tbl_customers.rowCount()
        print("    仅走 order/get 回读 → OK")

        print("[6] 所有区域站都看不到这批订单 → 预检后即中止（整批 SKIPPED）")
        page = FAKE.session.page
        page.blocked_origins = set()
        page.seller_counts = {"7496000161034505075": 12}  # 页面自己用的是「619 单」那家
        page.cookie_seller_id = "7494876782645708274"     # cookie 读到的是「13 单」那家
        page.denied_sellers = set()
        page.host_sellers = {}
        page.host_ok_only = set()
        page.host_jsfail = set()
        page.host_noroute = set()
        page.host_ok = set()
        page.redirect = {}
        # 每个站都回 98001008 —— 这批订单不在这份登录态名下的任何站
        page.host_denied = {
            "seller.tiktokshopglobalselling.com",
            "seller.eu.tiktokshopglobalselling.com",
            "seller.us.tiktokshopglobalselling.com",
            "api16-normal-no1a.tiktokshopglobalselling.com",
        }
        page.get_calls = 0

        win.tbl_customers.setRowCount(0)
        win._start_pipeline(do_orders=True, do_customers=True)
        ok = pump(app, lambda: win.worker is None, timeout=60)
        assert ok, "身份不一致用例超时"
        app.processEvents()

        # 「来源 × 写法」只扫一遍就单发：绝不把 618 条都跑一遍
        assert page.get_calls <= 20, f"预检没拦住，发了 {page.get_calls} 次请求"
        _st_col = tc.CUSTOMER_COLUMNS.index("Fetch Status")
        statuses = {win.tbl_customers.item(r, _st_col).text()
                    for r in range(win.tbl_customers.rowCount())}
        assert statuses == {"SKIPPED"}, statuses
        print(f"    只发 {page.get_calls} 次请求即中止；客户表状态={statuses}")
        assert win.lbl_selected.text(), "选中店铺标签丢了"

        print("[7] 真机场景：通用站 98001008 / eu 无路由 / us 同源通 → 锁定 us 站")
        page.seller_counts = {}
        page.cookie_seller_id = "7496000161034505075"
        page.denied_sellers = set()
        page.host_sellers = {}
        page.host_denied = {"seller.tiktokshopglobalselling.com"}
        page.host_noroute = {"seller.eu.tiktokshopglobalselling.com"}
        page.host_jsfail = {"api16-normal-no1a.tiktokshopglobalselling.com"}
        page.host_ok = {"seller.us.tiktokshopglobalselling.com"}
        page.redirect = {"seller.eu.tiktokshopglobalselling.com":
                         "seller.us.tiktokshopglobalselling.com"}
        page.get_calls = 0
        FAKE.session.visited.clear()
        win.tbl_customers.setRowCount(0)
        win._start_pipeline(do_orders=False, do_customers=True)
        ok = pump(app, lambda: win.worker is None, timeout=60)
        assert ok, "区域站用例超时"
        app.processEvents()
        statuses = {win.tbl_customers.item(r, _st_col).text()
                    for r in range(win.tbl_customers.rowCount())}
        assert statuses == {"OK"}, statuses
        host = zb._host_of(FAKE.session.page.url)
        assert host == "seller.us.tiktokshopglobalselling.com", \
            f"页面最终应停在 us 站，实际 {host}"
        assert page.get_calls <= 3 * len(ORDER_IDS) + 5, f"扫站太啰嗦：{page.get_calls}"
        print(f"    3 条全部成功；页面最终停在 {host}；"
              f"共 {page.get_calls} 次请求 → OK")

        print("[8] 同源写法打不通 → 自动换「绝对 host」写法")
        page.seller_counts = {}
        page.denied_sellers = set()
        page.host_sellers = {}
        page.cookie_seller_id = "7494876782645708274"
        page.host_denied = set()
        page.host_noroute = set()
        page.host_jsfail = set()
        page.host_ok = set()
        page.redirect = {}
        # 只有抓包留档里的 api16-* host 通；同源 seller.* 一律回 10000
        page.host_ok_only = {"api16-normal-no1a.tiktokshopglobalselling.com"}
        page.get_calls = 0
        win.tbl_customers.setRowCount(0)
        win._start_pipeline(do_orders=False, do_customers=True)
        ok = pump(app, lambda: win.worker is None, timeout=60)
        assert ok, "换写法用例超时"
        app.processEvents()
        statuses = {win.tbl_customers.item(r, _st_col).text()
                    for r in range(win.tbl_customers.rowCount())}
        assert statuses == {"OK"}, statuses
        # 第一条会先试同源（失败）再试绝对（成功），之后记住写法单发
        assert page.get_calls <= 2 * len(ORDER_IDS) + 2, f"换写法太啰嗦：{page.get_calls}"
        print(f"    同源不通 → 换绝对 host 后成功（共 {page.get_calls} 次请求，状态={statuses}）")

        page.denied_sellers = set()
        page.host_sellers = {}
        page.host_ok_only = set()
        page.host_denied = set()
        page.host_noroute = set()
        page.host_jsfail = set()
        page.host_ok = set()
        page.redirect = {}
        page.cookie_seller_id = "7494876782645708274"
        print("    恢复桩状态")

    # 冒烟跑出来的运行日志别留在工作区（它含本机路径与账号名；.gitignore 也挡了）
    try:
        leftover = getattr(win, "_log_path", None)
        if leftover and Path(leftover).is_file():
            Path(leftover).unlink()
            print("    已清理冒烟产生的运行日志")
    except Exception:
        pass

    print("\n全部通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())
