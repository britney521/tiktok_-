"""离屏冒烟测试：验证界面构建、布局、四步流水线（桩化紫鸟，零网络）。

    QT_QPA_PLATFORM=offscreen python devtools/smoke_gui.py

覆盖（13 段）：
    [1]  主窗口构建（「强制重启客户端」默认关）
    [2]  布局溢出检查 + 截图
    [2b] 输出框的「浏览…」：只换存放目录、文件名不变（④选输入文件那个「浏览…」不算）
    [3]  店铺列表 → 表格列自动推导 → 选中
    [4]  流水线（导出订单 + 抓客户信息）
         ③订单：桩化「导出订单」面板（面板文案/导出历史/下载按钮都由桩给出），
           但**真实跑** tiktok_export 的整套流程与解析（含跳过字段说明行）；
         ④客户：真实走 BrowserTransport 的 JS 构造 / 响应解析
    [4b] 页面来源不对 → 自动换来源重试（客户阶段）
    [5]  仅抓客户信息（从订单 Excel 回读 Order ID 列）
    [6]  所有区域站都看不到这批订单 → 预检后即中止（整批 SKIPPED，
         绝不把 618 条都重复请求一遍）
    [7]  真机场景（2026-10-14）：通用站回 98001008 / eu 站 "No matching route"
         / seller.us.* 同源通 → 自动换到 us 站并**把组合锁定**，之后每条只发 1 次
    [8]  同源写法打不通 → 自动换「绝对 host」写法（并记住，之后单发）
    [9]  商品评分 → 筛 1~3 星覆盖保存 → ⑥ 抓差评客户（端到端）
    [10] 单独执行⑥：读已有评分表，不再开评分页
    [11] 后台店铺：按 shop_region 直接导航（秋图2 = Spain + United States）
    [12] 商品评分真机形状：异步任务 / CSV / 自带订单号
    [13] ⑥ 前置检查：评分表自带订单号 → 没有订单表也不拦
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
from PyQt6.QtWidgets import QApplication, QPushButton, QWidget  # noqa: E402

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
# 接口响应里的 sku_module[].sku_id（客户表 SKU ID 的**主源**）
API_SKUS = {oid: f"1732505384{i:04d}" for i, oid in enumerate(ORDER_IDS)}
# 订单表里的 SKU ID 列（兜底用，和接口值不同，好区分数据到底来自哪条路）
ORDER_SKUS = {oid: f"SKU-{oid[-4:]}" for oid in ORDER_IDS}
# 订单表里另一个同形状的列：商家自定义编码 —— 千万别拿它当 SKU ID
ORDER_SELLER_SKUS = {oid: f"SELLER-{oid[-4:]}" for oid in ORDER_IDS}
# 接口响应里的 sku_module[].seller_sku_name（客户表「商家 SKU」的**主源**）
# 同样刻意和订单表那列取不同的值：两列都得能看出自己走的是哪条路
API_SELLER_SKUS = {oid: f"API-SELLER-{oid[-4:]}" for oid in ORDER_IDS}

# 导出文件的表头（真实导出是 63 列，这里取带客户字段的一小撮）
EXPORT_COLUMNS = [
    "Order ID", "Order Status", "SKU ID", "Seller SKU", "Buyer Nickname",
    "Recipient", "Phone #", "Country", "State", "City", "Zipcode",
    "Address Line 1",
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
               "Platform SKU ID.",
               "Seller sku input by the seller in the product system.",
               "Buyer's nickname on the platform.", None, None, None, None, None,
               None])  # 字段说明行 —— 解析时应被跳过
    for i, oid in enumerate(ORDER_IDS):
        ws.append([oid, "AWAITING_SHIPMENT", ORDER_SKUS[oid],
                   ORDER_SELLER_SKUS[oid], f"buyer_{oid[-4:]}",
                   f"G*** R***{i}", "(+44)751*****87", "英国", "北爱尔兰",
                   "贝尔法斯特", "BT14 8GL", f"*** street {i}"])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# 商品评分导出文件（与订单文件同形状：表头 + 字段说明行 + 数据）
# 真机确认：评分文件里**没有订单号**，只有买家昵称和商品 —— ⑥ 靠昵称回订单表换
RATING_COLUMNS = ["Product ID", "Product Name", "Rating Score", "Buyer Nickname",
                  "Review Count"]
RATING_FILE_NAME = "商品评分-2026-10-15.xlsx"
# 三条评价：5 星（不抓）、2 星、1 星 → 后两条进差评客户表
RATING_SCORES = [5, 2, 1]


def build_rating_xlsx() -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ProductRating"
    ws.append(RATING_COLUMNS)
    ws.append(["Unique ID of the product.", "Name of the product.",
               "Average rating score.", "Buyer's nickname on the platform.",
               "Total number of reviews."])
    for i, score in enumerate(RATING_SCORES):
        # 第 i 条评价就是第 i 个订单那个买家留的
        ws.append([f"17293{i:04d}", f"圣诞树 {i+1}.8m", score,
                   f"buyer_{ORDER_IDS[i][-4:]}", 120 + i])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---- 真机形状的评分导出（2026-10-15 连店铺浏览器抓到的）----------------------
# 是 **CSV**，而且**自带 order_id** —— ⑥ 不用再拿昵称去订单表反查。
RATING_CSV_COLUMNS = ["rating", "review_text", "review_date", "order_id",
                      "product_id", "sku_id"]
RATING_JOB_URL = ("https://seller.us.tiktokshopglobalselling.com/api/v1/review/"
                  "biz_backend/export/job?oec_seller_id=1&fp=verify_mozc")
# 接口给的 file_url **没有协议头**，而且域名是另一个（seller-us.tiktok.com）
RATING_FILE_URL = ("seller-us.tiktok.com/easesafe/seller_review_export/object/"
                   "abc123safe?expire=1792145585&skipCookie=true&sign=deadbeef")


def build_rating_csv() -> bytes:
    lines = [",".join(RATING_CSV_COLUMNS)]
    for i, (oid, score) in enumerate(zip(ORDER_IDS, RATING_SCORES)):
        lines.append(f"{score},,2026-10-1{i} 07:14:13,{oid},"
                     f"17293{i:04d}0,173217294580846263{i}")
    return ("\n".join(lines) + "\n").encode("utf-8")


def _order_get_payload(oid: str) -> dict:
    return {
        "code": 0,
        "message": "success",
        "data": {"main_order": [{
            "main_order_id": oid,
            # 客户表「SKU ID」的主源：sku_module 里的 sku_id（平台 ID）；
            # 「商家 SKU」的主源：同一个节点里的 seller_sku_name（商家自定义编码）
            "sku_module": [{
                "sku_id": API_SKUS[oid],
                "seller_sku_name": API_SELLER_SKUS[oid],
            }],
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
        self.rating_clicks = 0
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
        # ---- 后台店铺（账号下多店铺，如秋图2 = Spain + United States）----
        # 店铺由 URL 里的 shop_region 决定：导航到哪个 shop_region 就是哪家店。
        self.shop_current = "ES\nSpain"
        # ---- 商品评分导出：真机是**异步任务**（2026-10-15 连店铺浏览器抓包）----
        #   False = 早先的假想（点「导出」就同步下载一个 xlsx）
        #   True  = 真机形状：export/start → export/job 回 file_url（裸 host）→
        #           跨域（seller-us.tiktok.com）→ 借同域 robots.txt 当壳子取 CSV
        self.rating_async = False
        self.rating_job_snapshots: list[list[dict]] = []
        self.rating_last_job: list[dict] = []
        self.rating_csv = b""

    def land(self, url: str) -> None:
        """页面落地到某个后台 URL → 右上角切换器就显示那家店（真机行为）。"""
        code = te.shop_region_of(url)
        if not code:
            return
        self.shop_current = {
            "US": "US\nUnited States (HAWKRAY-Vision)",
            "ES": "ES\nSpain",
        }.get(code, f"{code}\n{code} shop")

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

    def click_rating_export(self) -> str:
        """商品评分页的「导出」。

        ``rating_async=False``（老假想）：同步下载型，点了就出 xlsx。
        ``rating_async=True``（真机形状）：只发起一个异步导出任务，不发下载请求。
        """
        self.rating_clicks += 1
        if self.rating_async:
            return "clicked-label-1"
        self.files[RATING_FILE_NAME] = build_rating_xlsx()
        return "clicked-xpath-0"

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
        self.page.land(url)                 # 落到哪个 shop_region 就是哪家店
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
        # --- 后台店铺：只读切换器胶囊文本（ensure_shop 用它把日志写成人话）---
        if "[A-Z]{2}" in s:
            return self.page.shop_current
        # --- 同源壳子：借 seller-us.tiktok.com/robots.txt 当文档看 origin ---
        if "location.origin" in s:
            return self.page.origin
        # --- ⑤ 商品评分：评分页是「点导出 → 直接下载」，没有面板 ---
        if "core-tabs-0-panel-0" in s or "clicked-xpath" in s:
            return self.page.click_rating_export()
        # --- ⑤ 真机形状：拿页面那条 export/job URL 自己续轮询 ---
        if "export/job" in s and "fetch(" in s:
            last = self.page.rating_last_job[-1] if self.page.rating_last_job else {}
            return json.dumps({"code": 0, "status": last.get("status", 2),
                               "job_id": last.get("job_id", 0),
                               "file_url": last.get("file_url", "")})
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
        if "biz_backend/export/job" in expr:      # tiktok_export.RATING_JOB_JS
            # 页面 hook 是累积的：读完了还是老内容（真机行为）
            if self.page.rating_job_snapshots:
                self.page.rating_last_job = self.page.rating_job_snapshots.pop(0)
            return self.page.rating_last_job
        if "__tk_exp" in expr:
            return [self.page.export_records_payload()]
        return default

    # -- 下载 ---------------------------------------------------------------
    def intercept_url(self, url_pattern, click_js, timeout=25.0, marker=None):  # noqa: ARG002
        # 评分导出按钮的 JS 里第一个字面量是 XPath 而不是文件名，单独认
        if "core-tabs-0-panel-0" in click_js or "clicked-xpath" in click_js:
            self.page.click_rating_export()   # 点了才会生成文件（同步下载型）
            if self.page.rating_async:
                return []                     # 真机：评分导出**没有**浏览器下载请求
            name = RATING_FILE_NAME
        else:
            name = _js_literal(click_js)
        if self.page.download(name):
            return [f"https://seller.us.tiktokglobalshop.com/easesafe/"
                    f"oec_trade_order_export/object/{name}?use-disposable=1"]
        return []

    def fetch_bytes(self, url, timeout=180.0, chunk=120_000, tag="__ziniao_dl"):  # noqa: ARG002
        if "easesafe" in url and self.page.rating_async:
            return self.page.rating_csv
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

    # IP 检测：桩里没有紫鸟那个扩展检测页，直接判「通过」。
    # 不桩掉的话，check_ip 会导航 + 轮询 30s 找「检测成功」按钮 —— 假会话永远找不到，
    # 于是**每一段流水线都白等 30s**（冒烟因此从 3 分钟涨到十几分钟，看着像卡死）。
    def _stub_check_ip(browser, ip_url, log=None, timeout=30.0):  # noqa: ARG001
        (log or (lambda *_a: None))("IP 检测通过（冒烟桩：假会话没有紫鸟扩展页）")
        return True

    zb.check_ip = _stub_check_ip


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
    # 「强制重启客户端」默认**开** —— 已经开着紫鸟再启动一个会串到旧进程上，
    # startBrowser 返回的调试端口是旧实例的 → CDP 连错目标。所以默认先清干净。
    assert win.chk_kill.isChecked(), "「强制重启客户端」必须默认勾选"
    assert "关掉已开启的紫鸟" in win.chk_kill.text(), win.chk_kill.text()
    print(f"    ①区默认值: 强制重启={win.chk_kill.isChecked()}｜{win.chk_kill.text()}")

    print("[1b] ①启动环境：勾了「强制重启」→ 先关掉已开启的紫鸟，再起新的")

    class _KillClient:
        """假客户端：只记调用顺序，顺便谎报「IPC 端口已就绪」。"""

        def __init__(self) -> None:
            self.calls: list[str] = []

        def port_open(self, port=None):        # noqa: ARG002
            self.calls.append("port_open")
            return True                        # 故意说「已就绪」，看还会不会去杀

        def kill_process(self, *a, **k):       # noqa: ARG002
            self.calls.append("kill_process")
            return True

        def start_browser(self):
            self.calls.append("start_browser")
            return True

        def update_core(self, on_progress=None):  # noqa: ARG002
            self.calls.append("update_core")
            return True

    real_make, real_kill = zb.make_client, zb.kill_all_clients
    made: list[_KillClient] = []
    killed: list[_KillClient] = []

    def _fake_make(cfg, log=None, root=None):  # noqa: ARG001
        c = _KillClient()
        made.append(c)
        return c

    def _fake_kill_all(cli, log=None):         # noqa: ARG001
        killed.append(cli)
        (log or (lambda *_a: None))("强制关闭已开启的紫鸟客户端…（冒烟桩）")
        return 1

    zb.make_client, zb.kill_all_clients = _fake_make, _fake_kill_all
    try:
        cg.EnvThread(None, None, True).run()
        c1 = made[-1]
        assert killed and killed[-1] is c1, "勾了强制重启却没去关已开启的紫鸟"
        assert "start_browser" in c1.calls, c1.calls
        assert "port_open" not in c1.calls, f"勾了就不该再走「复用」判断：{c1.calls}"
        print("    勾选 → kill_all_clients 之后 start_browser（不看端口是否就绪）")

        cg.EnvThread(None, None, False).run()
        c2 = made[-1]
        assert not any(x is c2 for x in killed), "没勾就只走「复用」判断，不杀"
        assert "port_open" in c2.calls and "start_browser" not in c2.calls, c2.calls
        print("    未勾选 + 端口就绪 → 直接复用，不重启")
    finally:
        zb.make_client, zb.kill_all_clients = real_make, real_kill

    print("[2] 布局溢出检查")
    print("    无溢出" if check_layout(win) else "    布局存在问题")
    shot = SHOTS / "main_window.png"
    win.grab().save(str(shot))
    print(f"    截图: {shot}")

    print("[2b] 输出框的「浏览…」：只换存放目录，文件名不变")
    out_dir = Path(tempfile.mkdtemp(prefix="smoke_out_dir_"))
    dialog_calls: list[str] = []
    picked = {"dir": ""}

    def _fake_pick_dir(parent=None, title="", directory="", *a, **k):
        """桩掉模态目录选择框（offscreen 下真会弹的话就永久阻塞了）。"""
        dialog_calls.append(str(directory))
        return picked["dir"]

    cg.QFileDialog.getExistingDirectory = staticmethod(_fake_pick_dir)
    texts = [b.text() for b in win.findChildren(QPushButton)]
    assert "另存为…" not in texts, f"输出框不该再用「另存为…」：{texts}"
    assert texts.count("浏览…") == 5, texts     # ④选输入文件 1 + ③④⑤⑥输出目录 4
    print(f"    「浏览…」{texts.count('浏览…')} 个，界面已无「另存为…」")

    saved = {n: getattr(win, n).text() for n in
             ("ed_order_out", "ed_cust_out", "ed_rating_out", "ed_bad_out")}
    # ① 换个目录，文件名必须原样带过去（probe 故意放在**另一个**子目录里，
    #    否则「什么都没改」也会让断言假通过）
    src_dir = out_dir / "旧的存放位置"
    src_dir.mkdir()
    probe = src_dir / "某店_订单.xlsx"
    probe.write_text("x")
    picked["dir"] = str(out_dir)
    win.ed_order_out.setText(str(probe))
    win._pick_output_dir(win.ed_order_out, cg.DEFAULT_ORDER_OUT)
    assert win.ed_order_out.text() == str(out_dir / "某店_订单.xlsx"), win.ed_order_out.text()
    assert dialog_calls and dialog_calls[0] == str(src_dir), dialog_calls
    # 上一次选过的目录会成为下一次的起始目录（连着改四个不用重复找）
    win._pick_output_dir(win.ed_rating_out, cg.DEFAULT_RATING_OUT)
    assert win.ed_rating_out.text() == str(out_dir / cg.DEFAULT_RATING_OUT.name), \
        win.ed_rating_out.text()
    assert dialog_calls[-1] == str(out_dir), dialog_calls
    # 框是空的 → 用默认文件名，绝不写成空路径
    win.ed_bad_out.setText("")
    win._pick_output_dir(win.ed_bad_out, cg.DEFAULT_BAD_CUSTOMER_OUT)
    assert win.ed_bad_out.text() == str(out_dir / cg.DEFAULT_BAD_CUSTOMER_OUT.name), \
        win.ed_bad_out.text()
    # 取消（返回空串）→ 一个字都不动
    picked["dir"] = ""
    before = win.ed_cust_out.text()
    win._pick_output_dir(win.ed_cust_out, cg.DEFAULT_CUSTOMER_OUT)
    assert win.ed_cust_out.text() == before, win.ed_cust_out.text()
    print("    只换目录 / 空框用默认名 / 记住上次目录 / 取消不改动 → OK")
    for _n, _v in saved.items():
        getattr(win, _n).setText(_v)

    print("[2c] ③订单 / ⑤评分的「浏览…」：浏览选中一个 xlsx / csv 文件")
    tips = {b.toolTip() for b in win.findChildren(QPushButton) if b.text() == "浏览…"}
    assert any("浏览选择" in t for t in tips), tips        # ③⑤ 弹「打开」式浏览框
    assert any("存放的文件夹" in t for t in tips), tips    # ④⑥ 只选文件夹
    print("    ③⑤=浏览选中文件，④⑥=选文件夹 → OK")

    open_calls: list[tuple[str, str]] = []
    picked_f = {"file": ""}

    def _fake_open_file(parent=None, title="", path="", *a, **k):
        """桩掉模态浏览框（同上，offscreen 下真弹会永久阻塞）。"""
        open_calls.append((str(title), str(path)))
        return picked_f["file"], ""

    cg.QFileDialog.getOpenFileName = staticmethod(_fake_open_file)
    target = out_dir / "手工选的评分.csv"
    target.write_text("x")
    win._last_out_dir = str(out_dir)          # 上一步刚选过这个目录
    win.ed_order_out.setText(str(src_dir / "某店_订单.xlsx"))
    picked_f["file"] = str(target)
    win._pick_output_file(win.ed_order_out, cg.DEFAULT_ORDER_OUT)
    assert win.ed_order_out.text() == str(target), win.ed_order_out.text()
    # 弹的是浏览框（不是「另存为」），起始路径 = 上次选过的目录 + 框里当前的文件名
    assert open_calls[-1][0] == "选择文件", open_calls[-1]
    assert open_calls[-1][1] == str(out_dir / "某店_订单.xlsx"), open_calls[-1]
    # 选中 .csv → 真按 csv 写（不是把 xlsx 内容塞进 .csv 里骗人）
    te.write_rows(["a", "b"], [["1", "2"]], Path(win.ed_order_out.text()))
    assert Path(win.ed_order_out.text()).read_bytes()[:2] != b"PK"
    # ⑤ 走的是同一套
    picked_f["file"] = str(target)
    win._pick_output_file(win.ed_rating_out, cg.DEFAULT_RATING_OUT)
    assert win.ed_rating_out.text() == str(target), win.ed_rating_out.text()
    # 取消（返回空串）→ 一个字都不动
    picked_f["file"] = ""
    before_f = win.ed_rating_out.text()
    win._pick_output_file(win.ed_rating_out, cg.DEFAULT_RATING_OUT)
    assert win.ed_rating_out.text() == before_f, win.ed_rating_out.text()
    print("    选中即填 / 起始路径带原名 / csv 按内容写 / 取消不改动 → OK")
    target.unlink()
    for _n, _v in saved.items():
        getattr(win, _n).setText(_v)

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
        # 按**列名**取，别写死列号 —— 客户表加列时会把这些断言全打歪（踩过）
        cmap = {str(c): i for i, c in enumerate(crows[0])}
        assert [str(c) for c in crows[0]] == tc.CUSTOMER_COLUMNS, "客户表头不一致"
        assert crows[1][cmap["Buyer Nickname"]] == "buyer_8290", \
            crows[1][cmap["Buyer Nickname"]]
        assert crows[3][cmap["Fetch Status"]] == "OK", crows[3][cmap["Fetch Status"]]
        # SKU ID 取的是**接口响应里的 sku_id**，不是订单表的 SKU 列、更不是 Seller SKU
        for r, oid in enumerate(ORDER_IDS, start=1):
            got_sku = crows[r][cmap["SKU ID"]]
            assert got_sku == API_SKUS[oid], f"{oid} 的 SKU ID：{got_sku!r}"
            # 商家 SKU 取的是**同一个节点里的 seller_sku_name**，
            # 不是订单表那列（ORDER_SELLER_SKUS）—— 两列各有各的源，不能串
            got_seller = crows[r][cmap["商家 SKU"]]
            assert got_seller == API_SELLER_SKUS[oid], \
                f"{oid} 的商家 SKU：{got_seller!r}（订单表那列是 " \
                f"{ORDER_SELLER_SKUS[oid]!r}，取错源了）"
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
        # 店铺是**按 URL 直接导航**过去的，所以流水线会先拿目标店区域站（US）当来源。
        # 就堵这个来源（不再写死通用站）：第一个来源被回 code=10000 → 应自动换下一个。
        blocked = zb.origin_of(te.shop_home_url(win.cb_shop.currentText()))
        assert blocked, f"目标店铺没解析出来源：{win.cb_shop.currentText()!r}"
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
        # （visited 里存的是完整 URL，所以按 origin 比 —— 店铺是按带路径的 URL 导航的）
        assert any(zb.origin_of(u) == blocked for u in FAKE.session.visited), \
            FAKE.session.visited
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
        # 这条没有③的 rows 可用 → 兜底表要能回读订单 Excel（且认的是 SKU ID 列）
        again = tc.read_sku_map(order_out, "订单", "Order ID")
        assert again == ORDER_SKUS, f"回读 SKU 不一致：{again}"
        print(f"    仅走 order/get 回读 → OK（SKU ID 取自响应，兜底表 {len(again)} 条）")

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
        # 一条响应都没拿到 → 两列都只能由订单表补：
        # SKU ID 认「SKU ID」列，商家 SKU 认「Seller SKU」列，各认各的
        _sku_col = tc.CUSTOMER_COLUMNS.index("SKU ID")
        skus = {win.tbl_customers.item(r, _sku_col).text()
                for r in range(win.tbl_customers.rowCount())}
        assert skus == set(ORDER_SKUS.values()), f"占位行的 SKU ID 没补上：{skus}"
        _seller_col = tc.CUSTOMER_COLUMNS.index("商家 SKU")
        sellers = {win.tbl_customers.item(r, _seller_col).text()
                   for r in range(win.tbl_customers.rowCount())}
        assert sellers == set(ORDER_SELLER_SKUS.values()), \
            f"占位行的商家 SKU 没补上（或补成了 SKU ID 那列）：{sellers}"
        print(f"    只发 {page.get_calls} 次请求即中止；客户表状态={statuses}"
              f"（占位行 SKU ID / 商家 SKU 已分别由订单表两列补齐）")
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

        print("[9] 商品评分 → 筛选 1~3 星覆盖保存 → ⑥ 抓差评客户")
        rating_out = Path(td) / "商品评分.xlsx"
        bad_out = Path(td) / "客户差评信息.xlsx"
        page.rating_clicks = 0
        page.get_calls = 0
        page.scripts.clear()
        FAKE.session.visited.clear()
        win.ed_rating_out.setText(str(rating_out))
        win.ed_bad_out.setText(str(bad_out))
        win.tbl_rating.setRowCount(0)
        win.tbl_bad.setRowCount(0)
        win._start_pipeline(do_orders=False, do_customers=False,
                            do_rating=True, do_bad=True)
        ok = pump(app, lambda: win.worker is None, timeout=90)
        assert ok, "商品评分流水线超时"
        app.processEvents()
        assert page.rating_clicks >= 1, "没点到评分页的「导出」按钮"
        assert any("/product/rating" in u for u in FAKE.session.visited), \
            f"没有导航到评分页：{FAKE.session.visited}"
        assert rating_out.exists(), "商品评分 Excel 未生成"

        # ⑤：落盘的**只有 1~3 星**（完整版仍在 _评分原始.xlsx）
        rheader, rrows = te.parse_rating(rating_out)
        assert len(rrows) == 2, f"筛选没生效：{len(rrows)} 行"
        assert [r[2] for r in rrows] == [2, 1], [r[2] for r in rrows]
        assert win.tbl_rating.rowCount() == 2, win.tbl_rating.rowCount()
        rcols = [win.tbl_rating.horizontalHeaderItem(i).text()
                 for i in range(win.tbl_rating.columnCount())]
        assert rcols == RATING_COLUMNS, rcols
        assert "1~3 星" in win.log_box.toPlainText(), "日志没写筛了多少条"

        # ⑥：差评 → 昵称 → 订单 → 客户
        assert bad_out.exists(), "客户差评 Excel 未生成"
        bcols = [win.tbl_bad.horizontalHeaderItem(i).text()
                 for i in range(win.tbl_bad.columnCount())]
        assert bcols == cg.BAD_CUSTOMER_COLUMNS, bcols
        assert win.tbl_bad.rowCount() == 2, win.tbl_bad.rowCount()
        got = []
        for r in range(win.tbl_bad.rowCount()):
            got.append((win.tbl_bad.item(r, cg.BAD_CUSTOMER_COLUMNS.index("Order ID")).text(),
                        win.tbl_bad.item(r, cg.BAD_CUSTOMER_COLUMNS.index("评分")).text(),
                        win.tbl_bad.item(r, cg.BAD_CUSTOMER_COLUMNS.index("Fetch Status")).text()))
        assert got == [(ORDER_IDS[1], "2", "OK"), (ORDER_IDS[2], "1", "OK")], got
        # 4 星那条不该出现在差评客户表里
        assert ORDER_IDS[0] not in [g[0] for g in got], got
        # ⑥ 只抓这两单，不多发请求
        assert page.get_calls == 2, f"⑥ 请求次数不对：{page.get_calls}"

        # 落盘文件与表格一致（回读一遍）
        bheader, brows = te.parse_rating(bad_out)
        assert bheader == cg.BAD_CUSTOMER_COLUMNS, bheader
        assert len(brows) == 2, f"差评客户表行数不对：{len(brows)}"
        print(f"    导出点击 {page.rating_clicks} 次；评分 3 → 差评 {len(rrows)} 行 / "
              f"{len(rheader)} 列 → {rating_out.name}")
        print(f"    差评客户 {len(brows)} 行（请求 {page.get_calls} 次）→ {bad_out.name}")

        # ---- 段10：⑤ 和 ⑥ 拆开 —— 只跑「抓取差评客户」 ----
        print("[10] 单独执行⑥：读已有评分表，不再开评分页")
        assert win.btn_rating.text() == "导出商品评分", win.btn_rating.text()
        assert win.btn_bad.text().startswith("抓取差评客户"), win.btn_bad.text()
        raw_out = rating_out.with_name(f"_评分原始{rating_out.suffix}")
        assert raw_out.exists(), "完整版 _评分原始.xlsx 没留下"
        page.rating_clicks = 0
        page.get_calls = 0
        FAKE.session.visited.clear()
        win.tbl_bad.setRowCount(0)
        win._start_pipeline(do_orders=False, do_customers=False, do_bad=True)
        ok = pump(app, lambda: win.worker is None, timeout=90)
        assert ok, "单独跑⑥超时"
        app.processEvents()
        # 关键：没有再点评分页的导出按钮，也没导航到评分页
        assert page.rating_clicks == 0, f"⑥ 不该再导出评分（点了 {page.rating_clicks} 次）"
        assert not any("/product/rating" in u for u in FAKE.session.visited), \
            f"⑥ 不该导航到评分页：{FAKE.session.visited}"
        assert "读取商品评分" in win.log_box.toPlainText(), "日志没写读取评分表"
        # 差评客户表照样出 2 行
        assert win.tbl_bad.rowCount() == 2, win.tbl_bad.rowCount()
        got2 = []
        for r in range(win.tbl_bad.rowCount()):
            got2.append((win.tbl_bad.item(r, cg.BAD_CUSTOMER_COLUMNS.index("Order ID")).text(),
                         win.tbl_bad.item(r, cg.BAD_CUSTOMER_COLUMNS.index("评分")).text()))
        assert got2 == [(ORDER_IDS[1], "2"), (ORDER_IDS[2], "1")], got2
        # 幂等：完整版仍是 3 行，筛过的 out 仍是 2 行（没被盖成筛过的筛过）
        _, raw_rows = te.parse_rating(raw_out)
        _, out_rows = te.parse_rating(rating_out)
        assert len(raw_rows) == 3, f"完整版被改了：{len(raw_rows)} 行"
        assert len(out_rows) == 2, f"评分表行数不对：{len(out_rows)} 行"
        print(f"    未重导评分（点击 0 次）；差评客户 {len(got2)} 行"
              f"（请求 {page.get_calls} 次）；完整版仍 3 行")

        # ---- 段11：后台店铺 —— 按 URL 直接导航到目标店铺（不点右上角切换器）----
        print("[11] 后台店铺：按 shop_region 直接导航（秋图2 = Spain + United States）")
        page.url = "https://seller.eu.tiktokshopglobalselling.com/homepage?shop_region=ES"
        page.shop_current = "ES\nSpain"
        win.cb_shop.setCurrentText("United States")
        FAKE.session.visited.clear()
        win._start_pipeline(do_orders=False, do_customers=False)
        ok = pump(app, lambda: win.worker is None, timeout=90)
        assert ok, "店铺导航段超时"
        app.processEvents()
        assert any("seller.us.tiktokshopglobalselling.com/homepage?shop_region=US" in u
                   for u in FAKE.session.visited), FAKE.session.visited
        assert not any("shop_region=ES" in u for u in FAKE.session.visited), \
            FAKE.session.visited
        assert page.url.startswith("https://seller.us."), page.url
        assert "United States" in page.shop_current, page.shop_current
        text = win.log_box.toPlainText()
        assert "shop_region=US" in text, "日志没写导航到哪个 URL"
        assert "United States" in text, "日志没写当前店铺"
        # 已在目标店铺 → 幂等：只导航 1 次（就是"打开店铺后台"那一次），不再跳第二遍
        FAKE.session.visited.clear()
        win._start_pipeline(do_orders=False, do_customers=False)
        ok = pump(app, lambda: win.worker is None, timeout=90)
        assert ok, "店铺导航段（幂等）超时"
        app.processEvents()
        navs = [u for u in FAKE.session.visited if "shop_region=US" in u]
        assert len(navs) == 1, f"已在 US 却重复导航：{FAKE.session.visited}"
        assert "后台当前店铺已是目标" in win.log_box.toPlainText(), "日志没写已是目标"
        # 认不出的店铺名 → 只告警，不乱导航
        FAKE.session.visited.clear()
        win.cb_shop.setCurrentText("Nowhere Land")
        win._start_pipeline(do_orders=False, do_customers=False)
        ok = pump(app, lambda: win.worker is None, timeout=90)
        assert ok, "店铺导航段（认不出）超时"
        app.processEvents()
        assert not any("Nowhere" in u for u in FAKE.session.visited), \
            FAKE.session.visited
        assert "认不出后台店铺" in win.log_box.toPlainText(), "没有对认不出的店铺名告警"
        print("    ES→US 按 URL 导航 1 次；已在目标 → 幂等；认不出 → 只告警不乱跑")

        # ---- 段12：⑤ 的真机形状（异步导出任务 + CSV + 自带订单号）----
        # 这一段的桩**照着 2026-10-15 真机抓包改的**：
        #   点「导出」→ 没有任何浏览器下载请求（拦不到）→ export/job 回 file_url
        #   → file_url 在 seller-us.tiktok.com（跨域）→ 借同域 robots.txt 当壳子
        #   → 取回来是 CSV（rating,review_text,review_date,order_id,product_id,sku_id）
        #   → ⑥ 直接拿 order_id 抓客户（不再靠昵称反查）
        print("[12] 商品评分：真机形状（异步任务 / CSV / 自带订单号）")
        rating_out2 = Path(td) / "商品评分2.xlsx"
        bad_out2 = Path(td) / "客户差评信息2.xlsx"
        page.rating_async = True
        page.rating_csv = build_rating_csv()
        page.rating_clicks = 0
        page.get_calls = 0
        page.rating_last_job = []
        page.rating_job_snapshots = [
            [{"url": RATING_JOB_URL, "code": 0, "status": 2, "job_id": 21}],
            [{"url": RATING_JOB_URL, "code": 0, "status": 2, "job_id": 21},
             {"url": RATING_JOB_URL, "code": 0, "status": 3, "job_id": 21,
              "file_url": RATING_FILE_URL}],
        ]
        FAKE.session.visited.clear()
        win.ed_rating_out.setText(str(rating_out2))
        win.ed_bad_out.setText(str(bad_out2))
        win.tbl_rating.setRowCount(0)
        win.tbl_bad.setRowCount(0)
        win._start_pipeline(do_orders=False, do_customers=False,
                            do_rating=True, do_bad=True)
        ok = pump(app, lambda: win.worker is None, timeout=90)
        assert ok, "真机形状的评分流水线超时"
        app.processEvents()
        text = win.log_box.toPlainText()
        assert "异步导出任务" in text, "没按异步任务那条路走"
        assert "下载地址" in text, "日志没写拿到的 file_url"
        assert any("seller-us.tiktok.com/robots.txt" in u
                   for u in FAKE.session.visited), \
            f"没借同域静态文件当壳子：{FAKE.session.visited}"
        assert rating_out2.exists(), "商品评分（真机形状）未生成"
        raw2 = rating_out2.with_name("_评分原始.csv")
        assert raw2.exists(), "CSV 完整版没按 .csv 落盘"
        r2header, r2rows = te.parse_rating(rating_out2)
        assert r2header == ["rating", "review_text", "review_date", "order_id",
                            "product_id", "sku_id"], r2header
        assert [r[0] for r in r2rows] == ["2", "1"], [r[0] for r in r2rows]
        assert page.get_calls == 2, f"⑥ 请求次数不对：{page.get_calls}"
        got2 = []
        for r in range(win.tbl_bad.rowCount()):
            got2.append((win.tbl_bad.item(r, cg.BAD_CUSTOMER_COLUMNS.index("Order ID")).text(),
                         win.tbl_bad.item(r, cg.BAD_CUSTOMER_COLUMNS.index("评分")).text()))
        assert got2 == [(ORDER_IDS[1], "2"), (ORDER_IDS[2], "1")], got2
        assert "自带订单号列" in text, "没说明⑥走的是文件自带的订单号"
        print(f"    拦不到下载请求 → 读 export/job 的 file_url → 借 robots.txt 同源取 "
              f"{len(r2rows)} 条差评 → ⑥ 按 order_id 抓 {page.get_calls} 单 → OK")
        page.rating_async = False

        # ---- 段13：⑥ 的前置检查 —— 评分表自带订单号时**不需要订单表** ----
        # 用户真机踩的坑：⑤ 导出的评分表自带 order_id，⑥ 拿订单号直接 order/get 就够，
        # 可旧版守卫**无条件**要求「订单 Excel」存在 → 明明有评分表也被拦下，弹窗还
        # 指着 <店铺>_订单.xlsx，看起来就像⑥读错了文件（其实是检查本身错了）。
        print("[13] ⑥ 前置检查：评分表自带订单号 → 没有订单表也不拦")
        td13 = Path(td) / "no_order"
        td13.mkdir(exist_ok=True)
        rating13 = td13 / "秋图2_商品评分.xlsx"
        te.write_rows(["rating", "review_text", "order_id", "product_id", "sku_id"],
                      [["2", "一般", ORDER_IDS[1], "172930000000001", "1732172945808462634"],
                       ["1", "很差", ORDER_IDS[2], "172930000000001", "1732172945808462635"]],
                      rating13, sheet="商品评分")
        win.ed_rating_out.setText(str(rating13))
        win.ed_bad_out.setText(str(td13 / "秋图2_客户差评信息.xlsx"))
        gone = str(td13 / "秋图2_订单.xlsx")      # 故意指向不存在的订单表
        win.ed_order_out.setText(gone)
        win.ed_cust_in.setText(gone)
        print(f"    「订单 Excel」故意指向不存在的 {Path(gone).name}")

        warns: list[str] = []
        orig_warn = cg.QMessageBox.warning
        cg.QMessageBox.warning = lambda *a, **k: warns.append(
            str(a[2]) if len(a) > 2 else str(a))
        try:
            page.get_calls = 0
            win.tbl_bad.setRowCount(0)
            win._start_pipeline(do_orders=False, do_customers=False, do_bad=True)
            ok = pump(app, lambda: win.worker is None, timeout=90)
            assert ok, "⑥（自带订单号、无订单表）超时"
            app.processEvents()
            assert not warns, f"有评分表（自带订单号）却被拦下来：{warns}"
            assert win.tbl_bad.rowCount() == 2, win.tbl_bad.rowCount()
            assert page.get_calls == 2, f"⑥ 请求次数不对：{page.get_calls}"
            # 连 SKU 兜底列也不该依赖订单表（它来自接口响应的 sku_module）
            skus13 = [win.tbl_bad.item(r, cg.BAD_CUSTOMER_COLUMNS.index("SKU ID")).text()
                      for r in range(win.tbl_bad.rowCount())]
            assert skus13 == [API_SKUS[ORDER_IDS[1]], API_SKUS[ORDER_IDS[2]]], skus13
            # 差评表也带「商家 SKU」——差评是哪个 SKU 吃的，一眼能看出来
            sellers13 = [win.tbl_bad.item(
                r, cg.BAD_CUSTOMER_COLUMNS.index("商家 SKU")).text()
                for r in range(win.tbl_bad.rowCount())]
            assert sellers13 == [API_SELLER_SKUS[ORDER_IDS[1]],
                                 API_SELLER_SKUS[ORDER_IDS[2]]], sellers13
            assert "自带订单号列" in win.log_box.toPlainText(), \
                "没说明⑥走的是文件自带的订单号"

            # 反向：老版「仅昵称」评分表 + 没有订单表 → **才**该拦，且文案要说对原因
            legacy13 = td13 / "老版评分.xlsx"
            te.write_rows(["Product Name", "Rating", "Buyer Nickname"],
                          [["圣诞树", "2", "bob"]], legacy13, sheet="商品评分")
            win.ed_rating_out.setText(str(legacy13))
            warns.clear()
            win._start_pipeline(do_orders=False, do_customers=False, do_bad=True)
            app.processEvents()
            assert len(warns) == 1, f"老版「仅昵称」+ 无订单表 应当拦下：{warns}"
            assert "订单号列" in warns[0], warns[0]
            assert win.worker is None, "被拦下后不该起流水线"
        finally:
            cg.QMessageBox.warning = orig_warn
        print("    自带 order_id + 无订单表 → 不弹窗、直接抓 2 单（SKU 来自响应）；"
              "老版「仅昵称」+ 无订单表 → 才拦，且说明原因")

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
