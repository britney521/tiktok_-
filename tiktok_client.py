"""TikTok Shop（全球跨境店）订单与客户信息客户端。

两条链路：
  1) 订单列表  POST /api/fulfillment/order/list  → 分页拉全量订单，写成订单 Excel
  2) 订单详情  POST /api/fulfillment/order/get   → 按 Order ID 取买家收货信息

凭据来源有两种：
  - 浏览器会话（推荐）：``OrderClient.from_cookies(cookies, transport=...)``，
    cookie 由紫鸟店铺浏览器实时提供，不会过期
  - 本地文件：``OrderClient.from_cookie_file()``，读同目录 cookies.json（或内置默认值）

关键结论（实测）：
  - ``order/get``（订单详情）**不需要签名参数**，本地 requests 直连即可
    （aid=4068 + seller_id → code=0），所以客户信息那一段有可靠兜底
  - ``order/list``（订单列表）本地直连**恒定**返回裸 ``code=10000``
    （aid=4068 / 6556 / 不带 aid 都试过），必须在店铺浏览器页面内带签名发起，
    所以要用 ``transport=BrowserTransport``
  - seller_id 直接取自 cookie ``oec_seller_id_unified_seller_env``；
    aid 取自 ``app_id_unified_seller_env``
"""

from __future__ import annotations

import json
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Iterable, Sequence

import requests

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_COOKIE_FILE = BASE_DIR / "cookies.json"

# 跨境店卖家后台的 API 宿主。各区域不同，实测本账号可用的是 no1a（EU/UK）。
DEFAULT_API_BASE = "https://api16-normal-no1a.tiktokshopglobalselling.com"
ORDER_GET_PATH = "/api/fulfillment/order/get"
ORDER_LIST_PATH = "/api/fulfillment/order/list"

ORDER_GET_URL = DEFAULT_API_BASE + ORDER_GET_PATH
ORDER_LIST_URL = DEFAULT_API_BASE + ORDER_LIST_PATH

# ----------------------------------------------------------------------------
# 内置默认凭据（来源：detail.py / list.py，失效时请更新 cookies.json）
# ----------------------------------------------------------------------------
DEFAULT_COOKIES: dict[str, str] = {
    "_m4b_theme_": "new",
    "passport_csrf_token": "14561a093b0e05f59c6991c33e44a6bc",
    "passport_csrf_token_default": "14561a093b0e05f59c6991c33e44a6bc",
    "d_ticket": "568930f491d12b640d4351937f68ed47ba84a",
    "odin_tt": "ebcde5d695a4a24b6c330f4d99b3e666cb343e3f2a1cf5146a4a0789bf8b930a4b870f72c0f7ec0103dc803d3b8eaa438e2ea94981b0fce401385ca779e0db32",
    "sid_guard": "69fbe6bb691149b5ec894aeb8191712a%7C1791863692%7C259200%7CFri%2C+16-Oct-2026+03%3A54%3A52+GMT",
    "uid_tt": "14615c6b84b79429127ff6962b56c46eddd160dc7c8b299b2125b69e9a6b55a2",
    "uid_tt_ss": "14615c6b84b79429127ff6962b56c46eddd160dc7c8b299b2125b69e9a6b55a2",
    "sid_tt": "69fbe6bb691149b5ec894aeb8191712a",
    "sessionid": "69fbe6bb691149b5ec894aeb8191712a",
    "sessionid_ss": "69fbe6bb691149b5ec894aeb8191712a",
    "msToken": "EsGOrjPfGxRAZVTZ5qQSafaU25KWu-t5pk_S1FfFehR6ckFZOOoiHt9UXgCVq4y9Pqy_Hr8XfhFyi2JBglNAXyAMPjSsvVo4OP9G1he0sX8f2TOxMy3iTm3yIq7HU9LNsS917_ajDcWBPg==",
    "global_seller_id_unified_seller_env": "7494876782645708274",
    "app_id_unified_seller_env": "6556",
    "oec_seller_id_unified_seller_env": "7494876782645708274",
}

DEFAULT_HEADERS: dict[str, str] = {
    "accept": "*/*",
    "accept-language": "en-GB,en;q=0.9",
    "content-type": "application/json",
    "origin": "https://seller.eu.tiktokshopglobalselling.com",
    "priority": "u=1, i",
    "referer": "https://seller.eu.tiktokshopglobalselling.com/",
    "sec-ch-ua": '"Not;A=Brand";v="8", "Chromium";v="150", "Google Chrome";v="150"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"macOS"',
    "sec-fetch-dest": "empty",
    "sec-fetch-mode": "cors",
    "sec-fetch-site": "same-site",
    "user-agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/150.0.0.0 Safari/537.36",
}

DEFAULT_SELLER_ID = "7494876782645708274"

# 应用 id：两个接口用的不是同一个（实测）
DETAIL_AID = "4068"     # order/get 用这个（detail.py 实测可通）
LIST_AID = "6556"       # order/list 用这个（list.py 抓到的后台真实请求）

# 客户信息导出的列（顺序即写表顺序）
# 「SKU ID」= 商品 SKU 的平台 ID（``sku_module[].sku_id``）。
# 主源是接口响应本身；订单表那条路只是兜底（见 ``read_sku_map``）。
CUSTOMER_COLUMNS: list[str] = [
    "Order ID",
    "SKU ID",
    "Buyer Nickname",
    "Recipient",
    "First Name",
    "Last Name",
    "Phone",
    "Email",
    "Zipcode",
    "Country",
    "Country Code",
    "Province",
    "City",
    "District",
    "Address",
    "Address Detail",
    "Address ID",
    "Fetch Status",
]

# 订单导出的列（取自 order/list 响应里字段语义明确的那些）
ORDER_COLUMNS: list[str] = [
    "Order ID",
    "Order Status",
    "Sub Status",
    "SKU Display Status",
    "Sale Region",
    "Buyer Nickname",
    "Order Amount",
    "Sub Total",
    "Shipping Fee",
    "Payment Total",
    "Currency",
    "Created Time",
    "Paid Time",
    "Updated Time",
    "Warehouse Name",
    "Warehouse Region",
    "Tracking ID",
    "Shipping Provider",
    "Product Name",
    "SKU Name",
    "Seller SKU",
    "SKU ID",
    "Quantity",
    "SKU Unit Price",
    "SKU Total Price",
    "Package ID",
    "Main Order Type",
    "Fulfillment Type",
]

TransportFn = Callable[[str, dict | None, dict | None], dict[str, Any]]

# 兜底来源：订单表里「商品 SKU 的平台 ID」可能的列名（随后台语言/版本变），
# 按优先级取第一个命中的。注意**不是**「Seller SKU」（那是商家自定义编码）。
SKU_SOURCE_COLUMNS: tuple[str, ...] = ("SKU ID", "商品 SKU", "商品SKU", "Sku ID")
# 接口响应里 sku 节点的字段名：只要**平台 SKU ID**
SKU_ID_KEYS: tuple[str, ...] = ("sku_id", "skuid", "id")
# 一个订单多个 SKU 时的分隔符
SKU_JOIN = " | "


def norm_col(name: Any) -> str:
    """列名归一化：去空格、连字符、下划线后转小写。"""
    return str(name or "").strip().lower().replace(" ", "").replace("-", "").replace("_", "")


class ApiError(RuntimeError):
    """接口返回非 0 code 时抛出。"""

    # 常见错误码的人话解释，便于界面直接显示有效信息
    HINTS: dict[str, str] = {
        "10000": "接口返回裸 code=10000，表示这次请求没被放行。\n"
                 "  实测：order/list 在本地直连下恒定如此（aid=4068 / 6556 / 不带 aid 都试过），\n"
                 "  只有带上浏览器侧签名的「浏览器内」通道才通。\n"
                 "  → 若仍报此码：① 确认页面已导航到店铺后台首页（seller.*.tiktokshopglobalselling.com）；\n"
                 "    ② 稍等几秒重试（可能只是限流）；③ 确认店铺浏览器里已登录后台",
        "98001004": "seller_id 与当前店铺不匹配（或缺少店铺上下文参数）。\n"
                    "  → 本工具已改成从店铺 cookie 自动推导 seller_id；若手动填过，请清空重试。\n"
                    "  → 注意：硬改 cookie 里的 seller_id 字段无效（服务端会拦），必须换真的那份 cookie。",
        "98001008": "当前登录态没有权限看这批订单 —— cookie 不是这些订单所属店铺的。\n"
                    "  一个店铺的登录态只能看它自己的订单（实测：汉维2-圣诞树 cookie + 嘉恒业订单\n"
                    "  必报此码；换回汉维自己的订单立刻 code=0）。\n"
                    "  → 请确认抓的订单与店铺浏览器里打开的店铺是同一家。",
        "1000003": "会话可能已失效，请在店铺浏览器里重新登录后台",
        "10037": "请求被风控拦截",
    }

    #: 重试没有意义、且必然是配置/凭据问题的错误码 —— 命中即中止整批抓取。
    FATAL_CODES: frozenset[str] = frozenset(
        {"98001001", "98001004", "98001008", "no_seller"})

    def __init__(self, code: Any, message: str, raw: Any = None):
        hint = self.HINTS.get(str(code))
        text = f"[code={code}] {message}".strip()
        if hint:
            text += f"\n  → {hint}"
        super().__init__(text)
        self.code = code
        self.message = message
        self.raw = raw

    @property
    def is_fatal(self) -> bool:
        """是否为「重试无用」的凭据/店铺身份类错误。"""
        return str(self.code) in self.FATAL_CODES


def seller_id_from_cookies(cookies: dict[str, str]) -> str:
    """从 cookie 推导 seller_id。"""
    for key in ("oec_seller_id_unified_seller_env",
                "global_seller_id_unified_seller_env",
                "oec_seller_id",
                "seller_id"):
        val = cookies.get(key)
        if val and str(val).strip():
            return str(val).strip()
    return ""


@dataclass
class OrderClient:
    """订单列表 / 详情请求客户端。"""

    cookies: dict[str, str] = field(default_factory=lambda: dict(DEFAULT_COOKIES))
    headers: dict[str, str] = field(default_factory=lambda: dict(DEFAULT_HEADERS))
    # 默认空 —— 绝不默认成某个店铺的 id（那是另一个店铺，会 98001004/98001008）
    seller_id: str = ""
    timeout: float = 30.0
    retries: int = 2
    retry_delay: float = 1.5
    api_base: str = DEFAULT_API_BASE
    # 可插拔通道：fn(url, params, body) -> dict。None 时走本地 requests。
    transport: TransportFn | None = None

    def __post_init__(self) -> None:
        self.session = requests.Session()
        if not self.seller_id:
            # 只从 cookie 推导；**绝不**回退到内置默认值 —— 那是另一个店铺的 id，
            # 会让接口回 code=98001004（实测踩过）。
            self.seller_id = seller_id_from_cookies(self.cookies)

    # -- 地址 ---------------------------------------------------------------
    @property
    def get_url(self) -> str:
        return self.api_base.rstrip("/") + ORDER_GET_PATH

    @property
    def list_url(self) -> str:
        return self.api_base.rstrip("/") + ORDER_LIST_PATH

    # -- 凭据 ----------------------------------------------------------------
    @classmethod
    def from_cookie_file(cls, path: Path | str = DEFAULT_COOKIE_FILE, **kw) -> "OrderClient":
        """若存在 cookies.json 则覆盖内置 cookie。

        这条路径读的是**本机保存的、已知属于某个店铺**的凭据，所以允许回退到
        DEFAULT_SELLER_ID；浏览器路径（``from_cookies``）不允许。
        """
        cookies = dict(DEFAULT_COOKIES)
        p = Path(path)
        if p.exists():
            try:
                loaded = json.loads(p.read_text("utf-8"))
                if isinstance(loaded, dict):
                    cookies.update({str(k): str(v) for k, v in loaded.items()})
            except Exception:  # 坏文件不该阻塞主流程
                pass
        kw.setdefault("seller_id", seller_id_from_cookies(cookies) or DEFAULT_SELLER_ID)
        return cls(cookies=cookies, **kw)

    @classmethod
    def from_cookies(cls, cookies: dict[str, str], **kw) -> "OrderClient":
        """用浏览器会话 cookie 构造（推荐）。"""
        merged = {str(k): str(v) for k, v in (cookies or {}).items()}
        kw.setdefault("seller_id", seller_id_from_cookies(merged))
        return cls(cookies=merged, **kw)

    def save_cookie_file(self, path: Path | str = DEFAULT_COOKIE_FILE) -> Path:
        p = Path(path)
        p.write_text(json.dumps(self.cookies, ensure_ascii=False, indent=2), "utf-8")
        return p

    # -- 参数 ----------------------------------------------------------------
    def aid(self) -> str:
        """应用 id：cookie 里有就用它（同一店铺会随后台版本变化），否则用跨境店默认值。"""
        val = (self.cookies.get("app_id_unified_seller_env")
               or self.cookies.get("aid") or "").strip()
        return val or LIST_AID

    def base_params(self) -> dict[str, str]:
        if not self.seller_id:
            raise ApiError(
                "no_seller",
                "没有 seller_id：店铺浏览器里没有该平台的登录 cookie，"
                "无法确定是哪个店铺。\n"
                "请先在紫鸟里打开这个店铺并登录 TikTok 商家后台"
                "（或在界面上手动填写 seller_id）。",
            )
        return {
            "locale": "zh-CN",
            "language": "zh-CN",
            "oec_seller_id": self.seller_id,
            "seller_id": self.seller_id,
        }

    def detail_params(self) -> dict[str, str]:
        """order/get：实测只要 aid + seller_id 就通（与 detail.py 一致）。"""
        params = self.base_params()
        params["aid"] = DETAIL_AID
        return params

    def list_params(self) -> dict[str, str]:
        """order/list：对齐后台真实请求（list.py 抓到的参数集）。

        注意 order/list 在本地直连下**恒定** code=10000（与 aid 无关），
        只有带浏览器侧签名的「浏览器内」通道才通；这里补齐浏览器环境参数，
        是为了让浏览器内那次请求与真实页面尽可能一致。
        """
        params = self.base_params()
        params.update({
            "aid": self.aid(),
            "app_name": "i18n_ecom_shop",
            "device_platform": "web",
            "cookie_enabled": "true",
            "screen_width": "1920",
            "screen_height": "1080",
            "browser_language": "en-GB",
            "browser_platform": "MacIntel",
            "browser_name": "Mozilla",
            "browser_version": self.headers.get("user-agent", ""),
            "browser_online": "true",
            "timezone_name": "Europe/London",
        })
        return params

    # -- 预检 ----------------------------------------------------------------
    def probe(self, order_ids: Sequence[str], max_ids: int = 3) -> dict[str, Any]:
        """店铺身份 / 凭据预检：先拿前几条订单试水，别等跑完几百条才发现全错。

        返回 ``{"ok", "partial", "tried", "results": [{id, code, message, hint, ok}],
        "fatal", "seller_id", "buyers": [...]}``。
        有一条成功即 ``ok``；全部失败且都命中凭据类错误码时给出 ``fatal`` 说明。
        """
        ids = [str(i) for i in (order_ids or []) if i][:max(1, int(max_ids))]
        results: list[dict[str, Any]] = []
        fatal = ""
        for oid in ids:
            item: dict[str, Any] = {"id": oid, "code": None, "message": "",
                                    "hint": "", "ok": False, "buyer": ""}
            try:
                payload = self._post(self.get_url, self.detail_params(),
                                     {"main_order_id": [oid]})
            except ApiError as exc:
                item.update(code=exc.code, message=exc.message,
                            hint=ApiError.HINTS.get(str(exc.code), ""))
                results.append(item)
                if exc.is_fatal and not fatal:
                    fatal = str(exc)
                continue
            code = payload.get("code")
            if code not in (0, None):
                item.update(code=code, message=payload.get("message") or "",
                            hint=ApiError.HINTS.get(str(code), ""))
                if str(code) in ApiError.FATAL_CODES and not fatal:
                    fatal = f"[code={code}] {payload.get('message') or ''}"
            else:
                orders = ((payload.get("data") or {}).get("main_order")) or []
                buyer = ""
                if orders:
                    buyer = str(((orders[0].get("buyer_info_module") or {})
                                 .get("buyer_nickname") or ""))
                item.update(ok=True, code=0, message="success", buyer=buyer)
            results.append(item)

        oks = [r for r in results if r["ok"]]
        return {"ok": bool(oks), "partial": bool(oks) and len(oks) < len(results),
                "tried": ids, "results": results,
                "fatal": "" if oks else fatal, "seller_id": self.seller_id,
                "buyers": [r["buyer"] for r in oks]}

    # -- 请求 ----------------------------------------------------------------
    def _post(self, url: str, params: dict[str, Any], body: dict[str, Any]) -> dict[str, Any]:
        last_exc: Exception | None = None
        for attempt in range(self.retries + 1):
            try:
                if self.transport is not None:
                    return self.transport(url, params, body)
                return self._post_requests(url, params, body)
            except ApiError:
                raise
            except Exception as exc:  # 网络 / 解析异常 → 重试
                last_exc = exc
                # ★ 页面请求超时（多一半是验证码把请求挂住了）**不在这里重试** ★
                # 墙没拆掉，重试只是把 60s 变成 3×60s。交给上层「先过验证码、
                # 再重试这一单」。按类名判断，免得为这一个判断反向依赖桥接层。
                if type(exc).__name__ == "ZiniaoTimeout":
                    raise ApiError("timeout", str(exc)) from exc
                if attempt < self.retries:
                    time.sleep(self.retry_delay)
        raise ApiError("network", f"{type(last_exc).__name__}: {last_exc}")

    def _post_requests(self, url: str, params: dict[str, Any],
                       body: dict[str, Any]) -> dict[str, Any]:
        resp = self.session.post(
            url,
            params=params,
            cookies=self.cookies,
            headers=self.headers,
            json=body,
            timeout=self.timeout,
        )
        resp.raise_for_status()
        return resp.json()

    @staticmethod
    def _check(payload: dict[str, Any]) -> dict[str, Any]:
        code = payload.get("code")
        if code not in (0, None):
            raise ApiError(code, payload.get("message") or payload.get("msg") or "", payload)
        return payload

    # -- 业务 ----------------------------------------------------------------
    def fetch(self, order_ids: Sequence[str], chunk_size: int = 50) -> dict[str, Any]:
        """按 order_id 列表请求详情，内部按 chunk_size 分批，返回原始 JSON。"""
        ids = list(order_ids)
        if not ids:
            return {"code": 0, "data": {"main_order": []}}

        merged: dict[str, Any] = {"code": 0, "data": {"main_order": []}}
        for i in range(0, len(ids), chunk_size):
            chunk = ids[i : i + chunk_size]
            payload = self._check(self._post(self.get_url, self.detail_params(),
                                             {"main_order_id": chunk}))
            orders = ((payload.get("data") or {}).get("main_order")) or []
            merged["data"]["main_order"].extend(orders)
        return merged

    def fetch_order_list(self, page_size: int = 50, max_orders: int = 0,
                         sort_info: str = "6",
                         on_progress: Callable[[int, int], None] | None = None,
                         should_stop: Callable[[], bool] | None = None,
                         on_log: Callable[[str], None] | None = None,
                         throttle_retries: int = 3,
                         ) -> list[dict[str, Any]]:
        """分页拉取订单列表。

        :param page_size: 每页条数
        :param max_orders: 最多拉多少条（0 = 不限）
        :param on_progress: 回调(已拉取数, 总数)
        :param should_stop: 返回 True 时提前中止，已拉到的照常返回
        :param on_log: 进度日志回调
        :param throttle_retries: 命中限流(code=10000)时的退避重试次数
        """
        log = on_log or (lambda _m: None)
        params = self.list_params()
        orders: list[dict[str, Any]] = []
        offset = 0
        total: int | None = None

        while True:
            if should_stop is not None and should_stop():
                break
            body = {
                "search_condition": {"condition_list": {}},
                "offset": offset,
                "count": page_size,
                "sort_info": sort_info,
                "pagination_type": 0,
            }

            payload: dict[str, Any] | None = None
            for attempt in range(max(0, throttle_retries) + 1):
                try:
                    payload = self._check(self._post(self.list_url, params, body))
                    break
                except ApiError as exc:
                    if str(exc.code) == "10000" and attempt < throttle_retries:
                        wait = self.retry_delay * (2 ** attempt)
                        log(f"接口限流(code=10000)，{wait:.0f}s 后重试"
                            f"（第 {attempt + 1}/{throttle_retries} 次）")
                        time.sleep(wait)
                        continue
                    raise
            if payload is None:  # pragma: no cover - 上面必然赋值或抛错
                break

            data = payload.get("data") or {}
            batch = data.get("main_orders") or []
            if total is None:
                try:
                    total = int(data.get("total_count"))
                except (TypeError, ValueError):
                    total = None

            orders.extend(batch)
            if on_progress is not None:
                on_progress(len(orders), total if total is not None else 0)

            if not batch:
                break
            if max_orders and len(orders) >= max_orders:
                orders = orders[:max_orders]
                break
            if total is not None and len(orders) >= total:
                break
            if len(batch) < page_size:
                break  # 最后一页
            offset += len(batch)

        return orders


# ----------------------------------------------------------------------------
# 解析
# ----------------------------------------------------------------------------
def _first(seq: Any) -> dict[str, Any]:
    if isinstance(seq, list) and seq and isinstance(seq[0], dict):
        return seq[0]
    return {}


def _price(node: Any) -> str:
    if isinstance(node, dict):
        return str(node.get("format_price") or "")
    return ""


def _fmt_time(value: Any) -> str:
    try:
        ts = int(value)
    except (TypeError, ValueError):
        return ""
    if ts <= 0:
        return ""
    if ts > 10_000_000_000:  # 毫秒
        ts //= 1000
    try:
        return datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M:%S")
    except (OverflowError, OSError, ValueError):
        return ""


def _text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "true" if value else "false"
    return str(value)


def parse_order_row(order: dict[str, Any]) -> dict[str, str]:
    """把 order/list 里的一条 main_order 摊平成一行订单记录。"""
    trade = order.get("trade_order_module") or {}
    status = _first(order.get("order_status_module"))
    price = order.get("price_module") or {}
    delivery = _first(order.get("delivery_module"))
    skus = order.get("sku_module") or []
    sku = _first(skus)
    pkg = _first(order.get("fulfill_unit_id_mapper"))
    buyer = order.get("buyer_info_module") or {}
    ship_provider = delivery.get("shipment_provider_info") or {}

    names = [_text(s.get("product_name")) for s in skus if isinstance(s, dict)]
    sku_names = [_text(s.get("sku_name")) for s in skus if isinstance(s, dict)]
    total_qty = 0
    for s in skus:
        try:
            total_qty += int(s.get("quantity") or 0)
        except (TypeError, ValueError):
            continue

    return {
        "Order ID": _text(order.get("main_order_id")),
        "Order Status": _text(status.get("main_order_status")),
        "Sub Status": _text(status.get("main_sub_order_status")),
        "SKU Display Status": _text(status.get("sku_display_status")),
        "Sale Region": _text(trade.get("sale_region")),
        "Buyer Nickname": _text(buyer.get("buyer_nickname")),
        "Order Amount": _price(price.get("grand_total")),
        "Sub Total": _price(price.get("sub_total")),
        "Shipping Fee": _price(trade.get("shipping_fee")),
        "Payment Total": _price(delivery.get("payment_total")),
        "Currency": _text((price.get("grand_total") or {}).get("currency")
                          if isinstance(price.get("grand_total"), dict) else ""),
        "Created Time": _fmt_time(trade.get("create_time")),
        "Paid Time": _fmt_time(trade.get("payment_time")),
        "Updated Time": _fmt_time(trade.get("update_time")),
        "Warehouse Name": _text(delivery.get("warehouse_name")),
        "Warehouse Region": _text(delivery.get("warehouse_region")),
        "Tracking ID": _text(delivery.get("last_tracking_no")),
        "Shipping Provider": _text(ship_provider.get("name")),
        "Product Name": " | ".join([n for n in names if n]),
        "SKU Name": " | ".join([n for n in sku_names if n]),
        "Seller SKU": _text(sku.get("seller_sku_name")),
        "SKU ID": _text(sku.get("sku_id")),
        "Quantity": _text(total_qty),
        "SKU Unit Price": _price(sku.get("sku_unit_price")),
        "SKU Total Price": _price(sku.get("sku_total_price")),
        "Package ID": _text(pkg.get("package_id")),
        "Main Order Type": _text(trade.get("main_order_type")),
        "Fulfillment Type": _text(trade.get("fulfillment_type")),
    }


def _addr_items(shipping_address: dict[str, Any]) -> dict[str, str]:
    out: dict[str, str] = {}
    for item in shipping_address.get("items") or []:
        key = item.get("key")
        if key:
            out[str(key)] = item.get("value") or ""
    return out


def _sku_ids_from_order(main_order: dict[str, Any]) -> str:
    """从 order/get 的响应里取**商品 SKU 的平台 ID**（``sku_module[].sku_id``）。

    这是客户表「SKU ID」列的**主源**：请求本身就是按 Order ID 打的，响应里带着
    这一单的 sku 列表，不用再回订单表关联。

    注意不要拿 ``seller_sku_name`` 顶上：那是商家自己填的编码，和平台 SKU ID
    是两回事。取不到就返回空串（由订单表兜底），不猜。
    """
    node = main_order.get("sku_module")
    if node is None:
        node = main_order.get("sku_list")
    if isinstance(node, dict):
        node = node.get("sku_list") or node.get("skus") or [node]
    if not isinstance(node, list):
        return ""
    found: list[str] = []
    for item in node:
        if not isinstance(item, dict):
            continue
        for key in SKU_ID_KEYS:
            val = item.get(key)
            if val:
                found.append(str(val))
                break
    return join_skus(found)


def join_skus(values: Sequence[Any], sep: str = SKU_JOIN) -> str:
    """同一订单多个 SKU 合成一格：去重保序。"""
    out: list[str] = []
    for v in values or []:
        text = str(v or "").strip()
        if text and text.lower() != "none" and text not in out:
            out.append(text)
    return sep.join(out)


def parse_customer(main_order: dict[str, Any], sku_ids: str = "") -> dict[str, str]:
    """从单条 main_order 里抽出客户信息，缺失字段填空串。

    :param sku_ids: 兜底用的 SKU ID（由订单表按 Order ID 关联而来），
        只在响应里没有 sku 节点时顶上
    """
    bim = main_order.get("buyer_info_module") or {}
    addr = bim.get("shipping_address") or {}
    items = _addr_items(addr)
    districts = [str(d.get("name") or "") for d in (addr.get("districts") or [])]
    region = addr.get("region") or {}

    def at(idx: int) -> str:
        return districts[idx] if idx < len(districts) else ""

    return {
        "Order ID": str(main_order.get("main_order_id") or ""),
        "SKU ID": _sku_ids_from_order(main_order) or sku_ids,
        "Buyer Nickname": str(bim.get("buyer_nickname") or ""),
        "Recipient": items.get("name", ""),
        "First Name": items.get("first_name", ""),
        "Last Name": items.get("last_name", ""),
        "Phone": items.get("phone", ""),
        "Email": items.get("email", ""),
        "Zipcode": items.get("zipcode", ""),
        "Country": str(region.get("name") or ""),
        "Country Code": str(region.get("code") or ""),
        "Province": at(0),
        "City": at(1),
        "District": at(2),
        "Address": items.get("address", ""),
        "Address Detail": items.get("address_detail", ""),
        "Address ID": str(addr.get("address_id") or ""),
        "Fetch Status": "OK",
    }


def blank_customer(order_id: str, status: str, sku_ids: str = "") -> dict[str, str]:
    """请求失败时的占位行，保证输出行数与原表一致。

    接口没回来的那些行，SKU ID 由订单表按 Order ID 补上（便于拿它回订单表核对）。
    """
    row = {col: "" for col in CUSTOMER_COLUMNS}
    row["Order ID"] = str(order_id)
    row["SKU ID"] = sku_ids
    row["Fetch Status"] = status
    return row


# ----------------------------------------------------------------------------
# Excel 读写
# ----------------------------------------------------------------------------
def iter_sheet_rows(path: Path | str, sheet: str | None = None) -> list[tuple]:
    """读整表为行元组列表。

    这类平台导出文件的 dimensions 元数据常被写成 A1:A1，read_only 模式会据此
    截断行宽（每行只剩首列）。先探测，异常时回退到普通加载。
    """
    import openpyxl

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet] if sheet else wb.active
    if ws is None:
        wb.close()
        raise ValueError(f"找不到工作表：{sheet}")

    bad_dim = False
    try:
        dim = ws.calculate_dimension()
        if dim in ("A1:A1", "A1"):
            bad_dim = True
    except ValueError:
        bad_dim = True  # unsized

    if bad_dim:
        wb.close()
        wb2 = openpyxl.load_workbook(path, data_only=True)
        ws2 = wb2[sheet] if sheet else wb2.active
        rows = [tuple(c.value for c in r) for r in ws2.iter_rows()]
        wb2.close()
        return rows

    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    return rows


def read_header(path: Path | str, sheet: str | None = None, row: int = 1) -> list[str]:
    """读取第 row 行作为表头。"""
    rows = iter_sheet_rows(path, sheet)
    if len(rows) < row:
        return []
    return [str(c) for c in rows[row - 1] if c is not None and str(c) != ""]


def read_order_ids(path: Path | str, sheet: str | None = None,
                   column: str = "Order ID") -> list[str]:
    """从 Excel 读 Order ID 列：定位表头 → 取数字型 ID → 去重保序。"""
    rows = iter_sheet_rows(path, sheet)
    if not rows:
        return []

    header_idx, col_idx = _locate_column(rows, column)
    seen: set[str] = set()
    result: list[str] = []
    for row in rows[header_idx + 1 :]:
        if col_idx >= len(row):
            continue
        val = row[col_idx]
        if val is None:
            continue
        text = str(val).strip()
        if not text.isdigit() or len(text) < 10:  # 跳过描述行 / 空行 / 杂质
            continue
        if text in seen:
            continue
        seen.add(text)
        result.append(text)
    return result


def read_sku_map(path: Path | str, sheet: str | None = None,
                 order_column: str = "Order ID",
                 sku_columns: Sequence[str] = SKU_SOURCE_COLUMNS,
                 ) -> dict[str, str]:
    """从订单表读「Order ID → 商品 SKU 的 SKU ID」映射（**兜底用**）。

    主源是接口响应里的 ``sku_module``，只有那条路拿不到（或是请求失败的占位行）
    才用这张表补。一个订单可能有多行（多个 SKU）→ 合成一格；表里的字段说明行
    第 1 列不是数字，按最后的数字校验天然会被跳过。表里没有 SKU ID 列时返回
    空 dict（调用方据此提示，不该当作错误中断流水线）。
    """
    rows = iter_sheet_rows(path, sheet)
    if not rows:
        return {}

    o_row, o_col = _locate_column(rows, order_column)
    s_col = -1
    for name in sku_columns:
        try:
            _, s_col = _locate_column(rows, name)
            break
        except ValueError:
            continue
    if s_col < 0:
        return {}

    parts: dict[str, list[str]] = {}
    for row in rows[o_row + 1 :]:
        if max(o_col, s_col) >= len(row):
            continue
        oid = str(row[o_col] or "").strip()
        if not oid.isdigit() or len(oid) < 10:   # 说明行 / 空行 / 表尾
            continue
        parts.setdefault(oid, []).append(row[s_col])
    return {oid: join_skus(vals) for oid, vals in parts.items()}


def _locate_column(rows: list[tuple], column: str) -> tuple[int, int]:
    target = norm_col(column)
    for r, row in enumerate(rows):
        for c, cell in enumerate(row):
            if cell is None:
                continue
            if norm_col(cell) == target:
                return r, c
    raise ValueError(f"未找到列：{column}")


def list_sheets(path: Path | str) -> list[str]:
    import openpyxl

    wb = openpyxl.load_workbook(path, read_only=True)
    names = list(wb.sheetnames)
    wb.close()
    return names


def _write_sheet(rows: Iterable[dict[str, str]], out_path: Path | str,
                 columns: Sequence[str], title: str) -> Path:
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = title
    ws.append(list(columns))
    for row in rows:
        ws.append([row.get(col, "") for col in columns])
    path = Path(out_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def write_customers(rows: Iterable[dict[str, str]], out_path: Path | str,
                    columns: Sequence[str] = CUSTOMER_COLUMNS) -> Path:
    """把客户信息写成 xlsx（覆盖写）。"""
    return _write_sheet(rows, out_path, columns, "Customers")


def write_orders(rows: Iterable[dict[str, str]], out_path: Path | str,
                 columns: Sequence[str] = ORDER_COLUMNS) -> Path:
    """把订单列表写成 xlsx（覆盖写），Order ID 固定在第一列。"""
    return _write_sheet(rows, out_path, columns, "Orders")
