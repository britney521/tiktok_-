"""TikTok Shop 客户信息（买家收货信息）抓取客户端 —— 只做一件事。

链路:
    读订单 Excel 的 ``Order ID`` 列 → 逐个 POST ``/api/fulfillment/order/get``
    → 解析买家收货信息 → 写 ``客户.xlsx``

为什么这一层可以「本地直连」（不需要紫鸟浏览器）:
    实测 ``order/get`` 只要 ``aid=4068`` + ``seller_id`` 就能通（``code=0``），
    不需要 X-Bogus / X-Tts-Oec-Bsid 这类签名头。所以「获取客户」这一个功能
    可以做成纯 ``requests`` 的独立小工具。

边界（很重要）:
    ``order/list``（订单列表）**不适用**本结论 —— 它本地直连恒定返回裸
    ``code=10000``，必须在店铺浏览器里带签名发起。要用订单导出功能，
    请用上级目录的 ``customer_gui.py``。

★ cookie 必须与被抓订单**同属一个店铺**（2026-10-13 实测）:
    一个店铺的登录态只能看它自己的订单，cookie 里的
    ``oec_seller_id_unified_seller_env`` 决定了「当前在哪个店」。
        · 汉维2-圣诞树 cookie + 汉维2-圣诞树订单 → ``code=0`` success
        · 汉维2-圣诞树 cookie + **嘉恒业**订单   → ``code=98001008`` 需要权限
        · 把 cookie 里的 seller_id 字段硬改成另一个店 → ``code=98001004``（服务端校验，改不动）
    所以「报了无权限」几乎一定不是 cookie 过期，而是 **cookie 是别的店铺的**。
    换店铺就得换那份 cookie。跑整批前建议先 ``CustomerClient.probe(order_id)`` 预检。

凭据:
    按优先级取：显式传入 → 同目录 cookies.json → 内置默认值。
    会话过期时，把新 cookie 存成同目录 cookies.json 即可（或在界面里粘贴）。
"""

from __future__ import annotations

import json
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Iterable, Sequence

import requests

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_COOKIE_FILE = BASE_DIR / "cookies.json"

# 跨境店卖家后台的 API 宿主（实测本账号可用的是 no1a，EU/UK）
DEFAULT_API_BASE = "https://api16-normal-no1a.tiktokshopglobalselling.com"
ORDER_GET_PATH = "/api/fulfillment/order/get"
ORDER_GET_URL = DEFAULT_API_BASE + ORDER_GET_PATH

DETAIL_AID = "4068"  # order/get 用这个（已验证 code=0）

# 「汉维2-圣诞树」的 seller_id。**仅作对照参考**，绝不自动兜底 ——
# 拿它去请求别的店铺的订单会得到 98001008 需要权限（实测踩过）。
DEFAULT_SELLER_ID = "7494876782645708274"

#: 已知店铺名 → seller_id 对照表（写日志/报错时用来把 id 翻译成人话）
KNOWN_STORES: dict[str, str] = {
    "7494876782645708274": "汉维2-圣诞树",
    "7496000161034505075": "嘉恒业",
}

# 内置默认凭据（失效时请更新同目录 cookies.json）
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

# 客户信息导出的列（顺序即写表顺序）
CUSTOMER_COLUMNS: list[str] = [
    "Order ID",
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

# 界面表格里适合展示的列（网络请求结果不易读的列放后面）
PREVIEW_COLUMNS: list[str] = [
    "Order ID",
    "Buyer Nickname",
    "Recipient",
    "Phone",
    "Email",
    "Country",
    "Province",
    "City",
    "Address",
    "Address Detail",
    "Fetch Status",
]

TransportFn = Callable[[str, dict | None, dict | None], dict[str, Any]]


class ApiError(RuntimeError):
    """接口返回非 0 code，或传输层失败时抛出。"""

    HINTS: dict[str, str] = {
        # 实测（2026-10-13，汉维2-圣诞树 cookie × 嘉恒业订单）：
        #   汉维 cookie + 汉维订单 → code=0 success
        #   汉维 cookie + 嘉恒业订单 → 98001008 需要权限
        #   cookie 与 seller_id 参数不一致 → 98001004 部分信息填写错误
        # 结论：98001008 = 这批订单不属于 cookie 所在店铺。
        "98001008": "这个登录态**没有权限访问这批订单** —— 典型原因：cookie 不是"
                    "导出该订单表的那个店铺的。一个店铺的 cookie 只能看它自己的订单，"
                    "请换成对应店铺的 cookie（或改用上级 customer_gui.py，它会自动"
                    "取当前店铺的 cookie）。",
        "98001004": "seller_id 与 cookie 所在店铺不一致（或 Cookie 与请求参数对不上）。"
                    "把界面上的 seller_id 清空，让它从 cookie 自动取。",
        "98001001": "该登录态无此接口权限（子账号角色缺 order 查看权限）",
        "1000003": "会话可能已失效，请更新 cookies.json（重新登录后台后导出 cookie）",
        "10000": "接口未放行（限流 / 风控）。稍等几秒重试；若持续如此，"
                 "请更新 cookies.json 里的登录态",
        "10037": "请求被风控拦截，请降低频率或更换登录态",
    }

    #: 重试没有意义、且必然是配置/凭据问题的错误码 —— 命中即中止整批抓取。
    #: 98001008=该登录态看不到这批订单；98001004=seller_id 与 cookie 不匹配。
    FATAL_CODES: frozenset[str] = frozenset(
        {"98001004", "98001008", "98001001", "no_seller"})

    @property
    def is_fatal(self) -> bool:
        """是否为「重试无用」的凭据/权限类错误。"""
        return str(self.code) in self.FATAL_CODES

    def __init__(self, code: Any, message: str, raw: Any = None):
        hint = self.HINTS.get(str(code))
        text = f"[code={code}] {message}".strip()
        if hint:
            text += f"\n  → {hint}"
        super().__init__(text)
        self.code = code
        self.message = message
        self.raw = raw


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


def parse_cookie_text(text: str) -> dict[str, str]:
    """把「粘贴来的 cookie」解析成 dict。

    支持两种格式：
      · Netscape/浏览器复制出来的 ``name=value; name2=value2``
      · JSON 对象 ``{"name": "value"}``
    """
    raw = (text or "").strip()
    if not raw:
        return {}
    if raw.startswith("{"):
        try:
            loaded = json.loads(raw)
            if isinstance(loaded, dict):
                return {str(k): str(v) for k, v in loaded.items()}
        except Exception:
            pass
    out: dict[str, str] = {}
    for part in raw.replace("\n", ";").split(";"):
        piece = part.strip()
        if not piece or "=" not in piece:
            continue
        name, _, value = piece.partition("=")
        name = name.strip()
        if name:
            out[name] = value.strip()
    return out


def load_cookies(path: Path | str | None = DEFAULT_COOKIE_FILE) -> dict[str, str]:
    """内置默认值 + （可选的）cookies.json 覆盖。"""
    cookies = dict(DEFAULT_COOKIES)
    if path:
        p = Path(path)
        if p.exists():
            try:
                loaded = json.loads(p.read_text("utf-8"))
                if isinstance(loaded, dict):
                    cookies.update({str(k): str(v) for k, v in loaded.items()})
            except Exception:  # 坏文件不该阻塞主流程
                pass
    return cookies


def save_cookies(cookies: dict[str, str],
                 path: Path | str = DEFAULT_COOKIE_FILE) -> Path:
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(cookies, ensure_ascii=False, indent=2), "utf-8")
    return p


# ----------------------------------------------------------------------------
# 客户端
# ----------------------------------------------------------------------------
@dataclass
class CustomerClient:
    """按 Order ID 取买家收货信息。"""

    cookies: dict[str, str] = field(default_factory=lambda: dict(DEFAULT_COOKIES))
    headers: dict[str, str] = field(default_factory=lambda: dict(DEFAULT_HEADERS))
    seller_id: str = ""
    timeout: float = 30.0
    retries: int = 2
    retry_delay: float = 1.5
    api_base: str = DEFAULT_API_BASE
    # 可插拔通道：fn(url, params, body) -> dict。None 时走本地 requests。
    transport: TransportFn | None = None

    def __post_init__(self) -> None:
        self.session = requests.Session()
        self.session.trust_env = False  # 本地直连不要走系统代理
        if not self.seller_id:
            # 只从 cookie 推导。**绝不**回退到 DEFAULT_SELLER_ID —— 那是汉维店的 id，
            # 拿它去请求别的店铺的订单只会得到 98001008 / 98001004（实测踩过）。
            self.seller_id = seller_id_from_cookies(self.cookies)

    @property
    def get_url(self) -> str:
        return self.api_base.rstrip("/") + ORDER_GET_PATH

    @classmethod
    def from_cookies(cls, cookies: dict[str, str], **kw) -> "CustomerClient":
        merged = {str(k): str(v) for k, v in (cookies or {}).items()}
        kw.setdefault("seller_id", seller_id_from_cookies(merged))
        return cls(cookies=merged, **kw)

    @classmethod
    def from_cookie_file(cls, path: Path | str = DEFAULT_COOKIE_FILE,
                         **kw) -> "CustomerClient":
        return cls(cookies=load_cookies(path), **kw)

    def detail_params(self) -> dict[str, str]:
        """order/get 参数：实测只要 aid + seller_id。"""
        if not self.seller_id:
            raise ApiError(
                "no_seller",
                "没能确定 seller_id：cookie 里没有该平台的登录态"
                "（oec_seller_id_unified_seller_env）。\n"
                "  → 请换成目标店铺的 cookie（`python sync_cookie.py --store N` 可一键同步），"
                "或在界面上手动填 seller_id。",
            )
        return {
            "aid": DETAIL_AID,
            "locale": "zh-CN",
            "language": "zh-CN",
            "oec_seller_id": self.seller_id,
            "seller_id": self.seller_id,
        }

    # -- 请求 ----------------------------------------------------------------
    def _post(self, params: dict[str, Any], body: dict[str, Any]) -> dict[str, Any]:
        url = self.get_url
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
                if attempt < self.retries:
                    time.sleep(self.retry_delay)
        raise ApiError("network", f"{type(last_exc).__name__}: {last_exc}")

    def _post_requests(self, url: str, params: dict[str, Any],
                       body: dict[str, Any]) -> dict[str, Any]:
        resp = self.session.post(url, params=params, cookies=self.cookies,
                                 headers=self.headers, json=body, timeout=self.timeout)
        resp.raise_for_status()
        return resp.json()

    @staticmethod
    def _check(payload: dict[str, Any]) -> dict[str, Any]:
        code = payload.get("code")
        if code not in (0, None):
            raise ApiError(code, payload.get("message") or payload.get("msg") or "", payload)
        return payload

    def probe(self, order_id: str) -> dict[str, Any]:
        """凭据预检：只拿一条订单试水，判断当前 cookie 能否看这批订单。

        返回 ``{"ok": bool, "code": ..., "message": str, "hint": str,
        "buyer": str}``。跑整批前先调它，能立刻暴露「cookie 属于别的店铺」。
        """
        oid = str(order_id or "").strip()
        out: dict[str, Any] = {"ok": False, "code": None, "message": "",
                               "hint": "", "buyer": "", "order_id": oid}
        if not oid:
            out["message"] = "没有可用于预检的 Order ID"
            return out
        try:
            payload = self._post(self.detail_params(), {"main_order_id": [oid]})
        except ApiError as exc:
            out.update(code=exc.code, message=exc.message,
                       hint=ApiError.HINTS.get(str(exc.code), ""))
            return out
        except Exception as exc:
            out.update(code="network", message=f"{type(exc).__name__}: {exc}")
            return out

        code = payload.get("code")
        if code not in (0, None):
            out.update(code=code, message=payload.get("message") or "",
                       hint=ApiError.HINTS.get(str(code), ""))
            return out
        orders = ((payload.get("data") or {}).get("main_order")) or []
        if not orders:
            out.update(code=0, message="接口返回 code=0 但没有订单数据",
                       hint="订单号可能不属于任何店铺，或已被删除")
            return out
        bim = orders[0].get("buyer_info_module") or {}
        out.update(ok=True, code=0, message="success",
                   buyer=str(bim.get("buyer_nickname") or ""))
        return out

    def fetch(self, order_ids: Sequence[str], chunk_size: int = 50) -> dict[str, Any]:
        """按 order_id 列表请求详情，内部按 chunk_size 分批，返回原始 JSON。"""
        ids = [str(i) for i in order_ids if i]
        if not ids:
            return {"code": 0, "data": {"main_order": []}}

        merged: dict[str, Any] = {"code": 0, "data": {"main_order": []}}
        for i in range(0, len(ids), chunk_size):
            chunk = ids[i: i + chunk_size]
            payload = self._check(self._post(self.detail_params(),
                                             {"main_order_id": chunk}))
            orders = ((payload.get("data") or {}).get("main_order")) or []
            merged["data"]["main_order"].extend(orders)
        return merged


# ----------------------------------------------------------------------------
# 解析
# ----------------------------------------------------------------------------
def _addr_items(shipping_address: dict[str, Any]) -> dict[str, str]:
    out: dict[str, str] = {}
    for item in shipping_address.get("items") or []:
        key = item.get("key")
        if key:
            out[str(key)] = item.get("value") or ""
    return out


def parse_customer(main_order: dict[str, Any]) -> dict[str, str]:
    """从单条 main_order 里抽出客户信息，缺失字段填空串。"""
    bim = main_order.get("buyer_info_module") or {}
    addr = bim.get("shipping_address") or {}
    items = _addr_items(addr)
    districts = [str(d.get("name") or "") for d in (addr.get("districts") or [])]
    region = addr.get("region") or {}

    def at(idx: int) -> str:
        return districts[idx] if idx < len(districts) else ""

    return {
        "Order ID": str(main_order.get("main_order_id") or ""),
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


def blank_customer(order_id: str, status: str) -> dict[str, str]:
    """请求失败时的占位行，保证输出行数与原表一致。"""
    row = {col: "" for col in CUSTOMER_COLUMNS}
    row["Order ID"] = str(order_id)
    row["Fetch Status"] = status
    return row


# ----------------------------------------------------------------------------
# 抓取主流程（GUI 与测试共用）
# ----------------------------------------------------------------------------
def fetch_customers(client: CustomerClient, order_ids: Sequence[str], *,
                    on_row: Callable[[int, dict[str, str]], None] | None = None,
                    on_progress: Callable[[int, int], None] | None = None,
                    should_stop: Callable[[], bool] | None = None,
                    interval: float = 0.0,
                    flush_every: int = 10,
                    out_path: Path | str | None = None,
                    on_log: Callable[[str], None] | None = None) -> dict[str, Any]:
    """逐个 Order ID 抓客户信息，边抓边落盘。

    :param on_row: 回调(下标, 客户行)，便于界面实时显示
    :param on_progress: 回调(已完成, 总数)
    :param should_stop: 返回 True 时提前中止，已抓到的照常写出
    :param interval: 每条之间的间隔秒数（防限流）
    :param flush_every: 每抓多少条落一次盘
    :param out_path: 客户 Excel 输出路径；None 则不写文件
    :return: {"rows": [...], "ok": n, "fail": n, "out": str|"", "stopped": bool}
    """
    log = on_log or (lambda _m: None)
    ids = [str(i) for i in order_ids if i]
    out = Path(out_path) if out_path else None

    results: dict[str, dict[str, str]] = {}
    rows: list[dict[str, str]] = []
    ok = fail = 0
    stopped = False
    fatal: str = ""      # 命中凭据/权限类错误时的说明，非空表示已中止

    if on_progress is not None:
        on_progress(0, len(ids))

    for i, oid in enumerate(ids):
        if should_stop is not None and should_stop():
            stopped = True
            log("已手动停止，写出已抓取的数据…")
            break
        if fatal:
            # 已判定是凭据/权限问题：后面每条都会同样失败，直接补齐占位行收工。
            row = blank_customer(oid, "SKIPPED")
            results[oid] = row
            if on_row is not None:
                on_row(i, row)
            if on_progress is not None:
                on_progress(i + 1, len(ids))
            continue
        try:
            payload = client.fetch([oid], chunk_size=1)
            orders = ((payload.get("data") or {}).get("main_order")) or []
            if not orders:
                row = blank_customer(oid, "EMPTY")
                fail += 1
                log(f"[{i + 1}/{len(ids)}] {oid} → 无数据返回")
            else:
                row = parse_customer(orders[0])
                ok += 1
                log(f"[{i + 1}/{len(ids)}] {oid} → "
                    f"{row.get('Buyer Nickname', '')} {row.get('Recipient', '')}")
        except ApiError as exc:
            row = blank_customer(oid, "API_ERROR")
            fail += 1
            log(f"[{i + 1}/{len(ids)}] {oid} → 失败：{exc}")
            if exc.is_fatal and not fatal:
                fatal = str(exc)
                log("—" * 58)
                log(f"已中止：这是凭据/权限问题（code={exc.code}），"
                    f"后面 {len(ids) - i - 1} 条会同样失败，不再继续请求。")
                log("请换成本店铺的 cookie 后重跑（界面③区「凭据预检」可先验证）。")
                log("—" * 58)
        except Exception as exc:
            row = blank_customer(oid, "EXCEPTION")
            fail += 1
            log(f"[{i + 1}/{len(ids)}] {oid} → 异常：{type(exc).__name__}: {exc}")

        results[oid] = row
        if on_row is not None:
            on_row(i, row)
        if on_progress is not None:
            on_progress(i + 1, len(ids))

        if out is not None and flush_every and (i + 1) % flush_every == 0:
            _flush(results, ids, out)

        if (should_stop is None or not should_stop()) and i + 1 < len(ids) and interval > 0:
            time.sleep(interval)

    rows = [results[i] for i in ids if i in results]
    final = _flush(results, ids, out) if out is not None else None
    if out is not None:
        log(f"抓取结束：成功 {ok}，失败 {fail}，已写入 {final}")
    else:
        log(f"抓取结束：成功 {ok}，失败 {fail}")
    if fatal:
        log(f"[提示] 本批因凭据问题提前中止：{fatal.splitlines()[0]}")

    return {"rows": rows, "ok": ok, "fail": fail,
            "out": str(final) if final else "", "stopped": stopped,
            "fatal": fatal}


def _flush(results: dict[str, dict[str, str]], ids: list[str], out: Path) -> Path:
    ordered = [results[i] for i in ids if i in results]
    return write_customers(ordered, out)


# ----------------------------------------------------------------------------
# Excel 读写
# ----------------------------------------------------------------------------
def iter_sheet_rows(path: Path | str, sheet: str | None = None) -> list[tuple]:
    """读整表为行元组列表。

    这类平台导出的文件 dimensions 元数据常被写成 A1:A1，read_only 模式会据此
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


def list_sheets(path: Path | str) -> list[str]:
    import openpyxl

    wb = openpyxl.load_workbook(path, read_only=True)
    names = list(wb.sheetnames)
    wb.close()
    return names


def read_header(path: Path | str, sheet: str | None = None, row: int = 1) -> list[str]:
    """读取第 row 行作为表头。"""
    rows = iter_sheet_rows(path, sheet)
    if len(rows) < row:
        return []
    return [str(c) for c in rows[row - 1] if c is not None and str(c) != ""]


def _locate_column(rows: list[tuple], column: str) -> tuple[int, int]:
    target = column.strip().lower().replace(" ", "")
    for r, row in enumerate(rows):
        for c, cell in enumerate(row):
            if cell is None:
                continue
            if str(cell).strip().lower().replace(" ", "") == target:
                return r, c
    raise ValueError(f"未找到列：{column}")


def read_order_ids(path: Path | str, sheet: str | None = None,
                   column: str = "Order ID") -> list[str]:
    """从 Excel 读 Order ID 列：定位表头 → 取数字型 ID → 去重保序。"""
    rows = iter_sheet_rows(path, sheet)
    if not rows:
        return []

    header_idx, col_idx = _locate_column(rows, column)
    seen: set[str] = set()
    result: list[str] = []
    for row in rows[header_idx + 1:]:
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


def write_customers(rows: Iterable[dict[str, str]], out_path: Path | str,
                    columns: Sequence[str] = CUSTOMER_COLUMNS) -> Path:
    """把客户信息写成 xlsx（覆盖写）。"""
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Customers"
    ws.append(list(columns))
    for row in rows:
        ws.append([row.get(col, "") for col in columns])
    path = Path(out_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path
