"""一次性诊断：把「当前页面所在的店 / cookie 的店 / 订单属于谁」一次问清楚。

为什么需要它:
    ``order/get`` 报 ``98001008 需要权限`` 时，有三种完全不同的原因，光看错误码
    分不出来，靠猜会浪费很多时间：

      ① 页面所在店 ≠ 订单所属店（切错店了）
      ② cookie 取串味了（统一商家环境下同名 cookie 多值覆盖）
      ③ 这批订单根本不在当前登录账号名下（导出这批订单的是另一个账号）

    实测（2026-10-14，本地只用得到「汉维」那份 cookie）：
      · 只改 params 里的 seller_id / oec_seller_id → **无效**
        （连「整段删掉 seller_id 参数」都还是 98001008）
      · 把 cookie 字段和 params 一起改成另一家店 → 98001004 / 13009001
      所以「改参数」不是解法，必须先搞清身份。

    本脚本会**同时**用三种身份各发一次请求做对照，一眼看出是哪种原因。

用法:
    python devtools/diagnose_store.py                    # 交互式选店
    python devtools/diagnose_store.py --store 2          # 第 2 个店铺（与主 GUI 同序）
    python devtools/diagnose_store.py --port 9222        # 浏览器已在跑，直连调试端口
    python devtools/diagnose_store.py --orders ../订单.xlsx --sample 3

只读：不关店、不改文件、不动浏览器。
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent          # .../tiktok商家/devtools
PROJECT_DIR = BASE_DIR.parent                       # .../tiktok商家
sys.path.insert(0, str(PROJECT_DIR))

import tiktok_client as tc          # noqa: E402
import tiktok_export as te          # noqa: E402
import ziniao_bridge as zb          # noqa: E402

HOSTS = [
    "https://seller.tiktokshopglobalselling.com",
    "https://seller.us.tiktokshopglobalselling.com",
    "https://seller.eu.tiktokshopglobalselling.com",
]
DETAIL_URL = tc.DEFAULT_API_BASE.rstrip("/") + tc.ORDER_GET_PATH


def log(msg: str = "") -> None:
    print(msg, flush=True)


def hr(title: str) -> None:
    log("")
    log("=" * 74)
    log(f"  {title}")
    log("=" * 74)


# ---------------------------------------------------------------------------
# 参数 / 订单
# ---------------------------------------------------------------------------
def detail_params(sid: str, aid: str = "4068") -> dict[str, str]:
    p: dict[str, str] = {"aid": aid, "locale": "zh-CN", "language": "zh-CN"}
    if sid:
        p["oec_seller_id"] = sid
        p["seller_id"] = sid
    return p


def read_order_ids(path: Path, limit: int) -> list[str]:
    try:
        import openpyxl
    except ImportError:
        log(f"（没装 openpyxl，跳过读 {path}）")
        return []
    if not path.exists():
        log(f"（找不到 {path}）")
        return []
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        return []
    hdr = [str(h) for h in rows[0]]
    idx = hdr.index("Order ID") if "Order ID" in hdr else 0
    ids = [str(r[idx]).strip() for r in rows[1:] if idx < len(r) and r[idx]]
    log(f"{path.name}：读到 {len(ids)} 条 Order ID（列「{hdr[idx]}」），取前 {limit} 条做对照")
    return ids[:max(1, limit)]


# ---------------------------------------------------------------------------
# 拿端口 / 店铺
# ---------------------------------------------------------------------------
def pick_store(stores: list[dict], index: int | None, name: str | None) -> dict:
    if index is not None:
        if not 1 <= index <= len(stores):
            raise SystemExit(f"--store 超范围：1-{len(stores)}")
        return stores[index - 1]
    if name:
        hit = [s for s in stores if name in zb.store_label(s)]
        if len(hit) != 1:
            raise SystemExit(f"店铺名 {name!r} 匹配到 {len(hit)} 个，请改用 --store")
        return hit[0]
    log("可用店铺：")
    for i, st in enumerate(stores, 1):
        log(f"  [{i}] {zb.store_label(st)}   key={zb.browser_key(st)}")
    raw = input("选择序号（回车取消）：").strip()
    if not raw.isdigit() or not 1 <= int(raw) <= len(stores):
        raise SystemExit("已取消")
    return stores[int(raw) - 1]


def resolve(args) -> tuple[int, str]:
    if args.port:
        return int(args.port), zb.TIKTOK_SELLER_HOME
    cfg = zb.load_config()
    client = zb.make_client(cfg, log=log)
    stores = list(client.get_browser_list() or [])
    if not stores:
        raise SystemExit("紫鸟没返回店铺（客户端起了吗？配置对吗？）")
    store = pick_store(stores, args.store, args.name)
    log(f"启动/接管店铺：{zb.store_label(store)}")
    ret = client.open_store(store) or {}
    port = ret.get("debuggingPort")
    if not port:
        raise SystemExit("startBrowser 没返回 debuggingPort；若浏览器在跑请用 --port")
    return int(port), (ret.get("launcherPage") or zb.TIKTOK_SELLER_HOME)


# ---------------------------------------------------------------------------
# 三种身份各打一次
# ---------------------------------------------------------------------------
def probe_in_browser(browser, origins: list[str], params: dict[str, str],
                     oid: str) -> str:
    """在页面里发一次 order/get，返回 ``code/msg`` 文本。"""
    bt = zb.BrowserTransport(browser, log=lambda _m: None, origins=origins)
    try:
        payload = bt(DETAIL_URL, params, {"main_order_id": [oid]})
    except Exception as exc:
        return f"EXC {type(exc).__name__}: {str(exc)[:90]}"
    code = payload.get("code")
    msg = payload.get("message")
    orders = ((payload.get("data") or {}).get("main_order")) or []
    buyer = ""
    if orders:
        buyer = str(((orders[0].get("buyer_info_module") or {}).get("buyer_nickname") or ""))
    tail = f"  buyer={buyer}" if buyer else ""
    return f"code={code}  msg={msg}{tail}"


def probe_local(cookies: dict[str, str], params: dict[str, str], oid: str) -> str:
    import requests
    s = requests.Session()
    s.trust_env = False
    try:
        r = s.post(DETAIL_URL, params=params, cookies=cookies,
                   headers={"accept": "*/*", "content-type": "application/json",
                            "origin": HOSTS[0], "referer": HOSTS[0] + "/",
                            "user-agent": tc.DEFAULT_HEADERS["user-agent"]},
                   json={"main_order_id": [oid]}, timeout=30)
        payload = r.json()
    except Exception as exc:
        return f"EXC {type(exc).__name__}: {str(exc)[:90]}"
    orders = ((payload.get("data") or {}).get("main_order")) or []
    buyer = ""
    if orders:
        buyer = str(((orders[0].get("buyer_info_module") or {}).get("buyer_nickname") or ""))
    return (f"code={payload.get('code')}  msg={payload.get('message')}"
            + (f"  buyer={buyer}" if buyer else ""))


# ---------------------------------------------------------------------------
def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="诊断 98001008：页面 / cookie / 订单 三者身份")
    ap.add_argument("--store", type=int, default=None, help="店铺序号（1 起）")
    ap.add_argument("--name", default=None, help="店铺名关键词")
    ap.add_argument("--port", type=int, default=None, help="调试端口（跳过紫鸟 IPC）")
    ap.add_argument("--orders", default=str(PROJECT_DIR / "订单.xlsx"),
                    help="订单 Excel（默认 ../订单.xlsx）")
    ap.add_argument("--sample", type=int, default=3, help="取前几条订单做对照（默认 3）")
    args = ap.parse_args(argv)

    order_ids = read_order_ids(Path(args.orders), args.sample)
    if not order_ids:
        raise SystemExit("没读到订单号，无法对照（用 --orders 指向订单表）")

    port, home = resolve(args)
    log(f"接管 127.0.0.1:{port}")
    browser = zb.open_session(port, log=log)

    # -- 1. 页面 ----------------------------------------------------------
    hr("1) 打开店铺后台，看它实际落到哪个站")
    try:
        zb.goto(browser, home, log=log)
    except Exception as exc:
        log(f"打开后台失败：{exc}")
    url = browser.current_url()
    host = "https://" + url.split("//")[-1].split("/")[0] if url.startswith("http") else ""
    log(f"当前页面：{url[:130]}")
    log(f"当前站  ：{host or '(不是 http 页面)'}")
    origins = [o for o in ([host] + HOSTS + list(zb.TIKTOK_PAGE_ORIGINS)) if o]
    seen: set[str] = set()
    origins = [o for o in origins if not (o in seen or seen.add(o))]

    # -- 2. cookie 里的身份 ------------------------------------------------
    hr("2) cookie 里的店铺身份（按站取 vs 全量）")
    per_host: dict[str, str] = {}
    for h in origins:
        try:
            ck = browser.cookies("tiktokshopglobalselling", url=h + "/")
        except Exception as exc:
            log(f"  {h}  → 取 cookie 失败：{exc}")
            continue
        sid = zb.seller_id_from_cookies(ck)
        per_host[h] = sid
        log(f"  {h}")
        log(f"      cookie {len(ck)} 条；seller_id={sid or '(无)'}"
            f"{f'（{zb.store_name_of(sid)}）' if sid and zb.store_name_of(sid) else ''}")
    try:
        raw = browser.call("Storage.getCookies", {}) or {}
        items = list(raw.get("cookies") or [])
    except Exception as exc:
        items, _ = [], log(f"全量取 cookie 失败：{exc}")
    log("")
    log(f"  全量 Storage.getCookies：{len(items)} 条；同名店铺 cookie 明细：")
    for c in items:
        name = str(c.get("name") or "")
        if "seller_id" in name or name in ("app_id_unified_seller_env",):
            log(f"      {name} = {str(c.get('value'))[:24]}…  domain={c.get('domain')}")

    # -- 3. 页面自己请求里的身份 -------------------------------------------
    hr("3) 页面自己发出的请求里用的 seller_id（唯一可信来源）")
    try:
        browser.add_init_script(te.HOOK_JS)
        browser.navigate(host + te.ORDER_PAGE_SUFFIX, retry=2, interval=2.0, wait=45)
        import time
        time.sleep(6)
        counts = browser.read_json(te.PAGE_SELLER_JS, default={}) or {}
    except Exception as exc:
        counts = {}
        log(f"读页面请求身份失败：{exc}")
    page_sid = ""
    if isinstance(counts, dict) and counts:
        page_sid = str(max(counts.items(), key=lambda kv: kv[1])[0])
        log("  出现次数：" + "，".join(
            f"{k}×{v}（{zb.store_name_of(k) or '未登记'}）"
            for k, v in sorted(counts.items(), key=lambda kv: -kv[1])[:6]))
        log(f"  → 页面自己用的是 seller_id={page_sid}（{zb.store_name_of(page_sid) or '未登记'}）")
    else:
        log("  没读到（订单页还没发出带 seller_id 的请求，或 hook 注入晚于首屏）")
    ck_home = per_host.get(host) or zb.seller_id_from_cookies(
        browser.cookies("tiktokshopglobalselling", url=host + "/"))
    log(f"  → cookie 推导的是 seller_id={ck_home or '(无)'}"
        f"（{zb.store_name_of(ck_home) or '未登记'}）")
    if page_sid and ck_home and page_sid != ck_home:
        log("  ⚠ 两者不一致 —— 就是它：cookie 串味 / 页面还没把当前店铺写进 cookie")

    # -- 4. 三种身份实测 ---------------------------------------------------
    hr("4) 用订单表里的单，三种身份各打一次")
    cookies_all = browser.cookies("tiktokshopglobalselling")
    seller_cookie = zb.seller_id_from_cookies(cookies_all)
    variants: list[tuple[str, str]] = [
        (f"页面自己的 seller_id = {page_sid or '(未取到)'}", page_sid),
        (f"cookie 里的   seller_id = {seller_cookie or '(无)'}", seller_cookie),
        ("不带 seller_id 参数（让服务端按会话判断）", ""),
    ]
    for oid in order_ids:
        log("")
        log(f"  订单 {oid}")
        for tag, sid in variants:
            if tag.startswith("页面") and not page_sid:
                continue
            res = probe_in_browser(browser, origins, detail_params(sid), oid)
            log(f"      [浏览器内] {tag:<44} → {res}")
        res = probe_local(cookies_all, detail_params(seller_cookie), oid)
        log(f"      [本地直连] cookie 里的 seller_id"
            f"{'':<29} → {res}")

    # -- 5. 结论 -----------------------------------------------------------
    hr("5) 怎么读上面的结果")
    log("  · 「页面自己的 seller_id」返回 code=0")
    log("      → 订单属于页面这家店。抓客户就该用它（主 GUI 已自动改用页面的值）。")
    log("        若主 GUI 仍整批 98001008，说明预检时页面身份还没取到 → 把 ③ 导出后")
    log("        再看日志里「页面请求里的 seller_id」那行是否出现。")
    log("  · 「页面」报 98001008，但 cookie 的 id 不一样")
    log("      → cookie 串味。用 获取客户/sync_cookie.py 重新按站同步即可。")
    log("  · 三种身份全是 98001008")
    log("      → 这批订单不在当前登录账号名下的任何店里。导出这批订单的是**另一个")
    log("        账号/店铺**，必须换那个店铺的登录态（紫鸟里换店/换账号）再抓。")
    log("      → 只改 params 里的 seller_id 不会有用（实测：整段删掉也还是 98001008）。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
