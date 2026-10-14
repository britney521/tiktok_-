"""从紫鸟店铺浏览器抓取**当前店铺**的 cookie，写入本目录 cookies.json。

为什么需要它:
    ``order/get`` 用的是店铺登录态，而一个店铺的登录态只能看它自己的订单。
    cookie 拿错店 → ``code=98001008 需要权限``（不是过期，是「不是这家店」）。
    所以抓哪个店的订单，就要同步哪个店的 cookie。

用法:
    python sync_cookie.py                  # 列出店铺，交互式选择
    python sync_cookie.py --store 2        # 直接用第 2 个店铺（与上级 GUI 同序）
    python sync_cookie.py --name 嘉恒业     # 按店铺名模糊匹配
    python sync_cookie.py --port 9222      # 浏览器已在跑，直连调试端口，跳过紫鸟 IPC
    python sync_cookie.py --store 2 --check-orders ../订单.xlsx
                                           # 顺手验证：拿表里第一条订单试一次 order/get

说明:
    · 只读取，不关店、不动浏览器（浏览器仍归紫鸟管理）。
    · 依赖上级目录的 ziniao_bridge / tiktok_client（即「紫鸟浏览器」项目已就位）。
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent          # .../tiktok商家/获取客户
PROJECT_DIR = BASE_DIR.parent                        # .../tiktok商家
sys.path.insert(0, str(PROJECT_DIR))

import tiktok_client as tc          # noqa: E402
import ziniao_bridge as zb          # noqa: E402
import customer_api as api          # noqa: E402

COOKIE_FILE = BASE_DIR / "cookies.json"


def log(msg: str) -> None:
    print(msg, flush=True)


def _pick_interactive(stores: list[dict]) -> dict:
    log("可用店铺：")
    for i, st in enumerate(stores, 1):
        log(f"  [{i}] {zb.store_label(st)}   key={zb.browser_key(st)}")
    raw = input("选择序号（回车取消）：").strip()
    if not raw.isdigit():
        raise SystemExit("已取消")
    idx = int(raw)
    if not 1 <= idx <= len(stores):
        raise SystemExit(f"序号超出范围（1-{len(stores)}）")
    return stores[idx - 1]


def _match_store(stores: list[dict], index: int | None, name: str | None) -> dict:
    if index is not None:
        if not 1 <= index <= len(stores):
            raise SystemExit(f"--store 超范围：1-{len(stores)}")
        return stores[index - 1]
    if name:
        hit = [s for s in stores if name in zb.store_label(s)]
        if not hit:
            raise SystemExit(f"没有店铺名包含 {name!r}")
        if len(hit) > 1:
            log(f"{name!r} 匹配到 {len(hit)} 个：")
            for s in hit:
                log(f"  - {zb.store_label(s)}  key={zb.browser_key(s)}")
            raise SystemExit("请用 --store 指定序号")
        return hit[0]
    return _pick_interactive(stores)


def _resolve(args) -> tuple[int, str]:
    """拿到 (调试端口, 后台首页)：优先 --port，否则走紫鸟 startBrowser。"""
    if args.port:
        return int(args.port), zb.TIKTOK_SELLER_HOME
    cfg = zb.load_config()
    client = zb.make_client(cfg, log=log)
    stores = list(client.get_browser_list() or [])
    if not stores:
        raise SystemExit("紫鸟没有返回任何店铺（客户端起了吗？配置对吗？）")
    store = _match_store(stores, args.store, args.name)
    log(f"启动/接管店铺：{zb.store_label(store)}")
    ret = client.open_store(store) or {}
    port = ret.get("debuggingPort")
    if not port:
        raise SystemExit(
            "startBrowser 没返回 debuggingPort；若浏览器已在运行，"
            "请用 --port <端口> 直接接管。返回：" + json.dumps(ret, ensure_ascii=False)[:300])
    home = ret.get("launcherPage") or zb.TIKTOK_SELLER_HOME
    return int(port), home


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="同步当前店铺 cookie 到 cookies.json")
    ap.add_argument("--store", type=int, default=None, help="店铺序号（1 起，与上级 GUI 同序）")
    ap.add_argument("--name", default=None, help="店铺名关键词")
    ap.add_argument("--port", type=int, default=None, help="浏览器调试端口（跳过紫鸟 IPC）")
    ap.add_argument("--check-orders", default=None,
                    help="订单 Excel：同步后拿第一条 Order ID 预检一次")
    args = ap.parse_args(argv)

    port, home = _resolve(args)
    log(f"接管调试端口 127.0.0.1:{port}")
    browser = zb.open_session(port, log=log)

    # 与上级 GUI 同一条链路：先导航到店铺后台，再取 cookie（确保登录态已就位）
    try:
        log(f"打开店铺后台：{home}")
        zb.goto(browser, home, log=log)
        log(f"当前页面：{browser.current_url()[:120]}")
    except Exception as exc:
        log(f"打开后台页面失败（继续尝试取 cookie）：{exc}")

    cookies = zb.grab_cookies(browser, log=log)
    if not cookies:
        raise SystemExit("没取到 cookie：店铺可能没登录该平台，或页面还没起来")

    seller = zb.seller_id_from_cookies(cookies)
    api.save_cookies(cookies, COOKIE_FILE)
    log(f"已写入 {COOKIE_FILE}（{len(cookies)} 条），seller_id={seller or '(未取到)'}")
    log("★ 请确认这个 seller_id 就是你订单表所属的店铺。")
    log("  对照：汉维2-圣诞树=7494876782645708274，嘉恒业=7496000161034505075")

    if args.check_orders:
        src = Path(args.check_orders)
        if not src.exists():
            log(f"跳过预检：找不到 {src}")
            return 1
        ids = api.read_order_ids(src, None, "Order ID")
        if not ids:
            log("跳过预检：订单表里没读到 Order ID")
            return 1
        client = api.CustomerClient.from_cookies(cookies)
        res = client.probe(ids[0])
        if res.get("ok"):
            log(f"✅ 预检通过：cookie 能看 {(src.name)} 的订单（买家 {res.get('buyer')}）")
            return 0
        log(f"❌ 预检未通过：code={res.get('code')} {res.get('message')}")
        if res.get("hint"):
            log(f"   → {res['hint']}")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
