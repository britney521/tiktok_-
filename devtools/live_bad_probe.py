"""真机验证：评分 CSV 里的 order_id 能不能直接换出客户（⑥ 的新链路）。

背景（2026-10-15 真机实测）：评分导出其实是 CSV，列是
``rating,review_text,review_date,order_id,product_id,sku_id`` —— **自带 order_id**，
所以 ⑥ 不必再拿买家昵称去订单表反查（那条路在这份文件上根本走不通：没有昵称列）。

这个脚本把最后一跳单独验一遍：CSV 的 1~3 星订单号 → order/get → 客户行。

用法::

    python devtools/live_bad_probe.py --port 36246 --csv /tmp/live_rating.csv --n 3
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_DIR))

import tiktok_client as tc      # noqa: E402
import tiktok_export as te      # noqa: E402
import ziniao_bridge as zb      # noqa: E402


def log(msg: str = "") -> None:
    print(msg, flush=True)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--store", default="秋图2")
    ap.add_argument("--port", default="")
    ap.add_argument("--want", default="United States")
    ap.add_argument("--csv", default="/tmp/live_rating.csv")
    ap.add_argument("--n", type=int, default=3, help="抓前几条差评")
    args = ap.parse_args()

    port = int(args.port) if args.port else 0
    if not port:
        mods = zb.load_ziniao()
        cfg = mods["client"].ZiniaoConfig.from_yaml(mods["config_path"])
        client = zb.make_client(cfg, log=lambda m: log(f"  [client] {m}"))
        hit = [s for s in client.get_browser_list() if args.store in zb.store_label(s)]
        port = int(client.open_store(hit[0]).get("debuggingPort") or 0)
    browser = zb.open_session(port, log=lambda m: log(f"  [cdp] {m}"),
                              home_urls=[zb.TIKTOK_SELLER_HOME])

    shop = te.ensure_shop(browser, args.want, log=log)
    home = shop.get("url") or browser.current_url()
    log(f"店铺落地：{home[:110]}")

    ctx = browser.store_context(home, expect="", timeout=25)
    cookies = ctx.get("cookies") or {}
    seller = ctx.get("seller_id") or ""
    log(f"cookie {len(cookies)} 条，seller_id={seller}")

    tmpl = zb.load_captured_template()
    chain = zb.TransportChain(
        [("浏览器内", zb.BrowserTransport(browser, log=log, origins=[zb.origin_of(home)],
                                          template=tmpl,
                                          page_path=te.ORDER_PAGE_SUFFIX))],
        log=log, fallthrough_codes=("10000",))
    cli = tc.OrderClient.from_cookies(cookies, transport=chain, seller_id=seller,
                                     retries=2, timeout=30)

    src = Path(args.csv)
    header, rows = te.parse_rating(src)
    score_col, bad, _ = te.filter_low_ratings(header, rows)
    order_col = te.pick_column(header, te.RATING_ORDER_COLUMNS)
    log(f"评分文件：{src.name}　{len(header)} 列 × {len(rows)} 行"
        f"　评分列={score_col!r}　订单号列={order_col!r}")
    log(f"1~3 星：{len(bad)} 条")
    if not order_col:
        log("✖ 这份文件里没有订单号列 → ⑥ 只能走昵称反查那条老路")
        return 1

    seen: list[str] = []
    for r in bad:
        oid = str(r[header.index(order_col)] or "").strip()
        if oid.isdigit() and oid not in seen:
            seen.append(oid)
    log(f"去重后待抓订单：{len(seen)} 个，取前 {args.n} 个：{seen[:args.n]}")

    ok = 0
    for i, oid in enumerate(seen[:args.n], 1):
        payload = cli.fetch([oid], chunk_size=1)
        orders = ((payload.get("data") or {}).get("main_order")) or []
        if not orders:
            log(f"[{i}] {oid} → 无数据（code={payload.get('code')}）")
            continue
        cust = tc.parse_customer(orders[0], "")
        ok += 1
        log(f"[{i}] {oid} → 昵称={cust.get('Buyer Nickname')!r}　"
            f"收件人={cust.get('Recipient')!r}　SKU ID={cust.get('SKU ID')!r}")
    log("")
    log(f"结果：{ok}/{min(args.n, len(seen))} 条拿到客户信息")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
