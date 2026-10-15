"""真机探测：后台店铺是不是按 URL（shop_region）导航对了。

用法（紫鸟 web_driver 模式客户端已在跑、店铺浏览器已开）：

    python devtools/live_shop_probe.py                       # 只读：现在在哪家店 + 目标会去哪个 URL
    python devtools/live_shop_probe.py --store 秋图2 --want Spain
    python devtools/live_shop_probe.py --port 27206 --want Spain --go   # 真的导航过去
    python devtools/live_shop_probe.py --url <后台URL> --go             # 直接贴 URL 导航

默认**只读**：读当前页面 URL 与右上角切换器上的店铺名，并打印「要切到 X 会导航到哪个
URL」。只有加 ``--go`` 才真的导航（会整页刷新）——导航后轮询落地 URL，确认是否真的
停在目标 ``shop_region``（真机见过「开 eu 落在 us」这种重定向）。
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent
sys.path.insert(0, str(PROJECT_DIR))

import tiktok_export as te      # noqa: E402
import ziniao_bridge as zb      # noqa: E402


def log(msg: str = "") -> None:
    print(msg, flush=True)


def resolve_port(args) -> int:
    """返回调试端口。--port 直接给；否则走客户端 API 拿。"""
    if args.port:
        return int(args.port)
    mods = zb.load_ziniao()
    cfg = mods["client"].ZiniaoConfig.from_yaml(mods["config_path"])
    client = zb.make_client(cfg, log=lambda m: log(f"  [client] {m}"))
    if not client.port_open():
        raise SystemExit("紫鸟客户端 IPC 端口没开：请用 web_driver 模式启动客户端")
    stores = client.get_browser_list()
    hit = [s for s in stores if args.store and args.store in zb.store_label(s)]
    if not hit:
        log("可用店铺：")
        for i, s in enumerate(stores, 1):
            log(f"  [{i}] {zb.store_label(s)}")
        raise SystemExit(f"没找到含 {args.store!r} 的店铺")
    store = hit[0]
    ret = client.open_store(store)          # 已开则复用，返回同一端口
    port = ret.get("debuggingPort")
    if not port:
        raise SystemExit(f"startBrowser 没返回调试端口：{ret}")
    log(f"店铺「{zb.store_label(store)}」→ 调试端口 {port}")
    return int(port)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--store", default="秋图2", help="店铺名（模糊匹配）")
    ap.add_argument("--port", default="", help="直接指定调试端口，绕过客户端 API")
    ap.add_argument("--want", default="United States", help="目标后台店铺名（显示名/区域码）")
    ap.add_argument("--url", default="", help="直接给后台 URL（优先于 --want）")
    ap.add_argument("--go", action="store_true", help="真的导航过去（不加则只读）")
    ap.add_argument("--settle", type=float, default=3.0)
    ap.add_argument("--wait", type=float, default=30.0)
    args = ap.parse_args()

    want = args.url or args.want
    target = te.shop_home_url(want)

    port = resolve_port(args)
    browser = zb.open_session(port, log=lambda m: log(f"  [cdp] {m}"),
                              home_urls=[zb.TIKTOK_SELLER_HOME])

    log("")
    log("=" * 74)
    log("  ① 现在在哪（只读）")
    log("=" * 74)
    here = browser.current_url()
    log(f"页面 URL    ：{here[:130]}")
    log(f"shop_region ：{te.shop_region_of(here) or '(无)'}")
    log(f"切换器写着  ：{te.shop_pill_text(browser).strip() or '(读不到)'}")

    log("")
    log("=" * 74)
    log("  ② 目标是哪家（店铺名 → URL，纯查表，不联网）")
    log("=" * 74)
    log(f"目标输入    ：{want}")
    log(f"目标 URL    ：{target or '(认不出 —— 可填 ' + '、'.join(te.SHOP_CHOICES[:6]) + ' …)'}")
    if not target:
        log("→ ensure_shop 会拒绝导航并告警（不乱猜区域码）。")
        return 1

    log("")
    log("=" * 74)
    log("  ③ 判定")
    log("=" * 74)
    code = te.shop_region_of(target)
    if te.shop_region_of(here) == code:
        log(f"→ 已经在目标店铺（shop_region={code}），ensure_shop 会**不动页面**。")
    elif not args.go:
        log(f"→ 需要导航。加 --go 才会真的跳（会整页刷新）。")
    else:
        log(f"→ 真的导航到 {target}")
        r = te.ensure_shop(browser, want, log=log, settle=args.settle, wait=args.wait)
        log("")
        log(f"结果：ok={r['ok']}　switched={r['switched']}　reason={r['reason'] or '-'}")
        log(f"落地：{r['url'][:130]}")
        log(f"页面显示：{r['current'].splitlines()[0] if r['current'] else '(读不到)'}")
        if not r["ok"]:
            log("→ 没确认切成功：看上面落地 URL 的 shop_region，别急着往下跑抓取。")
            return 1
        log("→ 切成功了。可以接着跑流水线（取 cookie 会按这个域名精确取）。")
        return 0
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
