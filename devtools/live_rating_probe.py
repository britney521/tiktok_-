"""真机探测：商品评分页点「导出」时，**浏览器到底发了哪些请求**。

为什么需要它：「拦截不到下载请求」有太多种可能 —— URL 不含 ``easesafe``、请求由
**新标签页/新 target** 发出（Fetch 拦截只挂在当前 target 上）、还是页面根本没发请求
（文件在本地由 blob 拼出来）。离线 mock 里那条下载 URL 是**测试桩自己编的**，证明不了
真机行为，所以这里直接连真实店铺浏览器，把点「导出」之后 60 秒内的**全部**网络事件
打出来。

用法（紫鸟 web_driver 模式客户端已在跑）::

    # 只读：打开评分页，打印页面上的按钮结构（不点任何东西）
    python devtools/live_rating_probe.py

    # 真的点「导出」，并把 60s 内所有网络事件打出来（会真的生成一次导出）
    python devtools/live_rating_probe.py --go --seconds 90

    python devtools/live_rating_probe.py --store 秋图2 --want United States --port 27206

默认**只读**：只导航到评分页并 dump 页面结构。只有 ``--go`` 才会真的点「导出」。
点了之后就等事件，不拦任何请求 —— 宁可让浏览器真下载一次（导出一条记录而已），
也要先看清真实 URL 长什么样。
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent
sys.path.insert(0, str(PROJECT_DIR))

import tiktok_export as te      # noqa: E402
import ziniao_bridge as zb      # noqa: E402


def log(msg: str = "") -> None:
    print(msg, flush=True)


# 页面上能看到什么按钮 / 面板：判断「导出」按钮是不是还能按 XPath 找到
DUMP_JS = r"""
(() => {
  const vis = (e) => e && e.offsetParent !== null;
  const buttons = [];
  document.querySelectorAll('button').forEach((b) => {
    if (!vis(b)) { return; }
    buttons.push({
      text: (b.innerText || '').trim().replace(/\s+/g, ' ').slice(0, 40),
      aria: b.getAttribute('aria-label') || '',
      log: b.getAttribute('data-log_click_for') || '',
      cls: String(b.className || '').split(' ')[0].slice(0, 40),
      id: b.id || '',
    });
  });
  const xp = %s;
  let xpHit = null;
  try {
    const r = document.evaluate(xp, document, null,
                                XPathResult.FIRST_ORDERED_NODE_TYPE, null);
    const el = r.singleNodeValue;
    xpHit = el ? {tag: el.tagName, text: (el.innerText || '').trim().slice(0, 40),
                  visible: vis(el)} : null;
  } catch (e) { xpHit = {err: String(e)}; }
  return JSON.stringify({url: location.href, title: document.title,
                         buttons: buttons, xpath_hit: xpHit});
})()
""" % json.dumps(list(te.RATING_EXPORT_XPATHS), ensure_ascii=False)

# 页面 hook 缓存里都收了哪些接口（判断「导出历史」接口在评分页叫什么）
HOOK_DUMP_JS = r"""
JSON.stringify((window.__tk_exp || []).map(function (x) {
  return {url: String(x.url).slice(0, 200), status: x.status,
          len: String(x.text || '').length};
}))
"""

# 延迟点击：先让 eval 立刻返回，再去点 —— 这样点击引发的**第一条**网络事件也不会
# 因为「等 eval 响应」而漏掉（CDP 事件在 call() 里会被丢弃，这是已知坑）。
ARM_CLICK_JS = (
    "setTimeout(function(){try{window.__probe_click=" + te.CLICK_RATING_EXPORT_JS +
    ";}catch(e){window.__probe_click='ERR '+String(e);}},1500); 'armed'"
)

# 临时挂在**当前页面**上的响应记录器：只收评分导出那几个接口，把响应体留下来。
# （临时 hook 只影响这次探测，不落盘；App 自己的 HOOK_JS 是 add_init_script 注入的。）
API_HOOK_JS = r"""
(() => {
  if (window.__probe_api) { return 'already'; }
  window.__probe_api = [];
  const KEYS = ['biz_backend/export', 'export/job', 'export/start',
                'export/update', 'easesafe', 'export_record'];
  const keep = (u) => {
    const s = String(u);
    for (let i = 0; i < KEYS.length; i++) { if (s.indexOf(KEYS[i]) >= 0) { return true; } }
    return false;
  };
  const push = (url, status, text) => {
    try {
      if (window.__probe_api.length > 60) { return; }
      window.__probe_api.push({url: String(url), status: status,
                               len: String(text || '').length,
                               text: String(text || '').slice(0, 8000)});
    } catch (e) {}
  };
  const of = window.fetch;
  window.fetch = function (...a) {
    let url = '';
    try {
      const u = a[0];
      url = (typeof u === 'string') ? u : ((u && u.url) || '');
    } catch (e) {}
    const p = of.apply(this, a);
    try {
      if (keep(url)) {
        p.then((r) => r.clone().text().then((t) => push(url, r.status, t))
                          .catch(() => {})).catch(() => {});
      }
    } catch (e) {}
    return p;
  };
  const xo = XMLHttpRequest.prototype.open, xs = XMLHttpRequest.prototype.send;
  XMLHttpRequest.prototype.open = function (m, u, ...r) {
    this.__probe_u = String(u);
    return xo.call(this, m, u, ...r);
  };
  XMLHttpRequest.prototype.send = function (b) {
    const self = this;
    try {
      this.addEventListener('load', () => {
        try { if (keep(self.__probe_u)) {
          push(self.__probe_u, self.status, self.responseText || '');
        } } catch (e) {}
      });
    } catch (e) {}
    return xs.call(this, b);
  };
  return 'ok';
})()
"""

API_DUMP_JS = "JSON.stringify(window.__probe_api || [])"


def resolve_port(args) -> int:
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
    ret = client.open_store(hit[0])
    port = ret.get("debuggingPort")
    if not port:
        raise SystemExit(f"startBrowser 没返回调试端口：{ret}")
    log(f"店铺「{zb.store_label(hit[0])}」→ 调试端口 {port}")
    return int(port)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--store", default="秋图2", help="店铺名（模糊匹配）")
    ap.add_argument("--port", default="", help="直接指定调试端口，绕过客户端 API")
    ap.add_argument("--want", default="United States", help="后台店铺（显示名/区域码）")
    ap.add_argument("--go", action="store_true", help="真的点「导出」（不加则只读）")
    ap.add_argument("--api", action="store_true",
                    help="真的点「导出」并留下接口响应体（看记录/下载 URL 长什么样）")
    ap.add_argument("--app", action="store_true",
                    help="端到端跑一遍 App 的 RatingExporter.export()（验证修复）")
    ap.add_argument("--out", default="/tmp/live_rating.xlsx", help="--app 的落盘路径")
    ap.add_argument("--seconds", type=float, default=60.0, help="点击后收事件多久")
    args = ap.parse_args()

    port = resolve_port(args)
    browser = zb.open_session(port, log=lambda m: log(f"  [cdp] {m}"),
                              home_urls=[zb.TIKTOK_SELLER_HOME])

    log("")
    log("=" * 74)
    log("  ① 切到目标店铺 + 打开商品评分页（= App 的 ⑤ 走的那条路）")
    log("=" * 74)
    shop = te.ensure_shop(browser, args.want, log=log)
    log(f"店铺：ok={shop['ok']} 落地={shop['url'][:110]}")
    exporter = te.RatingExporter(browser, log=log)
    exporter.base = zb.origin_of(shop.get("url") or browser.current_url())
    log(f"评分页站：{exporter.base}")
    exporter.prepare()
    log(f"现在在：{browser.current_url()[:130]}")

    log("")
    log("=" * 74)
    log("  ② 页面结构（只读）：可见按钮 + 导出按钮 XPath 命中情况")
    log("=" * 74)
    raw = browser.eval(DUMP_JS, await_promise=False, timeout=30)
    try:
        dump = json.loads(str(raw))
    except Exception:
        log(f"读页面结构失败：{raw!r}")
        return 1
    log(f"URL  ：{dump.get('url')}")
    log(f"标题 ：{dump.get('title')}")
    log(f"XPath 命中：{dump.get('xpath_hit')}")
    log(f"可见按钮 {len(dump.get('buttons') or [])} 个：")
    for b in (dump.get("buttons") or []):
        log(f"  - text={b['text']!r} aria={b['aria']!r} "
            f"data-log={b['log']!r} class={b['cls']!r} id={b['id']!r}")

    log("")
    log("=" * 74)
    log("  ③ 页面 hook 已经收到的接口（__tk_exp）")
    log("=" * 74)
    for item in (browser.read_json(HOOK_DUMP_JS, default=[]) or []):
        log(f"  [{item.get('status')}] {item.get('len')}B  {item.get('url')}")

    if args.app:
        log("")
        log("=" * 74)
        log("  ⑤ 端到端：直接跑 App 的 RatingExporter.export()（真机）")
        log("=" * 74)
        out = Path(args.out)
        exporter = te.RatingExporter(browser, log=log)
        exporter.base = zb.origin_of(browser.current_url())
        try:
            res = exporter.export(out_path=out, timeout=args.seconds * 3,
                                  poll=5.0, direct_wait=20.0)
        except Exception as exc:
            log(f"❌ 导出失败：{type(exc).__name__}: {exc}")
            return 1
        header = res.get("header") or []
        rows = res.get("rows") or []
        log(f"✅ 导出成功：{len(header)} 列 × {len(rows)} 行 → {res.get('raw_path')}")
        log(f"列名：{header}")
        for r in rows[:3]:
            log(f"  {r}")
        return 0

    if not (args.go or args.api):
        log("")
        log("→ 只读模式结束。加 --go / --api 才会真的点「导出」。")
        return 0

    if args.api:
        log("")
        log("=" * 74)
        log("  ④b 点「导出」→ 把评分导出接口的**响应体**留下来（看清记录/下载 URL）")
        log("=" * 74)
        log(f"临时 hook：{browser.eval(API_HOOK_JS, await_promise=False, timeout=20)}")
        got = browser.eval(ARM_CLICK_JS, await_promise=False, timeout=20)
        log(f"点击已布防：{got}（1.5s 后真的点）")
        browser.pump(args.seconds, on_event=lambda msg: None)
        log(f"按钮点击返回值：{browser.eval('String(window.__probe_click)', await_promise=False, timeout=20)}")
        items = browser.read_json(API_DUMP_JS, default=[]) or []
        log(f"共留下 {len(items)} 条响应：")
        for it in items:
            log("")
            log("-" * 70)
            log(f"[{it.get('status')}] {it.get('len')}B  {str(it.get('url'))[:170]}")
            log(str(it.get("text") or "")[:2500])
        log("")
        log("当前标签页：")
        for pg in (browser.pages() or []):
            log(f"  {str(pg.get('type')):8s} {str(pg.get('url'))[:150]}")
        return 0

    log("")
    log("=" * 74)
    log("  ④ 真的点「导出」→ 抓网络事件（不拦请求，看清真实 URL）")
    log("=" * 74)
    for cmd, params in (("Page.enable", {}), ("Network.enable", {}),
                        ("Target.setDiscoverTargets", {"discover": True})):
        try:
            browser.call(cmd, params, timeout=20)
        except Exception as exc:
            log(f"  {cmd} 失败（继续）：{exc}")

    got = browser.eval(ARM_CLICK_JS, await_promise=False, timeout=20)
    log(f"点击已布防：{got}（1.5s 后真的点）")

    seen: list[str] = []
    hits: list[str] = []

    def on_event(msg: dict) -> None:
        m = msg.get("method") or ""
        p = msg.get("params") or {}
        if m == "Network.requestWillBeSent":
            req = p.get("request") or {}
            url = str(req.get("url") or "")
            kind = str(p.get("type") or "")
            line = f"[REQ/{kind}] {url[:180]}"
            if url not in seen:
                seen.append(url)
                log(line)
            if "easesafe" in url:
                hits.append(url)
                log("  ^^^ 含 easesafe（现在的拦截模式认的就是这个）")
        elif m == "Page.downloadWillBegin":
            log(f"[DOWNLOAD] {p.get('suggestedFilename')} ← {str(p.get('url'))[:180]}")
            hits.append(str(p.get("url") or ""))
        elif m in ("Page.downloadProgress",):
            log(f"[DOWNLOAD-PROGRESS] {p}")
        elif m in ("Target.targetCreated", "Target.targetInfoChanged"):
            info = p.get("targetInfo") or {}
            if str(info.get("type")) in ("page", "other"):
                log(f"[TARGET/{m.split('.')[-1]}] {info.get('type')} "
                    f"{str(info.get('url'))[:140]}")
        elif m in ("Page.frameNavigated", "Page.navigatedWithinDocument"):
            fr = p.get("frame") or {}
            if not fr.get("parentId"):
                log(f"[NAV] {str(fr.get('url') or p.get('url'))[:160]}")

    browser.pump(args.seconds, on_event=on_event)

    log("")
    log(f"点击后 {args.seconds:.0f}s 内共 {len(seen)} 个不同请求。")
    log(f"命中的下载 URL：{hits or '(一个都没命中)'}")
    log(f"按钮点击返回值：{browser.eval('String(window.__probe_click)', await_promise=False, timeout=20)}")
    log("")
    log("页面 hook 现在收到的接口：")
    for item in (browser.read_json(HOOK_DUMP_JS, default=[]) or []):
        log(f"  [{item.get('status')}] {item.get('len')}B  {item.get('url')}")
    log("")
    log("当前标签页：")
    for pg in (browser.pages() or []):
        log(f"  {str(pg.get('type')):8s} {str(pg.get('url'))[:150]}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
