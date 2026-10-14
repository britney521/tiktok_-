#!/usr/bin/env python3.10
"""抓「店铺后台自己发的订单请求」，并定位它的发送位置 —— 一次跑清。

为什么需要它:
    我们本地拼参数（``aid=4068`` + ``seller_id``）在跨区域店上恒定
    ``code=98001008 需要权限``。实测把 seller_id 整段删掉也照样 98001008，
    说明「这个请求该怎么发」只有页面自己知道，猜不得:
      · 页面发的可能是**同源** ``seller.<region>.tiktokshopglobalselling.com/api/...``，
        而不是我们硬编码的 ``api16-normal-no1a...``；
      · URL 前缀（如北美线的 ``/na/``）、aid、locale、页码/排序字段、
        以及必要的请求头，都必须**照抄页面真实请求**。

它做四件事:
    [1] 用 CDP ``Network`` 域抓真实请求：完整 URL / query / headers / postData
        ＋ ``initiator.stack.callFrames``（浏览器自己给出的「哪段 JS 发的」）；
    [2] 页面内 hook（fetch/XHR）兜底记录 URL / body / 调用栈 —— 即使
        导航把 CDP 会话打断、事件丢了，也还能拿到；
    [3] 顺着调用栈去那个 JS 文件里搜接口与参数关键词，打印构造片段
        —— 这就是「发送请求位置」；
    [4] ``--probe-get``：把抓到的 ``order/list`` 当模板，推导出若干候选
        ``order/get`` URL，在页面内原样重放，看哪个 code=0。

用法:
    # 接管第 2 个店铺（嘉恒业），打开订单列表，等 30s 抓包
    python devtools/capture_order_request.py --store 2 --wait 30

    # 浏览器已经在跑：直接接管端口，不导航，靠手动点击触发
    python devtools/capture_order_request.py --port 16896 --no-nav --wait 45

    # 打开指定页面（美国站订单列表），并自动点第一个含「详情」的元素
    python devtools/capture_order_request.py --store 2 \\
        --url 'https://seller.us.tiktokshopglobalselling.com/order?selected_sort=6&tab=all' \\
        --click 详情 --wait 25

    # 抓完顺手用订单表第 1 条做候选 URL 重放，找出能通的写法
    python devtools/capture_order_request.py --store 2 --probe-get ../订单.xlsx

说明:
    · 只读：不关店、不改文件（除 ``--out`` 写的模板 JSON）。
    · 依赖上级目录的 ziniao_bridge / tiktok_client。
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import time
import urllib.parse
from pathlib import Path
from typing import Any

BASE_DIR = Path(__file__).resolve().parent           # .../tiktok商家/devtools
PROJECT_DIR = BASE_DIR.parent                         # .../tiktok商家
sys.path.insert(0, str(PROJECT_DIR))

import tiktok_client as tc      # noqa: E402
import ziniao_bridge as zb      # noqa: E402

DEFAULT_OUT = BASE_DIR / "captured_order_request.json"

# 只关心这些接口
NEEDLE = "fulfillment"

# 重放时不能自己设的请求头（浏览器会接管，或属于禁设名单）
FORBIDDEN_HEADERS = {
    "cookie", "content-length", "host", "connection", "accept-encoding",
    "origin", "referer", "user-agent", "accept-language", "priority",
    "sec-fetch-mode", "sec-fetch-site", "sec-fetch-dest",
    "sec-ch-ua", "sec-ch-ua-mobile", "sec-ch-ua-platform",
}

# 列表接口里的分页/排序字段，做 get 时不该带
PAGE_KEYS = {
    "page_size", "pageSize", "page_token", "pageToken", "cursor", "sort",
    "sort_by", "sort_field", "sort_order", "page", "page_number",
    "pageNumber", "offset", "limit", "tab", "selected_sort", "search_type",
    "create_time_ge", "create_time_lt", "order_status", "sort_type",
}


# ---------------------------------------------------------------------------
# 页面内 hook：fetch / XHR 只要命中 NEEDLE 就记下来（含调用栈）
# ---------------------------------------------------------------------------
HOOK_JS = r"""
(() => {
  const KEY = 'fulfillment';
  if (window.__CAP_HOOKED) return 'already';
  window.__CAP_HOOKED = true;
  window.__CAP = [];
  const st = () => { try { return String((new Error()).stack || ''); } catch (e) { return ''; } };
  const push = (r) => { try { window.__CAP.push(r); if (window.__CAP.length > 300) window.__CAP.shift(); } catch (e) {} };

  const of = window.fetch;
  if (of) {
    window.fetch = function (input, init) {
      try {
        const url = (typeof input === 'string') ? input : ((input && input.url) || '');
        if (String(url).indexOf(KEY) >= 0) {
          const method = String((init && init.method) || (input && input.method) || 'GET').toUpperCase();
          let body = init && init.body;
          if (body && typeof body !== 'string') { try { body = JSON.stringify(body); } catch (e) { body = String(body); } }
          push({ kind: 'fetch', url: String(url), method: method, body: body || null,
                 headers: (init && init.headers) || null, stack: st(), ts: Date.now() });
        }
      } catch (e) {}
      return of.apply(this, arguments);
    };
  }

  const XO = XMLHttpRequest.prototype.open;
  const XS = XMLHttpRequest.prototype.send;
  const XH = XMLHttpRequest.prototype.setRequestHeader;
  XMLHttpRequest.prototype.open = function (m, u) {
    try { this.__cap = { kind: 'xhr', method: String(m || 'GET').toUpperCase(),
                         url: String(u || ''), headers: {}, body: null,
                         stack: st(), ts: Date.now() }; } catch (e) {}
    return XO.apply(this, arguments);
  };
  XMLHttpRequest.prototype.setRequestHeader = function (k, v) {
    try { if (this.__cap) this.__cap.headers[String(k)] = String(v); } catch (e) {}
    return XH.apply(this, arguments);
  };
  XMLHttpRequest.prototype.send = function (b) {
    try {
      if (this.__cap && String(this.__cap.url).indexOf(KEY) >= 0) {
        this.__cap.body = (typeof b === 'string') ? b : (b ? String(b) : null);
        push(this.__cap);
      }
    } catch (e) {}
    return XS.apply(this, arguments);
  };
  return 'installed';
})()
"""

CLICK_JS = r"""
((needle) => {
  const els = Array.from(document.querySelectorAll('button,a,div[role="button"],span,td,tr'));
  const hit = els.find(e => (e.innerText || '').indexOf(needle) >= 0);
  if (!hit) return 'not-found';
  hit.click();
  return 'clicked:' + String(hit.innerText || '').slice(0, 40).replace(/\n/g, ' ');
})
"""

READ_CAP_JS = "JSON.stringify(window.__CAP || [])"


def log(msg: str) -> None:
    print(msg, flush=True)


def js(value: Any) -> str:
    """把 Python 值变成合法的 JS 字面量（json.dumps 的输出即是）。"""
    return json.dumps(value, ensure_ascii=False)


# ---------------------------------------------------------------------------
# 端口 / 店铺
# ---------------------------------------------------------------------------
def _pick_interactive(stores: list[dict]) -> dict:
    log("可用店铺：")
    for i, st in enumerate(stores, 1):
        log(f"  [{i}] {zb.store_label(st)}")
    raw = input("选择序号（回车取消）：").strip()
    if not raw.isdigit() or not 1 <= int(raw) <= len(stores):
        raise SystemExit("已取消 / 序号无效")
    return stores[int(raw) - 1]


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
            raise SystemExit(f"{name!r} 匹配到 {len(hit)} 个，请用 --store 指定序号")
        return hit[0]
    return _pick_interactive(stores)


def _resolve_port(args) -> tuple[int, str]:
    """返回 (调试端口, 后台首页)。"""
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
        raise SystemExit("startBrowser 没返回 debuggingPort；"
                         "若浏览器已在跑，请用 --port 直接接管。")
    return int(port), str(ret.get("launcherPage") or zb.TIKTOK_SELLER_HOME)


# ---------------------------------------------------------------------------
# 抓包
# ---------------------------------------------------------------------------
def _short(url: str, n: int = 150) -> str:
    return url if len(url) <= n else url[:n] + "…"


def _frames_of_initiator(params: dict) -> list[str]:
    """从 Network 事件的 initiator 里取出调用栈（浏览器视角的发送位置）。"""
    out: list[str] = []
    init = (params.get("initiator") or {})
    stack = init.get("stack") or {}
    for f in (stack.get("callFrames") or []):
        u = str(f.get("url") or "")
        if not u:
            continue
        out.append(f"{u}:{f.get('lineNumber')}:{f.get('columnNumber')}"
                   f"  [{f.get('functionName') or '(anonymous)'}]")
    if not out and init.get("url"):
        out.append(f"{init.get('url')}:{init.get('lineNumber')}:{init.get('columnNumber')}")
    return out


def _frames_of_stack(stack: str) -> list[str]:
    """从 JS ``new Error().stack`` 里取出 ``*.js:line:col``。"""
    out: list[str] = []
    for line in str(stack or "").splitlines():
        m = re.search(r"(https?://[^\s)]+?\.js(?:\?[^\s)]*)?:(\d+):(\d+))", line)
        if m:
            out.append(m.group(1))
    return out


def capture(browser, *, url: str | None, wait: float, clicks: list[str],
            no_nav: bool) -> tuple[list[dict], list[dict]]:
    """返回 (Network 事件抓到的请求, 页面 hook 记录的请求)。"""
    net: list[dict] = []

    def on_event(msg: dict) -> None:
        if msg.get("method") != "Network.requestWillBeSent":
            return
        p = msg.get("params") or {}
        r = p.get("request") or {}
        u = str(r.get("url") or "")
        if NEEDLE not in u:
            return
        net.append({
            "url": u,
            "method": str(r.get("method") or ""),
            "headers": dict(r.get("headers") or {}),
            "post_data": r.get("postData"),
            "post_data_entries": r.get("postDataEntries"),
            "resource_type": p.get("type"),
            "frames": _frames_of_initiator(p),
        })
        log(f"  ⚡ {r.get('method')} {_short(u, 120)}")
        pd = r.get("postData")
        if pd:
            log(f"        body: {_short(str(pd), 160)}")

    # [1] CDP 网络域
    try:
        browser.call("Network.enable", {"maxPostDataSize": 1_000_000}, timeout=20)
        log("Network 域已启用（抓真实请求头与 initiator 调用栈）")
    except Exception as exc:
        log(f"Network.enable 失败（降级为只用页面 hook）：{exc}")

    # [2] 页面 hook（当前文档）
    try:
        browser.eval(HOOK_JS, await_promise=False, timeout=20)
        log("页面 hook 已注入（fetch/XHR 记录调用栈）")
    except Exception as exc:
        log(f"页面 hook 注入失败：{exc}")

    # [3] 导航（可选）
    if url and not no_nav:
        log(f"打开：{url}")
        try:
            zb.goto(browser, url, log=log)
        except Exception as exc:
            log(f"导航失败（继续抓）：{exc}")
        # 导航会重建文档 / 可能断 CDP：补装 hook + Network
        try:
            browser.eval(HOOK_JS, await_promise=False, timeout=20)
        except Exception as exc:
            log(f"导航后补装 hook 失败：{exc}")
        try:
            browser.call("Network.enable", {"maxPostDataSize": 1_000_000}, timeout=20)
        except Exception as exc:
            log(f"导航后补 Network.enable 失败：{exc}")

    # [4] 自动点击（触发详情请求）
    for text in clicks:
        try:
            res = browser.eval(f"({CLICK_JS})({js(text)})", await_promise=False, timeout=20)
            log(f"点击含 {text!r} 的元素 → {res}")
        except Exception as exc:
            log(f"点击 {text!r} 失败：{exc}")
        time.sleep(1.0)

    # [5] 收事件
    if wait > 0:
        log(f"开始收事件 {wait:.0f}s —— 需要的话现在在浏览器里点开一个订单 / 刷新列表…")
        deadline = time.time() + wait
        while time.time() < deadline:
            remaining = deadline - time.time()
            try:
                browser.pump(min(2.0, max(0.3, remaining)), on_event=on_event)
            except Exception:
                time.sleep(0.3)
        log("收事件结束")

    # [6] 读页面 hook 记录
    recs: list[dict] = []
    try:
        got = browser.read_json(READ_CAP_JS, default=[])
        if isinstance(got, list):
            recs = [x for x in got if isinstance(x, dict)]
    except Exception as exc:
        log(f"读取页面 hook 记录失败：{exc}")
    return net, recs


# ---------------------------------------------------------------------------
# 定位：顺调用栈去 JS 里找构造片段
# ---------------------------------------------------------------------------
SCRIPT_SCAN_JS = r"""
(async (spec) => {
  const out = [];
  let txt = '';
  try {
    const r = await fetch(spec.url, {credentials: 'include'});
    txt = await r.text();
  } catch (e) {
    return JSON.stringify([{error: String(e), url: spec.url}]);
  }
  const size = txt.length;
  for (const n of spec.needles) {
    let i = txt.indexOf(n), c = 0;
    while (i >= 0 && c < spec.per) {
      out.push({ needle: n, idx: i, size: size,
                 ctx: txt.slice(Math.max(0, i - 240), i + 240) });
      i = txt.indexOf(n, i + 1); c++;
    }
  }
  return JSON.stringify(out);
})
"""

SCAN_NEEDLES = ["fulfillment/order/get", "fulfillment/order/list",
                "oec_seller_id", "main_order_id", "fulfillment/na/"]


def _site_key(netloc: str) -> str:
    """seller.us.tiktokshopglobalselling.com → tiktokshopglobalselling.com。"""
    host = str(netloc or "").split(":")[0]
    parts = host.split(".")
    return ".".join(parts[-2:]) if len(parts) > 2 else host


def _same_site(url: str, page: str) -> bool:
    a = _site_key(urllib.parse.urlsplit(str(url)).netloc)
    b = _site_key(urllib.parse.urlsplit(str(page)).netloc)
    return bool(a) and a == b


_SCRIPT_RE = re.compile(r"(https?://[^\s'\"]+?\.js(?:\?[^\s'\"]*)?)")


def _script_url(frame: str) -> str:
    m = _SCRIPT_RE.search(str(frame or ""))
    return m.group(1) if m else ""


def locate_source(browser, frames: list[str], page: str, per: int = 3) -> list[dict]:
    """在调用栈指向的 JS 文件里搜接口/参数关键词。"""
    seen: set[str] = set()
    out: list[dict] = []
    for fr in frames:
        u = _script_url(fr)
        if not u or u in seen:
            continue
        seen.add(u)
        if not _same_site(u, page):
            continue
        try:
            raw = browser.eval(f"({SCRIPT_SCAN_JS})({js({'url': u, 'needles': SCAN_NEEDLES, 'per': per})})",
                               await_promise=True, timeout=60)
            hits = json.loads(raw) if isinstance(raw, str) else (raw or [])
        except Exception as exc:
            out.append({"script": u, "error": str(exc)})
            continue
        if hits:
            out.append({"script": u, "hits": hits})
        if len(out) >= 3:
            break
    return out


# ---------------------------------------------------------------------------
# 重放：用 order/list 当模板，推导 order/get 候选
# ---------------------------------------------------------------------------
def derive_get_paths(path: str) -> list[str]:
    """由列表接口路径推导详情接口路径。"""
    cands: list[str] = []
    p = path or ""

    def add(x: str) -> None:
        if x and x not in cands:
            cands.append(x)

    if p.endswith("/list"):
        add(p[:-len("/list")] + "/get")
        add(p[:-len("/list")] + "/detail")
    if "/na/" in p:
        add(p.replace("/na/order/list", "/na/order/get"))
        add(p.replace("/na/", "/"))
    if "/fulfillment/order/list" in p:
        add(p.replace("/fulfillment/order/list", "/fulfillment/order/get"))
    add("/api/fulfillment/order/get")
    add("/api/fulfillment/na/order/get")
    return cands


REPLAY_JS = r"""
(async (spec) => {
  const headers = Object.assign({'content-type': 'application/json'}, spec.headers || {});
  try {
    const r = await fetch(spec.url, {
      method: 'POST', headers: headers,
      credentials: 'include', body: JSON.stringify(spec.body || {})
    });
    const t = await r.text();
    let code = null, msg = null, nick = null;
    try {
      const j = JSON.parse(t);
      code = j.code; msg = j.message;
      const mo = (j.data && j.data.main_order) || [];
      if (mo[0] && mo[0].buyer_info_module) nick = mo[0].buyer_info_module.buyer_nickname;
    } catch (e) {}
    return JSON.stringify({status: r.status, code: code, message: msg, buyer: nick,
                           text: t.slice(0, 300)});
  } catch (e) {
    return JSON.stringify({error: String(e)});
  }
})
"""


def read_order_ids(path: Path, limit: int = 1) -> list[str]:
    import openpyxl
    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        return []
    hdr = [str(h) for h in rows[0]]
    idx = hdr.index("Order ID") if "Order ID" in hdr else 0
    out = [str(r[idx]).strip() for r in rows[1:] if r and r[idx] not in (None, "")]
    return out[:limit]


def probe_get(browser, net: list[dict], order_ids: list[str], seller_id: str) -> None:
    log("")
    log("=" * 72)
    log("候选 order/get 重放（页面内 fetch，照抄页面真实 host/query/头）")
    log("=" * 72)
    if not order_ids:
        log("没有订单号，跳过")
        return

    tmpl = None
    for r in net:
        if "/fulfillment/" in str(r.get("url")) and str(r.get("method")).upper() in ("GET", "POST"):
            tmpl = r
            break
    if tmpl:
        u = urllib.parse.urlsplit(tmpl["url"])
        base_q = {k: v for k, v in urllib.parse.parse_qsl(u.query, keep_blank_values=True)
                  if k not in PAGE_KEYS}
        base_q.setdefault("aid", tc.DETAIL_AID)
        host_prefix = f"{u.scheme}://{u.netloc}"
        paths = derive_get_paths(u.path)
        skip = FORBIDDEN_HEADERS
        extra = {k: v for k, v in (tmpl.get("headers") or {}).items()
                 if k.lower() not in skip and k.lower() not in ("content-type", "accept")}
        log(f"模板来自：{tmpl['method']} {_short(tmpl['url'], 120)}")
        log(f"  host={host_prefix}  保留 query={list(base_q)}")
        log(f"  照抄的自定义头：{list(extra) or '（无）'}")
    else:
        log("没抓到任何 fulfillment 请求 —— 回落到硬编码 host 试")
        host_prefix = tc.DEFAULT_API_BASE
        base_q = {"aid": tc.DETAIL_AID, "locale": "zh-CN", "language": "zh-CN"}
        paths = ["/api/fulfillment/order/get"]
        extra = {}

    oid = order_ids[0]
    for sid in ([seller_id] if seller_id else [""]):
        q = dict(base_q)
        if sid:
            q["oec_seller_id"] = sid
            q["seller_id"] = sid
        for p in paths:
            full = host_prefix.rstrip("/") + p + "?" + urllib.parse.urlencode(q)
            spec = {"url": full, "headers": extra, "body": {"main_order_id": [oid]}}
            script = "(" + REPLAY_JS + ")(" + js(spec) + ")"
            try:
                raw = browser.eval(script, await_promise=True, timeout=60)
                res = json.loads(raw) if isinstance(raw, str) else (raw or {})
            except Exception as exc:
                log(f"  ✗ {p:<42} 异常 {exc}")
                continue
            code = res.get("code")
            flag = "✅" if str(code) == "0" else "  "
            tail = f"buyer={res.get('buyer')}" if res.get("buyer") else \
                (res.get("message") or res.get("error") or res.get("text", "")[:60])
            log(f"  {flag} {p:<42} code={code} {tail}")
        break  # 只试一个 seller_id


# ---------------------------------------------------------------------------
def report(net: list[dict], recs: list[dict], page: str) -> None:
    log("")
    log("=" * 72)
    log(f"[1] CDP Network 抓到 {len(net)} 条 fulfillment 请求")
    log("=" * 72)
    for i, r in enumerate(net[:12], 1):
        u = urllib.parse.urlsplit(r["url"])
        log(f"{i}. {r['method']} {u.scheme}://{u.netloc}{u.path}")
        if u.query:
            for k, v in urllib.parse.parse_qsl(u.query, keep_blank_values=True):
                log(f"     · {k} = {_short(str(v), 110)}")
        if r.get("post_data"):
            log(f"   body: {_short(str(r['post_data']), 300)}")
        custom = {k: v for k, v in (r.get("headers") or {}).items()
                  if k.lower() not in ("cookie", "content-length", "accept-encoding",
                                       "accept-language", "user-agent", "priority",
                                       "sec-ch-ua", "sec-ch-ua-mobile",
                                       "sec-ch-ua-platform", "sec-fetch-mode",
                                       "sec-fetch-site", "sec-fetch-dest")}
        if custom:
            log(f"   headers: {custom}")
        for fr in (r.get("frames") or [])[:5]:
            log(f"   ↳ {_short(fr, 150)}")

    log("")
    log("=" * 72)
    log(f"[2] 页面 hook 记录 {len(recs)} 条（URL / body / 调用栈）")
    log("=" * 72)
    for i, r in enumerate(recs[:12], 1):
        log(f"{i}. [{r.get('kind')}] {r.get('method')} {_short(str(r.get('url')), 130)}")
        if r.get("body"):
            log(f"   body: {_short(str(r['body']), 200)}")
        for fr in _frames_of_stack(str(r.get("stack") or ""))[:4]:
            log(f"   ↳ {_short(fr, 150)}")

    log("")
    log("=" * 72)
    log("[3] 发送请求位置（在该 JS 里搜接口/参数构造片段）")
    log("=" * 72)
    log(f"当前页面：{page}")


def show_source(blocks: list[dict]) -> None:
    if not blocks:
        log("（没能定位到同站脚本；把上面 [1]/[2] 的输出贴出来即可继续分析）")
        return
    for b in blocks:
        if b.get("error"):
            log(f"· {_short(b['script'], 120)} → 取源码失败：{b['error']}")
            continue
        log(f"· {_short(b['script'], 130)}  （脚本 {b['hits'][0].get('size')} 字符）")
        for h in b["hits"]:
            log(f"  ▸ 命中 {h['needle']!r} @ {h['idx']}")
            ctx = str(h["ctx"]).replace("\n", " ")
            log(f"    …{ctx}…")


# ---------------------------------------------------------------------------
def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="抓后台真实订单请求并定位发送位置")
    ap.add_argument("--store", type=int, default=None, help="店铺序号（1 起）")
    ap.add_argument("--name", default=None, help="店铺名关键词")
    ap.add_argument("--port", type=int, default=None, help="调试端口（跳过紫鸟 IPC）")
    ap.add_argument("--url", default=None, help="先导航到这个页面再抓")
    ap.add_argument("--no-nav", action="store_true", help="不导航，只抓当前页面")
    ap.add_argument("--wait", type=float, default=30.0, help="收事件秒数（默认 30）")
    ap.add_argument("--click", action="append", default=[], help="注入 hook 后点击含该文本的元素（可重复）")
    ap.add_argument("--probe-get", default=None, help="订单 Excel：抓完用第 1 条重放候选 order/get")
    ap.add_argument("--seller-id", default="", help="重放时用的 seller_id（默认取页面 cookie）")
    ap.add_argument("--out", default=str(DEFAULT_OUT), help="抓到的请求模板写到哪")
    args = ap.parse_args(argv)

    port, home = _resolve_port(args)
    log(f"接管调试端口 127.0.0.1:{port}")
    browser = zb.open_session(port, log=log)
    page = browser.current_url()
    log(f"当前页面：{page or '(空)'}")

    url = None if args.no_nav else (args.url or home)
    net, recs = capture(browser, url=url, wait=args.wait,
                        clicks=list(args.click), no_nav=args.no_nav)

    page = browser.current_url()
    report(net, recs, page)

    # 定位源码
    frames: list[str] = []
    for r in net:
        frames.extend(r.get("frames") or [])
    for r in recs:
        frames.extend(_frames_of_stack(str(r.get("stack") or "")))
    blocks = locate_source(browser, frames, page) if frames else []
    show_source(blocks)

    # 存模板
    try:
        Path(args.out).write_text(json.dumps(
            {"page": page, "network": net, "page_hook": recs},
            ensure_ascii=False, indent=2), encoding="utf-8")
        log("")
        log(f"请求模板已写入：{args.out}")
    except Exception as exc:
        log(f"写模板失败：{exc}")

    # 重放候选
    if args.probe_get:
        src = Path(args.probe_get)
        ids = read_order_ids(src, 1) if src.exists() else []
        sid = args.seller_id
        if not sid:
            try:
                ck = zb.grab_cookies(browser, log=lambda _m: None, url=page)
                sid = zb.seller_id_from_cookies(ck)
            except Exception:
                sid = ""
        log(f"重放用 seller_id={sid or '(无)'}，订单={ids or '(无)'}")
        probe_get(browser, net, ids, sid)

    log("")
    log("完成。把上面的输出贴回来即可定案。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
