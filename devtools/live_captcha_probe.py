"""真机探测：抓客户时弹出来的「验证码」到底长什么样（**默认只读**）。

用途：写自动过验证码的代码之前，先把真机 DOM 摸清楚 —— 是哪一层 iframe、
背景图/拼图块的 URL 与尺寸、滑块轨道的元素与坐标、事件坐标要不要加上 iframe 偏移。

用法（紫鸟店铺浏览器已开、页面上正显示验证码时）：

    python devtools/live_captcha_probe.py                 # 只读：扫所有 target 找验证码元素
    python devtools/live_captcha_probe.py --port 17700    # 直接指定 CDP 端口
    python devtools/live_captcha_probe.py --dump          # 顺便把图片落到 devtools/_captcha_dump/
    python devtools/live_captcha_probe.py --shot          # 主页面整屏截图到 /tmp

只读保证：不点、不拖、不改页面状态；`--dump` 只写本地文件。
"""

from __future__ import annotations

import argparse
import base64
import json
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent
sys.path.insert(0, str(PROJECT_DIR))

import websocket                      # noqa: E402
import ziniao_bridge as zb            # noqa: E402

# 关键词：验证码相关的 class/id 多半带这些串
KEY = rb"captcha|verify|verification|slider|slide|drag|puzzle|jigsaw|secsdk|ucenter|nc_|geetest|yidun|shumei|check"

SCAN_JS = r"""
(() => {
  const R = (r) => ({x: Math.round(r.left), y: Math.round(r.top),
                     w: Math.round(r.width), h: Math.round(r.height)});
  const vis = (el) => {
    const r = el.getBoundingClientRect();
    if (r.width < 6 || r.height < 6) return null;
    if (r.right < -200 || r.bottom < -200 || r.left > innerWidth + 200 || r.top > innerHeight + 200) return null;
    const cs = getComputedStyle(el);
    if (cs.display === 'none' || cs.visibility === 'hidden') return null;
    if (parseFloat(cs.opacity || '1') < 0.05) return null;
    return r;
  };
  const sig = (el) => {
    const c = (el.className && el.className.baseVal !== undefined) ? el.className.baseVal : (el.className || '');
    return ((el.id || '') + ' ' + String(c)).trim();
  };
  const out = {url: location.href, title: document.title,
               vw: innerWidth, vh: innerHeight, dpr: devicePixelRatio,
               nodes: [], imgs: [], bgs: [], canvases: [], iframes: []};

  const KEY = /captcha|verify|verification|slider|slide|drag|puzzle|jigsaw|secsdk|ucenter|nc_|geetest|yidun|shumei|check/i;
  for (const el of document.querySelectorAll('*')) {
    const s = sig(el);
    if (!KEY.test(s)) continue;
    const r = vis(el);
    out.nodes.push({tag: el.tagName, id: el.id, cls: s.slice(0, 140),
                    rect: r ? R(r) : null,
                    text: (el.innerText || '').trim().slice(0, 50)});
    if (out.nodes.length > 120) break;
  }
  for (const im of document.querySelectorAll('img')) {
    const r = vis(im);
    if (!r) continue;
    const src = String(im.currentSrc || im.src || '');
    out.imgs.push({rect: R(r), nw: im.naturalWidth, nh: im.naturalHeight,
                   kind: src.slice(0, 30), len: src.length,
                   alt: String(im.alt || '').slice(0, 40),
                   cls: sig(im).slice(0, 100), src: src.length < 400 ? src : src.slice(0, 200) + '…'});
  }
  for (const el of document.querySelectorAll('*')) {
    const bg = getComputedStyle(el).backgroundImage;
    if (!bg || bg === 'none' || bg.indexOf('url(') < 0) continue;
    const r = vis(el);
    if (!r) continue;
    out.bgs.push({rect: R(r), cls: sig(el).slice(0, 100), bg: bg.slice(0, 120)});
  }
  for (const cv of document.querySelectorAll('canvas')) {
    const r = vis(cv);
    if (!r) continue;
    out.canvases.push({rect: R(r), w: cv.width, h: cv.height, cls: sig(cv).slice(0, 100)});
  }
  for (const f of document.querySelectorAll('iframe')) {
    const r = vis(f);
    if (!r) continue;
    out.iframes.push({rect: R(r), src: String(f.src || '').slice(0, 160), cls: sig(f).slice(0, 80)});
  }
  out.text = (document.body ? document.body.innerText : '').replace(/\n{2,}/g, '\n').trim().slice(0, 600);
  return JSON.stringify(out);
})()
"""


def log(msg: str = "") -> None:
    print(msg, flush=True)


def resolve_port(args) -> int:
    if args.port:
        return int(args.port)
    mods = zb.load_ziniao()
    cfg = mods["client"].ZiniaoConfig.from_yaml(mods["config_path"])
    client = zb.make_client(cfg, log=lambda m: log(f"  [client] {m}"))
    if not client.port_open():
        raise SystemExit("紫鸟客户端 IPC 端口没开：请用 web_driver 模式启动客户端")
    for s in client.get_browser_list():
        log(f"  [店铺] {zb.store_label(s)}")
    raise SystemExit("请用 --port 指定 CDP 端口（或用 --store 让脚本自己开店铺）")


def list_targets(port: int) -> list[dict]:
    import urllib.request
    raw = urllib.request.urlopen(f"http://127.0.0.1:{port}/json/list", timeout=5).read()
    return json.loads(raw.decode("utf-8", "ignore"))


def eval_on(ws_url: str, script: str, timeout: float = 20.0,
            await_promise: bool = False) -> str:
    """在某个 target 上执行 JS 并取回字符串值（异步 IIFE 记得 ``await_promise=True``）。"""
    ws = websocket.create_connection(ws_url, suppress_origin=True, timeout=timeout)
    try:
        ws.send(json.dumps({"id": 1, "method": "Runtime.evaluate",
                            "params": {"expression": script, "returnByValue": True,
                                       "awaitPromise": bool(await_promise)}}))
        while True:
            msg = json.loads(ws.recv())
            if msg.get("id") == 1:
                res = msg.get("result") or {}
                if res.get("exceptionDetails"):
                    return json.dumps({"err": str(res["exceptionDetails"])[:300]},
                                      ensure_ascii=False)
                return (res.get("result") or {}).get("value") or ""
    finally:
        try:
            ws.close()
        except Exception:
            pass


def shot(ws_url: str, out: Path) -> bool:
    """整屏截图（只读）。"""
    ws = websocket.create_connection(ws_url, suppress_origin=True, timeout=30)
    try:
        ws.send(json.dumps({"id": 1, "method": "Page.captureScreenshot",
                            "params": {"format": "png"}}))
        while True:
            msg = json.loads(ws.recv())
            if msg.get("id") == 1:
                data = (msg.get("result") or {}).get("data")
                if not data:
                    return False
                out.write_bytes(base64.b64decode(data))
                return True
    finally:
        try:
            ws.close()
        except Exception:
            pass


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", default="", help="CDP 调试端口（不填则用客户端 API 找）")
    ap.add_argument("--dump", action="store_true", help="把图片落盘到 devtools/_captcha_dump/")
    ap.add_argument("--shot", action="store_true", help="主页面整屏截图到 /tmp/captcha_page.png")
    args = ap.parse_args()

    port = resolve_port(args)
    targets = list_targets(port)
    dump_dir = BASE_DIR / "_captcha_dump"
    if args.dump:
        dump_dir.mkdir(parents=True, exist_ok=True)

    log("")
    log("=" * 78)
    log(f"  CDP 端口 {port} —— 共 {len(targets)} 个 target")
    log("=" * 78)

    main_page = next((t for t in targets
                      if t.get("type") == "page" and str(t.get("url", "")).startswith("http")),
                     None)
    if args.shot and main_page:
        ok = shot(main_page["webSocketDebuggerUrl"], Path("/tmp/captcha_page.png"))
        log(f"整屏截图：{'/tmp/captcha_page.png' if ok else '失败'}")

    for t in targets:
        if t.get("type") not in ("page", "iframe"):
            continue
        url = str(t.get("url") or "")
        if url.startswith("chrome-extension://") or url in ("about:blank", ""):
            continue
        log("")
        log("-" * 78)
        log(f"[{t.get('type')}] {url[:150]}")
        log("-" * 78)
        raw = eval_on(t["webSocketDebuggerUrl"], SCAN_JS)
        try:
            d = json.loads(raw)
        except Exception:
            log(f"  取不到（{str(raw)[:120]}）")
            continue
        if isinstance(d, dict) and d.get("err"):
            log(f"  JS 报错：{d['err'][:200]}")
            continue
        log(f"  视口 {d.get('vw')}×{d.get('vh')} dpr={d.get('dpr')}  标题 {str(d.get('title'))[:50]}")
        log(f"  命中关键词的节点 {len(d.get('nodes') or [])} 个：")
        for n in (d.get("nodes") or [])[:25]:
            log(f"    <{n['tag']}> id={n['id']!r} cls={n['cls']!r} rect={n['rect']} text={n['text']!r}")
        log(f"  可见图片 {len(d.get('imgs') or [])} 张：")
        for i, im in enumerate((d.get("imgs") or [])[:25]):
            log(f"    #{i} rect={im['rect']} 原图 {im['nw']}×{im['nh']} {im['kind']} len={im['len']} "
                f"cls={im['cls']!r}")
            if args.dump:
                put_dump(im, dump_dir, t, i)
        log(f"  背景图元素 {len(d.get('bgs') or [])} 个 / canvas {len(d.get('canvases') or [])} 个：")
        for b in (d.get("bgs") or [])[:10]:
            log(f"    bg rect={b['rect']} cls={b['cls']!r} {b['bg'][:80]}")
        for c in (d.get("canvases") or [])[:10]:
            log(f"    cv rect={c['rect']} {c['w']}×{c['h']} cls={c['cls']!r}")
        log(f"  iframe {len(d.get('iframes') or [])} 个：")
        for f in (d.get("iframes") or [])[:8]:
            log(f"    rect={f['rect']} {f['src'][:110]}")
        if d.get("text"):
            log("  正文片段：")
            for line in str(d["text"]).splitlines()[:12]:
                log(f"    | {line[:110]}")
    log("")
    log("（只读探测结束：没有点、没有拖、没有改页面状态）")
    return 0


def put_dump(im: dict, dump_dir: Path, target: dict, idx: int) -> None:
    """把 data: URL 的图片落盘，便于本地看形状。"""
    src = im.get("src") or ""
    if not src.startswith("data:") or "," not in src or len(src) < 300:
        return
    head, b64 = src.split(",", 1)
    ext = "png" if "png" in head else ("jpg" if "jpeg" in head or "jpg" in head else "bin")
    name = f"{idx}_{im['rect']['w']}x{im['rect']['h']}_nat{im['nw']}x{im['nh']}.{ext}"
    try:
        (dump_dir / name).write_bytes(base64.b64decode(b64))
        log(f"      → 已落盘 {name}")
    except Exception as exc:
        log(f"      → 落盘失败：{exc}")


if __name__ == "__main__":
    raise SystemExit(main())
