"""真机驱动：检测 → 取图 → 算缺口 → （可选）真的拖。

用法（紫鸟店铺浏览器已开、页面上正显示验证码时）：

    python devtools/live_captcha_solve.py                # 干跑：只算距离，不动页面
    python devtools/live_captcha_solve.py --port 17700   # 指定 CDP 端口
    python devtools/live_captcha_solve.py --go           # 真的拖
    python devtools/live_captcha_solve.py --go --watch 60  # 盯 60s：一出现就自动解

默认**干跑**：取图、算距离、落盘 devtools/_captcha_dump/，不点不拖。
"""

from __future__ import annotations

import argparse
import json
import sys
import time
import urllib.request
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent
sys.path.insert(0, str(PROJECT_DIR))

import captcha_solver as cs      # noqa: E402
import ziniao_bridge as zb       # noqa: E402

DUMP = BASE_DIR / "_captcha_dump"


def log(msg: str = "") -> None:
    print(msg, flush=True)


def find_port(arg: str) -> int:
    """--port 优先；否则扫本机 CDP 端口，挑页面在 tiktokshopglobalselling 的那个。"""
    if arg:
        return int(arg)
    hits = []
    for port in list(range(9000, 40000)):
        try:
            raw = urllib.request.urlopen(
                f"http://127.0.0.1:{port}/json/list", timeout=0.3).read()
            items = json.loads(raw.decode("utf-8", "ignore"))
        except Exception:
            continue
        for t in items:
            if t.get("type") == "page" and "tiktokshopglobalselling" in str(t.get("url")):
                hits.append(port)
                break
    if not hits:
        raise SystemExit("没找到开着商家后台的 CDP 端口（--port 手动指定）")
    if len(hits) > 1:
        log(f"[提示] 多个端口都有后台页面：{hits}，用第一个")
    return hits[0]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", default="", help="CDP 调试端口")
    ap.add_argument("--go", action="store_true", help="真的拖（不加只算不动）")
    ap.add_argument("--watch", type=float, default=0,
                    help="盯 N 秒：验证码一出现就自动解（需配合 --go）")
    args = ap.parse_args()

    port = find_port(args.port)
    log(f"CDP 端口 {port}")
    browser = zb.open_session(port, log=lambda m: log(f"  [cdp] {m}"),
                              home_urls=[zb.TIKTOK_SELLER_HOME])
    log(f"页面：{browser.current_url()[:120]}")

    if args.watch and args.go:
        log(f"盯 {args.watch:.0f}s，验证码一出现就自动解…")
        deadline = time.time() + args.watch
        while time.time() < deadline:
            det = cs.detect(browser)
            if det is not None:
                log("→ 检测到验证码！")
                r = cs.solve(browser, log=log, dump_dir=DUMP)
                log(f"结果：{r}")
                return 0 if r.get("ok") else 1
            time.sleep(1.5)
        log("时限内没等到验证码。")
        return 0

    det = cs.detect(browser)
    if det is None:
        log("页面上没有验证码（容器不存在）。")
        return 0
    log(f"检测到验证码：bg={json.dumps(det.get('bg'), ensure_ascii=False)[:160]}")
    log(f"              piece={json.dumps(det.get('piece'), ensure_ascii=False)[:160]}")
    log(f"              icon={det.get('icon')}")

    bg_bytes, pc_bytes = cs.grab_images(browser, det)
    DUMP.mkdir(parents=True, exist_ok=True)
    (DUMP / "bg.png").write_bytes(bg_bytes)
    (DUMP / "piece.png").write_bytes(pc_bytes)
    log(f"图片已落盘：{DUMP}/bg.png、piece.png")

    gap = cs.find_gap(bg_bytes, pc_bytes)
    dist = cs.drag_distance(det, gap)
    dddd = gap.get("dddd_x")
    log(f"缺口：left={gap['x']}（原图坐标，{gap['method']}）"
        f"{'，ddddocr 参考 ' + str(dddd)}" if dddd else
        f"缺口：left={gap['x']}（原图坐标，{gap['method']}）")
    log(f"→ 拼图需移动 {dist:.1f}px（显示坐标）")

    if not args.go:
        log("（干跑结束：没点、没拖。要真拖加 --go）")
        return 0
    log("→ 真的拖…")
    cs._drag(browser, det, dist, log)
    deadline = time.time() + 12
    while time.time() < deadline:
        time.sleep(0.8)
        if cs.detect(browser) is None:
            log("✅ 验证通过（验证框消失）")
            return 0
        log("  …验证框还在")
    log("✖ 12s 内没通过。")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
