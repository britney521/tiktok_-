"""TikTok 商家后台「拼图滑块」验证码 —— 自动识别 + 拖动。

真机结构（2026-10-15 实测，seller.us 后台弹出的 OEC/Lucifer 验证码）::

    body.captcha-disable-scroll
    └ div.captcha_verify_container            [770,232,380,384]
        ├ div.captcha_verify_bar--title        「请完成下列验证后继续:」
        ├ div.captcha_verify_img--wrapper      [790,300,340,212]   ← 背景图区域
        │   ├ img#captcha-verify-image         显示 340×212 / 原图 552×344（JPEG）
        │   └ img.captcha_verify_img_slide
        │       .react-draggable               显示 68×68 / 原图 110×110（PNG 带透明）
        │                                      style="left:8px; transform:translate(0,0)"
        ├ div.captcha_verify_slide--slidebar   [790,516,340,44]    ← 轨道
        │   └ div.secsdk-captcha-drag-icon     [792,517,64,40]     ← 要按住的按钮
        └ div.captcha_verify_action                                （刷新 / 反馈 / 单号）

两张图片都在 ``p16-oec-ttp.tiktokcdn-us.com``，页面内 ``fetch(url,{credentials:'omit'})``
实测 HTTP 200 直接拿到字节（CDN 带 CORS 头），不需要借同源那套。

缺口识别（``find_gap``）三级，取第一个成功的：

  1. **cv2 暗色剪影**（主，最准）：缺口是拼图形状的深色剪影，阈值 + 连通域找暗块，
     与拼图 PNG 的 alpha 包围盒对齐 —— 「拼图不透明区左缘 → 缺口左缘」。
  2. **掩码最小亮度**（兜底）：拿 alpha 当掩码在背景上滑窗，平均亮度最低处即缺口。
  3. **ddddocr** ``slide_match(simple_target=True)``（末选 + 交叉参考）：
     真机实测它给的 x 比几何真值偏左 ~19 原图像素（≈12 显示像素，足以判失败），
     所以只当参考值打进日志 / 前两级都失败时才用。

拖动（``_drag``）：CDP ``Input.dispatchMouseEvent``（坐标 = 顶层视口 CSS 像素，
验证码就在主文档里，没有 iframe 偏移问题）。先拖 24px **实测校准**
「按钮位移 → 拼图位移」的比值（不假设 1:1），再按剩余距离走一条
先快后慢、带 2~5px 过冲和回拉的拟人轨迹。松手后轮询验证码容器是否消失。
"""

from __future__ import annotations

import base64
import json
import random
import time
from typing import Any, Callable

LogFn = Callable[[str], None]

# 页面上验证码的「指纹」选择器（一个都不 hardcode 坐标，全按 rect 现场读）
SEL_CONTAINER = ".captcha_verify_container"
SEL_BG = "#captcha-verify-image"
SEL_PIECE = "img.captcha_verify_img_slide"
SEL_ICON = ".secsdk-captcha-drag-icon"

DETECT_JS = r"""
(() => {
  const box = (el) => { if (!el) return null; const r = el.getBoundingClientRect();
    return {x: r.left, y: r.top, w: r.width, h: r.height}; };
  const src = (el) => el ? String(el.currentSrc || el.src || '') : '';
  const cont = document.querySelector('.captcha_verify_container');
  const bg = document.querySelector('#captcha-verify-image');
  const piece = document.querySelector('img.captcha_verify_img_slide');
  const icon = document.querySelector('.secsdk-captcha-drag-icon');
  if (!cont && !bg && !piece) return 'null';
  return JSON.stringify({
    container: !!cont,
    bg: bg ? {src: src(bg), nw: bg.naturalWidth, nh: bg.naturalHeight, rect: box(bg)} : null,
    piece: piece ? {src: src(piece), nw: piece.naturalWidth, nh: piece.naturalHeight,
                    rect: box(piece)} : null,
    icon: box(icon),
  });
})()
"""

FETCH_JS = r"""
(async (url) => {
  try {
    const r = await fetch(url, {credentials: 'omit'});
    const b = await r.arrayBuffer(); const u8 = new Uint8Array(b);
    let s = ''; const CH = 0x8000;
    for (let i = 0; i < u8.length; i += CH) {
      s += String.fromCharCode.apply(null, u8.subarray(i, i + CH));
    }
    return JSON.stringify({status: r.status, len: u8.length, b64: btoa(s)});
  } catch (e) {
    return JSON.stringify({err: String(e)});
  }
})
"""


class CaptchaError(RuntimeError):
    """验证码处理失败（原因写在消息里，调用方只管打日志）。"""


# 松手后等多久确认结果（容器消失=通过 / 换图=失败重试）
POLL_TIMEOUT = 12.0


# -- 检测 -------------------------------------------------------------------
def detect(browser: Any) -> dict[str, Any] | None:
    """页面上有没有验证码。有 → 结构信息；没有 → None（页面不在/读不到也算没有）。"""
    try:
        raw = browser.eval(DETECT_JS, await_promise=False, timeout=15)
    except Exception:
        return None
    text = str(raw or "").strip()
    if not text or text == "null":
        return None
    try:
        det = json.loads(text)
    except Exception:
        return None
    if not isinstance(det, dict) or not det.get("container"):
        return None
    return det


# -- 取图 -------------------------------------------------------------------
def _grab(browser: Any, url: str) -> bytes:
    """页面内 fetch 图片字节（CDN 带 CORS 头，真机实测可行）。"""
    if not url.startswith("http"):
        raise CaptchaError(f"图片地址不像 URL：{url[:80]}")
    raw = browser.eval(f"({FETCH_JS})({json.dumps(url)})", await_promise=True, timeout=60)
    try:
        d = json.loads(str(raw or "{}"))
    except Exception as exc:
        raise CaptchaError(f"图片取回结果解析失败：{exc}") from exc
    if d.get("err"):
        raise CaptchaError(f"页面内 fetch 图片失败：{d['err']}（跨域被拦时试 --dump 看细节）")
    if int(d.get("status") or 0) != 200 or int(d.get("len") or 0) <= 0:
        raise CaptchaError(f"图片响应异常：HTTP {d.get('status')}，{d.get('len')} 字节")
    data = base64.b64decode(d["b64"])
    if len(data) != int(d["len"]):
        raise CaptchaError(f"图片字节数不符：{d['len']} ≠ {len(data)}")
    return data


def grab_images(browser: Any, det: dict[str, Any]) -> tuple[bytes, bytes]:
    """（背景图, 拼图块）的字节。"""
    bg, piece = det.get("bg") or {}, det.get("piece") or {}
    if not bg.get("src") or not piece.get("src"):
        raise CaptchaError("页面上没找到背景图/拼图块的 src（验证码结构变了？跑 "
                           "devtools/live_captcha_probe.py 看看）")
    return _grab(browser, bg["src"]), _grab(browser, piece["src"])


# -- 缺口识别（纯函数，离线可测） ---------------------------------------------
def find_gap(bg_png: bytes, piece_png: bytes,
             piece_nat_y: float | None = None) -> dict[str, Any]:
    """算拼图应落在背景图的**原图坐标**位置（返回拼图整张图的 left）。

    ``piece_nat_y``：拼图当前在背景图**原图坐标系**里的 y（由 ``solve`` 用两个
    rect 换算传入）。缺口一定和拼图在同一水平带上 —— 真机踩过：暗色阈值在
    深色海面/山体背景图上会捞出一堆同尺寸暗块，没有 y 约束就会选错（选错 =
    VerifyErr，还白耗一次尝试次数）。

    返回 ``{"x": 缺口左缘-拼图透明边距, "y": 同理, "w": 拼图不透明宽,
    "method": "cv2-notch(thr=N)"|"cv2-mask"|"ddddocr", "dddd_x": 参考值}``。
    """
    import numpy as np
    import cv2

    bg = cv2.imdecode(np.frombuffer(bg_png, np.uint8), cv2.IMREAD_COLOR)
    pc = cv2.imdecode(np.frombuffer(piece_png, np.uint8), cv2.IMREAD_UNCHANGED)
    if bg is None or pc is None:
        raise CaptchaError("图片解码失败（拿到的可能不是图片）")
    if pc.shape[2] < 4:
        raise CaptchaError("拼图块没有 alpha 通道，无法定位不透明区")
    alpha = pc[:, :, 3]
    ys, xs = np.where(alpha > 10)
    if len(xs) == 0:
        raise CaptchaError("拼图块整张透明？")
    ax1, ay1 = int(xs.min()), int(ys.min())
    pw, ph = int(xs.max()) - ax1 + 1, int(ys.max()) - ay1 + 1
    gray = cv2.cvtColor(bg, cv2.COLOR_BGR2GRAY)
    bh, bw = gray.shape
    # y 带约束：真机两例里「缺口中心 vs 拼图中心」纵向差 5~35 原图像素（拼图初始
    # y 是随机的，服务端只校验 X）—— 中心 ±80px，只挡「离谱远」的暗块（深色海面/
    # 山体、底部整片暗区），不挡正常缺口。
    y_lo = y_hi = None
    if piece_nat_y is not None:
        cy_piece = piece_nat_y + ph / 2
        y_lo, y_hi = cy_piece - 80, cy_piece + 80

    def in_band(cy: float) -> bool:
        return y_lo is None or (y_lo <= cy <= y_hi)

    # --- ① cv2 暗色剪影：跨阈值收集候选 → 按「最暗」选 ---
    # 为什么不「第一个有候选的阈值就用」：真机踩过 thr=80 时把深色海面的一块
    # 当成了缺口（VerifyErr）。缺口是整图**最暗**的区域（0.3 倍压暗），跨阈值
    # 收齐再比平均灰度，稳得多。
    cands: list[tuple[float, int, int, int, int]] = []
    for thr in (120, 100, 80, 65):
        mask = (gray < thr).astype(np.uint8)
        n, _, stats, _ = cv2.connectedComponentsWithStats(mask, 8)
        for x, y, w, h, area in stats[1:]:
            if not (pw * 0.5 <= w <= pw * 1.4 and ph * 0.5 <= h <= ph * 1.4):
                continue
            if area < pw * ph * 0.25:
                continue
            if not in_band(y + h / 2):
                continue
            blob = gray[y:y + h, x:x + w][mask[y:y + h, x:x + w] > 0]
            cands.append((float(blob.mean()), int(x), int(y), int(w), int(h)))
    # 位置相近的候选（不同阈值捞到的同一块）去重
    uniq: list[tuple[float, int, int, int, int]] = []
    for c in sorted(cands, key=lambda s: s[0]):
        if all(abs(c[1] - u[1]) > 20 or abs(c[2] - u[2]) > 20 for u in uniq):
            uniq.append(c)
    if uniq:
        g, x, y, w, h = uniq[0]
        return {"x": x - ax1, "y": y - ay1, "w": pw, "h": ph,
                "method": f"cv2-notch", "dddd_x": None}

    # --- ② 掩码最小亮度：alpha 加权平均亮度最低的窗口（限制在 y 带内） ---
    # 注意 mask 要裁到**不透明区**（和窗口同尺寸），整张 alpha 是 110×110 而窗口
    # 是 96×104，直接乘会 broadcast 报错。
    mcrop = (alpha[ay1:ay1 + ph, ax1:ax1 + pw] > 10).astype(np.float32)
    mcnt = float(mcrop.sum())
    y0s = 0 if y_lo is None else max(0, int(y_lo))
    y1s = (bh - ph) if y_hi is None else min(bh - ph, int(y_hi))
    best, best_score = None, None
    step = 2
    for yy in range(max(0, y0s), max(1, y1s), step):
        for xx in range(0, max(1, bw - pw), step):
            win = gray[yy:yy + ph, xx:xx + pw]
            if win.shape[:2] != (ph, pw):
                continue
            score = float((win * mcrop).sum()) / mcnt
            if best_score is None or score < best_score:
                best_score, best = score, (xx, yy)
    if best is not None:
        return {"x": best[0] - ax1, "y": best[1] - ay1, "w": pw, "h": ph,
                "method": "cv2-mask", "dddd_x": None}

    # --- ③ ddddocr（末选）---
    dddd = _ddddocr_slide(bg_png, piece_png)
    if dddd is not None:
        return {"x": int(dddd), "y": 0, "w": pw, "h": ph,
                "method": "ddddocr", "dddd_x": int(dddd)}
    raise CaptchaError("三种方法都没定位到缺口（背景图上没有明显的深色剪影？）")


_DDDDOCR: Any = None


def _ddddocr_slide(bg_png: bytes, piece_png: bytes) -> int | None:
    """ddddocr slide_match（懒加载模型，第一次约 1~2s）。"""
    global _DDDDOCR
    try:
        import ddddocr
    except Exception:
        return None
    try:
        if _DDDDOCR is None:
            _DDDDOCR = ddddocr.DdddOcr(det=False, ocr=False, show_ad=False)
        r = _DDDDOCR.slide_match(piece_png, bg_png, simple_target=True)
        x = int((r or {}).get("target_x") or 0)
        return x if x > 0 else None
    except Exception:
        return None


# -- 距离换算（纯函数，离线可测） ---------------------------------------------
def drag_distance(det: dict[str, Any], gap: dict[str, Any]) -> float:
    """缺口位置（原图坐标）→ 拼图要移动的**显示像素**距离。

    缩放比按背景图「显示宽 / 原图宽」现场算（真机 340/552≈0.616），
    起点用拼图当前 rect 相对背景图 rect 的偏移（真机 8px），不写死。
    """
    bg, piece = det.get("bg") or {}, det.get("piece") or {}
    bgr, pcr = bg.get("rect") or {}, piece.get("rect") or {}
    nw = float(bg.get("nw") or 0)
    if not nw or not bgr.get("w") or not pcr.get("x") or not bgr.get("x"):
        raise CaptchaError("检测信息不全（bg/piece 的 rect 或 naturalWidth 缺失）")
    scale = float(bgr["w"]) / nw
    target_left = float(gap["x"]) * scale
    init_left = float(pcr["x"]) - float(bgr["x"])
    return target_left - init_left


def trajectory(dist: float, seed: int | None = None) -> list[tuple[float, float, float]]:
    """拟人轨迹：[(累计dx, dy抖动, 步间隔秒), ...]。

    先快后慢（smoothstep），末端 2~5px 过冲再回拉，全程 0.7~1.2s、25~40 步。
    """
    rng = random.Random(seed)
    n = rng.randint(25, 40)
    over = rng.uniform(2.0, 5.0)
    steps: list[tuple[float, float, float]] = []
    total = dist + over
    yphase = rng.uniform(0, 6.28)
    for i in range(1, n + 1):
        t = i / n
        eased = t * t * (3 - 2 * t)                    # smoothstep：先加速后减速
        dx = total * eased
        dy = 1.5 * (1 - t) * rng.uniform(-0.6, 0.6) + 0.8 * __import__("math").sin(yphase + t * 9)
        dt = rng.uniform(0.012, 0.030) * (1.0 + 2.2 * t)   # 越接近终点越慢
        steps.append((dx, dy, dt))
    for i in range(1, 7):                              # 回拉：over → 0
        t = i / 7
        steps.append((dist + over * (1 - t), rng.uniform(-0.4, 0.4),
                      rng.uniform(0.02, 0.05)))
    return steps


# -- 拖动 -------------------------------------------------------------------
def _mouse(browser: Any, etype: str, x: float, y: float, *, down: bool) -> None:
    browser.call("Input.dispatchMouseEvent",
                 {"type": etype, "x": round(x, 1), "y": round(y, 1),
                  "button": "left", "buttons": 1 if down else 0,
                  "clickCount": 1 if etype in ("mousePressed", "mouseReleased") else 0},
                 timeout=20)


def _piece_x(browser: Any) -> float:
    raw = browser.eval(
        "(function(){const e=document.querySelector('%s');"
        "return e?e.getBoundingClientRect().left:-1;})()" % SEL_PIECE,
        await_promise=False, timeout=15)
    try:
        v = float(raw)
    except (TypeError, ValueError):
        return -1.0
    return v


def _drag(browser: Any, det: dict[str, Any], dist: float, log: LogFn | None) -> None:
    """按住轨道按钮，把拼图拖到 dist（显示像素）。带现场校准，不假设 1:1。"""
    log = log or (lambda _m: None)
    icon = det.get("icon") or {}
    if not icon.get("w"):
        raise CaptchaError("找不到滑块按钮（.secsdk-captcha-drag-icon）")
    x0 = float(icon["x"]) + float(icon["w"]) / 2
    y0 = float(icon["y"]) + float(icon["h"]) / 2
    px_start = _piece_x(browser)
    if px_start < 0:
        raise CaptchaError("拖动开始前读不到拼图位置")

    _mouse(browser, "mousePressed", x0, y0, down=True)
    try:
        # 校准：先拖 24px，实测「按钮 → 拼图」的位移比
        cal = 24.0
        for dx in (8.0, 16.0, cal):
            _mouse(browser, "mouseMoved", x0 + dx, y0, down=True)
            time.sleep(0.03)
        ratio = (_piece_x(browser) - px_start) / cal
        if not (0.2 < ratio < 3.0):
            log(f"[验证码] 校准比值异常（{ratio:.3f}），按 1:1 处理")
            ratio = 1.0
        total = dist / ratio
        log(f"[验证码] 校准：按钮 24px → 拼图 {ratio * cal:.1f}px（比值 {ratio:.3f}），"
            f"目标按钮位移 {total:.1f}px")
        for dx, dy, dt in trajectory(total - cal):
            _mouse(browser, "mouseMoved", x0 + dx, y0 + dy, down=True)
            time.sleep(dt)
        _mouse(browser, "mouseMoved", x0 + total, y0, down=True)
        time.sleep(0.12)
    finally:
        _mouse(browser, "mouseReleased", x0 + total, y0, down=False)
    log(f"[验证码] 已松手（按钮位移 {total:.1f}px，拼图目标位移 {dist:.1f}px），等结果…")


# -- 编排 -------------------------------------------------------------------
def solve(browser: Any, log: LogFn | None = None,
          max_attempts: int = 3, dump_dir: Any = None) -> dict[str, Any]:
    """发现验证码 → 取图 → 算缺口 → 拖动 → 等验证结果。最多试 ``max_attempts`` 次。"""
    log = log or (lambda _m: None)
    for attempt in range(1, max_attempts + 1):
        det = detect(browser)
        if det is None:
            return {"ok": True, "skipped": True, "attempts": attempt - 1}
        bg, pc = det.get("bg") or {}, det.get("piece") or {}
        log(f"[验证码] 第 {attempt}/{max_attempts} 次：背景图 {bg.get('nw')}×{bg.get('nh')}"
            f"（显示 {round((bg.get('rect') or {}).get('w', 0))}px 宽），"
            f"拼图 {pc.get('nw')}×{pc.get('nh')}")
        try:
            bg_bytes, pc_bytes = grab_images(browser, det)
            # 拼图当前位置换算到原图坐标（给 find_gap 做 y 带约束）
            bgr, pcr = bg.get("rect") or {}, pc.get("rect") or {}
            nat_y = None
            if bgr.get("y") is not None and pcr.get("y") is not None and bg.get("nw"):
                scale0 = float(bgr.get("w") or 0) / float(bg["nw"])
                if scale0 > 0:
                    nat_y = (float(pcr["y"]) - float(bgr["y"])) / scale0
            gap = find_gap(bg_bytes, pc_bytes, piece_nat_y=nat_y)
            dist = drag_distance(det, gap)
        except CaptchaError as exc:
            log(f"[验证码] ✖ {exc}")
            return {"ok": False, "error": str(exc), "attempts": attempt}
        dddd = gap.get("dddd_x")
        log(f"[验证码] 缺口 left={gap['x']}px（原图坐标，{gap['method']}）"
            f"{'，ddddocr 参考 ' + str(dddd) + 'px' if dddd else ''}"
            f" → 拼图需移动 {dist:.1f}px")
        if dump_dir is not None:
            try:
                from pathlib import Path
                d = Path(dump_dir)
                d.mkdir(parents=True, exist_ok=True)
                (d / "bg.png").write_bytes(bg_bytes)
                (d / "piece.png").write_bytes(pc_bytes)
            except Exception:
                pass
        _drag(browser, det, dist, log)
        # 等验证结果：容器消失 = 通过；图片换了 = 被打回，重试
        deadline = time.time() + POLL_TIMEOUT
        while time.time() < deadline:
            time.sleep(0.8)
            now = detect(browser)
            if now is None:
                log("[验证码] ✅ 已通过（验证框消失）")
                return {"ok": True, "attempts": attempt, "dist": dist,
                        "method": gap["method"]}
            new_bg = ((now.get("bg") or {}).get("src") or "")
            if new_bg and new_bg != (bg.get("src") or ""):
                log("[验证码] 图片已刷新（上一把没过），重试…")
                break
        else:
            log("[验证码] ✖ 12s 内验证框没消失，也没换图 —— 可能没触发校验，重试…")
    return {"ok": False, "error": f"{max_attempts} 次都没通过", "attempts": max_attempts}


def guard(browser: Any, log: LogFn | None = None) -> bool:
    """给抓取循环用的闸门：没有验证码 → False（调用方照常发请求）；
    有 → 解掉后返回 True（解不掉返回 False，抓取继续但请求多半会失败）。"""
    if browser is None:
        return False
    if detect(browser) is None:
        return False
    log and log("⚠ 页面上出现了验证码，开始自动处理…")
    try:
        r = solve(browser, log=log)
        return bool(r.get("ok"))
    except Exception as exc:
        log and log(f"[验证码] 处理异常：{type(exc).__name__}: {exc}")
        return False
