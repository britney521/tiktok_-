"""验证码自动处理（captcha_solver）—— 离线回归测试（不联网、不碰 Qt）。

覆盖四块：

  1. ``find_gap``：用**程序合成的验证码图**（渐变背景 + 拼图形状挖暗 + 描边 +
     拼图 PNG），断言识别出的缺口位置与合成时的真值一致（±3 原图像素）。
     合成时可塞「远处同尺寸暗块」当干扰项，锁住 y 带约束（真机踩过：深色
     海面/山体背景会捞出一堆同尺寸暗块，选错 = VerifyErr 还白耗尝试次数）。
  2. ``drag_distance``：显示/原图坐标换算、初始偏移。
  3. ``trajectory``：单调推进、终点回到目标、步间隔为正。
  4. ``detect`` / ``guard`` / ``solve`` / ``_drag``：假浏览器（CDP 桩）走通
     「检测 → 取图 → 算缺口 → 拖动 → 通过」全链路，断言鼠标事件序列完整
     （press → 多次 move → release）。

用法：``python devtools/test_captcha.py``
"""

from __future__ import annotations

import base64
import json
import math
import sys
import tempfile
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent
sys.path.insert(0, str(PROJECT_DIR))

import numpy as np
import cv2

import captcha_solver as cs

OK = 0
FAIL = 0
TMP = Path(tempfile.mkdtemp(prefix="test_captcha_"))


def check(name: str, cond: bool, detail: str = "") -> None:
    global OK, FAIL
    if cond:
        OK += 1
        print(f"  ✓ {name}")
    else:
        FAIL += 1
        print(f"  ✗ {name}  {detail}")


# -- 合成验证码 ---------------------------------------------------------------
def puzzle_mask(w: int = 110, h: int = 110,
                ox: int = 7, oy: int = 3) -> np.ndarray:
    """拼图形状的 alpha 掩码（方块 + 上侧圆凸起 + 右侧圆凹口），uint8 0/255。"""
    m = np.zeros((h, w), np.uint8)
    cv2.rectangle(m, (ox, oy), (ox + 82, oy + 88), 255, -1)
    cv2.circle(m, (ox + 41, oy + 4), 17, 255, -1)          # 上凸
    cv2.circle(m, (ox + 84, oy + 44), 15, 0, -1)           # 右凹
    return m


def synth_captcha(seed: int, gap_x: int, gap_y: int,
                  decoy_xy: tuple[int, int] | None = None):
    """合成（背景图 BGR, 拼图 PNG bytes, 真值信息）。

    真值：拼图**整张图**应放在 (gap_x, gap_y)（原图坐标）——即不透明区左上角
    落在 (gap_x+ox, gap_y+oy)。

    背景刻意做成中亮度（105~195，模拟照片）：只有缺口（0.32 倍压暗）和干扰块
    （0.15 倍）是深色 —— 干扰块必须被 y 带约束排除，否则它会以「更暗」赢过缺口。
    """
    rng = np.random.default_rng(seed)
    bh, bw = 344, 552
    gx_ = np.linspace(105, 195, bw, dtype=np.uint8)
    gy_ = np.linspace(190, 120, bh, dtype=np.uint8)
    bgc = np.dstack([np.tile(gx_, (bh, 1)),
                     np.tile(gy_[:, None], (1, bw)),
                     np.full((bh, bw), 150, np.uint8)]).astype(np.int16)
    for _ in range(8):
        x0, y0 = rng.integers(0, bw - 120), rng.integers(0, bh - 120)
        w0, h0 = rng.integers(60, 120), rng.integers(40, 100)
        bgc[y0:y0 + h0, x0:x0 + w0] += rng.integers(-12, 12)
    bgc += rng.integers(-6, 6, (bh, bw, 3))
    bg = np.clip(bgc, 0, 255).astype(np.uint8)

    mask = puzzle_mask()
    mh, mw = mask.shape
    # 拼图 = 原图在该位置的裁剪（带 alpha）
    piece = np.zeros((mh, mw, 4), np.uint8)
    piece[:, :, :3] = bg[gap_y:gap_y + mh, gap_x:gap_x + mw]
    piece[:, :, 3] = mask
    # 缺口 = 同一形状区域压暗 + 深色描边
    region = bg[gap_y + 3:gap_y + 3 + mh - 6, gap_x + 7:gap_x + 7 + mw - 14]
    dark = (region.astype(np.float32) * 0.32).clip(0, 255).astype(np.uint8)
    ys, xs = np.where(mask[3:-3, 7:-7] > 0)
    sub = bg[gap_y + 3:gap_y + 3 + mh - 6, gap_x + 7:gap_x + 7 + mw - 14]
    sub[ys, xs] = dark[ys, xs]
    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    cv2.drawContours(bg, [c + np.array([gap_x, gap_y]) for c in contours],
                     -1, (10, 10, 10), 3)
    # 干扰暗块（远处，同尺寸档位）
    if decoy_xy:
        dx, dy = decoy_xy
        bg[dy:dy + 84, dx:dx + 76] = (bg[dy:dy + 84, dx:dx + 76] * 0.15).astype(np.uint8)
    ok, png = cv2.imencode(".png", piece)
    assert ok
    return bg, png.tobytes(), {"x": gap_x, "y": gap_y}


def main() -> int:
    print("== ① find_gap：合成图识别缺口 ==")
    truth = []
    for i, (gx, gy, decoy) in enumerate([
            (211, 120, None), (330, 150, (470, 255)), (120, 90, (60, 10)),
            (400, 180, (450, 20))]):
        bg, piece, t = synth_captcha(42 + i, gx, gy, decoy)
        ok, bg_png = cv2.imencode(".png", bg)
        # 真机 solve 传的 piece_nat_y = 拼图在背景里的当前 y（合成时就是 gap_y）
        gap = cs.find_gap(bg_png.tobytes(), piece, piece_nat_y=float(gy))
        close = abs(gap["x"] - gx) <= 4 and abs(gap["y"] - gy) <= 4
        check(f"合成#{i} 缺口=({gx},{gy}) → 识别=({gap['x']},{gap['y']})",
              close, f"method={gap['method']}")
        truth.append((bg, piece, t))

    print("== ② drag_distance：坐标换算 ==")
    det = {"bg": {"nw": 552, "nh": 344, "rect": {"x": 790, "y": 300, "w": 340, "h": 212}},
           "piece": {"nw": 110, "nh": 110, "rect": {"x": 798, "y": 373, "w": 68, "h": 68}},
           "icon": {"x": 792, "y": 517, "w": 64, "h": 40}}
    d = cs.drag_distance(det, {"x": 211, "y": 0, "w": 96, "h": 104})
    check("缺口211(原图) → 距离≈122(显示)", abs(d - 122) < 2, f"{d:.2f}")
    det2 = dict(det, piece={"nw": 110, "nh": 110,
                            "rect": {"x": 793, "y": 373, "w": 68, "h": 68}})
    d2 = cs.drag_distance(det2, {"x": 211, "y": 0, "w": 96, "h": 104})
    check("起点偏 5px → 距离同步 +5", abs((d2 - d) - 5) < 0.6, f"{d2:.2f} vs {d:.2f}")

    print("== ③ trajectory ==")
    tr = cs.trajectory(122.0, seed=7)
    check("步数 25~47", 25 <= len(tr) <= 47, str(len(tr)))
    check("步间隔全为正", all(s[2] > 0 for s in tr))
    xs = [s[0] for s in tr]
    # 推进段（除末尾 7 步回拉）必须单调不减
    adv = xs[:-7]
    check("推进段 x 单调不减", all(b >= a - 0.51 for a, b in zip(adv, adv[1:])))
    check("回拉段只向后（≤峰值）", all(v <= max(xs) + 0.6 for v in xs[-7:]))
    check("终点回到目标(±1)", abs(xs[-1] - 122.0) <= 1.0, f"{xs[-1]:.2f}")
    check("有过冲（峰值>目标）", max(xs) > 122.0, f"{max(xs):.2f}")
    total_t = sum(s[2] for s in tr)
    check("总时长 0.6~3s", 0.6 <= total_t <= 3.0, f"{total_t:.2f}s")

    # -- 假浏览器（CDP 桩） ---------------------------------------------------
    class FakeBrowser:
        """detect / fetch / piece_x / Input.dispatchMouseEvent 全桩。

        ``det_seq``：每次 DETECT_JS 依次返回的值；传 ``None`` 表示永远返回
        ``det_forever``（默认 = 一直有验证码，测「三把都没过」的重试路径）。
        """

        def __init__(self, det_seq, images, piece_xs, det_forever=None):
            self.det_forever = det_forever
            self.det_seq = list(det_seq) if det_seq is not None else None
            self.images = images                  # url → bytes
            self.piece_xs = list(piece_xs)
            self.events = []
            self.eval_n = 0

        def eval(self, script, await_promise=False, timeout=60):
            self.eval_n += 1
            if "captcha_verify_container" in script:      # DETECT_JS
                if self.det_seq is None:
                    return self.det_forever
                return self.det_seq.pop(0) if self.det_seq else "null"
            if "captcha_verify_img_slide" in script and "left" in script:
                return self.piece_xs.pop(0) if self.piece_xs else -1.0
            if "fetch(url" in script or "FETCH" in script.upper():
                import re
                m = re.search(r'\)\("([^"]+)"\)', script)
                url = m.group(1) if m else ""
                data = self.images.get(url)
                if data is None:
                    return json.dumps({"err": "no such url"})
                return json.dumps({"status": 200, "len": len(data),
                                   "b64": base64.b64encode(data).decode()})
            return "null"

        def call(self, method, params=None, timeout=None, retry=True):
            self.events.append((method, params or {}))
            return {}

    print("== ④ detect / guard（假浏览器） ==")
    fb = FakeBrowser(["null"], {}, [])
    check("没验证码 → detect None", cs.detect(fb) is None)
    fb2 = FakeBrowser([json.dumps({"container": False})], {}, [])
    check("container=False → detect None", cs.detect(fb2) is None)

    bg, piece, t = truth[0]
    scale = 340 / 552
    det_live = {
        "container": True,
        "bg": {"src": "https://x/bg.png", "nw": 552, "nh": 344,
               "rect": {"x": 790, "y": 300, "w": 340, "h": 212}},
        "piece": {"src": "https://x/piece.png", "nw": 110, "nh": 110,
                  "rect": {"x": 798, "y": 300 + t["y"] * scale,
                           "w": 68, "h": 68}},
        "icon": {"x": 792, "y": 517, "w": 64, "h": 40}}
    det_str = json.dumps(det_live)
    fb3 = FakeBrowser([det_str], {}, [])
    check("有验证码 → detect 出结构", (cs.detect(fb3) or {}).get("container") is True)

    logs = []
    fb4 = FakeBrowser(["null"], {}, [])
    check("guard：没验证码 → False（不发拖动）",
          cs.guard(fb4, log=logs.append) is False and not fb4.events)

    print("== ⑤ solve 全链路（假浏览器）：检测→取图→算缺口→拖→通过 ==")
    fb5 = FakeBrowser(
        [det_str, "null"],                       # 第一次检测到，第二次（松手后）没了
        {"https://x/bg.png": cv2.imencode(".png", bg)[1].tobytes(),
         "https://x/piece.png": piece},
        [798.0, 798.0, 798.0, 798.0])            # piece_x: 起点/校准中/松手前
    r = cs.solve(fb5, log=logs.append)
    check("solve ok=True", r.get("ok") is True, str(r))
    mouse = [p for m, p in fb5.events if m == "Input.dispatchMouseEvent"]
    kinds = [p["type"] for p in mouse]
    check("鼠标事件 press→move…→release",
          kinds[0] == "mousePressed" and kinds[-1] == "mouseReleased"
          and kinds.count("mouseMoved") > 15, f"{len(kinds)} 个事件")
    check("press 时 buttons=1", mouse[0].get("buttons") == 1)
    check("release 时 buttons=0", mouse[-1].get("buttons") == 0)
    moves = [p for p in mouse if p["type"] == "mouseMoved"]
    check("move 的 x 递增到目标", moves[-1]["x"] > moves[0]["x"], )
    x_press = 792 + 64 / 2                       # icon 中心 = mousePressed 的 x
    check("拖动总位移 ≈ 缺口换算距离",
          abs((moves[-1]["x"] - x_press) - (t["x"] * scale - 8)) < 6,
          f"{moves[-1]['x'] - x_press:.1f} vs {t['x'] * scale - 8:.1f}")

    print("== ⑥ solve：三把都没过 → 如实报失败 ==")
    old_poll = cs.POLL_TIMEOUT
    cs.POLL_TIMEOUT = 1.5                        # 缩短等待，别把测试拖到 30s
    try:
        fb6 = FakeBrowser(None,                  # 永远有验证码
                          {"https://x/bg.png": cv2.imencode(".png", bg)[1].tobytes(),
                           "https://x/piece.png": piece},
                          [798.0] * 60, det_forever=det_str)
        r6 = cs.solve(fb6, log=logs.append, max_attempts=2)
    finally:
        cs.POLL_TIMEOUT = old_poll
    check("ok=False", r6.get("ok") is False)
    check("attempts=max_attempts", r6.get("attempts") == 2, str(r6))

    print(f"\n通过 {OK} 项，失败 {FAIL} 项（临时目录 {TMP}）")
    return 1 if FAIL else 0


if __name__ == "__main__":
    raise SystemExit(main())
