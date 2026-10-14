"""离线验证「页面来源 × 写法」的择路逻辑（ziniao_bridge.BrowserTransport）。

为什么需要它：抓客户能不能通，本质只取决于**页面停在哪个区域站** —— 而这一点
在真机上要跑一遍才知道（用户那边一跑就是几分钟，还常常只拿到一段日志）。
这里用假浏览器把 2026-10-14 真机日志那条链路逐条复刻出来：

    通用站 seller.tiktokshopglobalselling.com      → code=98001008「需要权限」
    eu 站   seller.eu.tiktokshopglobalselling.com  → HTTP 200 "No matching route"
    us 站   seller.us.tiktokshopglobalselling.com  → code=0 ✅

任何平台都能跑（不碰 Qt / 不碰网络），CI 每次构建都执行。

跑法：
    python devtools/test_transport.py
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

# 这个脚本全程 print 中文。Windows 控制台默认 cp1252，不切编码就 UnicodeEncodeError
# （CI 里 workflow 已设 PYTHONUTF8/PYTHONIOENCODING，这层防本机直接跑）。
for _name in ("stdout", "stderr"):
    _s = getattr(sys, _name, None)
    if _s is not None:
        try:
            _s.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import ziniao_bridge as zb  # noqa: E402

GEN = "https://seller.tiktokshopglobalselling.com"
EU = "https://seller.eu.tiktokshopglobalselling.com"
US = "https://seller.us.tiktokshopglobalselling.com"
API = "https://api16-normal-no1a.tiktokshopglobalselling.com"
ORDER_GET = API + "/api/fulfillment/order/get"
PARAMS = {"seller_id": "7496000161034505075", "oec_seller_id": "7496000161034505075",
          "aid": 4068}

# 假浏览器是按 **host** 判定「这个站行不行」的，别拿完整 URL 去比
H_GEN, H_EU, H_US, H_API = (zb._host_of(x) for x in (GEN, EU, US, API))

# 真机上「页面自己跳站」时 Runtime.evaluate 回的就是这个（-32000）
STALE_MSG = ("Runtime.evaluate: {'code': -32000, "
             "'message': 'Inspected target navigated or closed'}")

OK = 0
FAIL = 0


def check(name: str, cond: bool, extra: str = "") -> None:
    global OK, FAIL
    if cond:
        OK += 1
        print(f"  [OK ] {name}")
    else:
        FAIL += 1
        print(f"  [FAIL] {name}  {extra}")


def title(text: str) -> None:
    print(f"\n== {text} ==")


# ---------------------------------------------------------------------------
# 假浏览器：只实现 BrowserTransport 真正用到的东西
# ---------------------------------------------------------------------------
class FakeBrowser:
    """按 host 决定「这个站能不能看这批订单」，并可模拟页面自己跳站。"""

    def __init__(self, url: str = GEN):
        self.url = url
        self.nav: list[str] = []
        self.sent: list[tuple[str, str]] = []
        self.denied: set[str] = set()      # 98001008：这个站的登录态看不到
        self.noroute: set[str] = set()     # HTTP 200 "No matching route"
        self.ok: set[str] = set()          # 只有这些站能通
        self.redirect: dict[str, str] = {}  # 打开 A 站会落在 B 站
        self.stale_once = False            # 下一次求值撞「目标已失效」（页面在跳）

    def current_url(self) -> str:
        return self.url

    def navigate(self, url: str, **kw):  # noqa: ARG002
        host = zb._host_of(url)
        target = self.redirect.get(host, host)
        self.url = url.replace(host, target, 1) if target != host else url
        self.nav.append(url)
        return self

    def reconnect(self) -> bool:
        return True

    def alive(self) -> bool:
        return True

    def run_js(self, script: str, timeout: float = 120) -> str:  # noqa: ARG002
        m = re.search(r'fetch\("([^"]+)"', script)
        url = m.group(1) if m else ""
        host = zb._host_of(url)
        self.sent.append((host, url))
        if self.stale_once:
            # 页面在我们发请求的这一刻自己跳去了区域站 —— 真机日志的原样复刻。
            # 真代码里 BrowserSession.run_js 会据 _is_ws_error 判定重连重试。
            self.stale_once = False
            self.url = US
            raise RuntimeError(STALE_MSG)
        if host in self.noroute:
            return json.dumps({"status": 200, "text": "No matching route"})
        if host in self.denied:
            return json.dumps({"status": 200, "text": json.dumps(
                {"code": 98001008, "data": None, "message": "需要权限"})})
        if host not in self.ok:
            # 跨域（api16-* 不是页面站的同源域）→ 页面只报 Failed to fetch
            raise RuntimeError("页面 JS 执行失败：TypeError: Failed to fetch")
        return json.dumps({"status": 200, "text": json.dumps(
            {"code": 0, "data": {"main_order": [{"order_id": "5770000000000"}]},
             "message": "success"})})


class _ScriptedSession(zb.BrowserSession):
    """只覆盖 current_url/eval，其余（``_wait_url`` 的判定）走**真代码**。

    故意不调 ``BrowserSession.__init__``：那会去连调试端口。
    """

    def __init__(self, urls: list[str], ready: str = "complete"):
        self._urls = list(urls)
        self._i = 0
        self._ready = ready
        self.log = lambda _m: None

    def current_url(self) -> str:
        u = self._urls[min(self._i, len(self._urls) - 1)]
        self._i += 1
        return u

    def eval(self, script: str, await_promise: bool = True,
             timeout: float = 120):  # noqa: ARG002
        return self._ready


def main() -> int:
    title("1. 「目标已失效」必须算连接层故障（否则通道会一直失败下去）")
    stale = RuntimeError(STALE_MSG)
    check("Inspected target navigated or closed → 可重连重试",
          zb._is_ws_error(stale) is True)
    check("Cannot find default execution context → 可重连重试",
          zb._is_ws_error(RuntimeError(
              "Runtime.evaluate: {'code': -32000, 'message': "
              "'Cannot find default execution context'}")) is True)
    check("跨域 Failed to fetch → 不算连接层故障（该换写法，不是重连）",
          zb._is_ws_error(RuntimeError("TypeError: Failed to fetch")) is False)
    check("业务码报错 → 不算连接层故障",
          zb._is_ws_error(RuntimeError("[同源] 返回 code=98001008")) is False)

    title("2. 候选顺序：**页面当前所在站排第一**（不再从通用站盲试起）")
    logs: list[str] = []
    br = FakeBrowser(US)
    t = zb.BrowserTransport(br, log=logs.append, origins=[GEN, EU, US, API])
    check("页面在 us → 顺序 [us, 通用, eu]，接口主机被排除",
          t._order() == [US, GEN, EU], str(t._order()))

    br2 = FakeBrowser("https://seller.uk.tiktokshopglobalselling.com")
    t2 = zb.BrowserTransport(br2, log=logs.append, origins=[GEN, EU, US])
    check("页面落在候选表之外的区域站 → 照样排第一（不硬套列表）",
          t2._order()[0] == "https://seller.uk.tiktokshopglobalselling.com",
          str(t2._order()))

    br3 = FakeBrowser("chrome-extension://abcdefghij/detection.html")
    t3 = zb.BrowserTransport(br3, log=logs.append, origins=[GEN, EU, US])
    check("紫鸟的 chrome-extension 检测页 → 不能当页面来源，回退候选表",
          t3._order() == [GEN, EU, US], str(t3._order()))

    title("3. 已经在对的站上 → **一次都不导航**（导航本身就会掉线）")
    br4 = FakeBrowser(US)
    t4 = zb.BrowserTransport(br4, log=logs.append, origins=[GEN, EU, US])
    hit = t4._ensure_origin(US)
    check("已在 us：判定为命中", hit is True)
    check("已在 us：没有发生导航", br4.nav == [], str(br4.nav))

    title("4. 切站后要记**实际落地**的站（区域站会互相重定向）")
    logs2: list[str] = []
    br5 = FakeBrowser(GEN)
    br5.redirect = {H_EU: H_US}                        # 打开 eu 会落在 us
    t5 = zb.BrowserTransport(br5, log=logs2.append, origins=[GEN, EU, US])
    hit5 = t5._ensure_origin(EU)
    check("请求 eu、落在 us → 判定为「没到 eu」", hit5 is False)
    check("但记下的是**实测**落地站 us", t5.page_origin == US, t5.page_origin)
    check("日志里写明了实际落在哪（排查 98001008 的第一现场）",
          any("实际落在" in line and "us" in line for line in logs2),
          str(logs2))

    title("5. 导航后要等**重定向跑完**，不能在链中途就认定到了")
    s = _ScriptedSession([EU, US, US])
    got = s._wait_url(EU, timeout=10.0, settle=1)
    check("eu→us 的重定向：等到站稳定后才返回 True", got is True)
    check("没有在第 1 次取值就返回（旧逻辑会在重定向中途认定「到了」）",
          s._i >= 3, f"取值次数={s._i}")

    s2 = _ScriptedSession([EU, US, EU, US, EU, US])
    check("地址一直变（停不下来）→ 如实返回 False",
          s2._wait_url(EU, timeout=2.4, settle=1) is False)

    s3 = _ScriptedSession([US, US, US, US], ready="interactive")
    check("readyState 一直不是 complete → 给足次数后放行，别把慢页面卡死",
          s3._wait_url(US, timeout=10.0, settle=1) is True)

    title("6. 端到端①：③ 导出订单之后页面已在 us（本次真机的实际状态）")
    logs3: list[str] = []
    br6 = FakeBrowser(US)
    br6.ok = {H_US}
    br6.denied = {H_GEN}
    t6 = zb.BrowserTransport(br6, log=logs3.append, origins=[GEN, EU, US, API])
    payload = t6(ORDER_GET, dict(PARAMS), {})
    check("拿到业务数据（code=0）", payload.get("code") == 0, str(payload)[:120])
    check("★ 一次请求就通（没有先跑去通用站试）", len(br6.sent) == 1,
          str(br6.sent))
    check("★ 全程没有导航（旧逻辑这里会把页面从 us 导航去通用站）",
          br6.nav == [], str(br6.nav))
    check("锁定组合 = 同源 @ us",
          t6.winner_label == "同源" and t6.winner_origin == US,
          f"{t6.winner_label} @ {t6.winner_origin}")
    logs3.clear()
    t6(ORDER_GET, dict(PARAMS), {})
    check("第 2 条订单仍然只发 1 次请求（不再白试一轮）",
          len(br6.sent) == 2 and br6.sent[-1][0] == H_US,
          str(br6.sent))

    title("7. 端到端②：真机那天（通用站 98001008 + 页面自己跳站 → us 通）")
    logs4: list[str] = []
    br7 = FakeBrowser(GEN)          # 页面停在通用站
    br7.denied = {H_GEN}              # 通用站 → 98001008
    br7.noroute = {H_EU}  # eu → No matching route
    br7.ok = {H_US}                   # 只有 us 能看这批订单
    br7.redirect = {H_EU: H_US}
    br7.stale_once = True           # 我们发第一条时就撞上页面自己跳站
    t7 = zb.BrowserTransport(br7, log=logs4.append, origins=[GEN, EU, US, API])
    payload7 = t7(ORDER_GET, dict(PARAMS), {})
    hosts = [h for h, _u in br7.sent]
    check("最终拿到数据", payload7.get("code") == 0, str(payload7)[:120])
    check("试的顺序：通用站(98001008) → api16(失败) → us(通)",
          hosts == [H_GEN, H_API, H_US],
          str(hosts))
    check("只导航了一次（去 eu，被重定向落在 us）", len(br7.nav) == 1, str(br7.nav))
    check("锁定组合 = 同源 @ us",
          t7.winner_label == "同源" and t7.winner_origin == US,
          f"{t7.winner_label} @ {t7.winner_origin}")
    check("日志里能看到「实际落在 us」", any("实际落在" in x for x in logs4),
          str(logs4))
    n_before = len(br7.sent)
    t7(ORDER_GET, dict(PARAMS), {})
    check("锁定后每条订单只发 1 次请求",
          len(br7.sent) - n_before == 1, str(br7.sent[n_before:]))

    title("8. 全都看不到（另一个登录账号）→ 只扫一遍并返回业务码")
    br8 = FakeBrowser(GEN)
    br8.denied = {H_GEN, H_EU, H_US}
    t8 = zb.BrowserTransport(br8, log=lambda _m: None, origins=[GEN, EU, US, API])
    p8 = t8(ORDER_GET, dict(PARAMS), {})
    check("如实返回业务码 98001008（不是「传输失败」）",
          str(p8.get("code")) == "98001008", str(p8)[:120])
    check("已标记 swept（后续不再无谓重试）", t8.swept is True)

    # ------------------------------------------------------------------
    # ★ 2026-10-14 用户真机定案：**站选对了也不够，页面必须停在订单页** ★
    #    站错页对 → 98001008；站对页错（停在首页/工作台）→ 照样 98001008。
    # ------------------------------------------------------------------
    title("9. 站对还不够：页面必须停在**订单页**上")
    ORDER_PATH = "/order?selected_sort=6&tab=all"

    # 9a 在 us 站、但停在首页 → 补一次导航到订单页
    br9 = FakeBrowser(US + "/")
    t9 = zb.BrowserTransport(br9, log=lambda _m: None, origins=[GEN, EU, US],
                             page_path=ORDER_PATH)
    hit9 = t9._ensure_origin(US)
    check("站对但页不对：host 层面判定为「已在目标站」", hit9 is True)
    check("★ 但确实补了一次导航，落到订单页", br9.nav == [US + ORDER_PATH],
          str(br9.nav))

    # 9b 已经在订单页 → 一次都不导航（导航本身就会掉线）
    br10 = FakeBrowser(US + ORDER_PATH)
    t10 = zb.BrowserTransport(br10, log=lambda _m: None, origins=[GEN, EU, US],
                              page_path=ORDER_PATH)
    t10._ensure_origin(US)
    check("已停在订单页 → 一次都不导航", br10.nav == [], str(br10.nav))

    # 9c 换站时直接开到「该站的订单页」，而不是只到站点根
    br11 = FakeBrowser(GEN)
    t11 = zb.BrowserTransport(br11, log=lambda _m: None, origins=[GEN, EU, US],
                              page_path=ORDER_PATH)
    t11._ensure_origin(EU)
    check("换站 → 导航目标是「该站 + 订单页」", br11.nav == [EU + ORDER_PATH],
          str(br11.nav))

    # 9d 路径判定：按首段比，子页 / 带查询串都算「在订单页」
    br12 = FakeBrowser(US + "/order/detail?id=1")
    t12 = zb.BrowserTransport(br12, log=lambda _m: None, origins=[GEN, EU, US],
                              page_path=ORDER_PATH)
    check("订单子页也算「在订单页」", t12._on_wanted_page() is True)
    br13 = FakeBrowser(US + "/home")
    t13 = zb.BrowserTransport(br13, log=lambda _m: None, origins=[GEN, EU, US],
                              page_path=ORDER_PATH)
    check("/home 不算订单页", t13._on_wanted_page() is False)
    br13b = FakeBrowser("chrome-extension://abcdefghij/detection.html")
    t13b = zb.BrowserTransport(br13b, log=lambda _m: None, origins=[GEN, EU, US],
                               page_path=ORDER_PATH)
    check("非 http 页面（紫鸟检测页）→ 不算「在订单页」",
          t13b._on_wanted_page() is False)

    # 9e 没配 page_path 时保持旧行为（只约束到站，不约束到页）
    br13c = FakeBrowser(US + "/home")
    t13c = zb.BrowserTransport(br13c, log=lambda _m: None, origins=[GEN, EU, US])
    t13c._ensure_origin(US)
    check("没配 page_path → 不因路径而多导航（向后兼容）",
          br13c.nav == [], str(br13c.nav))

    # 9f 端到端：通用站 98001008 → 一路扫站，**每次都停在订单页**，最后 us 订单页通
    br14 = FakeBrowser(GEN + "/")
    br14.denied = {H_GEN, H_API}
    br14.ok = {H_US}
    t14 = zb.BrowserTransport(br14, log=lambda _m: None,
                              origins=[GEN, EU, US, API], page_path=ORDER_PATH)
    p14 = t14(ORDER_GET, dict(PARAMS), {})
    check("端到端拿到数据（code=0）", p14.get("code") == 0, str(p14)[:110])
    check("★ 每一次换站都停在订单页上（不是站点根）",
          bool(br14.nav) and all(ORDER_PATH in u for u in br14.nav),
          str(br14.nav))
    check("锁定组合 = 同源 @ us",
          t14.winner_label == "同源" and t14.winner_origin == US,
          f"{t14.winner_label} @ {t14.winner_origin}")

    print(f"\n合计 {OK} 项通过，{FAIL} 项失败")
    return 1 if FAIL else 0


if __name__ == "__main__":
    sys.exit(main())
