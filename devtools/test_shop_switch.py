"""后台店铺（ensure_shop）—— 离线回归测试（不联网、不碰 Qt）。

2026-10-15 用户定案：**不再模拟点击右上角切换器**，改成直接导航
    https://seller.<区域>.tiktokshopglobalselling.com/homepage?shop_region=<代码>
``shop_region=US`` 就代表 United States。账号下挂了多个店铺（「秋图2」=
Spain + United States）时，URL 决定当前是哪家店，一步到位；不用开菜单、不用等
渲染、不会误点，也不会让浮动菜单挡住导出按钮。

要守住的分支：
  1. 认不出的店铺名 → ok=False，**一个页面都不导航**（不乱猜区域码）
  2. 已经是目标店 → 幂等：不动页面
  3. 不是目标 → 导航到目标店的后台首页
  4. 落错站（页面把 eu 重定向到别处）→ 如实标 switch-not-applied，
     并把**实际落地 URL** 回填给调用方（后面按域名取 cookie 要用它）
  5. 导航失败 / 读不到 URL / JS 报错 → ok=False 且不抛异常（不中断流水线）
  6. 别名解析：显示名 / 大小写 / 区域码 / 中文 / 带副标题 / 直接给 URL

用法：``python devtools/test_shop_switch.py``
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import tiktok_export as te  # noqa: E402

OK = 0
FAIL = 0

US_URL = ("https://seller.us.tiktokshopglobalselling.com"
          "/homepage?shop_region=US")
ES_URL = ("https://seller.eu.tiktokshopglobalselling.com"
          "/homepage?shop_region=ES")
GENERIC = "https://seller.tiktokshopglobalselling.com"


def check(name: str, got: bool, detail: str = "") -> None:
    global OK, FAIL
    if got:
        OK += 1
        print(f"  [OK  ] {name}")
    else:
        FAIL += 1
        print(f"  [FAIL] {name}" + (f"  ← {detail}" if detail else ""))


class FakePage:
    """最小浏览器桩：记下导航去了哪，按 URL 回答「现在在哪」。"""

    def __init__(self, url: str = GENERIC, pill: str = "", redirect_to: str = "",
                 nav_error: str = "", stale: int = 0, blank_url: bool = False):
        self.url = url
        self.pill = pill
        self.redirect_to = redirect_to     # 导航过去后实际落在别的 URL
        self.nav_error = nav_error
        self.stale = stale                 # 前 N 次读 URL 还是旧值（页面还没跳完）
        self.blank_url = blank_url         # 永远读不到 URL
        self.navigated: list[str] = []
        self._pending = ""
        self._reads = 0

    # -- BrowserSession 的对外接口 -----------------------------------------
    def navigate(self, url, retry=5, interval=2.0, wait=60.0):  # noqa: ARG002
        self.navigated.append(url)
        if self.nav_error:
            raise RuntimeError(self.nav_error)
        self._pending = self.redirect_to or url
        self._reads = 0
        return self

    def current_url(self) -> str:
        if self.blank_url:
            return ""
        if self._pending:
            self._reads += 1
            if self._reads <= self.stale:
                return self.url
            self.url, self._pending = self._pending, ""
        return self.url

    def eval(self, script, await_promise=True, timeout=120):  # noqa: ARG002
        return self.pill


class BoomPage(FakePage):
    """JS 通道全炸（CDP 断了）的桩。"""

    def eval(self, script, await_promise=True, timeout=120):  # noqa: ARG002
        raise RuntimeError("CDP 断了")


def main() -> int:
    print("1) 只认 URL：从店铺名解析出后台首页 URL")
    check("United States → us + US", te.shop_home_url("United States") == US_URL,
          te.shop_home_url("United States"))
    check("Spain → eu + ES", te.shop_home_url("Spain") == ES_URL,
          te.shop_home_url("Spain"))
    check("大小写不敏感", te.shop_home_url("united states") == US_URL)
    check("直接填区域码", te.shop_home_url("US") == US_URL, te.shop_home_url("US"))
    check("中文名", te.shop_home_url("美国") == US_URL, te.shop_home_url("美国"))
    check("带副标题的显示名",
          te.shop_home_url("United States (HAWKRAY-Vision)") == US_URL,
          te.shop_home_url("United States (HAWKRAY-Vision)"))
    check("直接给 URL 就照用", te.shop_home_url(US_URL) == US_URL)
    check("认不出的名字 → 空串",
          te.shop_home_url("Nowhere Land") == "", te.shop_home_url("Nowhere Land"))
    check("空 → 空串", te.shop_home_url("") == "")
    check("shop_region_of 抠参数",
          te.shop_region_of(US_URL + "&x=1") == "US", te.shop_region_of(US_URL))
    check("shop_region_of 没有则空", te.shop_region_of(GENERIC) == "")

    print("2) 认不出的店铺名 → 只告警，一个页面都不导航")
    p = FakePage(url=ES_URL)
    r = te.ensure_shop(p, "Nowhere Land", settle=0)
    check("ok=False", r["ok"] is False, str(r))
    check("reason=unknown-shop", r["reason"] == "unknown-shop", r["reason"])
    check("没导航", p.navigated == [], str(p.navigated))

    print("3) 空目标 → 直接返回")
    r = te.ensure_shop(FakePage(), "", settle=0)
    check("reason=empty-want", r["reason"] == "empty-want", r["reason"])

    print("4) 已经在目标店铺 → 幂等，不动页面")
    p = FakePage(url=US_URL, pill="US\nUnited States (HAWKRAY-Vision)")
    r = te.ensure_shop(p, "United States", settle=0)
    check("ok", r["ok"] is True, str(r))
    check("没切换", r["switched"] is False, str(r))
    check("没导航", p.navigated == [], str(p.navigated))
    check("带回当前店铺名", "United States" in r["current"], r["current"])
    check("url 是当前页面", r["url"] == US_URL, r["url"])

    print("5) 停在 Spain → 导航到 United States 的后台首页")
    p = FakePage(url=ES_URL, pill="ES\nSpain")
    r = te.ensure_shop(p, "United States", settle=0)
    check("ok", r["ok"] is True, str(r))
    check("发生了切换", r["switched"] is True, str(r))
    check("导航到 US 首页", p.navigated == [US_URL], str(p.navigated))
    check("url = 落地 URL", r["url"] == US_URL, r["url"])

    print("6) 反过来：United States → Spain")
    p = FakePage(url=US_URL)
    r = te.ensure_shop(p, "Spain", settle=0)
    check("导航到 ES 首页", p.navigated == [ES_URL], str(p.navigated))
    check("ok", r["ok"] is True, str(r))

    print("7) 页面还没跳完 → 轮询等到目标站为止")
    p = FakePage(url=ES_URL, stale=2)
    r = te.ensure_shop(p, "United States", settle=0, wait=8)
    check("ok（等到了）", r["ok"] is True, str(r))
    check("url 是等到的落地站", r["url"] == US_URL, r["url"])

    print("8) 落错站 → 如实标未生效，并回填实际落地 URL")
    p = FakePage(url=GENERIC, redirect_to=ES_URL)   # 要 US，页面落在 ES
    r = te.ensure_shop(p, "United States", settle=0, wait=4)
    check("ok=False", r["ok"] is False, str(r))
    check("reason=switch-not-applied",
          r["reason"] == "switch-not-applied", r["reason"])
    check("url 是**实际**落地站（不是我们请求的那个）", r["url"] == ES_URL, r["url"])
    check("switched=True（确实导航过）", r["switched"] is True, str(r))

    print("9) 一直读不到 URL → 不硬撑，标未确认")
    p = FakePage(url=GENERIC, blank_url=True)
    r = te.ensure_shop(p, "United States", settle=0, wait=2)
    check("ok=False", r["ok"] is False, str(r))
    check("reason=switch-unconfirmed",
          r["reason"] == "switch-unconfirmed", r["reason"])

    print("10) 导航失败 → 不抛异常，保留原页面")
    p = FakePage(url=GENERIC, nav_error="Navigation timeout")
    r = te.ensure_shop(p, "United States", settle=0)
    check("ok=False", r["ok"] is False, str(r))
    check("reason 里有 nav-failed", "nav-failed" in r["reason"], r["reason"])
    check("url 还是原来的页面", r["url"] == GENERIC, r["url"])

    print("11) 读胶囊的 JS 报错 → 不影响判定（只是日志少一行）")
    p = BoomPage(url=ES_URL)
    r = te.ensure_shop(p, "United States", settle=0)
    check("ok（URL 对了就算成功）", r["ok"] is True, str(r))
    check("current 空但不崩", r["current"] == "", r["current"])

    print("12) 导航前先落一步：不在目标就等页面稳定（不会读一次旧 URL 就下结论）")
    p = FakePage(url=GENERIC, stale=1)
    t0 = time.time()
    r = te.ensure_shop(p, "Spain", settle=0, wait=6)
    check("ok", r["ok"] is True, str(r))
    check("耗时在轮询区间内（>=1s，没把 wait 跑满）",
          1.0 <= (time.time() - t0) < 6.0, f"{time.time() - t0:.2f}s")

    print("13) 直接贴不带 shop_region 的 URL → 没法核对，落到了就算成功")
    p = FakePage(url="https://seller.tiktokshopglobalselling.com")
    r = te.ensure_shop(p, "https://seller.tiktokshopglobalselling.com/homepage",
                       settle=0, wait=2)
    check("导航了", p.navigated == ["https://seller.tiktokshopglobalselling.com/homepage"],
          str(p.navigated))
    check("ok（不误报未确认）", r["ok"] is True, str(r))

    print(f"\n通过 {OK} 项，失败 {FAIL} 项")
    return 1 if FAIL else 0


if __name__ == "__main__":
    raise SystemExit(main())
