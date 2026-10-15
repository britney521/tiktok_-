"""TikTok 商家后台「导出订单」自动化 —— 走页面按钮，不碰 order/list 接口。

为什么不用接口（2026-10-13 真机实测）：
  · ``order/list`` 本地直连**恒定**返回裸 ``code=10000``（aid 4068/6556/不带都试过）；
  · 页面真实用的是同源 ``GET /api/fulfillment/na/order/list``，带签名参数，
    缺签名重放直接 ``404 page not found``，没法在页面外调。
  · 但后台订单页右上角有个「导出」图标按钮，点开面板可选导出范围与格式，
    点「导出」后**异步**生成，完成后出现在「导出历史记录」里，
    点那一行的「下载」即下载一份完整的订单 Excel（全部订单 63 列，含客户字段）。

下载环节为什么不交给浏览器：
  · 实测 ziniao 内核会**忽略** CDP 的 ``Browser.setDownloadBehavior``（``allow`` /
    ``allowAndName`` 都不落盘，文件下完就"消失"，全盘都搜不到）；
  · 下载 URL 带 ``use-disposable=1``，让浏览器真下载会把它消耗掉，之后重取返回
    ``[CheckDisposableURL] out of use times``。
  → 所以用 Fetch 域把这次请求**拦在发送前**（服务器看不到、URL 未消耗），
    再在页面内 ``fetch`` 取字节（走店铺出口 IP + 登录态），自己决定存哪。

导出文件形状：
  第 1 行 = 英文表头；第 2 行 = 字段说明（"Platform unique order ID." 之类）；
  第 3 行起 = 真实数据。
"""

from __future__ import annotations

import json
import time
import urllib.parse
from pathlib import Path
from typing import Any, Callable, Iterable, Sequence

import ziniao_bridge as zb

LogFn = Callable[[str], None]

# 订单页（用户实测可用；tab=all 表示包含全部状态）
ORDER_PAGE_SUFFIX = "/order?selected_sort=6&tab=all"

# 导出入口按钮：优先用稳定的 data-log_click_for，兜底用页面结构 XPath
EXPORT_ENTRY_XPATH = '//*[@id="main"]/div[1]/div[8]/div/div/div/div[2]/div[3]/button'

# 面板里出现这段文案即认为「导出订单」面板已展开
PANEL_MARKER = "导出订单范围"

# 导出记录接口（面板打开时页面自己会调它）
RECORD_API = "export_record/get"

SCOPE_FILTERED = "filtered"   # 筛选出的订单（默认）
SCOPE_TAB = "tab"             # 当前标签页的所有订单
FORMATS = ("excel", "csv")

# 等待节奏（模块级常量，方便测试里调小）
PAGE_SETTLE = 6.0    # 订单页加载后等首屏请求跑完
PANEL_SETTLE = 0.5   # 每次点击面板后的等待

_SCOPE_PREFIX = {SCOPE_FILTERED: "筛选出的订单", SCOPE_TAB: "当前标签页的所有订单"}
_FORMAT_LABEL = {"excel": "Excel", "csv": "CSV"}


# ---------------------------------------------------------------------------
# 注入页面里的 hook：记录 /api/fulfillment/ 的响应，用来读「导出历史记录」
# ---------------------------------------------------------------------------
HOOK_JS = r"""
(() => {
  if (window.__tk_exp) { return 'already'; }
  window.__tk_exp = [];
  // 订单走 /api/fulfillment/；商品评分的导出任务在 /api/v1/review/biz_backend/export/
  // 下面（真机实测 2026-10-15：export/start → export/job → export/update）。
  // ★ 别再写成 '/api/review/' ★ —— 真实路径是 '/api/v1/review/'，多一层 v1，
  //   子串匹配不上，评分页的接口就一条都收不到（踩过）。
  const KEYS = ['export_record', '/api/fulfillment/', '/api/product/',
                '/api/review/', '/api/v1/review/', 'biz_backend/export',
                'export/job'];
  const keep = (u) => {
    const s = String(u);
    for (let i = 0; i < KEYS.length; i++) {
      if (s.indexOf(KEYS[i]) >= 0) { return true; }
    }
    return false;
  };
  const push = (url, status, text) => {
    try {
      if (window.__tk_exp.length > 400) { return; }
      window.__tk_exp.push({url: String(url), status: status,
                            text: String(text).slice(0, 60000)});
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
        p.then((resp) => {
          try {
            resp.clone().text().then((t) => push(url, resp.status, t)).catch(() => {});
          } catch (e) {}
        }).catch(() => {});
      }
    } catch (e) {}
    return p;
  };
  const xo = XMLHttpRequest.prototype.open, xs = XMLHttpRequest.prototype.send;
  XMLHttpRequest.prototype.open = function (m, u, ...r) {
    this.__tk_u = String(u);
    return xo.call(this, m, u, ...r);
  };
  XMLHttpRequest.prototype.send = function (b) {
    const self = this;
    try {
      this.addEventListener('load', () => {
        try { if (keep(self.__tk_u)) { push(self.__tk_u, self.status, self.responseText || ''); } }
        catch (e) {}
      });
    } catch (e) {}
    return xs.call(this, b);
  };
  return 'ok';
})();
"""

# ---------------------------------------------------------------------------
# 页面「店铺身份」：从页面**自己发出的** /api/fulfillment/ 请求里读 seller_id
# ---------------------------------------------------------------------------
# 为什么不从 cookie 取：统一商家环境下同一个 profile 可能同时存在多个区域店铺
# （seller.us.* / seller.eu.*）的同名 cookie，合并成 dict 会串味（实测取到**别的
# 店铺**的 seller_id → 整批 98001008 需要权限）。而页面自己发出的请求里带的
# seller_id，就是**它当前真正所属店铺**，是唯一可信来源。
PAGE_SELLER_JS = r"""
(() => {
  const out = {};
  const KEYS = ['seller_id', 'oec_seller_id', 'sellerId', 'oecSellerId'];
  const add = (v) => {
    v = String(v || '').trim();
    if (/^\d{8,}$/.test(v)) { out[v] = (out[v] || 0) + 1; }
  };
  try {
    (window.__tk_exp || []).forEach((r) => {
      try {
        const u = new URL(r.url, location.origin);
        KEYS.forEach((k) => add(u.searchParams.get(k)));
      } catch (e) {}
    });
  } catch (e) {}
  return JSON.stringify(out);
})()
"""

OPEN_PANEL_JS = r"""
(() => {
  let btn = document.querySelector('button[data-log_click_for="export_entry"]');
  if (!btn) {
    try {
      const r = document.evaluate(%s, document, null,
                                  XPathResult.FIRST_ORDERED_NODE_TYPE, null);
      btn = r.singleNodeValue;
    } catch (e) {}
  }
  if (!btn) { return 'entry-not-found'; }
  btn.scrollIntoView({block: 'center'});
  btn.click();
  return 'clicked';
})()
""" % json.dumps(EXPORT_ENTRY_XPATH, ensure_ascii=False)

PANEL_OPEN_JS = r"""
(() => {
  let open = false;
  document.querySelectorAll('div').forEach((e) => {
    if (open || !e.offsetParent) { return; }
    if ((e.innerText || '').indexOf(%s) >= 0) { open = true; }
  });
  return open ? '1' : '';
})()
""" % json.dumps(PANEL_MARKER, ensure_ascii=False)

CLICK_LABEL_JS = r"""
(() => {
  const want = %s;
  let hit = null;
  document.querySelectorAll('label').forEach((e) => {
    if (hit || !e.offsetParent) { return; }
    const t = (e.innerText || '').trim().replace(/\s+/g, ' ');
    if (t.indexOf(want) === 0) { hit = e; }
  });
  if (!hit) { return 'label-not-found'; }
  hit.scrollIntoView({block: 'center'});
  hit.click();
  return 'clicked';
})()
"""

CHECKED_JS = r"""
(() => {
  const out = [];
  document.querySelectorAll('label').forEach((e) => {
    if (e.offsetParent && String(e.className).indexOf('checked') >= 0) {
      out.push((e.innerText || '').trim().replace(/\s+/g, ' ').slice(0, 40));
    }
  });
  return JSON.stringify(out);
})()
"""

# 点提交按钮：面板里唯一的 button 且文案恰好是「导出」（标题「导出订单」不是 button）
CLICK_EXPORT_JS = r"""
(() => {
  let n = 0;
  document.querySelectorAll('button').forEach((e) => {
    if (!e.offsetParent) { return; }
    if ((e.innerText || '').trim() === '导出') { e.click(); n += 1; }
  });
  return 'clicked-' + n;
})()
"""

# 按文件名找到对应记录那一行的「下载」按钮并点击
CLICK_DOWNLOAD_JS = r"""
(() => {
  const name = %s;
  let host = null, bestLen = 1e9;
  document.querySelectorAll('div,li,tr,td,span').forEach((e) => {
    if (!e.offsetParent) { return; }
    const t = (e.innerText || '');
    if (t.indexOf(name) >= 0 && t.length < bestLen) { host = e; bestLen = t.length; }
  });
  if (!host) { return JSON.stringify({ok: false, why: 'row-not-found'}); }
  let cur = host, btn = null;
  for (let i = 0; i < 5 && cur && !btn; i++) {
    btn = Array.from(cur.querySelectorAll('button')).find(
      (b) => (b.innerText || '').trim() === '下载');
    cur = cur.parentElement;
  }
  if (!btn) { return JSON.stringify({ok: false, why: 'btn-not-found'}); }
  btn.scrollIntoView({block: 'center'});
  btn.click();
  return JSON.stringify({ok: true});
})()
"""


# ----------------------------------------------------------------------------
# 后台店铺：**直接用 URL 选店**，不模拟点击右上角切换器。
# 店铺由 URL 里的 ``shop_region`` 决定（2026-10-15 用户指定）：
#   https://seller.us.tiktokshopglobalselling.com/homepage?shop_region=US
#   ⇔ United States
# 账号下挂了多个店铺（如「秋图2」= Spain + United States）时，**导航到目标店铺的
# 后台首页**即可 —— 页面自己会把该店铺的身份写进 cookie（随后 store_context 会等它
# 稳定）→ 后面的接口就是这家店的。
#
# 为什么不点菜单（真机实测的教训）：点开要等菜单渲染、点中后要轮询等胶囊重渲染
# （6~14s），中途菜单会浮在右上角挡住「导出」按钮，还可能误点同名的山寨店。
# URL 导航一步到位，而且落错站能立刻从 URL 看出来。
# ----------------------------------------------------------------------------
# 店铺显示名 / 别名（小写）→ (区域站前缀, shop_region)
SHOP_SITES: dict[str, tuple[str, str]] = {
    "united states": ("us", "US"), "usa": ("us", "US"), "us": ("us", "US"),
    "美国": ("us", "US"),
    "united kingdom": ("uk", "GB"), "uk": ("uk", "GB"), "gb": ("uk", "GB"),
    "英国": ("uk", "GB"),
    "spain": ("eu", "ES"), "es": ("eu", "ES"), "西班牙": ("eu", "ES"),
    "ireland": ("eu", "IE"), "ie": ("eu", "IE"),
    "germany": ("eu", "DE"), "de": ("eu", "DE"),
    "france": ("eu", "FR"), "fr": ("eu", "FR"),
    "italy": ("eu", "IT"), "it": ("eu", "IT"),
    "netherlands": ("eu", "NL"), "nl": ("eu", "NL"),
    "belgium": ("eu", "BE"), "be": ("eu", "BE"),
    "portugal": ("eu", "PT"), "pt": ("eu", "PT"),
    "poland": ("eu", "PL"), "pl": ("eu", "PL"),
    "sweden": ("eu", "SE"), "se": ("eu", "SE"),
    "singapore": ("sg", "SG"), "sg": ("sg", "SG"), "新加坡": ("sg", "SG"),
    "malaysia": ("my", "MY"), "my": ("my", "MY"),
    "thailand": ("th", "TH"), "th": ("th", "TH"),
    "vietnam": ("vn", "VN"), "vn": ("vn", "VN"),
    "philippines": ("ph", "PH"), "ph": ("ph", "PH"),
    "indonesia": ("id", "ID"), "id": ("id", "ID"),
    "japan": ("jp", "JP"), "jp": ("jp", "JP"),
    "brazil": ("br", "BR"), "br": ("br", "BR"),
    "mexico": ("mx", "MX"), "mx": ("mx", "MX"),
}

# 界面上「后台店铺」下拉的候选（显示名；下拉可编辑，也能直接填区域码或后台 URL）
SHOP_CHOICES: list[str] = [
    "United States", "United Kingdom", "Spain", "Ireland", "Germany", "France",
    "Italy", "Netherlands", "Belgium", "Portugal", "Poland", "Sweden",
    "Singapore", "Malaysia", "Thailand", "Vietnam", "Philippines",
    "Indonesia", "Japan", "Brazil", "Mexico",
]


def norm_shop_name(name: str) -> str:
    """归一化店铺名：去掉副标题括号（如 ``(HAWKRAY-Vision)``）、压空白、小写。"""
    text = str(name or "")
    for lo, hi in (("(", ")"), ("（", "）")):
        while True:
            start = text.find(lo)
            end = text.find(hi, start + 1) if start >= 0 else -1
            if start < 0 or end < 0:
                break
            text = text[:start] + " " + text[end + 1:]
    return " ".join(text.split()).strip().lower()


def shop_site(name: str) -> tuple[str, str]:
    """店铺名 / 别名 / 区域码 → ``(区域站前缀, shop_region)``；认不出返回 ``("", "")``。"""
    key = norm_shop_name(name)
    if not key:
        return ("", "")
    if key in SHOP_SITES:
        return SHOP_SITES[key]
    code = key.upper()                      # 直接填区域码（"US" / "ES" / "GB"）
    for site, c in SHOP_SITES.values():
        if c == code:
            return (site, code)
    return ("", "")


def shop_region_of(url: str) -> str:
    """从后台 URL 抠 ``shop_region``（``…?shop_region=US`` → ``"US"``）；没有则空串。"""
    try:
        q = urllib.parse.parse_qs(urllib.parse.urlsplit(str(url)).query)
        return str((q.get("shop_region") or [""])[0]).strip().upper()
    except Exception:
        return ""


def shop_home_url(name: str) -> str:
    """店铺名 / 别名 / 区域码 / 完整后台 URL → 该店铺的后台首页 URL（认不出=空串）。"""
    text = str(name or "").strip()
    if not text:
        return ""
    if "://" in text:                       # 直接贴 URL 就照用（界面上允许这么填）
        return text
    site, code = shop_site(text)
    if not site:
        return ""
    return (f"https://seller.{site}.tiktokshopglobalselling.com"
            f"/homepage?shop_region={code}")


# 读切换器胶囊上的文本（= 当前店铺）——**只读**，用来把日志写成人看得懂的样子。
SHOP_PILL_JS = r"""
(() => {
  // ① 真机结构：header 右侧区里可见、有文案的 ub-content-xxx
  let cont = null;
  document.querySelectorAll(
      '#header nav [class*="ub-rightSlot-"] [class*="ub-content-"]')
    .forEach((e) => {
      if (cont || !e.offsetParent) { return; }
      if ((e.innerText || '').trim()) { cont = e; }
    });
  if (cont) { return (cont.innerText || '').trim(); }
  // ② 兜底启发式：页面顶部 160px 内、以两位大写国家码开头的可见元素
  let hit = null;
  document.querySelectorAll('body *').forEach((e) => {
    if (hit || !e.offsetParent) { return; }
    const r = e.getBoundingClientRect();
    if (r.top < 0 || r.top > 160 || r.width < 60 || r.width > 700) { return; }
    const t = (e.innerText || '').trim();
    if (/^[A-Z]{2}[\s\n]/.test(t)) { hit = e; }
  });
  return hit ? (hit.innerText || '').trim() : '';
})()
"""


def shop_pill_text(browser: Any, log: LogFn | None = None) -> str:
    """读右上角切换器上的店铺名（**只读**，失败返回空串）—— 给日志/自检看。"""
    try:
        return str(browser.eval(SHOP_PILL_JS, await_promise=False, timeout=20) or "")
    except Exception:
        return ""


def _pill_line(text: str, fallback: str = "") -> str:
    """把切换器文本压成一行：真机是「店铺名\\n(副标题)」两行，日志里换行很脏。"""
    return (" ".join(str(text or "").split())[:60]) or fallback


def ensure_shop(browser: Any, want: str, log: LogFn | None = None,
                settle: float = 3.0, wait: float = 30.0) -> dict[str, Any]:
    """确保后台当前店铺 = ``want``（如 ``"United States"``）—— **直接导航，不点击**。

    店铺由 URL 里的 ``shop_region`` 决定：
    ``https://seller.us.tiktokshopglobalselling.com/homepage?shop_region=US``。
    已经在目标店就**不动页面**（幂等）；认不出的店铺名只告警、不乱猜、不乱点。

    导航后等页面落到目标 ``shop_region``（区域站之间会重定向，真机见过「开 eu 落在
    us」），确认后把**实际落地 URL** 回填到 ``url`` —— 调用方拿它去取 cookie（按域名
    精确取，站不对就拿不到这家店的身份）和当页面来源。所以这一步必须在取 cookie 前。

    任何失败都只返回 ``ok=False`` + 日志，**不抛异常**：切换失败不该中断流水线
    （后面的身份核对与 98001008 区域站扫描还会兜底）。

    返回 ``{"ok", "current", "url", "requested", "switched", "reason"}``。
    """
    log = log or _noop
    want = str(want or "").strip()
    out: dict[str, Any] = {"ok": False, "current": "", "url": "",
                           "requested": "", "switched": False, "reason": ""}
    if not want:
        out["reason"] = "empty-want"
        return out

    url = shop_home_url(want)
    if not url:
        out["reason"] = "unknown-shop"
        log(f"[警告] 认不出后台店铺「{want}」——不知道对应哪个 shop_region，"
            "保持当前店铺不动")
        log("        可填： " + "、".join(SHOP_CHOICES) + "（也可直接填区域码或后台 URL）")
        return out
    out["requested"] = url
    code = shop_region_of(url)

    def _cur() -> str:
        try:
            return str(browser.current_url() or "")
        except Exception:
            return ""

    here = _cur()
    if code and shop_region_of(here) == code:
        out["ok"] = True
        out["url"] = here
        out["current"] = shop_pill_text(browser, log)
        log(f"后台当前店铺已是目标（shop_region={code}："
            f"{_pill_line(out['current'], want)}），不动页面")
        return out

    log(f"后台当前店铺不是目标（{here[:100] or '(读不到当前页面)'}）"
        f"→ 导航到「{want}」：{url}")
    try:
        zb.goto(browser, url, log=log)
    except Exception as exc:
        out["reason"] = f"nav-failed: {exc}"
        out["url"] = here
        log(f"[警告] 导航到「{want}」失败（保持当前店铺不动）：{exc}")
        return out

    out["switched"] = True
    time.sleep(max(0.0, settle))
    # 轮询等到「落到目标站」或「稳定停在**另一个**站」为止。
    # 真机切站要整页刷新（几秒），期间读到的还是**旧 URL** —— 所以「连续两次读到
    # 同一个 URL」这条必须再排除「那个 URL 就是导航前的页面」，否则页面稍慢就会被
    # 误判成「落错站」。
    landed = ""
    prev = ""
    deadline = time.time() + max(0.0, wait)
    while True:
        landed = _cur() or landed
        if code and shop_region_of(landed) == code:
            break
        if (landed and landed == prev and landed != here
                and shop_region_of(landed)):
            break                     # 稳在别的站了（不是目标）→ 不用再等
        prev = landed
        if time.time() >= deadline:
            break
        time.sleep(1.0)

    out["url"] = landed or url
    out["current"] = shop_pill_text(browser, log)
    shown = _pill_line(out["current"], shop_region_of(landed) or "(读不到店铺名)")
    if code and shop_region_of(landed) == code:
        out["ok"] = True
        log(f"已导航到「{want}」，当前店铺：{shown}")
        log(f"        {landed[:110]}")
    elif not code:
        # 用户直接贴了个不带 shop_region 的 URL → 没法按区域核对，落在新页面就算成功
        out["ok"] = bool(landed)
        if landed:
            log(f"已导航到（URL 里没带 shop_region，无法核对是哪家店）：{landed[:110]}")
        else:
            out["reason"] = "switch-unconfirmed"
            log("[警告] 导航后读不到当前 URL，无法确认停在哪个店铺")
    else:
        out["ok"] = False
        out["reason"] = ("switch-not-applied" if shop_region_of(landed)
                         else "switch-unconfirmed")
        log(f"[警告] 导航到「{want}」后落在 {landed[:110] or '(读不到 URL)'}"
            f"（想要 shop_region={code}，页面显示 {shown}）——"
            "后面身份核对会兜底；若后续报 98001008，先看这里落在哪家店")
    return out


def _noop(_msg: str) -> None:  # pragma: no cover
    pass


def looks_like_order_id(value: Any) -> bool:
    """订单号判定：纯数字且够长（用来跳过「字段说明」行和杂质行）。"""
    text = "" if value is None else str(value).strip()
    return len(text) >= 8 and text.isdigit()


def parse_export(path: Path | str) -> tuple[list[str], list[list[Any]]]:
    """读导出文件 → (表头, 数据行)。

    第 2 行是字段说明（"Platform unique order ID." 之类），按「首列像不像订单号」
    过滤掉。**不用 read_only 模式**：这类文件常带坏掉的 dimensions 元数据
    （A1:A1），read_only 下会只迭代出一行。
    """
    import openpyxl

    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        ws = wb[wb.sheetnames[0]]
        all_rows = [list(r) for r in ws.iter_rows(values_only=True)]
    finally:
        wb.close()
    if not all_rows:
        return [], []
    header = ["" if c is None else str(c).strip() for c in all_rows[0]]
    rows: list[list[Any]] = []
    for row in all_rows[1:]:
        if not row or all(c is None or str(c).strip() == "" for c in row):
            continue
        if not looks_like_order_id(row[0]):
            continue
        rows.append(["" if c is None else c for c in row])
    return header, rows


def write_rows(header: Iterable[str], rows: Iterable[Iterable[Any]],
               out_path: Path | str, sheet: str = "订单") -> Path:
    """把解析出来的行写成干净的 xlsx（无说明行）。"""
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet[:31] or "订单"
    ws.append(list(header))
    for row in rows:
        ws.append(list(row))
    p = Path(out_path)
    p.parent.mkdir(parents=True, exist_ok=True)
    wb.save(p)
    return p


class OrderExporter:
    """驱动店铺后台的「导出订单」按钮，把导出文件取到本地。"""

    # 要停在哪个页面（子类换成商品评分页即可复用整条链路）
    PAGE_SUFFIX = ORDER_PAGE_SUFFIX

    def __init__(self, session: Any, log: LogFn | None = None,
                 home_url: str = zb.TIKTOK_SELLER_HOME):
        self.sess = session
        self.log = log or _noop
        self.home = home_url
        self.base = ""      # 强制用这个站（不填就从当前页面推）

    # -- 页面准备 ----------------------------------------------------------
    def host(self) -> str:
        cur = self.sess.current_url()
        parts = str(cur).split("//")
        if len(parts) < 2 or not cur.startswith("http"):
            return "https://seller.tiktokshopglobalselling.com"
        return "https://" + parts[-1].split("/")[0]

    def page_suffix(self, host: str = "") -> str:
        """站点下要打开的路径（子类可按站改，比如评分页的 ``shop_region``）。"""
        return self.PAGE_SUFFIX

    def prepare(self) -> str:
        """进后台 → 注入 hook → 打开目标页；返回站点 origin。

        注意顺序：hook 必须在订单页加载**之前**注入，否则首屏请求
        （含导出历史）抓不到。
        """
        cur = str(self.sess.current_url())
        if not (cur.startswith("http") and "tiktokshopglobalselling" in cur):
            try:
                self.sess.navigate(self.home, retry=2, interval=2.0)
            except Exception as exc:
                # 紫鸟有时会给一个 chrome-extension:// 检测页，从它导航会把 target
                # 弄丢；这时直接新建一个标签页更稳。
                self.log(f"打开后台首页失败（改新建标签页）：{exc}")
                newbie = getattr(self.sess, "new_page", None)
                if newbie is None:
                    raise
                newbie(self.home)
                time.sleep(PAGE_SETTLE)
        host = self.base.rstrip("/") or self.host()
        if "tiktokshopglobalselling" not in host:
            host = "https://seller.tiktokshopglobalselling.com"
        self.sess.add_init_script(HOOK_JS)
        self.sess.navigate(host + self.page_suffix(host), retry=3, interval=2.0, wait=45)
        time.sleep(PAGE_SETTLE)
        self.log(f"目标页：{self.sess.current_url()}")
        return host

    # -- 面板 ---------------------------------------------------------------
    def panel_open(self) -> bool:
        try:
            return str(self.sess.eval(PANEL_OPEN_JS, await_promise=False,
                                      timeout=20) or "") == "1"
        except Exception:
            return False

    def _click_entry(self) -> str:
        return str(self.sess.eval(OPEN_PANEL_JS, await_promise=False,
                                  timeout=20) or "")

    def open_panel(self, tries: int = 4) -> bool:
        """确保导出面板展开（按钮是 toggle，所以先看再点）。"""
        for _ in range(max(1, tries)):
            if self.panel_open():
                return True
            try:
                self._click_entry()
            except Exception as exc:
                self.log(f"点导出入口失败：{exc}")
            for _ in range(10):
                time.sleep(PANEL_SETTLE)
                if self.panel_open():
                    return True
        return False

    def refresh_panel(self) -> bool:
        """关掉再打开面板，逼页面重新拉一次导出历史。

        面板只在**打开的那一刻**调 ``export_record/get``（之后靠页面自己轮询状态），
        所以要拿到最新列表就得让它重新开一次。
        """
        if self.panel_open():
            try:
                self._click_entry()      # 先收起来
            except Exception:
                pass
            time.sleep(PANEL_SETTLE)
        return self.open_panel()

    def _click_label(self, text: str) -> str:
        return str(self.sess.eval(CLICK_LABEL_JS % json.dumps(text, ensure_ascii=False),
                                  await_promise=False, timeout=20) or "")

    def select(self, scope: str = SCOPE_FILTERED, fmt: str = "excel") -> None:
        """选导出范围与格式（默认值与页面一致时直接跳过点击）。"""
        if not self.open_panel():
            raise zb.ZiniaoBridgeError("导出面板没打开（按钮结构可能变了）")
        fmt = (fmt or "excel").lower()
        if fmt not in FORMATS:
            raise ValueError(f"不支持的导出格式：{fmt}（可选 {FORMATS}）")
        want_scope = _SCOPE_PREFIX.get(scope, _SCOPE_PREFIX[SCOPE_FILTERED])
        for text, label in ((want_scope, f"范围={scope}"),
                            (_FORMAT_LABEL[fmt], f"格式={fmt}")):
            try:
                self._click_label(text)
            except Exception as exc:
                self.log(f"选 {label} 失败（沿用页面默认）：{exc}")
            time.sleep(1)
        try:
            checked = self.sess.read_json(CHECKED_JS, default=[]) or []
            self.log(f"当前选中：{checked}")
        except Exception:
            pass

    def click_export(self) -> None:
        got = str(self.sess.eval(CLICK_EXPORT_JS, await_promise=False,
                                 timeout=20) or "")
        self.log(f"点「导出」：{got}")
        if got in ("clicked-0", ""):
            raise zb.ZiniaoBridgeError("没找到「导出」按钮（面板可能已关闭）")

    # -- 导出历史 ----------------------------------------------------------
    def records(self, refresh: bool = True) -> list[dict[str, Any]]:
        """读导出历史记录（按时间倒序）。

        重新展开一次面板即可逼页面再调一次 ``export_record/get``。
        """
        if refresh:
            self.refresh_panel()
        blobs = self.sess.read_json(
            "JSON.stringify((window.__tk_exp||[])"
            f".filter(function(x){{return x.url.indexOf('{RECORD_API}')>=0;}})"
            ".map(function(x){return x.text;}))", default=[]) or []
        seen: set[str] = set()
        out: list[dict[str, Any]] = []
        for blob in blobs:
            try:
                obj = json.loads(blob)
            except Exception:
                continue
            for rec in ((obj.get("data") or {}).get("export_records") or []):
                key = str(rec.get("file_key") or "")
                if key and key not in seen:
                    seen.add(key)
                    out.append(rec)
        return out

    # -- 店铺身份 ----------------------------------------------------------
    def page_seller_id(self) -> dict[str, Any]:
        """页面自己请求里出现过的 seller_id（出现次数倒序）。

        返回 ``{"seller_id": 最多的那个, "counts": {id: 次数}, "has_data": bool}``。
        ``has_data`` 为 False 说明页面还没发出过带 seller_id 的请求（太早读）。
        """
        try:
            counts = self.sess.read_json(PAGE_SELLER_JS, default={}) or {}
        except Exception as exc:
            self.log(f"读页面店铺身份失败（忽略）：{exc}")
            return {"seller_id": "", "counts": {}, "has_data": False}
        if not isinstance(counts, dict) or not counts:
            return {"seller_id": "", "counts": {}, "has_data": False}
        best = max(counts.items(), key=lambda kv: kv[1])[0]
        return {"seller_id": str(best),
                "counts": {str(k): int(v) for k, v in counts.items()},
                "has_data": True}

    @staticmethod
    def record_name(rec: dict[str, Any]) -> str:
        """文件名（接口里是 URL 编码的）。"""
        raw = str(rec.get("file_name") or "")
        try:
            return urllib.parse.unquote(raw)
        except Exception:
            return raw

    def wait_record(self, known_keys: set[str], timeout: float = 300.0,
                    poll: float = 6.0, refresh: bool = True) -> dict[str, Any]:
        """等新导出出现（file_key 不在 known_keys 里）。"""
        deadline = time.time() + max(10.0, float(timeout))
        last_note = 0.0
        first = True
        while time.time() < deadline:
            if not first:
                time.sleep(max(1.0, poll))
            first = False
            try:
                recs = self.records(refresh=refresh)
            except Exception as exc:
                self.log(f"读导出历史失败（继续等）：{exc}")
                continue
            for rec in recs:
                key = str(rec.get("file_key") or "")
                if key and key not in known_keys:
                    self.log(f"新导出已就绪：{self.record_name(rec)}")
                    return rec
            if time.time() - last_note > 20:
                last_note = time.time()
                self.log(f"等待导出生成中…（历史 {len(recs)} 条，剩余 "
                         f"{int(deadline - time.time())}s）")
        raise zb.ZiniaoBridgeError("等待导出生成超时（可调大 timeout 或稍后重试）")

    # -- 下载 ---------------------------------------------------------------
    def download(self, rec: dict[str, Any], out_path: Path | str,
                 timeout: float = 25.0) -> Path:
        """点该记录的「下载」→ 拦下一次性 URL → 页面内取字节 → 落盘。"""
        name = self.record_name(rec)
        self.open_panel()
        click_js = CLICK_DOWNLOAD_JS % json.dumps(name, ensure_ascii=False)
        urls = self.sess.intercept_url("*easesafe*", click_js,
                                       timeout=timeout, marker="easesafe")
        if not urls:
            raise zb.ZiniaoBridgeError(
                f"没拦到下载请求（导出可能还在生成中，或文件名不匹配）：{name}")
        data = self.sess.fetch_bytes(urls[0])
        if data[:2] != b"PK" and not str(name).lower().endswith(".csv"):
            # 不是标准 xlsx：只要它不是一张 HTML 错误页就照原样存（CSV / 加密包等
            # 都可能不是 PK 开头，直接抛错会盖掉真正的原因）。
            if data[:200].lstrip()[:1] == b"<":
                raise zb.ZiniaoBridgeError(
                    f"下载回来的不是 Excel（前 200 字节：{data[:200]!r}）")
            self.log("⚠ 下载回来的不是标准 xlsx 头（可能是 CSV / 其它格式），照原样保存")
        p = Path(out_path)
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_bytes(data)
        self.log(f"已保存：{p}（{len(data)} 字节）")
        return p

    # -- 主流程 -------------------------------------------------------------
    def export(self, out_path: Path | str, scope: str = SCOPE_FILTERED,
               fmt: str = "excel", timeout: float = 300.0, poll: float = 6.0,
               retries: int = 3) -> dict[str, Any]:
        """完整跑一轮：打开订单页 → 导出 → 等生成 → 下载 → 解析。"""
        host = self.prepare()
        known = {str(r.get("file_key") or "") for r in self.records(refresh=True)}
        self.log(f"现有导出记录 {len(known)} 条")
        self.select(scope=scope, fmt=fmt)
        self.click_export()

        rec = self.wait_record(known, timeout=timeout, poll=poll)
        last: Exception | None = None
        for i in range(1, max(1, retries) + 1):
            try:
                raw = self.download(rec, out_path)
                break
            except Exception as exc:
                last = exc
                self.log(f"[下载 {i}/{retries}] 失败：{exc}")
                time.sleep(5)
        else:
            raise zb.ZiniaoBridgeError(f"下载导出文件失败：{last}")

        header, rows = parse_export(raw)
        self.log(f"解析完成：{len(header)} 列 × {len(rows)} 行")
        return {"raw_path": raw, "header": header, "rows": rows,
                "file_name": self.record_name(rec),
                "file_key": rec.get("file_key"), "host": host}


# ---------------------------------------------------------------------------
# 商品评分导出
# ---------------------------------------------------------------------------
# 页面：https://seller.<region>.tiktokshopglobalselling.com/product/rating?shop_region=US
#
# ★ 与订单导出**不是同一条路**（2026-10-15 真机连店铺浏览器抓包实测）★
#   订单页：点「下载」→ 浏览器真的发一个 easesafe 下载请求 → Fetch 拦下来取字节。
#   评分页：点「导出」→ **异步任务**，全程没有任何浏览器下载请求：
#       ① POST /api/v1/review/biz_backend/export/start          → {"code":0}
#       ② GET  .../review/biz_backend/export/job（页面自己轮询）→ {"status":2,"job_id":21}
#       ③ GET  .../review/biz_backend/export/job                → {"status":3,
#            "file_url":"seller-us.tiktok.com/easesafe/seller_review_export/object/
#                        196de912073b4366safe?expire=…&skipCookie=true&sign=…"}
#       ④ POST .../review/biz_backend/export/update             → 收尾
#   所以「拦 ``*easesafe*``」在评分页**永远拦不到东西**（曾经就是这么错的：等 45s
#   什么也没有）。正解是**从 ②③ 的响应里拿 file_url**，再在页面内 fetch 它。
#   真机还踩过一个：hook 的关键词写成 ``/api/review/``，而真实路径是
#   ``/api/v1/review/``（多一层 v1）→ 一条响应都收不到。
#   file_url 没有协议头，要么 ``//host/…`` 要么 ``host/…``，统一补 https。
RATING_PAGE_SUFFIX = "/product/rating?shop_region=US"

# 评分导出的异步任务接口（真机实测）
RATING_START_API = "biz_backend/export/start"
RATING_JOB_API = "biz_backend/export/job"
RATING_DONE_STATUS = 3          # job.status = 3 → 文件生成好了，响应里带 file_url

# 老的 XPath 只作兼容：真机 2026-10-15 已失效（页面结构变了）→ 靠按钮文案兜底
RATING_EXPORT_XPATHS = (
    '//*[@id="core-tabs-0-panel-0"]/div/div/div[3]/div[1]/div[3]/div[2]/button[1]',
)
RATING_EXPORT_LABELS = ("导出", "Export", "导出报表", "导出数据")
RATING_HISTORY_LABELS = ("导出历史", "下载历史", "导出记录",
                         "Export history", "Download history")

CLICK_RATING_EXPORT_JS = r"""
(() => {
  const paths = %s;
  const labels = %s;
  const clickIt = (el) => {
    if (!el) { return false; }
    try { el.scrollIntoView({block: 'center'}); } catch (e) {}
    el.click();
    return true;
  };
  let i = 0;
  for (i = 0; i < paths.length; i++) {
    try {
      const r = document.evaluate(paths[i], document, null,
                                  XPathResult.FIRST_ORDERED_NODE_TYPE, null);
      const el = r.singleNodeValue;
      if (el && el.offsetParent !== null && clickIt(el)) { return 'clicked-xpath-' + i; }
    } catch (e) {}
  }
  let n = 0;
  document.querySelectorAll('button').forEach((b) => {
    if (!b.offsetParent) { return; }
    const t = (b.innerText || '').trim();
    const aria = b.getAttribute('aria-label') || '';
    for (let j = 0; j < labels.length; j++) {
      if (t.indexOf(labels[j]) >= 0 || aria.indexOf(labels[j]) >= 0) {
        if (clickIt(b)) { n += 1; }
        break;
      }
    }
  });
  return n ? ('clicked-label-' + n) : 'not-found';
})()
""" % (json.dumps(list(RATING_EXPORT_XPATHS), ensure_ascii=False),
       json.dumps(list(RATING_EXPORT_LABELS), ensure_ascii=False))

# 评分页没有可开关的面板，只能靠点「导出历史」之类的入口让它重新拉一次列表
RATING_REFRESH_JS = r"""
(() => {
  const words = %s;
  let hit = null;
  document.querySelectorAll('button,span,div,a').forEach((e) => {
    if (hit || !e.offsetParent) { return; }
    const t = (e.innerText || '').trim();
    if (!t || t.length > 40) { return; }
    for (let i = 0; i < words.length; i++) {
      if (t.indexOf(words[i]) >= 0) { hit = e; break; }
    }
  });
  if (!hit) { return 'no-entry'; }
  try { hit.scrollIntoView({block: 'center'}); } catch (e) {}
  hit.click();
  return 'clicked';
})()
""" % json.dumps(list(RATING_HISTORY_LABELS), ensure_ascii=False)


# 评分页的导出任务状态（从页面 hook 收到的 export/job 响应里读）。
# 返回的 URL 是**页面自己调过的那条**（带 fp / signer 参数），拿它可以在页面里
# 继续轮询 —— 这些参数是会话绑定的，外面拼一条出来会被服务端拒。
RATING_JOB_JS = r"""
(() => {
  const key = %s;
  const out = [];
  (window.__tk_exp || []).forEach((r) => {
    try {
      if (String(r.url).indexOf(key) < 0) { return; }
      const obj = JSON.parse(r.text);
      out.push({url: String(r.url), code: obj.code, status: obj.status,
                job_id: obj.job_id, file_url: obj.file_url || '',
                message: obj.message || ''});
    } catch (e) {}
  });
  return JSON.stringify(out);
})()
""" % json.dumps(RATING_JOB_API, ensure_ascii=False)


def _is_note_row(row: Iterable[Any]) -> bool:
    """是不是「字段说明」行（第 2 行常见：'The unique ID of the product.' 之类）。

    订单文件可以按「首列是不是订单号」判断，评分文件没有这种锚点，所以用
    「整行大部分单元格都以句号结尾」来认 —— 说明行几乎必然满足，真实数据不会。
    """
    cells = [str(c).strip() for c in row if c is not None and str(c).strip()]
    if not cells:
        return True
    dotted = sum(1 for c in cells if c.endswith(".") and len(c) <= 120)
    return len(cells) >= 2 and dotted / len(cells) >= 0.6


# ----------------------------------------------------------------------------
# 差评（1/2/3 星）：筛选 + 反查订单
# ----------------------------------------------------------------------------
# 评分文件里**没有订单号**（用户真机确认：只有买家昵称和商品），所以要拿差评的
# 客户信息，只能拿昵称回订单表换 Order ID。平台各版本列名不同，一律按别名认。

# 「评分」列
RATING_SCORE_COLUMNS: tuple[str, ...] = (
    "Rating", "Rating Score", "Review Rating", "Overall Rating", "Rating Star",
    "Ratings", "Score", "评分", "商品评分", "星级", "评分星级",
)
# 「买家昵称」列（评分文件侧）
RATING_BUYER_COLUMNS: tuple[str, ...] = (
    "Buyer Nickname", "Buyer Username", "Nickname", "User Name", "Username",
    "买家昵称", "昵称", "用户名", "买家", "用户",
)
# 「商品」列（跟着带进差评客户表，方便一眼看出是哪个商品被差评）
RATING_PRODUCT_COLUMNS: tuple[str, ...] = (
    "Product Name", "Product", "Product Title", "商品名称", "商品", "产品名称",
    # 真机 CSV 里只有 ID（没有商品名）→ 放最后当兜底，总比空着强
    "Product ID", "product_id", "Product Id",
)
# 「订单号」列：真机 CSV 的评分文件**自带 order_id**（2026-10-15 实测），
# 有它就**不用**再拿昵称去订单表反查（那条路只在老版「只有昵称」的文件上才需要）。
RATING_ORDER_COLUMNS: tuple[str, ...] = (
    "Order ID", "Order Id", "OrderID", "order_id", "Order Number",
    "订单号", "订单 ID", "订单ID",
)
# 订单表侧的买家列（两个都收：昵称与账号名评分表里可能只给其中一个）
ORDER_BUYER_COLUMNS: tuple[str, ...] = (
    "Buyer Nickname", "Buyer Username", "买家昵称", "昵称", "买家",
)


def pick_column(header: Sequence[str], candidates: Sequence[str]) -> str:
    """在表头里找第一个候选列（归一化后相等），没有则返回空串。

    平台导出文件的列名会随后台语言/版本变（``Rating``/``评分``/``Rating Score``），
    还有人手写时多敲个空格，所以一律归一化后再比。
    """
    for cand in candidates:
        want = norm_key(cand)
        for col in header:
            if norm_key(col) == want:
                return col
    return ""


def norm_key(name: Any) -> str:
    return str(name or "").strip().lower().replace(" ", "").replace(
        "-", "").replace("_", "")


def _norm_value(value: Any) -> str:
    """值归一化：昵称大小写/首尾空格差异不算两个人（脱敏的 ***** 除外）。"""
    return str(value or "").strip().lower()


def rating_value(value: Any) -> int:
    """把评分单元格变成整数 1~5；不认识的返回 0（不当差评、也不当好评）。

    实测过：平台可能给 ``5`` / ``5.0`` / ``"5星"`` / ``"★5"`` 甚至图片占位符，
    拿不到数字一律返回 0，由调用方决定怎么处理（我们是不纳入）。
    """
    text = str(value or "").strip()
    if not text:
        return 0
    for ch in ("星", "star", "★", "☆", "分", "score"):
        text = text.replace(ch, "").replace(ch.upper(), "")
    text = text.strip()
    for part in ("-", "_", "/"):
        text = text.split(part)[0].strip()
    try:
        return int(round(float(text)))
    except (TypeError, ValueError):
        return 0


def filter_low_ratings(header: Sequence[str], rows: Sequence[Sequence[Any]],
                       max_score: int = 3,
                       ) -> tuple[str, list[list[Any]], list[int]]:
    """筛出评分 1~max_score 的行。

    返回 ``(实际评分列名, 保留的行, 这些行的评分值)``。
    认不出评分列时不筛选（返回空列名 + 原样全部 + 全 0），由调用方决定报不报错 ——
    **绝不因为认不出列名就把文件写成空的**。
    """
    col = pick_column(header, RATING_SCORE_COLUMNS)
    if not col:
        return "", [list(r) for r in rows], [0] * len(rows)
    idx = list(header).index(col)
    kept: list[list[Any]] = []
    scores: list[int] = []
    for row in rows:
        if idx >= len(row):
            continue
        score = rating_value(row[idx])
        if 1 <= score <= max_score:
            kept.append(list(row))
            scores.append(score)
    return col, kept, scores


def read_buyer_order_map(path: Path | str, sheet: str | None = None
                         ) -> dict[str, list[str]]:
    """订单表 → {买家昵称(归一化): [Order ID, ...]}。

    一个买家可能下过多单，所以值是列表。订单表里的昵称是明文
    （``Kristi`` / ``kristihernandez7582``），评分表里若被平台脱敏成 ``K*****``
    就对不上 —— 那种情况只能认下来，由调用方报告匹配率。
    """
    import tiktok_client as tc

    rows = tc.iter_sheet_rows(path, sheet)
    if not rows:
        return {}
    o_row, o_col = tc._locate_column(rows, "Order ID")
    buyer_cols: list[int] = []
    for name in ORDER_BUYER_COLUMNS:
        try:
            _, c = tc._locate_column(rows, name)
        except ValueError:
            continue
        if c not in buyer_cols:
            buyer_cols.append(c)
    if not buyer_cols:
        return {}

    out: dict[str, list[str]] = {}
    for row in rows[o_row + 1:]:
        if o_col >= len(row):
            continue
        oid = str(row[o_col] or "").strip()
        if not (oid.isdigit() and len(oid) >= 8):   # 说明行 / 空行
            continue
        for c in buyer_cols:
            if c >= len(row):
                continue
            key = _norm_value(row[c])
            if not key:
                continue
            bucket = out.setdefault(key, [])
            if oid not in bucket:
                bucket.append(oid)
    return out


def _decode_text(raw: bytes) -> str:
    """导出文件可能是 CSV：优先 utf-8-sig（平台带 BOM），再退 gbk。"""
    for enc in ("utf-8-sig", "utf-8", "gbk"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def _parse_text_table(text: str) -> tuple[list[str], list[list[Any]]]:
    """纯文本表格 → (表头, 数据行)。分隔符自动分辨（制表符 / 逗号）。"""
    import csv
    import io

    sample = text[:4096]
    delim = "\t" if sample.count("\t") > sample.count(",") else ","
    reader = csv.reader(io.StringIO(text), delimiter=delim)
    all_rows = [list(r) for r in reader]
    if not all_rows:
        return [], []
    header = [str(c).strip() for c in all_rows[0]]
    rows: list[list[Any]] = []
    for row in all_rows[1:]:
        if not row or all(not str(c).strip() for c in row):
            continue
        if _is_note_row(row):
            continue
        rows.append(["" if c is None else c for c in row])
    return header, rows


def parse_rating(path: Path | str) -> tuple[list[str], list[list[Any]]]:
    """读商品评分导出文件 → (表头, 数据行)。跳过空行与字段说明行。

    ★ 真机实测（2026-10-15）：评分页导出的其实是 **CSV**（``rating,review_text,
    review_date,order_id,product_id,sku_id``），不是订单页那种 xlsx —— 所以这里
    按内容判型：``PK`` 开头当 xlsx，其余当分隔符文本解析。以前只会 openpyxl，
    遇到 CSV 直接抛 BadZipFile。
    """
    raw = Path(path).read_bytes()
    if raw[:2] == b"PK":
        import openpyxl

        wb = openpyxl.load_workbook(path, data_only=True)
        try:
            ws = wb[wb.sheetnames[0]]
            all_rows = [list(r) for r in ws.iter_rows(values_only=True)]
        finally:
            wb.close()
        if not all_rows:
            return [], []
        header = ["" if c is None else str(c).strip() for c in all_rows[0]]
        rows: list[list[Any]] = []
        for row in all_rows[1:]:
            if not row or all(c is None or str(c).strip() == "" for c in row):
                continue
            if _is_note_row(row):
                continue
            rows.append(["" if c is None else c for c in row])
        return header, rows
    return _parse_text_table(_decode_text(raw))


# 「seller.<区域>.tiktokshopglobalselling.com」里的区域 → 评分页的 shop_region
_SHOP_REGION = {"us": "US", "uk": "GB", "eu": "EU", "sg": "SG", "id": "ID",
                "th": "TH", "vn": "VN", "ph": "PH", "my": "MY", "jp": "JP",
                "br": "BR", "mx": "MX", "es": "ES", "ie": "IE", "de": "DE",
                "fr": "FR", "it": "IT"}


def rating_page_suffix(host: str = "") -> str:
    """按站点拼评分页路径（通用站 / 认不出的站退回 US —— 真机就是在美区跑的）。"""
    region = "US"
    try:
        parts = str(host).split("//")[-1].split(".")
        if len(parts) >= 3 and parts[0].lower() == "seller":
            region = _SHOP_REGION.get(parts[1].lower(), "US")
    except Exception:
        pass
    return f"/product/rating?shop_region={region}"


class RatingExporter(OrderExporter):
    """驱动「商品评分」页的导出按钮，把导出文件取到本地。"""

    PAGE_SUFFIX = RATING_PAGE_SUFFIX

    def page_suffix(self, host: str = "") -> str:  # noqa: D102
        return rating_page_suffix(host)

    # -- 面板：评分页没有，别去点订单页那个入口 -----------------------------
    def open_panel(self, tries: int = 1) -> bool:  # noqa: D102
        return False

    def refresh_panel(self) -> bool:
        """刷新导出历史 —— 只能点页面上的「导出历史」入口。

        评分页没有可开关的面板，找不到入口时返回 False（调用方照旧读 hook 缓存，
        页面自己有时会轮询刷新，所以拿不到入口也不算致命）。
        """
        try:
            got = str(self.sess.eval(RATING_REFRESH_JS, await_promise=False,
                                     timeout=20) or "")
        except Exception as exc:
            self.log(f"刷新导出历史失败（忽略）：{exc}")
            return False
        if got == "no-entry":
            return False
        time.sleep(PANEL_SETTLE * 2)
        return True

    def click_export(self) -> None:
        got = str(self.sess.eval(CLICK_RATING_EXPORT_JS, await_promise=False,
                                 timeout=20) or "")
        self.log(f"点「导出」：{got}")
        if got == "not-found":
            raise zb.ZiniaoBridgeError(
                "没找到商品评分页的「导出」按钮（页面结构可能变了，见 README 排错表）")

    @staticmethod
    def _check_bytes(data: bytes, what: str) -> None:
        """下载内容体检：空 / HTML 错误页直接报错，其它格式放行。"""
        if not data:
            raise zb.ZiniaoBridgeError(f"{what}：下载回来是空的（0 字节）")
        if data[:200].lstrip()[:1] == b"<":
            raise zb.ZiniaoBridgeError(
                f"{what}：拿回来的是网页不是文件（前 200 字节：{data[:200]!r}）")

    # -- 导出任务（评分页的正解）--------------------------------------------
    @staticmethod
    def abs_file_url(url: str) -> str:
        """接口给的 ``file_url`` 没有协议头（``seller-us.tiktok.com/…``）→ 补 https。"""
        text = str(url or "").strip()
        if not text:
            return ""
        if text.startswith("//"):
            return "https:" + text
        if not text.startswith("http"):
            return "https://" + text.lstrip("/")
        return text

    def job_states(self) -> list[dict[str, Any]]:
        """读页面 hook 收到的 export/job 响应（按发生顺序）。"""
        try:
            got = self.sess.read_json(RATING_JOB_JS, default=[]) or []
        except Exception as exc:
            self.log(f"读评分导出任务失败（继续）：{exc}")
            return []
        return [dict(x) for x in got if isinstance(x, dict)]

    def job_file_url(self) -> tuple[str, str]:
        """已收到的响应里有没有「生成好」的：``(file_url, 页面那条 job URL)``。

        两个返回值都可能是空串：``file_url`` 空 = 还没生成好；job URL 空 = 页面还
        没调过 ``export/job``（点完按钮之后才会有）。
        """
        states = self.job_states()
        if not states:
            return ("", "")
        seen_job = str(states[-1].get("url") or "")
        for st in reversed(states):
            if (int(st.get("status") or 0) == RATING_DONE_STATUS
                    or st.get("file_url")):
                return (self.abs_file_url(str(st.get("file_url") or "")), seen_job)
        return ("", seen_job)

    def poll_job(self, job_url: str) -> dict[str, Any]:
        """用**页面自己那条** export/job URL 在页面里再拉一次（走店铺登录态）。

        URL 里带 ``fp=`` / signer 之类的会话参数，外面自己拼一条会被拒 ——
        所以直接复用页面刚发出去的那条。
        """
        js = ("fetch(%s, {credentials: 'include'}).then((r) => r.text())"
              ".then((t) => String(t)).catch((e) => 'ERR ' + String(e))"
              % json.dumps(job_url))
        text = str(self.sess.run_js(js, timeout=60) or "")
        if text.startswith("ERR "):
            raise zb.ZiniaoBridgeError(f"续轮询评分导出任务失败：{text[4:120]}")
        obj = json.loads(text)
        if not isinstance(obj, dict):
            raise zb.ZiniaoBridgeError(f"评分导出任务返回的不是 JSON：{text[:120]}")
        return obj

    def wait_file_url(self, timeout: float = 300.0, poll: float = 5.0) -> str:
        """等评分导出任务生成完 → 返回可下载的 ``file_url``（等不到返回空串）。

        先看页面自己轮询的结果；拿到 job URL 之后就**自己接着轮询**（页面可能只
        轮询一两次就撒手，等它不如自己来）。
        """
        deadline = time.time() + max(10.0, float(timeout))
        job_url = ""
        last_status: Any = None
        first = True
        while time.time() < deadline:
            if not first:
                time.sleep(max(1.0, poll))
            first = False

            url, seen_job = self.job_file_url()
            job_url = job_url or seen_job
            if url:
                return url

            states = self.job_states()
            if states and states[-1].get("status") != last_status:
                last_status = states[-1].get("status")
                self.log(f"评分导出任务 status={last_status}"
                         f"（job_id={states[-1].get('job_id')}）")
            if not job_url:
                continue                      # 页面还没发出 export/job，等它
            try:
                st = self.poll_job(job_url)
            except Exception as exc:
                self.log(f"续轮询失败（继续等页面自己的）：{exc}")
                continue
            if int(st.get("code") or 0) != 0:
                self.log(f"评分导出任务 code={st.get('code')}（{st.get('message')}）")
                continue
            if (int(st.get("status") or 0) == RATING_DONE_STATUS
                    or st.get("file_url")):
                self.log(f"评分导出已生成（status={st.get('status')}）")
                return self.abs_file_url(str(st.get("file_url") or ""))
        return ""

    @staticmethod
    def _save_and_parse(data: bytes, out: Path, host: str,
                        file_name: str = "", file_key: Any = "") -> dict[str, Any]:
        # 真机导出的是 CSV：把裸文件按真实类型落盘（别把 CSV 写成 .xlsx 误导人）
        if data[:2] != b"PK":
            csv_out = out.with_suffix(".csv")
            csv_out.write_bytes(data)
            raw_path = csv_out
        else:
            out.write_bytes(data)
            raw_path = out
        header, rows = parse_rating(raw_path)
        return {"raw_path": raw_path, "header": header, "rows": rows,
                "file_name": file_name or raw_path.name, "file_key": file_key,
                "host": host}

    # -- 取文件字节 ---------------------------------------------------------
    # file_url 在**另一个域名**（seller-us.tiktok.com），从店铺页 fetch 它是跨域：
    # 真机实测 cors 模式一律 ``TypeError: Failed to fetch``（no-cors 只能拿到
    # opaque，读不了 body）—— 这是 CORS，不是网络不通。
    # 解法：先把当前文档导航到**同一个域名**下的一个静态文件（robots.txt），
    # 这样文档 origin 就是它，再 fetch 就是**同源**、CORS 不参与，body 随便读。
    # 读完再导航回原来的页面。
    FILE_DOC_CANDIDATES = ("https://{host}/robots.txt",
                           "https://{host}/favicon.ico")

    def fetch_file_bytes(self, file_url: str, back_url: str = "",
                         timeout: float = 180.0) -> bytes:
        """跨域也能取：借用「同域静态文件」当文档，再同源 fetch 目标 URL。"""
        parts = urllib.parse.urlsplit(str(file_url))
        host = parts.netloc
        want_origin = f"{parts.scheme or 'https'}://{host}"
        if not host:
            raise zb.ZiniaoBridgeError(f"下载地址不合法：{file_url[:120]}")
        landed = ""
        for tpl in self.FILE_DOC_CANDIDATES:
            doc = tpl.format(host=host)
            try:
                self.sess.navigate(doc, retry=2, interval=2.0, wait=30)
            except Exception as exc:
                self.log(f"换文档到 {doc} 失败（试下一个）：{exc}")
                continue
            try:
                landed = str(self.sess.eval("location.origin", await_promise=False,
                                            timeout=20) or "")
            except Exception as exc:
                self.log(f"读 {doc} 的 origin 失败：{exc}")
                landed = ""
            if landed == want_origin:
                self.log(f"文档同源就位：{landed}（借 {doc} 当壳子）")
                break
            self.log(f"{doc} 落地 origin={landed or '(读不到)'}，不是 {want_origin}")
        else:
            raise zb.ZiniaoBridgeError(
                f"没法把页面换到 {want_origin} 下（跨域取字节的前提），"
                f"下载地址：{file_url[:120]}")
        try:
            return self.sess.fetch_bytes(file_url, timeout=timeout)
        finally:
            if back_url:
                try:
                    self.sess.navigate(back_url, retry=2, interval=2.0, wait=30)
                except Exception as exc:
                    self.log(f"取完文件后导航回原页面失败（不影响结果）：{exc}")

    # -- 主流程 -------------------------------------------------------------
    def export(self, out_path: Path | str, timeout: float = 300.0,
               poll: float = 6.0, retries: int = 3,
               direct_wait: float = 20.0) -> dict[str, Any]:
        """完整跑一轮：进评分页 → 点导出 → 读任务响应拿 file_url → 取字节 → 解析。

        ``direct_wait``：点完按钮后愿意等多久看它是不是「同步下载」型页面
        （真机评分页**不是**，所以默认只给 20s；订单页那种同步下载几秒就出请求）。
        主要时间花在 ``wait_file_url`` 上（异步生成），由 ``timeout`` 兜底。
        """
        host = self.prepare()
        out = Path(out_path)
        out.parent.mkdir(parents=True, exist_ok=True)
        known = {str(r.get("file_key") or "") for r in self.records(refresh=True)}
        self.log(f"现有导出记录 {len(known)} 条")

        # 路径 A：同步下载型页面（订单页那样，点了就发 easesafe 下载请求）
        urls: list[str] = []
        try:
            urls = self.sess.intercept_url("*easesafe*", CLICK_RATING_EXPORT_JS,
                                           timeout=direct_wait, marker="easesafe")
        except Exception as exc:
            self.log(f"点「导出」失败：{type(exc).__name__}: {exc}")
        if urls:
            self.log(f"拦到下载请求：{str(urls[0])[:110]}")
            data = self.sess.fetch_bytes(urls[0])
            self._check_bytes(data, "商品评分导出")
            res = self._save_and_parse(data, out, host)
            self.log(f"已保存：{out}（{len(data)} 字节）")
            self.log(f"解析完成：{len(res['header'])} 列 × {len(res['rows'])} 行")
            return res

        # 路径 B（真机实测的正解）：异步导出任务 —— 下载地址在 export/job 响应里
        self.log(f"点击后 {direct_wait}s 内没有浏览器下载请求 → 按「异步导出任务」处理，"
                 "读 export/job 的 file_url…")
        file_url = self.wait_file_url(timeout=timeout, poll=max(3.0, poll))
        if file_url:
            self.log(f"下载地址：{file_url[:130]}")
            back = ""
            try:
                back = str(self.sess.current_url() or "")
            except Exception:
                pass
            data = self.fetch_file_bytes(file_url, back_url=back)
            self._check_bytes(data, "商品评分导出")
            res = self._save_and_parse(data, out, host)
            self.log(f"已保存：{res['raw_path']}（{len(data)} 字节）")
            self.log(f"解析完成：{len(res['header'])} 列 × {len(res['rows'])} 行")
            return res

        # 路径 C：兜底 —— 万一哪天评分页也做成订单页那种「导出历史」面板
        if not self.refresh_panel():
            raise zb.ZiniaoBridgeError(
                "商品评分导出没拿到下载地址：点了「导出」但没等到 export/job 返回 "
                "file_url，页面上也没有「导出历史」入口。请看日志里的 status 变化；"
                "若一直是 status=2，说明服务端还在生成（可调大导出超时）")
        self.log("页面上有「导出历史」入口 → 按订单页那套等新记录…")
        rec = self.wait_record(known, timeout=min(180.0, timeout), poll=poll,
                              refresh=True)
        last: Exception | None = None
        for i in range(1, max(1, retries) + 1):
            try:
                raw = self.download(rec, out)          # download() 自己会落盘
                header, rows = parse_rating(raw)
                self.log(f"解析完成：{len(header)} 列 × {len(rows)} 行")
                return {"raw_path": raw, "header": header, "rows": rows,
                        "file_name": self.record_name(rec),
                        "file_key": rec.get("file_key"), "host": host}
            except Exception as exc:
                last = exc
                self.log(f"[下载 {i}/{retries}] 失败：{exc}")
                time.sleep(5)
        raise zb.ZiniaoBridgeError(f"下载商品评分导出文件失败：{last}")
