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
from typing import Any, Callable, Iterable

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
  const KEY = '/api/fulfillment/';
  const keep = (u) => String(u).indexOf(KEY) >= 0;
  const push = (url, status, text) => {
    try {
      if (window.__tk_exp.length > 200) { return; }
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

    def __init__(self, session: Any, log: LogFn | None = None,
                 home_url: str = zb.TIKTOK_SELLER_HOME):
        self.sess = session
        self.log = log or _noop
        self.home = home_url

    # -- 页面准备 ----------------------------------------------------------
    def host(self) -> str:
        cur = self.sess.current_url()
        parts = str(cur).split("//")
        if len(parts) < 2 or not cur.startswith("http"):
            return "https://seller.tiktokshopglobalselling.com"
        return "https://" + parts[-1].split("/")[0]

    def prepare(self) -> str:
        """进后台 → 注入 hook → 打开订单页；返回站点 origin。

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
        host = self.host()
        if "tiktokshopglobalselling" not in host:
            host = "https://seller.tiktokshopglobalselling.com"
        self.sess.add_init_script(HOOK_JS)
        self.sess.navigate(host + ORDER_PAGE_SUFFIX, retry=3, interval=2.0, wait=45)
        time.sleep(PAGE_SETTLE)
        self.log(f"订单页：{self.sess.current_url()}")
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
                    poll: float = 6.0) -> dict[str, Any]:
        """等新导出出现（file_key 不在 known_keys 里）。"""
        deadline = time.time() + max(10.0, float(timeout))
        last_note = 0.0
        first = True
        while time.time() < deadline:
            if not first:
                time.sleep(max(1.0, poll))
            first = False
            try:
                recs = self.records(refresh=True)
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
            raise zb.ZiniaoBridgeError(
                f"下载回来的不是 Excel（前 200 字节：{data[:200]!r}）")
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
