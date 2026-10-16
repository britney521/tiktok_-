"""TikTok 订单 / 客户信息导出 GUI（紫鸟店铺浏览器版）。

为什么不用 list.py 那种本地直连：
    TikTok 的 order/list 接口带签名头（X-Bogus / X-Gnarly / X-Tts-Oec-Bsid），
    并且会话与**店铺代理出口 IP** 绑定。本机 requests 直连只会拿到裸
    ``code=10000``（限流 / 风控），detail 接口也可能时通时不通。
    所以订单必须在「店铺自己的浏览器里」发起 —— 由紫鸟打开店铺，复用它的登录态
    和出口 IP，在页面内 fetch 接口，再把结果解析成 Excel。

四步流水线：
    ① 启动环境   —— 调紫鸟客户端（--run_type=web_driver --ipc_type=http）
    ② 获取店铺列表
    ③ 选择店铺 并 导出订单  → 订单.xlsx（order/list，分页拉全量）
    ④ 读订单 Excel 的 Order ID 列 → 抓客户信息 → 客户.xlsx（order/get）

复用（不重复造轮子）：
    · ``ziniao_bridge.py`` —— 找到并加载紫鸟项目的协议客户端、连店铺浏览器、
      在页面里 fetch（BrowserTransport）
    · ``tiktok_client.py`` —— 接口参数、分页、响应解析、Excel 读写

运行:
    python customer_gui.py
离屏自检（不联网，仅构建界面）:
    python customer_gui.py --check
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import threading
import time
import traceback
from pathlib import Path
from typing import Any, Callable, Sequence

from PyQt6.QtCore import QObject, Qt, QThread, pyqtSignal
from PyQt6.QtGui import QAction, QFont
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QDockWidget,
    QDoubleSpinBox,
    QFileDialog,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QSpinBox,
    QSplitter,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

import tiktok_client as tc
import tiktok_export as te
import ziniao_bridge as zb

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_ORDER_OUT = BASE_DIR / "订单.xlsx"
DEFAULT_CUSTOMER_OUT = BASE_DIR / "客户.xlsx"
DEFAULT_RATING_OUT = BASE_DIR / "商品评分.xlsx"
DEFAULT_BAD_CUSTOMER_OUT = BASE_DIR / "客户差评信息.xlsx"
DEFAULT_INPUT = BASE_DIR / "全部 笔订单-2026-10-13-06_43.xlsx"


def reveal_path(path: Any) -> None:
    """用系统默认的方式打开文件 / 文件夹（跨平台）。

    原来是硬编码 ``subprocess.Popen(["open", p])`` —— 那是 **macOS 专属**命令，
    在 Windows 上直接抛 ``FileNotFoundError``（找不到 open 这个程序），
    而「打开输出目录」「打开日志文件」恰恰是 Windows 用户最常用的两个按钮。
    """
    p = str(path)
    if not p:
        return
    try:
        if sys.platform.startswith("win"):
            os.startfile(p)  # type: ignore[attr-defined]  # 仅 Windows 有
        elif sys.platform == "darwin":
            subprocess.Popen(["open", p])
        else:
            subprocess.Popen(["xdg-open", p])
    except Exception:
        pass


def _safe_filename(name: Any) -> str:
    """把店铺名变成安全的文件名片段（去掉路径分隔符等非法字符）。"""
    text = str(name or "").strip()
    for ch in '/\\:*?"<>|\n\r\t':
        text = text.replace(ch, "_")
    return text.strip(" .")[:60]


def store_output_paths(store: dict[str, Any] | None
                       ) -> tuple[Path, Path, Path, Path]:
    """按店铺名生成默认输出路径：``<店铺名>_订单/客户/商品评分/客户差评信息.xlsx``。

    以前所有店铺都写同一个「订单.xlsx」「客户.xlsx」，导完一个店再导另一个就
    互相覆盖了。带上店铺名后一个店一份文件，也方便按店名找。
    店铺名取不到（未选店/字段缺失）时退回原名。
    """
    name = _safe_filename((store or {}).get("browserName") or "")
    if not name:
        return (DEFAULT_ORDER_OUT, DEFAULT_CUSTOMER_OUT, DEFAULT_RATING_OUT,
                DEFAULT_BAD_CUSTOMER_OUT)
    return (
        DEFAULT_ORDER_OUT.with_name(f"{name}_订单{DEFAULT_ORDER_OUT.suffix}"),
        DEFAULT_CUSTOMER_OUT.with_name(f"{name}_客户{DEFAULT_CUSTOMER_OUT.suffix}"),
        DEFAULT_RATING_OUT.with_name(f"{name}_商品评分{DEFAULT_RATING_OUT.suffix}"),
        DEFAULT_BAD_CUSTOMER_OUT.with_name(
            f"{name}_客户差评信息{DEFAULT_BAD_CUSTOMER_OUT.suffix}"),
    )

# 订单表格最多渲染多少行（导出文件可能几千行 × 63 列，全塞进 QTableWidget 很卡）
MAX_TABLE_ROWS = 800

# ③ 导出 / ④ 抓客户的固定参数 —— 界面上不再暴露这些框，一律走这里的默认值：
#   导出范围 = 页面当前筛选结果；格式 = Excel；传输 = 先浏览器内、失败再本地直连
DEFAULT_EXPORT_SCOPE = te.SCOPE_FILTERED
DEFAULT_EXPORT_FORMAT = "excel"
DEFAULT_TRANSPORT_MODE = "auto"

# 抓客户的最小请求间隔（秒）。太密容易被平台风控，界面上不允许调到比这更小。
MIN_CUSTOMER_INTERVAL = 1.5

# 后台店铺默认目标（账号下挂了多个店铺时用；按显示名/区域码解析成 shop_region 导航）
DEFAULT_TARGET_SHOP = "United States"

# ⑥ 差评客户表的列 = 客户列 + 差评自带的两列（评分、商品）
BAD_CUSTOMER_COLUMNS: list[str] = list(tc.CUSTOMER_COLUMNS) + ["评分", "商品"]

# 店铺列表优先展示的字段（名称取自紫鸟 v6 getBrowserList 的真实返回）
PREFERRED_FIELDS: list[tuple[str, str]] = [
    ("browserName", "店铺名称"),
    ("browserOauth", "店铺 ID"),
    ("siteId", "siteId"),
    ("platform_name", "平台"),
    ("store_username", "店铺账号"),
    ("browserIp", "IP"),
    ("proxyType", "代理类型"),
    ("isExpired", "已过期"),
]


def _cell(value: Any) -> str:
    """把任意字段值渲染成表格单元格文本。"""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "是" if value else "否"
    if isinstance(value, (list, dict)):
        return json.dumps(value, ensure_ascii=False)
    return str(value)


def dir_of(path: Any) -> Path:
    """取一个路径**所在的目录**（它本身就是目录就用它自己），不可用则退回项目目录。

    「浏览…」只让用户选目录，所以起始位置得从「框里那个文件的全路径」推出来；
    框里可能是空的、不存在的、或者只剩个文件名 —— 都不能让对话框崩，一律兜底。
    """
    text = str(path or "").strip()
    if not text:
        return BASE_DIR
    p = Path(text)
    d = p if p.is_dir() else p.parent
    return d if d.is_dir() else (d.parent if d.parent.is_dir() else BASE_DIR)


# ----------------------------------------------------------------------------
# 评分文件（⑤ 的产物）—— ⑥ 的依赖判断要用（MainWindow 的检查 + PipelineThread 的读取）
# ----------------------------------------------------------------------------

def rating_file_candidates(rating_out: Path) -> list[Path]:
    """评分文件的候选路径（按优先级，去重）：筛过的 out → 完整版 ``_评分原始``。

    真机⑤导出的完整版其实是 **CSV**（平台给的就是 CSV），所以 ``.csv`` 和
    ``.xlsx`` 两种后缀都认 —— 只看 ``.xlsx`` 会在「只有 CSV 完整版」时误报
    「找不到评分表」。
    """
    out = Path(rating_out)
    cands = [out,
             out.with_name(f"_评分原始{out.suffix or '.xlsx'}"),
             out.with_name("_评分原始.csv"),
             out.with_name("_评分原始.xlsx")]
    seen: set[str] = set()
    return [p for p in cands if not (str(p) in seen or seen.add(str(p)))]


def rating_order_column(rating_out: Path) -> str:
    """读评分文件表头，返回认出的「订单号」列名（认不出/读不到返回空串）。

    ⑥ 的依赖判断要用它：评分文件**自带 order_id** 时压根不需要订单表，那时不该
    再拿「找不到订单表」拦人（真机⑤导出的 CSV 就有 order_id）。认不出才退回
    「昵称 → 订单表」那条老路。

    只看**第一个存在的文件**：那正是 ⑥ 实际会读的那一份，多看了反而误导。
    """
    for p in rating_file_candidates(rating_out):
        if not p.exists():
            continue
        try:
            header, rows = te.parse_rating(p)
        except Exception:
            continue
        if rows:
            return te.pick_column(header, te.RATING_ORDER_COLUMNS) or ""
    return ""


# ----------------------------------------------------------------------------
# 日志总线
# ----------------------------------------------------------------------------
class LogBus(QObject):
    """把「任意线程的日志请求」转成 Qt 信号，自动排队回 GUI 线程。

    为什么需要它：紫鸟客户端的 ``log=`` 拿到的是一个**裸回调**，工作线程里
    ``kill_process()`` / ``start_browser()`` / ``update_core()`` 会**直接同步调用**它，
    绕过信号槽的队列机制。而 QPlainTextEdit 不是线程安全的，在工作线程里写它会
    直接 SIGSEGV（实测崩在 QTextEngine::shapeTextWithHarfbuzzNG）。
    """

    message = pyqtSignal(str)


# ----------------------------------------------------------------------------
# 后台线程
# ----------------------------------------------------------------------------
class EnvThread(QThread):
    """启动环境：强制关掉已开启的客户端 → 启动一个新的 web_driver 实例 → 更新内核。"""

    log = pyqtSignal(str)
    done = pyqtSignal(bool, str)

    def __init__(self, cfg: Any, root: str | Path | None, kill_first: bool, parent=None):
        super().__init__(parent)
        self.cfg = cfg
        self.root = root
        self.kill_first = kill_first

    def run(self) -> None:  # noqa: D102
        try:
            client = zb.make_client(self.cfg, log=self.log.emit, root=self.root)
        except Exception as exc:
            self.log.emit(traceback.format_exc(limit=5))
            self.done.emit(False, f"加载紫鸟项目失败：{exc}")
            return
        try:
            # ★ 先定「要不要重新拉起一个客户端」★
            reused = False
            if self.kill_first:
                # 强制重启（默认）：**不管现在有没有在跑，一律先关干净**。
                # 紫鸟客户端进程唯一 —— 用户手动开的那个实例还在时，新起的
                # web_driver 实例会被它接管/直接退出，`--run_type` / `--port`
                # 没生效，startBrowser 返回的 debuggingPort 是旧实例的，
                # CDP 连上去就会连错目标。所以先把所有紫鸟进程清掉再启。
                zb.kill_all_clients(client, log=self.log.emit)
            else:
                # 没勾：能复用就复用（少一次无谓重启），只有端口没开才启新实例。
                try:
                    reused = bool(client.port_open())
                except Exception:
                    reused = False
                if reused:
                    self.log.emit(
                        f"检测到紫鸟客户端已在运行（IPC 端口 "
                        f"{getattr(client, 'socket_port', '') or '(未知)'} 已就绪），"
                        "直接复用，不再重启。")

            if not reused:
                self.log.emit("启动客户端（web_driver + http IPC）…")
                if not client.start_browser():
                    if not client.config.client_executable():
                        # 「找不到程序」和「程序起来了但端口没开」是两回事，别混成一句
                        # 「端口未就绪」——那会把真正的配置问题盖掉。
                        self.done.emit(
                            False,
                            "启动环境失败：找不到紫鸟客户端程序，请在「编辑配置…」里"
                            "把「客户端路径」填成 ziniao.exe 的位置（详见上方日志）。",
                        )
                    else:
                        self.done.emit(
                            False,
                            "启动环境失败：客户端没进入 web_driver 模式（IPC 端口未就绪）。"
                            "请检查 --user-data-dir / 端口占用 / 启动日志（紫鸟项目 "
                            "logs/ziniao_launch.log）",
                        )
                    return
            self.log.emit("更新内核（需要时）…")
            if not client.update_core(on_progress=self.log.emit):
                self.done.emit(False, "更新内核失败，请升级紫鸟客户端")
                return
            self.done.emit(True, "环境就绪")
        except Exception as exc:
            self.log.emit(traceback.format_exc(limit=5))
            self.done.emit(False, f"启动环境异常：{type(exc).__name__}: {exc}")


class ListThread(QThread):
    """拉取店铺列表。"""

    log = pyqtSignal(str)
    done = pyqtSignal(bool, object, str)

    def __init__(self, cfg: Any, root: str | Path | None, parent=None):
        super().__init__(parent)
        self.cfg = cfg
        self.root = root

    def run(self) -> None:  # noqa: D102
        try:
            client = zb.make_client(self.cfg, log=self.log.emit, root=self.root)
            self.log.emit("正在获取店铺列表…")
            stores = client.get_browser_list()
            self.done.emit(True, stores, f"共 {len(stores)} 个店铺")
        except Exception as exc:
            self.log.emit(traceback.format_exc(limit=5))
            self.done.emit(False, [], f"获取店铺列表失败：{type(exc).__name__}: {exc}")


class PipelineThread(QThread):
    """打开店铺 → 导出订单 → 读订单 Excel 抓客户信息。"""

    log = pyqtSignal(str)
    stage = pyqtSignal(str)                    # 当前阶段名
    order_progress = pyqtSignal(int, int)      # 已拉取订单, 总数(0=未知)
    customer_progress = pyqtSignal(int, int)   # 已完成, 总数
    rating_progress = pyqtSignal(int, int)     # 已导出, 总数
    bad_progress = pyqtSignal(int, int)        # 差评客户已完成, 总数
    orders_ready = pyqtSignal(object)          # list[dict] 订单行
    customer_row = pyqtSignal(int, object)     # 行号, 客户信息
    rating_ready = pyqtSignal(object)          # list[dict] 商品评分行
    bad_ready = pyqtSignal(object)             # list[dict] 客户差评行
    done = pyqtSignal(dict)                    # 汇总
    failed = pyqtSignal(str)

    def __init__(self, cfg: Any, root: str | Path | None, store: dict[str, Any],
                 opts: dict[str, Any], parent=None):
        super().__init__(parent)
        self.cfg = cfg
        self.root = root
        self.store = dict(store)
        self.opts = dict(opts)
        self.stop_event = threading.Event()
        # ③ 导出订单成功时页面所在的站（= 这家店自己的区域站）。
        # 只用于排查与日志：④ 抓客户时通道会先问浏览器「现在在哪」，
        # 那条路等同于此，但万一页面漂了，日志里有这个对照值。
        self.export_origin = ""
        # ⑤→⑥ 之间的交接：评分表头 / 用到的列名 / 筛出的 1~3 星行
        self._rating_header: list[str] = []
        self._rating_cols: dict[str, str] = {"score": "", "buyer": "", "product": ""}
        self._rating_bad: list[dict[str, Any]] = []
        self._rating_all: list[dict[str, Any]] = []
        self._bad_hit = False
        # ④ 抓过的客户行缓存 {Order ID: 行} —— ⑥ 直接复用，同一单不重复请求
        self._cust_cache: dict[str, dict[str, str]] = {}
        # ④ 建好的 SKU 兜底表（⑥ 直接用，不重读订单表）：
        #   平台 SKU ID →「SKU ID」列；商家自定义 SKU →「商家 SKU」列
        self._sku_map_cache: dict[str, str] = {}
        self._seller_sku_cache: dict[str, str] = {}

    def stop(self) -> None:
        self.stop_event.set()

    # -- 主流程 -------------------------------------------------------------
    def run(self) -> None:  # noqa: D102
        o = self.opts
        summary: dict[str, Any] = {
            "order_count": 0, "order_out": "", "customer_count": 0,
            "customer_out": "", "rating_count": 0, "rating_out": "",
            "bad_count": 0, "bad_out": "",
            "errors": 0, "stopped": False,
        }
        client = None
        browser = None
        store_key = ""
        try:
            client = zb.make_client(self.cfg, log=self.log.emit, root=self.root)

            # ---- ③ 打开店铺 ----
            self.stage.emit("打开店铺")
            label = zb.store_label(self.store)
            self.log.emit(f"===== 打开店铺：{label} =====")
            self.log.emit("打开店铺…")
            ret = client.open_store(self.store)
            store_key = (ret.get("browserOauth") or ret.get("browserId")
                         or zb.browser_key(self.store))
            port = ret.get("debuggingPort")
            if not port:
                raise zb.ZiniaoBridgeError("startBrowser 未返回 debuggingPort，无法接管浏览器")

            home = ret.get("launcherPage") or o.get("seller_home") or zb.TIKTOK_SELLER_HOME
            origins = self._origins(home, o)
            browser = zb.open_session(port, log=self.log.emit, home_urls=origins)
            summary["browser_port"] = port
            summary["origins"] = origins

            # IP 检测：确认店铺代理已生效（可在界面上跳过；失败不中断）
            ip_url = ret.get("ipDetectionPage")
            if o.get("skip_ip_check"):
                self.log.emit("已跳过 IP 检测")
            else:
                self.stage.emit("IP 检测")
                zb.check_ip(browser, ip_url, log=self.log.emit)

            # ---- 打开店铺后台，建立会话 ----
            # 账号下挂了多个店铺时（如秋图2：Spain / United States），**直接导航到
            # 目标店铺的后台首页** —— URL 里的 shop_region 决定是哪家店，一步到位：
            #   https://seller.<区域>.tiktokshopglobalselling.com/homepage?shop_region=XX
            # 不模拟点击右上角切换器（要开菜单、等渲染、还可能挡住导出按钮）。
            # 切错店铺 → cookie 里 seller_id 是另一家的 → 后面整批 98001008，所以
            # home / origins 必须跟着换：cookie 是**按域名精确取**的
            # （seller.us.* ≠ seller.eu.*），站不对就拿不到这家店的身份。
            target_shop = (o.get("target_shop") or "").strip()
            shop_url = te.shop_home_url(target_shop) if target_shop else ""
            if shop_url:
                home = shop_url
                origins = self._origins(home, o)
                summary["origins"] = origins
            self.stage.emit("打开店铺后台")
            self.log.emit(f"打开店铺后台：{home}")
            try:
                zb.goto(browser, home, log=self.log.emit)
                self.log.emit(f"当前页面：{browser.current_url()[:120]}")
            except Exception as exc:
                self.log.emit(f"打开后台页面失败（继续尝试接口）：{exc}")

            # ---- 确认确实停在目标店铺（已在 = 幂等不动；认不出的店铺名只告警）----
            if target_shop:
                self.stage.emit("确认后台店铺")
                shop = te.ensure_shop(browser, target_shop, log=self.log.emit)
                new_home = str(shop.get("url") or "").strip()
                if new_home:
                    home = new_home
                    origins = self._origins(home, o)
                    summary["origins"] = origins
                if not shop.get("ok"):
                    self.log.emit("[警告] 没能确认已切到目标后台店铺，继续用当前页面跑；"
                                  "若下面报 98001008，先看上面的落地 URL 停在哪个店")

            # 取 cookie 必须**按店铺域名精确取**（store_context → Network.getCookies）。
            # Storage.getCookies 会把整个 profile 的 cookie 全给我们：统一商家环境下
            # 同一账号可能同时有多个区域店铺（seller.us.* / seller.eu.*），
            # oec_seller_id_unified_seller_env 这类同名 cookie 会有多个值，合并成 dict
            # 时「后写覆盖前写」→ 可能拿到**别的店铺**的身份，整批 98001008（实测踩过）。
            # 同时等它稳定：页面刚导航时前端还没把「当前店铺」写进 cookie。
            typed = (o.get("seller_id") or "").strip()
            ctx_fn = getattr(browser, "store_context", None)
            if callable(ctx_fn):
                self.stage.emit("确认店铺身份")
                ctx = ctx_fn(home, expect=typed,
                             timeout=float(o.get("store_wait") or 25))
            else:  # 桩/旧对象兜底
                cookies_ = zb.grab_cookies(browser, log=self.log.emit, url=home)
                ctx = {"cookies": cookies_,
                       "seller_id": zb.seller_id_from_cookies(cookies_),
                       "stable": True, "seen": [], "candidates": []}
            cookies = ctx.get("cookies") or {}
            cands = sorted({c["value"] for c in (ctx.get("candidates") or [])})
            if len(cands) > 1:
                self.log.emit(f"[警告] 浏览器里读到 {len(cands)} 个不同的店铺 id："
                              f"{[f'{v}（{zb.store_name_of(v) or chr(63)}）' for v in cands]}")
                self.log.emit("        已按店铺后台域名精确取用，若抓取时仍报 98001008，"
                              "请确认③里打开的店铺就是导出订单的那个店。")
            # seller_id 必须来自**该店铺自己**的仓库身份。绝不回退到内置默认值 ——
            # 那是另一个店铺的 id（实测踩过）。界面上手填的优先。
            seller = typed or ctx.get("seller_id") or ""
            if cookies:
                # 明确写清身份从哪来 —— 这条链路**不读任何本地 cookie 文件**，
                # 用的是紫鸟店铺浏览器里那份实时登录态（CDP Network.getCookies，
                # 且按店铺后台域名精确取）。换台电脑跑同一个 exe 也能用，
                # 前提是那台机器上的店铺浏览器登录的是同一家店。
                self.log.emit("身份来源：紫鸟店铺浏览器实时登录态（非本地文件、非内置默认值）")
                self.log.emit(f"从店铺浏览器取到 {len(cookies)} 条 cookie，"
                              f"店铺 seller_id={seller or '(未取到)'}"
                              f"{f'（{zb.store_name_of(seller)}）' if zb.store_name_of(seller) else ''}"
                              f"{'' if ctx.get('stable') else '［未稳定，可能取早了］'}")
            if not seller:
                self.log.emit(
                    "[警告] 没能确定 seller_id：店铺 cookie 里没有该平台的登录态。\n"
                    "        请先在紫鸟里打开这个店铺并登录 TikTok 商家后台，"
                    "或在界面上手动填写 seller_id。"
                )

            # ---- 组装请求通道 ----
            self.stage.emit("准备请求通道")
            mode = o.get("transport_mode") or "auto"
            chain = self._build_chain(browser, cookies, mode, origins)
            client_obj = tc.OrderClient.from_cookies(
                cookies, transport=chain, seller_id=seller,
                api_base=(o.get("api_base") or tc.DEFAULT_API_BASE).rstrip("/"),
                retries=int(o.get("retries") or 2),
            )
            self.log.emit(f"接口主机：{client_obj.api_base}；"
                          f"seller_id={client_obj.seller_id or '(未取到)'}")

            # ---- ③ 导出订单 ----
            rows: list[dict[str, str]] = []
            if o.get("do_orders"):
                rows = self._export_orders(browser)
                summary["order_count"] = len(rows)
                summary["order_out"] = str(o.get("order_out") or "")
                self.orders_ready.emit(rows)

                # 订单是页面自己导出的 → 「页面自己请求里用的 seller_id」是唯一
                # 可信的店铺身份（cookie 可能串味）。不一致就以页面为准。
                page = getattr(self, "last_page_store", None) or {}
                psid = str(page.get("seller_id") or "")
                if page.get("has_data") and psid:
                    if psid != seller:
                        self.log.emit("")
                        self.log.emit("⚠ 店铺身份不一致！（cookie 取到的和页面实际用的不同）")
                        self.log.emit(f"    页面自己发出的请求用的是 seller_id={psid}"
                                      f"（{zb.store_name_of(psid) or '未登记'}）")
                        self.log.emit(f"    我们原本准备用的是   seller_id={seller}"
                                      f"（{zb.store_name_of(seller) or '未登记'}）")
                        self.log.emit("    → 以页面为准（订单就是它导出的），已切换。")
                        seller = psid
                        client_obj.seller_id = psid
                    else:
                        self.log.emit(f"店铺身份已核对：seller_id={psid}"
                                      f"（{zb.store_name_of(psid) or '未登记'}）"
                                      f"，与页面一致")
                else:
                    self.log.emit("[提示] 页面还没发出带 seller_id 的请求，"
                                  "跳过店铺身份核对（继续用 cookie 推导的值）")

            # ---- ④ 抓客户信息 ----
            if o.get("do_customers") and not self.stop_event.is_set():
                got = self._fetch_customers(browser, client_obj, rows, origins)
                summary["customer_count"] = got.get("ok", 0)
                summary["errors"] += got.get("fail", 0)
                summary["customer_out"] = got.get("out", "")

            # ---- ⑤ 导出商品评分 ----
            # 放在最后：它会把页面导航到 /product/rating，而 ④ 抓客户依赖页面停在
            # 订单页。反过来就每次都要靠 _ensure_order_page 救一次，白白多一次导航。
            if o.get("do_rating") and not self.stop_event.is_set():
                rrows = self._export_ratings(browser)
                summary["rating_count"] = len(rrows)
                summary["rating_out"] = str(o.get("rating_out") or "")
                self.rating_ready.emit(rrows)

            # ---- ⑥ 差评客户 ----
            # 可以**单独跑**：⑤ 没跑就直接读「评分 Excel」框里那个文件（⑤ 已经筛
            # 过的是幂等的，没筛过的就地筛一遍再覆盖保存）。
            # 评分文件**自带订单号**（真机导出的 CSV 就有 order_id 列）→ 直接用它
            # 去 order/get，不碰订单表。只有老版「仅昵称」的文件才拿昵称回订单表
            # 换 Order ID（可能因脱敏 K***** 匹配不上）。跟在 ④ 后面时复用缓存，
            # 不重复发请求。
            if o.get("do_bad") and not self.stop_event.is_set():
                if not o.get("do_rating"):
                    self._load_rating_bad()
                if self._bad_hit:
                    got = self._fetch_bad_customers(browser, client_obj,
                                                    self._rating_bad,
                                                    origins, rows)
                    summary["bad_count"] = got.get("ok", 0)
                    summary["errors"] += got.get("fail", 0)
                    summary["bad_out"] = got.get("out", "")
                    self.bad_ready.emit(got.get("rows") or [])
                else:
                    self.log.emit("没有 1~3 星的评价 → ⑥ 不抓差评客户")

            summary["stopped"] = self.stop_event.is_set()
            self.done.emit(summary)

        except Exception as exc:
            self.log.emit(traceback.format_exc(limit=8))
            self.failed.emit(f"{type(exc).__name__}: {exc}")
        finally:
            if browser is not None:
                browser.close()  # 只断开 CDP，浏览器交给 stopBrowser 关
            if client is not None and store_key:
                self.log.emit(f"===== 关闭店铺：{zb.store_label(self.store)} =====")
                zb.close_store(client, str(store_key), log=self.log.emit)

    # -- 通道 ---------------------------------------------------------------
    @staticmethod
    def _origins(home: str, o: dict[str, Any]) -> list[str]:
        """候选页面来源：界面上手填的 → 店铺后台首页 → 内置候选 → 接口 origin。

        fetch 必须从其中一个 origin 发起，否则跨域预检失败（Failed to fetch）。

        只收 http/https：紫鸟给的 ``launcherPage`` 有时是它自己的
        ``chrome-extension://…`` 检测页，拿它当页面来源做 ``Page.navigate``
        会把 target 直接弄丢（之后连不上，报一堆 JS 错误）。
        """
        out: list[str] = []
        for cand in [o.get("page_origins") or "", home, *zb.TIKTOK_PAGE_ORIGINS]:
            for part in str(cand).replace(";", ",").split(","):
                origin = zb.origin_of(part.strip())
                if origin and zb._is_page_url(origin) and origin not in out:
                    out.append(origin)
        api_origin = zb.origin_of(o.get("api_base") or "")
        if api_origin and api_origin not in out:
            out.append(api_origin)  # 同源兜底：完全绕开 CORS
        return out

    def _build_chain(self, browser: Any, cookies: dict[str, str], mode: str,
                     origins: list[str]):
        items: list[tuple[str, Callable]] = []
        if mode in ("auto", "browser"):
            # 有抓包模板就照抄页面真实请求（host / 路径前缀 / 查询参数），
            # 否则至少改写成「当前页面的同源地址」—— 页面自己就是这么发的。
            tmpl = zb.load_captured_template()
            if tmpl and tmpl.get("url"):
                self.log.emit(f"检测到抓包模板：{tmpl['url'][:110]}")
                self.log.emit("  → 浏览器内请求将照抄它的 host / 路径前缀 / 查询参数")
            # page_path：同源 fetch **必须**从订单页发起（真机定案）。
            # 换区域站时会直接打开「该站的订单页」，而不是只到站点根。
            items.append(("浏览器内", zb.BrowserTransport(
                browser, log=self.log.emit, origins=origins, template=tmpl,
                page_path=te.ORDER_PAGE_SUFFIX)))
        if mode in ("auto", "direct"):
            items.append(("本地直连", zb.RequestsTransport(
                cookies, dict(tc.DEFAULT_HEADERS), log=self.log.emit)))
        if not items:
            raise zb.ZiniaoBridgeError("没有可用的请求通道（浏览器连接与 cookie 均不可用）")
        # 「换个通道试试」的信号，只在真有下一个通道时才启用（单通道等于不启用）：
        #   · 10000         —— order/list 本地直连恒定返回它（风控/限流）
        #   · 98001008 等    —— 浏览器内那条路可能只是**页面站/页**选得不对。
        #                      而 ``order/get``（客户信息）本地直连实测可用，
        #                      让它有机会兜住，不至于因为一个站选错就整批中止。
        codes = ["10000"] + (list(zb.PERMISSION_CODES) if len(items) > 1 else [])
        return zb.TransportChain(items, log=self.log.emit, fallthrough_codes=tuple(codes))

    # -- 订单 ---------------------------------------------------------------
    def _export_orders(self, browser: Any) -> list[dict[str, str]]:
        """③ 走后台「导出订单」按钮导出并下载（不调 order/list 接口）。

        为什么不用接口：``order/list`` 本地直连恒定 ``code=10000``，页面里那版
        ``/api/fulfillment/na/order/list`` 又缺签名参数（重放 404）。而页面上
        「导出」按钮能直接把全部订单导成 Excel（实测 63 列 × 616 行，客户字段都有）。
        下载环节用 Fetch 域拦一次性 URL + 页面内 fetch 取字节 —— 详见 tiktok_export。
        """
        o = self.opts
        self.stage.emit("导出订单")
        self.log.emit("===== 后台「导出订单」按钮导出 =====")
        exporter = te.OrderExporter(browser, log=self.log.emit)
        out = Path(o.get("order_out") or DEFAULT_ORDER_OUT)
        raw_out = out.with_name(f"_导出原始{out.suffix or '.xlsx'}")
        self.order_progress.emit(0, 0)

        result = exporter.export(
            out_path=raw_out,
            scope=o.get("export_scope") or te.SCOPE_FILTERED,
            fmt=o.get("export_format") or "excel",
            timeout=float(o.get("export_timeout") or 300),
            poll=float(o.get("export_poll") or 6),
        )
        header: list[str] = result["header"]
        rows: list[list[Any]] = result["rows"]
        if not rows:
            self.log.emit("⚠ 导出文件里没有数据行（可能是筛选结果为空）")
            return []

        te.write_rows(header, rows, out)
        self.log.emit(f"共 {len(rows)} 条订单，已写入：{out}（{len(header)} 列）")
        self.order_progress.emit(len(rows), len(rows))

        # 记下**导出成功时页面所在的站**。这是这家店自己的区域站（导出就是在它上面
        # 点出来的），④ 抓客户会优先按这个站发；不写这一行，之后 98001008 就没法
        # 一眼看出「当时页面到底停在哪个站」（2026-10-14 真机排查全靠它）。
        try:
            self.export_origin = exporter.host()
        except Exception:
            self.export_origin = ""
        if self.export_origin:
            self.log.emit(f"导出用的站：{self.export_origin}")

        # 顺手把「页面自己用的 seller_id」取回来 —— 这是唯一可信的店铺身份，
        # ④ 用它来核对 cookie 推导出来的值（cookie 可能多店铺串味）。
        try:
            self.last_page_store = exporter.page_seller_id()
            counts = self.last_page_store.get("counts") or {}
            if counts:
                items = sorted(counts.items(), key=lambda kv: -kv[1])
                self.log.emit("页面请求里的 seller_id："
                              + "，".join(f"{k}×{v}（{zb.store_name_of(k) or '未登记'}）"
                                          for k, v in items[:5]))
        except Exception as exc:
            self.last_page_store = {"has_data": False}
            self.log.emit(f"读页面店铺身份失败（忽略，继续）：{exc}")
        return [dict(zip(header, r)) for r in rows]

    def _export_ratings(self, browser: Any) -> list[dict[str, str]]:
        """⑤ 走商品评分页自己的「导出」按钮导出并下载。

        和订单是两条独立的路：评分页没有「导出范围 / 格式」面板，只有一颗按钮；
        点完可能立刻下载、也可能进导出历史异步生成 —— 两种都在
        ``te.RatingExporter`` 里实现了。
        """
        o = self.opts
        self.stage.emit("导出商品评分")
        self.log.emit("===== 后台「商品评分」导出 =====")
        exporter = te.RatingExporter(browser, log=self.log.emit)
        # 用③导出订单成功时那个站（= 这家店自己的区域站），别退回通用站
        exporter.base = self.export_origin or ""
        out = Path(o.get("rating_out") or DEFAULT_RATING_OUT)
        raw_out = out.with_name(f"_评分原始{out.suffix or '.xlsx'}")
        self.rating_progress.emit(0, 0)

        result = exporter.export(
            out_path=raw_out,
            timeout=float(o.get("export_timeout") or 300),
            poll=float(o.get("export_poll") or 6),
            retries=int(o.get("retries") or 3),
        )
        header: list[str] = result["header"]
        rows: list[list[Any]] = result["rows"]
        if not rows:
            self.log.emit("⚠ 商品评分导出文件里没有数据行")
            return []

        return self._apply_rating_filter(header, rows, out)

    def _apply_rating_filter(self, header: list[str], rows: list[list[Any]],
                             out: Path) -> list[dict[str, str]]:
        """⑤ 和 ⑥ 共用：筛出 1~3 星 → **覆盖保存**到 ``out`` → 记进 self。

        完整版由调用方各自另存（⑤ 存 ``_评分原始.xlsx``；单独跑⑥时若完整版
        已存在就别动它，免得把完整版盖成筛过的）。
        没筛到差评时**保留完整表**：宁可多存一份数据，也绝不往用户的文件里
        写一个空表。
        """
        score_col, bad_rows, _scores = te.filter_low_ratings(header, rows)
        self._rating_header = list(header)
        self._rating_cols = {
            "score": score_col,
            "buyer": te.pick_column(header, te.RATING_BUYER_COLUMNS),
            "product": te.pick_column(header, te.RATING_PRODUCT_COLUMNS),
            # ★ 真机 CSV 评分文件自带 order_id（2026-10-15 实测）→ 有它就直接用，
            #   不必拿昵称回订单表反查（老版「只有昵称」的文件才需要那条路）。
            "order": te.pick_column(header, te.RATING_ORDER_COLUMNS),
        }
        if not score_col:
            self.log.emit(f"⚠ 没认出评分列（文件列：{'，'.join(header[:8])}…），"
                          "本轮不筛选，评分表保留完整内容")
        hits = bool(bad_rows)
        if hits:
            te.write_rows(header, bad_rows, out, sheet="商品评分")
            self.log.emit(f"1~3 星 {len(bad_rows)}/{len(rows)} 条，"
                          f"已覆盖保存：{out}（{len(header)} 列）")
        else:
            te.write_rows(header, rows, out, sheet="商品评分")
            self.log.emit(f"共 {len(rows)} 条商品评分（其中没有 1~3 星），"
                          f"已写入：{out}（{len(header)} 列）")
        self._bad_hit = hits
        self._rating_bad = [dict(zip(header, r)) for r in bad_rows]
        self._rating_all = [dict(zip(header, r)) for r in rows]
        self.rating_progress.emit(len(bad_rows) if hits else len(rows),
                                  len(bad_rows) if hits else len(rows))
        return self._rating_bad if hits else self._rating_all

    def _load_rating_bad(self) -> bool:
        """单独跑 ⑥ 时的评分来源：直接读 ⑤ 已经导出的评分 Excel。

        读 ``rating_out``（⑤ 已把它筛成 1~3 星，再筛一遍是幂等的）；它不在就退
        回完整版 ``_评分原始``（真机评分导出其实是 **CSV**，所以 ``.csv`` 也认）。
        完整版不存在时才先备份一份，避免覆盖丢数据。
        返回 True 表示有差评可抓。
        """
        o = self.opts
        out = Path(o.get("rating_out") or DEFAULT_RATING_OUT)
        candidates = rating_file_candidates(out)
        raw_candidates = [p for p in candidates if p != out]
        src = next((p for p in candidates if p.exists()), out)
        raw_out = next((p for p in raw_candidates if p.exists()), raw_candidates[0])
        if not src.exists():
            self.log.emit(f"✖ 找不到评分文件：{out} / {' / '.join(str(p) for p in raw_candidates)}\n"
                          "   → 先点「导出商品评分」，或把文件路径填到「评分 Excel」框。")
            return False
        self.stage.emit("读取商品评分")
        self.log.emit(f"===== 读取商品评分：{src.name} =====")
        try:
            header, rows = te.parse_rating(src)
        except Exception as exc:
            self.log.emit(f"✖ 读评分文件失败：{type(exc).__name__}: {exc}")
            return False
        if not rows:
            self.log.emit("⚠ 评分文件里没有数据行 → ⑥ 跳过")
            return False
        # 读的是「已筛过的 out」且完整版还没存过 → 先另存一份再覆盖。
        # 读的是完整版（raw）就不用备份：它自己就是完整版。
        if src == out and not raw_out.exists():
            try:
                te.write_rows(header, rows, raw_out, sheet="商品评分")
                self.log.emit(f"完整版另存：{raw_out}")
            except Exception as exc:
                self.log.emit(f"[提示] 完整版另存失败（忽略）：{exc}")
        self._apply_rating_filter(header, rows, out)
        self.rating_ready.emit(self._rating_bad if self._bad_hit else self._rating_all)
        return bool(self._bad_hit)

    # -- 客户 ---------------------------------------------------------------
    def _resolve_customer_ids(self, rows: list[dict[str, str]]) -> list[str]:
        if rows:
            ids = [r.get("Order ID", "") for r in rows if r.get("Order ID")]
            if ids:
                return ids
        src = Path(self.opts.get("customer_input") or "")
        if not src.exists():
            raise zb.ZiniaoBridgeError(f"订单 Excel 不存在：{src}")
        ids = tc.read_order_ids(src, self.opts.get("customer_sheet") or None,
                                self.opts.get("order_column") or "Order ID")
        if not ids:
            raise zb.ZiniaoBridgeError(
                f"在「{self.opts.get('order_column') or 'Order ID'}」列没有读到有效 Order ID（{src}）")
        return ids

    @staticmethod
    def _sku_col_of(row: dict[str, str],
                    columns: tuple[str, ...] | None = None) -> str:
        """在导出行里找目标 SKU 列的实际列名（平台各版本叫法不同）。

        ``columns`` 不传 = 「平台 SKU ID」那组别名；传
        ``tc.SELLER_SKU_SOURCE_COLUMNS`` 就找「商家 SKU」那一列（``Seller SKU``）。
        """
        wanted = {tc.norm_col(c) for c in (columns or tc.SKU_SOURCE_COLUMNS)}
        for key in row.keys():
            if tc.norm_col(key) in wanted:
                return str(key)
        return ""

    @classmethod
    def _sku_map_from_rows(cls, rows: list[dict[str, str]],
                           columns: tuple[str, ...] | None = None) -> dict[str, str]:
        """把③刚导出的订单表折叠成 {Order ID: 某个 SKU 列}（一单多 SKU 合并）。"""
        parts: dict[str, list[str]] = {}
        for row in rows or []:
            col = cls._sku_col_of(row, columns)
            if not col:
                return {}
            oid = str(row.get("Order ID") or "").strip()
            if not oid.isdigit() or len(oid) < 10:
                continue
            val = row.get(col)
            if val:
                parts.setdefault(oid, []).append(val)
        return {oid: tc.join_skus(v) for oid, v in parts.items()}

    def _load_sku_map(self, rows: list[dict[str, str]],
                      columns: tuple[str, ...] | None = None,
                      label: str = "SKU ID") -> tuple[dict[str, str], str]:
        """订单 ID → 某个 SKU 列（**兜底用**），返回 ``(映射, 来源说明)``。

        ``columns`` 默认是「平台 SKU ID」那组别名；传
        ``tc.SELLER_SKU_SOURCE_COLUMNS`` 就改成读「商家 SKU」那列，其余逻辑一样
        （两条路都是「响应优先、订单表兜底」）。``label`` 只用于日志措辞。

        主源是接口响应里的 ``sku_module``（请求本身就带着），这张表只负责两种情形：
          ① 响应里没有 sku 节点；② 请求失败/中止的占位行 —— 它们没有响应可用。
        """
        if rows:
            got = self._sku_map_from_rows(rows, columns)
            if got:
                return got, "③导出的订单表"
        src = Path(self.opts.get("customer_input") or "")
        if src.exists():
            try:
                got = tc.read_sku_map(
                    src, self.opts.get("customer_sheet") or None,
                    self.opts.get("order_column") or "Order ID",
                    sku_columns=columns or tc.SKU_SOURCE_COLUMNS)
            except Exception as exc:  # SKU 是附加列，读不到不该中断抓客户
                self.log.emit(f"[提示] 读取订单表 {label} 失败（忽略，该列留空）："
                              f"{type(exc).__name__}: {exc}")
                got = {}
            if got:
                return got, f"订单文件 {src.name}"
        return {}, ""

    def _sku_pair(self, oid: str) -> tuple[str, str]:
        """订单表兜底的两个 SKU 值：``(平台 SKU ID, 商家 SKU)``。

        顺序就是 ``tc.parse_customer`` / ``tc.blank_customer`` 的
        ``(sku_ids, seller_skus)`` 参数顺序，调用处直接 ``*self._sku_pair(oid)``。
        响应里自带 sku 节点时，这两个值会被响应里的真实值盖掉。
        """
        return (self._sku_map_cache.get(oid, ""),
                self._seller_sku_cache.get(oid, ""))

    def _heal_store_identity(self, browser: Any, client_obj: tc.OrderClient,
                             ids: list[str], cur_sid: str,
                             origins: list[str]) -> dict[str, Any] | None:
        """预检失败时的自救 —— 让**浏览器内通道**把区域站扫完再复检。

        为什么是「扫站」而不是「换 cookie」：cookie 是**按站隔离**的
        （``seller.us.*`` / ``seller.eu.*`` / ``seller.*`` 各带一份），但同一账号在
        这些站里的 ``seller_id`` **常常是同一个** —— 所以按 seller_id 去重会把它们
        全部当成「＝当前身份」跳过（踩过）。真正决定能不能看这批订单的是
        **页面停在哪个区域站**：实测订单表 619 条全是美国订单，页面停在通用站
        → 全批 98001008；切到 ``seller.us.*`` 并用同源路径请求，立刻 code=0。

        ``BrowserTransport`` 自己就会做「区域站 × 写法」的扫描，并把打通的
        「写法 + 站点」组合锁定下来，所以这里只要给它一次机会并复检。
        （不再用本地直连：接口主机是 api16-*，那条通道拿不到区域站的身份。）
        """
        if cur_sid:
            self.log.emit(f"  · 当前身份 seller_id={cur_sid}"
                          f"（{zb.store_name_of(cur_sid) or '未登记'}），"
                          "让浏览器内通道把候选区域站扫完…")
        try:
            probe = client_obj.probe(ids, max_ids=3)
        except Exception as exc:
            self.log.emit(f"    复检异常：{type(exc).__name__}: {exc}")
            return None
        if not probe.get("ok"):
            first = str(probe.get("fatal") or "").splitlines()
            self.log.emit(f"    还是不行：{first[0] if first else '拿不到数据'}")
            return None
        sid = probe.get("seller_id") or ""
        self.log.emit(f"    ✅ 打通了 —— seller_id={sid}"
                      f"（{zb.store_name_of(sid) or '未登记'}）")
        return probe

    def _ensure_order_page(self, browser: Any) -> str:
        """把工作页面停到**订单页**上，返回实际落地的 origin（读不到就给空串）。

        ★ 这是抓客户能不能通的**第一前提**（2026-10-14 用户真机定案）★
        同源 fetch 只在订单页上成立：``<站>/order?selected_sort=6&tab=all``。
        注意它和「换区域站」是**两层**问题，都要满足：

          · 站错、页对 → 98001008（拿错站的身份发）
          · 站对、页错 → 照样 98001008（页面自己的会话上下文不在订单页上）

        优先用**导出订单时那个站**（``export_origin``）—— 导出就是在它上面成功的；
        没有就退回当前站，再没有就用内置首页。
        """
        base = (self.export_origin or "").rstrip("/")
        if not base:
            try:
                base = zb.origin_of(str(browser.current_url() or ""))
            except Exception:
                base = ""
        if not base:
            base = zb.TIKTOK_SELLER_HOME.rstrip("/")
        # 已经在「本站的订单页」上就别再导航 —— 导航会掉线，而且订单页首屏的
        # 那几秒等待也没必要（③ 导出刚走过一遍）。
        page_root = te.ORDER_PAGE_SUFFIX.split("?")[0]      # "/order"
        try:
            cur = str(browser.current_url() or "")
        except Exception:
            cur = ""
        already = (bool(cur) and cur.startswith("http")
                   and page_root in cur and zb.origin_of(cur) == base)
        if already:
            self.log.emit(f"页面已在订单页：{cur[:120]}")
            return zb.origin_of(cur)
        url = base + te.ORDER_PAGE_SUFFIX
        ok = False
        try:
            browser.navigate(url, retry=3, interval=2.0, wait=45)
            ok = True
        except Exception as exc:
            self.log.emit(f"导航到订单页失败（继续尝试接口）：{exc}")
        if ok:
            time.sleep(te.PAGE_SETTLE)      # 等订单页首屏请求跑完
        try:
            cur = str(browser.current_url() or "")
        except Exception:
            cur = ""
        landed = zb.origin_of(cur) if cur.startswith("http") else ""
        if cur:
            self.log.emit(f"页面已停在：{cur[:120]}")
        if landed and base and landed != base:
            self.log.emit(f"[注意] 请求的是 {base}，实际落在 {landed}"
                          "（区域站会互相重定向，抓客户按落地站发）")
        return landed

    def _fetch_customers(self, browser: Any, client_obj: tc.OrderClient,
                         rows: list[dict[str, str]],
                         origins: list[str] | None = None) -> dict[str, Any]:
        o = self.opts
        origins = list(origins or [])
        self.stage.emit("抓取客户信息")
        ids = self._resolve_customer_ids(rows)
        out = Path(o.get("customer_out") or DEFAULT_CUSTOMER_OUT)
        interval = float(o.get("interval") or 0.0)

        self.log.emit(f"===== 抓取客户信息：{len(ids)} 条，间隔 {interval}s =====")
        self.customer_progress.emit(0, len(ids))

        # SKU 兜底（两条路一样）：接口响应优先，没响应的行（失败/中止）由订单表补齐。
        # 「SKU ID」= 平台 ID；「商家 SKU」= 商家自定义编码（Seller SKU），两者都要。
        sku_map, sku_src = self._load_sku_map(rows)
        seller_map, seller_src = self._load_sku_map(
            rows, tc.SELLER_SKU_SOURCE_COLUMNS, "商家 SKU")
        self._sku_map_cache = sku_map          # ⑥ 直接用，不重复读订单表
        self._seller_sku_cache = seller_map
        if sku_map:
            self.log.emit(f"SKU ID 兜底表已就绪：{len(sku_map)} 条（来源：{sku_src}）")
        else:
            self.log.emit("[提示] 订单表里没有 SKU ID 列，SKU ID 只取接口响应里的值")
        if seller_map:
            self.log.emit(f"商家 SKU 兜底表已就绪：{len(seller_map)} 条"
                          f"（来源：{seller_src}）")
        else:
            self.log.emit("[提示] 订单表里没有 Seller SKU 列，"
                          "商家 SKU 只取接口响应里的值")

        # ★ 抓客户前必须把页面停回**订单页** ★
        # 2026-10-14 用户真机定案：同源 fetch 只在订单页上成立 —— 站选对了但页面
        # 停在首页/工作台，一样是 98001008「需要权限」。③ 导出本来就在订单页上，
        # 但下载与面板操作都可能把页面带偏，所以这里显式回一次。
        self._ensure_order_page(browser)

        # ---- 店铺身份预检：先拿前 3 条试水，别等跑完几百条才发现全错 ----
        try:
            probe = client_obj.probe(ids, max_ids=3)
        except Exception as exc:
            probe = {"ok": False, "fatal": "", "tried": [], "results": [],
                     "seller_id": client_obj.seller_id, "buyers": []}
            self.log.emit(f"预检异常（忽略，继续）：{type(exc).__name__}: {exc}")
        else:
            sid = probe.get("seller_id") or ""
            self.log.emit(f"店铺身份预检：seller_id={sid}"
                          f"（{zb.store_name_of(sid) or '未登记'}），"
                          f"试了 {len(probe.get('tried') or [])} 条，"
                          f"成功 {len(probe.get('buyers') or [])} 条")
        if not probe.get("ok"):
            if probe.get("fatal"):
                self.log.emit("")
                self.log.emit("✖ 预检未通过：这个站的登录态看不到这批订单")
                for line in str(probe["fatal"]).splitlines():
                    self.log.emit("   " + line)

            # 先别放弃。预检这轮并非白跑：浏览器内通道会在这一轮里把
            # **区域站 × 写法**逐个试过去，并把真正打通的组合锁定下来
            # （实测：通用站回 98001008 → 切到 seller.us.* 用同源路径立刻 code=0）。
            # 所以紧接着复检一次，往往直接就通。
            healed = None
            if browser is not None and origins:
                self.log.emit("")
                self.log.emit("→ 重扫区域站后复检：这批订单未必属于当前站…")
                healed = self._heal_store_identity(
                    browser, client_obj, ids, client_obj.seller_id, origins)

            if healed:
                probe = healed
                self.log.emit("")
                self.log.emit(f"✅ 已切换到能看这批订单的区域站：seller_id="
                              f"{client_obj.seller_id}"
                              f"（{zb.store_name_of(client_obj.seller_id) or '未登记'}）")
            else:
                results0 = {oid: tc.blank_customer(oid, "SKIPPED", *self._sku_pair(oid))
                            for oid in ids}
                self.log.emit("")
                self.log.emit("✖ 所有候选区域站都看不到这批订单，已中止抓客户"
                              "（后面的请求只会重复同样的错误）")
                self.log.emit("   → 这批订单很可能属于**另一个登录账号**。"
                              "一个店铺的登录态只能看它自己的订单，"
                              "换店铺必须换那个店铺的登录态（紫鸟里换店/换账号）。")
                self.log.emit("   → 别去改 params 里的 seller_id：实测整段删掉也还是 "
                              "98001008，服务端按会话判定。")
                self.log.emit("   → 定位：python devtools/diagnose_store.py --store N")
                final0 = self._flush_customers(results0, ids, out)
                for i, oid in enumerate(ids):
                    self.customer_row.emit(i, results0[oid])
                    self.customer_progress.emit(i + 1, len(ids))
                return {"ok": 0, "fail": len(ids), "out": str(final0), "aborted": True}

        results: dict[str, dict[str, str]] = {}
        ok = fail = 0
        fatal = ""
        healed_try = False          # 主循环里只自救一次
        for i, oid in enumerate(ids):
            if self.stop_event.is_set():
                self.log.emit("已手动停止，写出已抓取的数据…")
                break
            if fatal:  # 已判定是店铺/凭据问题：补齐占位行收工，不再发请求
                row = tc.blank_customer(oid, "SKIPPED", *self._sku_pair(oid))
                results[oid] = row
                self.customer_row.emit(i, row)
                self.customer_progress.emit(i + 1, len(ids))
                continue
            try:
                payload = client_obj.fetch([oid], chunk_size=1)
                orders = ((payload.get("data") or {}).get("main_order")) or []
                if not orders:
                    row = tc.blank_customer(oid, "EMPTY", *self._sku_pair(oid))
                    fail += 1
                    self.log.emit(f"[{i + 1}/{len(ids)}] {oid} → 无数据返回")
                else:
                    row = tc.parse_customer(orders[0], *self._sku_pair(oid))
                    ok += 1
                    self.log.emit(f"[{i + 1}/{len(ids)}] {oid} → "
                                  f"{row.get('Buyer Nickname', '')} {row.get('Recipient', '')}")
            except tc.ApiError as exc:
                row = tc.blank_customer(oid, "API_ERROR", *self._sku_pair(oid))
                fail += 1
                self.log.emit(f"[{i + 1}/{len(ids)}] {oid} → 失败：{exc}")
                if exc.is_fatal and not fatal:
                    # 保险：预检通过了但首条就撞凭据错误 —— 通道在这一路上还会
                    # 再扫一遍区域站并锁定可用组合，所以复检一次再决定中止与否。
                    repaired = False
                    if not healed_try and browser is not None and origins:
                        healed_try = True
                        self.log.emit("   → 疑似区域站没对：重扫并复检…")
                        again = self._heal_store_identity(
                            browser, client_obj, ids, client_obj.seller_id, origins)
                        repaired = bool(again and again.get("ok"))
                    if repaired:
                        self.log.emit("   → 复检通过，重试本条后继续。")
                        try:
                            payload = client_obj.fetch([oid], chunk_size=1)
                            orders = ((payload.get("data") or {}).get("main_order")) or []
                            if orders:
                                row = tc.parse_customer(orders[0],
                                                         *self._sku_pair(oid))
                                ok += 1
                                fail -= 1
                                self.log.emit(f"[{i + 1}/{len(ids)}] {oid} → "
                                              f"{row.get('Buyer Nickname', '')} "
                                              f"{row.get('Recipient', '')}")
                            else:
                                row = tc.blank_customer(oid, "EMPTY",
                                                          *self._sku_pair(oid))
                        except Exception as exc2:
                            self.log.emit(f"   重试仍失败：{type(exc2).__name__}: {exc2}")
                    else:
                        fatal = str(exc)
                        self.log.emit("—" * 58)
                        self.log.emit(f"已中止：这是店铺/凭据问题（code={exc.code}），"
                                      f"剩下的 {len(ids) - i - 1} 条不再请求。")
                        self.log.emit("请核对③里打开的店铺是否就是导出该订单表的店铺。")
                        self.log.emit("—" * 58)
            except Exception as exc:
                row = tc.blank_customer(oid, "EXCEPTION", *self._sku_pair(oid))
                fail += 1
                self.log.emit(f"[{i + 1}/{len(ids)}] {oid} → 异常："
                              f"{type(exc).__name__}: {exc}")

            # 存进缓存：⑥ 差评客户会命中同一批订单，不重复发请求
            self._cust_cache[oid] = row
            results[oid] = row
            self.customer_row.emit(i, row)
            self.customer_progress.emit(i + 1, len(ids))

            if (i + 1) % 10 == 0:
                self._flush_customers(results, ids, out)
            if not self.stop_event.is_set() and i + 1 < len(ids) and interval > 0:
                time.sleep(interval)

        final = self._flush_customers(results, ids, out)
        self.log.emit(f"客户信息抓取结束：成功 {ok}，失败 {fail}，已写入 {final}")
        return {"ok": ok, "fail": fail, "out": str(final)}

    # -- ⑥ 差评客户 ---------------------------------------------------------
    @staticmethod
    def _norm(value: Any) -> str:
        """昵称/商品名归一化：只消大小写与首尾空格，保留脱敏的 *****（对不上就是
        对不上，别伪装成匹配）。"""
        return str(value or "").strip().lower()

    @staticmethod
    def _cols_of(row: dict[str, Any], names: Sequence[str]) -> list[str]:
        """取行里这一族列的值（列名按导出文件实际写法，不写死）。"""
        out: list[str] = []
        for key in row.keys():
            if te.norm_key(key) in [te.norm_key(n) for n in names]:
                val = row.get(key)
                if val:
                    out.append(str(val))
        return out

    def _order_index(self, order_rows: list[dict[str, str]]) -> dict[str, Any]:
        """订单表 → {昵称: {"ids": [...], "prods": [(商品名, oid), ...]}}。

        **兜底路径专用**：评分文件里没有订单号列（老版「仅昵称」文件）时才用它，
        先在订单表里建这个索引，⑥ 才能把 1~3 星的评价换成订单号去抓客户。
        真机评分 CSV 自带 order_id，走不到这里。
        订单行优先用③刚导出的那批（内存里），没有才回读订单 Excel。
        """
        rows = list(order_rows or [])
        if not rows:
            src = Path(self.opts.get("customer_input")
                       or self.opts.get("order_out") or "")
            if src.exists():
                header, data = te.parse_export(src)
                rows = [dict(zip(header, r)) for r in data]
        index: dict[str, Any] = {}
        for row in rows:
            oid = str(row.get("Order ID") or "").strip()
            if not (oid.isdigit() and len(oid) >= 8):
                continue
            nicks = self._cols_of(row, te.ORDER_BUYER_COLUMNS)
            prods = self._cols_of(row, te.RATING_PRODUCT_COLUMNS)
            for nick in nicks:
                nk = self._norm(nick)
                if not nk:
                    continue
                node = index.setdefault(nk, {"ids": [], "prods": []})
                if oid not in node["ids"]:
                    node["ids"].append(oid)
                for prod in prods:
                    pk = self._norm(prod)
                    if pk:
                        node["prods"].append((pk, oid))
        return index

    @staticmethod
    def _order_id(value: Any) -> str:
        """订单号单元格 → 纯数字串（认不出返回空串）。

        表格里的订单号可能是文本、也可能是数字单元格（Excel 读出来是 float，
        甚至被写成 ``5.7758E+17`` 这种科学计数），一律先归一再验长度。
        """
        text = str(value or "").strip().replace("#", "").replace(" ", "")
        if not text:
            return ""
        if text.endswith(".0"):
            text = text[:-2]
        if not text.isdigit():
            # ★ 只有「不是纯数字」时才走 float 换算（科学计数 / 数字单元格）。
            #   已经是一串数字的就**绝不能**碰 —— float 只有 53 位精度，
            #   577589364214960772 会被改成 …768，订单号直接查不到。
            try:
                num = float(text)
                if num.is_integer():
                    text = str(int(num))
            except (TypeError, ValueError):
                return ""
        return text if (text.isdigit() and len(text) >= 8) else ""

    def _match_orders(self, index: dict[str, Any], nick: str,
                      product: str) -> list[str]:
        """先用「昵称+商品」精确锁定，锁不到再退回「只用昵称」。"""
        node = index.get(nick) or {}
        if not node:
            return []
        if product:
            hit: list[str] = []
            for p, oid in node.get("prods") or []:
                # 完全相等，或一方包含另一方（商品名常被截断）
                if p == product or (len(p) >= 8 and (p in product or product in p)):
                    if oid not in hit:
                        hit.append(oid)
            if hit:
                return hit
        return list(node.get("ids") or [])

    def _fetch_bad_customers(self, browser: Any, client_obj: tc.OrderClient,
                             bad_rows: list[dict[str, Any]],
                             origins: list[str] | None,
                             order_rows: list[dict[str, str]]) -> dict[str, Any]:
        """⑥ 把 1~3 星评价对应的客户抓出来，写 `<店铺名>_客户差评信息.xlsx`。

        链条：⑤ 筛出的差评行 →**订单号**→ order/get → 客户行 + 评分 + 商品。
        订单号优先取评分文件自带的那一列（真机 CSV 就有）；老版「仅昵称」的文件
        才退回去拿（昵称+商品）查订单表换 Order ID。④ 抓过的订单直接复用缓存，
        不发第二次请求。
        """
        o = self.opts
        self.stage.emit("抓取差评客户")
        out = Path(o.get("bad_out") or DEFAULT_BAD_CUSTOMER_OUT)
        interval = float(o.get("interval") or 0.0)
        cols = self._rating_cols
        buyer_col = cols.get("buyer") or ""
        product_col = cols.get("product") or ""
        score_col = cols.get("score") or ""
        order_col = cols.get("order") or ""

        self.log.emit(f"===== 抓取差评客户：1~3 星 {len(bad_rows)} 条 =====")

        # ---- 差评行 → 订单号（同一订单只留第一条差评）----
        # ① 评分文件自带订单号列（真机 CSV 就是这样）→ 直接用，最省事也最准；
        # ② 老版文件只有买家昵称 → 回订单表反查（可能因脱敏 K***** 匹配不上）。
        pairs: list[tuple[str, dict[str, Any]]] = []
        seen: set[str] = set()
        missed = 0
        if order_col:
            self.log.emit(f"评分文件自带订单号列「{order_col}」→ 直接取订单号，不查订单表")
            for row in bad_rows:
                oid = self._order_id(row.get(order_col))
                if not oid:
                    missed += 1
                    continue
                if oid in seen:
                    continue
                seen.add(oid)
                pairs.append((oid, row))
        else:
            if not buyer_col:
                self.log.emit("✖ 评分表里既没订单号列、也没认出买家昵称列 → "
                              "无法定位订单，⑥ 跳过")
                self.log.emit("   评分文件实际列名：" + "，".join(self._rating_header[:10]))
                self.log.emit("   → 把列名发我，加进 te.RATING_ORDER_COLUMNS / "
                              "te.RATING_BUYER_COLUMNS 即可")
                return {"ok": 0, "fail": 0, "out": "", "rows": []}

            index = self._order_index(order_rows)
            if not index:
                self.log.emit("✖ 订单表里没有可用的买家昵称列，⑥ 跳过"
                              "（③ 没跑过就先跑一次导出订单）")
                return {"ok": 0, "fail": 0, "out": "", "rows": []}
            self.log.emit(f"订单索引：{len(index)} 个买家")

            for row in bad_rows:
                nick = self._norm(row.get(buyer_col))
                if not nick:
                    missed += 1
                    continue
                ids = self._match_orders(index, nick, self._norm(row.get(product_col)))
                if not ids:
                    missed += 1
                    continue
                for oid in ids:
                    if oid in seen:
                        continue
                    seen.add(oid)
                    pairs.append((oid, row))
        if missed:
            self.log.emit(f"[提示] {missed}/{len(bad_rows)} 条差评没匹配到订单"
                          "（昵称可能在评分文件里被平台脱敏成 K*****）")
        if not pairs:
            self.log.emit("⚠ 没有一条差评能对应到订单 → 不写差评客户表")
            return {"ok": 0, "fail": 0, "out": "", "rows": []}
        self.log.emit(f"差评 → 订单：{len(pairs)} 单待抓")

        # ★ 抓客户前必须把页面停回**订单页**（和④同一个前提：站对、页不对照样
        #   98001008）。⑤ 结束时页面还在 /product/rating 上。
        self._ensure_order_page(browser)

        # 单独跑⑥时④没跑过，缓存是空的 —— 这里自己按订单表补一份（两张表都补）
        if not (self._sku_map_cache and self._seller_sku_cache) \
                and order_rows is not None:
            rows_l = list(order_rows)
            if not self._sku_map_cache:
                self._sku_map_cache, _src = self._load_sku_map(rows_l)
                if self._sku_map_cache:
                    self.log.emit(f"SKU ID 兜底表已就绪：{len(self._sku_map_cache)} 条"
                                  f"（来源：{_src}）")
            if not self._seller_sku_cache:
                self._seller_sku_cache, _src2 = self._load_sku_map(
                    rows_l, tc.SELLER_SKU_SOURCE_COLUMNS, "商家 SKU")
                if self._seller_sku_cache:
                    self.log.emit(f"商家 SKU 兜底表已就绪："
                                  f"{len(self._seller_sku_cache)} 条（来源：{_src2}）")
        results: list[dict[str, str]] = []
        ok = fail = 0
        fatal = False
        for i, (oid, rrow) in enumerate(pairs):
            if self.stop_event.is_set():
                self.log.emit("已手动停止，写出已抓取的数据…")
                break
            score = te.rating_value(rrow.get(score_col)) if score_col else 0
            extra = {"评分": str(score) if score else str(rrow.get(score_col) or ""),
                     "商品": str(rrow.get(product_col) or "") if product_col else ""}
            if fatal:
                results.append(self._bad_row(
                    tc.blank_customer(oid, "SKIPPED", *self._sku_pair(oid)), extra))
                continue
            cached = self._cust_cache.get(oid)
            if cached:                       # ④ 已经抓过 —— 直接复用，不重复请求
                results.append(self._bad_row(dict(cached), extra))
                ok += 1
                self.log.emit(f"[{i + 1}/{len(pairs)}] {oid} → 复用④的结果")
            else:
                try:
                    payload = client_obj.fetch([oid], chunk_size=1)
                    orders = ((payload.get("data") or {}).get("main_order")) or []
                    if orders:
                        results.append(self._bad_row(
                            tc.parse_customer(orders[0], *self._sku_pair(oid)), extra))
                        ok += 1
                        self._cust_cache[oid] = results[-1]
                        self.log.emit(f"[{i + 1}/{len(pairs)}] {oid} → "
                                      f"{results[-1].get('Buyer Nickname', '')}")
                    else:
                        results.append(self._bad_row(
                            tc.blank_customer(oid, "EMPTY", *self._sku_pair(oid)),
                            extra))
                        fail += 1
                        self.log.emit(f"[{i + 1}/{len(pairs)}] {oid} → 无数据返回")
                except tc.ApiError as exc:
                    results.append(self._bad_row(
                        tc.blank_customer(oid, "API_ERROR", *self._sku_pair(oid)),
                        extra))
                    fail += 1
                    self.log.emit(f"[{i + 1}/{len(pairs)}] {oid} → 失败：{exc}")
                    if exc.is_fatal and not fatal:
                        fatal = True
                        self.log.emit("—" * 58)
                        self.log.emit(f"⑥ 已中止：{exc.code} 是店铺/凭据问题，"
                                      f"剩下的 {len(pairs) - i - 1} 条不再请求。")
                except Exception as exc:
                    results.append(self._bad_row(
                        tc.blank_customer(oid, "EXCEPTION", *self._sku_pair(oid)),
                        extra))
                    fail += 1
                    self.log.emit(f"[{i + 1}/{len(pairs)}] {oid} → 异常："
                                  f"{type(exc).__name__}: {exc}")
            self.bad_progress.emit(len(results), len(pairs))
            if not self.stop_event.is_set() and i + 1 < len(pairs) and interval > 0:
                time.sleep(interval)

        final = self._flush_bad_customers(results, out)
        self.log.emit(f"差评客户抓取结束：成功 {ok}，失败 {fail}，已写入 {final}")
        return {"ok": ok, "fail": fail, "out": str(final), "rows": results}

    @staticmethod
    def _bad_row(cust: dict[str, str], extra: dict[str, str]) -> dict[str, str]:
        """客户行 + 差评信息（评分 / 商品），列顺序 = 客户列 + 评分 + 商品。"""
        out = dict(cust)
        out.update(extra)
        return out

    @staticmethod
    def _flush_bad_customers(rows: list[dict[str, str]], out: Path) -> Path:
        cols = BAD_CUSTOMER_COLUMNS
        return te.write_rows(cols, [[r.get(c, "") for c in cols] for r in rows],
                             out, sheet="客户差评")

    @staticmethod
    def _flush_customers(results: dict[str, dict[str, str]], ids: list[str],
                         out: Path) -> Path:
        ordered = [results[i] for i in ids if i in results]
        return tc.write_customers(ordered, out)


# ----------------------------------------------------------------------------
# 对话框
# ----------------------------------------------------------------------------
class ZiniaoConfigDialog(QDialog):
    """编辑紫鸟 config.yaml（企业账号 / 客户端路径 / 端口 / 数据目录）。"""

    def __init__(self, cfg: Any, parent=None, root: str | Path | None = None):
        super().__init__(parent)
        self.setWindowTitle("紫鸟配置")
        self.resize(680, 360)
        self.cfg = cfg
        self.root = root
        layout = QVBoxLayout(self)
        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)

        self.ed_version = QComboBox()
        self.ed_version.addItems(["v6", "v5"])
        self.ed_version.setCurrentText(cfg.version or "v6")
        form.addRow("客户端版本", self.ed_version)

        # 客户端路径 + 自动检测。Windows 上紫鸟的安装目录是安装时自选的，没有
        # 固定位置，手填很容易填错（真机就出现过把 exe 填进「数据目录」）。
        client_row = QWidget()
        crl = QHBoxLayout(client_row)
        crl.setContentsMargins(0, 0, 0, 0)
        crl.setSpacing(6)
        self.ed_client = QLineEdit(cfg.client_path or "")
        self.ed_client.setPlaceholderText(
            "留空自动检测；Windows 可填 D:\\紫鸟\\ziniao 或 ……\\ziniao.exe；macOS 填 ziniao")
        self.ed_client.setToolTip(
            "紫鸟浏览器**客户端程序**的位置（不是数据目录）。\n"
            "Windows 没有固定安装路径，填 exe 所在的文件夹也行，留空会自动去找。")
        self.ed_client.setMinimumWidth(160)
        crl.addWidget(self.ed_client, 1)
        self.btn_find_client = QPushButton("自动检测")
        self.btn_find_client.setToolTip("在常见安装位置、运行中的进程、注册表里找 ziniao.exe")
        self.btn_find_client.clicked.connect(self._auto_detect_client)
        crl.addWidget(self.btn_find_client)
        form.addRow("客户端路径", client_row)

        self.sp_port = QSpinBox()
        self.sp_port.setRange(1, 65535)
        self.sp_port.setValue(int(cfg.socket_port or 16851))
        form.addRow("IPC 端口", self.sp_port)

        self.ed_userdata = QLineEdit(cfg.user_data_dir or "")
        self.ed_userdata.setPlaceholderText("留空自动推导（v6 必填项）")
        self.ed_userdata.setToolTip(
            "存放 web_driver 实例数据的**文件夹**（不是客户端程序）。\n"
            "留空即可，会自动建到用户目录下。")
        form.addRow("数据目录", self.ed_userdata)

        self.ed_company = QLineEdit(cfg.company or "")
        form.addRow("企业名称", self.ed_company)
        self.ed_username = QLineEdit(cfg.username or "")
        form.addRow("登录账号", self.ed_username)
        self.ed_password = QLineEdit(cfg.password or "")
        self.ed_password.setEchoMode(QLineEdit.EchoMode.Password)
        form.addRow("登录密码", self.ed_password)

        layout.addLayout(form)
        box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        box.accepted.connect(self._accept)
        box.rejected.connect(self.reject)
        layout.addWidget(box)

    def _auto_detect_client(self) -> None:
        """在常见位置 / 运行中的进程 / 注册表里找出客户端 exe，填进输入框。"""
        raw = self.ed_client.text().strip()
        try:
            found = zb.find_client_executable(raw, root=self.root, force=True)
        except Exception as exc:
            QMessageBox.warning(self, "自动检测失败", f"加载紫鸟项目出错：\n{exc}")
            return
        if found:
            self.ed_client.setText(found)
            QMessageBox.information(self, "已找到客户端", found)
            return
        # 找不到：按平台给出「填成什么样」+「已经找过哪些位置」，比一句
        # 「没找到」有用得多（Windows 上填错的坑特别多）。
        QMessageBox.information(self, "没找到客户端", zb.client_hint(self.cfg))

    def _accept(self) -> None:
        c = self.cfg
        client = self.ed_client.text().strip()
        data = self.ed_userdata.text().strip()

        # 最常见的误填：把客户端 exe 填进了「数据目录」（两个框挨着）。拦住并帮它搬过去。
        if data and (data.lower().endswith(".exe") or Path(data).is_file()):
            ans = QMessageBox.question(
                self, "数据目录填错了",
                "「数据目录」要填一个文件夹，不是客户端程序。\n\n"
                f"你填的是：\n{data}\n\n"
                "这看起来是紫鸟客户端的路径 —— 要把它挪到上面的「客户端路径」吗？\n"
                "（数据目录留空即可，会自动推导）",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if ans != QMessageBox.StandardButton.Yes:
                return
            if not client:
                self.ed_client.setText(data)
            self.ed_userdata.clear()

        c.version = self.ed_version.currentText()
        c.client_path = self.ed_client.text().strip()
        c.socket_port = int(self.sp_port.value())
        c.user_data_dir = self.ed_userdata.text().strip()
        c.company = self.ed_company.text().strip()
        c.username = self.ed_username.text().strip()
        c.password = self.ed_password.text()
        self.accept()


class StoreJsonDialog(QDialog):
    def __init__(self, store: dict[str, Any], parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"店铺原始数据 - {store.get('browserName', '')}")
        self.resize(780, 540)
        layout = QVBoxLayout(self)
        edit = QPlainTextEdit()
        edit.setReadOnly(True)
        edit.setFont(QFont("Menlo", 11))
        edit.setPlainText(json.dumps(store, ensure_ascii=False, indent=2))
        layout.addWidget(edit)
        box = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        box.rejected.connect(self.reject)
        layout.addWidget(box)


# ----------------------------------------------------------------------------
# 主窗口
# ----------------------------------------------------------------------------
class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("TikTok 订单 / 客户信息导出（紫鸟店铺浏览器）")
        self.resize(1360, 900)

        self.cfg: Any = None
        # 最近一次装载紫鸟配置的失败原因（自检报告 / 排错用）
        self.cfg_error: str = ""
        self.stores: list[dict[str, Any]] = []
        self._fields: list[tuple[str, str]] = []
        self.env_ready = False
        self.worker: Any = None
        # 「浏览…」上次选过的存放目录 —— 让四个输出框能连着收到同一个文件夹里
        self._last_out_dir: str = ""

        self.log_bus = LogBus()
        self.log_bus.message.connect(self._append_log)

        self._build_ui()
        self._build_log_dock()
        self._build_menu()

        self._load_config(initial=True)
        self._set_busy("idle")
        self.log_line("就绪。顺序：①启动环境 → ②获取店铺列表 → ③选店铺并导出订单 "
                      "→ ④抓客户信息 → ⑤导出商品评分 → ⑥抓差评客户。"
                      "⑤⑥ 各有一颗按钮，可分开单独执行。")

    # -- 界面构建 -----------------------------------------------------------
    def _build_ui(self) -> None:
        left = QWidget()
        lv = QVBoxLayout(left)
        lv.setContentsMargins(8, 8, 8, 8)
        lv.setSpacing(8)
        lv.addWidget(self._build_env_box())
        lv.addWidget(self._build_store_box(), 1)

        right = QWidget()
        rv = QVBoxLayout(right)
        rv.setContentsMargins(8, 8, 8, 8)
        rv.setSpacing(8)
        rv.addWidget(self._build_order_box())
        rv.addWidget(self._build_customer_box())
        rv.addWidget(self._build_rating_box())
        rv.addLayout(self._build_run_bar())
        rv.addStretch(0)

        top = QSplitter(Qt.Orientation.Horizontal)
        top.addWidget(left)
        top.addWidget(right)
        top.setStretchFactor(0, 4)
        top.setStretchFactor(1, 5)

        self.tabs = QTabWidget()
        self.tbl_orders = self._make_table(tc.ORDER_COLUMNS)
        self.tbl_customers = self._make_table(tc.CUSTOMER_COLUMNS)
        self.tbl_rating = self._make_table([])   # 列由导出文件的表头决定
        self.tbl_bad = self._make_table(BAD_CUSTOMER_COLUMNS)
        self.tabs.addTab(self.tbl_orders, "订单列表")
        self.tabs.addTab(self.tbl_customers, "客户信息")
        self.tabs.addTab(self.tbl_rating, "商品评分")
        self.tabs.addTab(self.tbl_bad, "客户差评")

        main = QSplitter(Qt.Orientation.Vertical)
        main.addWidget(top)
        main.addWidget(self.tabs)
        main.setStretchFactor(0, 3)
        main.setStretchFactor(1, 2)
        self.setCentralWidget(main)

    @staticmethod
    def _make_table(columns: list[str]) -> QTableWidget:
        t = QTableWidget(0, len(columns))
        t.setHorizontalHeaderLabels(columns)
        t.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        t.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        t.setAlternatingRowColors(True)
        t.verticalHeader().setDefaultSectionSize(24)
        h = t.horizontalHeader()
        h.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        h.setStretchLastSection(True)
        return t

    def _build_env_box(self) -> QGroupBox:
        box = QGroupBox("① 紫鸟环境")
        grid = QGridLayout(box)
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(6)

        grid.addWidget(QLabel("紫鸟项目："), 0, 0)
        self.ed_root = QLineEdit()
        self.ed_root.setPlaceholderText("ziniao_webdriver_demo 目录")
        self.ed_root.setMinimumWidth(140)
        grid.addWidget(self.ed_root, 0, 1, 1, 2)
        btn_auto = QPushButton("自动检测")
        btn_auto.clicked.connect(lambda: self._load_config(initial=False))
        grid.addWidget(btn_auto, 0, 3)

        grid.addWidget(QLabel("配置文件："), 1, 0)
        self.ed_cfgpath = QLineEdit()
        self.ed_cfgpath.setReadOnly(True)
        self.ed_cfgpath.setMinimumWidth(140)
        grid.addWidget(self.ed_cfgpath, 1, 1, 1, 3)

        # 客户端程序的实际位置。Windows 上没有固定安装目录，常出现「填错框 /
        # 没填」导致启动失败，这里直接显示解析结果，找不到就红字给填法。
        grid.addWidget(QLabel("客户端："), 2, 0)
        self.lbl_client = QLabel("—")
        self.lbl_client.setWordWrap(True)
        self.lbl_client.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        grid.addWidget(self.lbl_client, 2, 1, 1, 3)

        self.lbl_account = QLabel("账号：—")
        self.lbl_account.setWordWrap(True)
        grid.addWidget(self.lbl_account, 3, 0, 1, 4)

        bar = QHBoxLayout()
        self.btn_env = QPushButton("启动环境")
        self.btn_env.clicked.connect(self._start_env)
        bar.addWidget(self.btn_env)
        self.btn_list = QPushButton("获取店铺列表")
        self.btn_list.clicked.connect(self._load_stores)
        bar.addWidget(self.btn_list)
        self.btn_exit = QPushButton("关闭客户端")
        self.btn_exit.clicked.connect(self._exit_client)
        bar.addWidget(self.btn_exit)
        self.btn_cfg = QPushButton("编辑配置…")
        self.btn_cfg.clicked.connect(self._edit_config)
        bar.addWidget(self.btn_cfg)
        bar.addStretch(1)
        grid.addLayout(bar, 4, 0, 1, 4)

        # ★ 默认**勾选**（2026-10-16 用户定案）★
        # 已经开着紫鸟、再让程序启动一个，两个实例会串到同一个进程上：新实例的
        # `--run_type=web_driver` / `--port` 不生效（Windows 单实例锁），startBrowser
        # 返回的 debuggingPort 是旧实例的 → CDP 连错目标、页面操作失灵。所以默认
        # 「启动环境」先把已开启的紫鸟全部关掉，再起一个干净的实例。
        self.chk_kill = QCheckBox("强制重启客户端（先关掉已开启的紫鸟，再起新的）")
        self.chk_kill.setChecked(True)
        self.chk_kill.setToolTip(
            "默认勾选：启动环境前先杀掉**所有**紫鸟进程（含你自己手动打开的那个），\n"
            "再拉起一个干净的 web_driver 实例。\n"
            "为什么需要：紫鸟客户端进程唯一，旧实例还在时新实例起不来 / 启动参数被吞，\n"
            "startBrowser 返回的调试端口可能是旧实例的 → CDP 连错进程。\n"
            "取消勾选：检测到客户端已在运行（IPC 端口已就绪）就直接复用，不重启。")
        grid.addWidget(self.chk_kill, 5, 0, 1, 4)

        grid.setColumnStretch(1, 1)
        return box

    def _build_store_box(self) -> QGroupBox:
        box = QGroupBox("② 店铺列表（选中一行）")
        v = QVBoxLayout(box)
        v.setSpacing(6)

        bar = QHBoxLayout()
        self.ed_filter = QLineEdit()
        self.ed_filter.setPlaceholderText("按店铺名称 / ID / 平台 关键词过滤")
        self.ed_filter.textChanged.connect(self._fill_store_table)
        bar.addWidget(self.ed_filter, 1)
        self.lbl_store_count = QLabel("共 0 个")
        bar.addWidget(self.lbl_store_count)
        v.addLayout(bar)

        self.tbl_stores = QTableWidget(0, 1)
        self.tbl_stores.setHorizontalHeaderLabels(["店铺（先点「获取店铺列表」）"])
        self.tbl_stores.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.tbl_stores.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.tbl_stores.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.tbl_stores.setAlternatingRowColors(True)
        self.tbl_stores.verticalHeader().setDefaultSectionSize(24)
        self.tbl_stores.itemSelectionChanged.connect(self._on_store_selected)
        self.tbl_stores.itemDoubleClicked.connect(
            lambda _i: self._show_store_json(self._current_store()))
        v.addWidget(self.tbl_stores, 1)

        self.lbl_selected = QLabel("已选店铺：—")
        self.lbl_selected.setWordWrap(True)
        v.addWidget(self.lbl_selected)

        # 账号下挂了多个后台店铺时（如「秋图2」= Spain + United States），
        # 打开后台后**按 URL 直接导航**到目标店铺的后台首页
        # （https://seller.<区域>.tiktokshopglobalselling.com/homepage?shop_region=XX）。
        # 下拉候选是内置的市场名；也可以手填区域码（US/ES/GB…）或直接贴后台 URL。
        shop_row = QHBoxLayout()
        shop_row.addWidget(QLabel("后台店铺："))
        self.cb_shop = QComboBox()
        self.cb_shop.setEditable(True)
        for name in te.SHOP_CHOICES:
            self.cb_shop.addItem(name)
        if DEFAULT_TARGET_SHOP not in te.SHOP_CHOICES:
            self.cb_shop.addItem(DEFAULT_TARGET_SHOP)
        self.cb_shop.setCurrentText(DEFAULT_TARGET_SHOP)
        self.cb_shop.setMinimumWidth(180)
        self.cb_shop.setToolTip(
            "账号下要用的后台店铺（如 United States）。\n"
            "做法是按店铺对应的 shop_region **直接导航**到该店的后台首页，"
            "不模拟点击右上角切换器。\n"
            "已经在目标店铺时不动页面。也可直接填区域码（US / ES / GB）或后台 URL。")
        shop_row.addWidget(self.cb_shop, 1)
        v.addLayout(shop_row)
        return box

    def _build_order_box(self) -> QGroupBox:
        """③ 导出订单 —— 界面上只留「订单 Excel」这一个框，其余参数全走默认。

        导出范围固定「筛选出的订单」、格式固定 Excel（``DEFAULT_EXPORT_*``）；
        传输通道 / 接口主机 / 页面来源 / seller_id 都是④抓客户才用到的，走自动
        （``DEFAULT_TRANSPORT_MODE`` / ``tc.DEFAULT_API_BASE``），不再让用户填。
        """
        box = QGroupBox("③ 导出订单（点后台的「导出」按钮，直接下载导出文件）")
        grid = QGridLayout(box)
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(6)

        grid.addWidget(QLabel("订单 Excel："), 0, 0)
        self.ed_order_out = QLineEdit(str(DEFAULT_ORDER_OUT))
        self.ed_order_out.setMinimumWidth(120)
        self.ed_order_out.setToolTip(
            "文件名默认按店铺命名（选中店铺后自动填入）。\n"
            "点「浏览…」在目录里**选中**一个 xlsx / csv 文件（导出就写它，已存在则覆盖）；\n"
            "想指定一个还不存在的路径，直接在这个框里敲完整路径。")
        grid.addWidget(self.ed_order_out, 0, 1, 1, 3)
        btn_out = QPushButton("浏览…")
        btn_out.setToolTip("浏览选择订单要写进的 xlsx / csv 文件（已存在则覆盖）")
        btn_out.clicked.connect(
            lambda: self._pick_output_file(self.ed_order_out, DEFAULT_ORDER_OUT))
        grid.addWidget(btn_out, 0, 4)

        bar = QHBoxLayout()
        self.btn_orders = QPushButton("导出订单")
        self.btn_orders.clicked.connect(lambda: self._start_pipeline(do_orders=True,
                                                                    do_customers=False))
        bar.addWidget(self.btn_orders)
        bar.addStretch(1)
        grid.addLayout(bar, 1, 0, 1, 5)

        grid.setColumnStretch(1, 1)
        return box

    def _build_customer_box(self) -> QGroupBox:
        box = QGroupBox("④ 读订单 Excel 抓客户信息（order/get）")
        grid = QGridLayout(box)
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(6)

        grid.addWidget(QLabel("订单 Excel："), 0, 0)
        self.ed_cust_in = QLineEdit(str(DEFAULT_ORDER_OUT))
        self.ed_cust_in.setMinimumWidth(120)
        self.ed_cust_in.textChanged.connect(self._refresh_columns)
        grid.addWidget(self.ed_cust_in, 0, 1, 1, 2)
        btn_in = QPushButton("浏览…")
        btn_in.clicked.connect(self._pick_customer_input)
        grid.addWidget(btn_in, 0, 3)
        btn_sync = QPushButton("用导出的订单")
        btn_sync.setToolTip("把上面「订单 Excel」的输出路径同步到这里")
        btn_sync.clicked.connect(
            lambda: self.ed_cust_in.setText(self.ed_order_out.text().strip()))
        grid.addWidget(btn_sync, 0, 4)

        grid.addWidget(QLabel("工作表："), 1, 0)
        self.cb_sheet = QComboBox()
        self.cb_sheet.setMinimumWidth(100)
        self.cb_sheet.currentTextChanged.connect(self._refresh_columns)
        grid.addWidget(self.cb_sheet, 1, 1)
        grid.addWidget(QLabel("Order ID 列："), 1, 2)
        self.cb_column = QComboBox()
        self.cb_column.setEditable(True)
        self.cb_column.setMinimumWidth(100)
        self.cb_column.addItem("Order ID")
        grid.addWidget(self.cb_column, 1, 3, 1, 2)

        grid.addWidget(QLabel("客户 Excel："), 2, 0)
        self.ed_cust_out = QLineEdit(str(DEFAULT_CUSTOMER_OUT))
        self.ed_cust_out.setMinimumWidth(120)
        self.ed_cust_out.setToolTip(
            "文件名默认按店铺命名（选中店铺后自动填入）。\n"
            "点「浏览…」选择**文件夹**，文件就存进这个文件夹（文件名不变）；\n"
            "想连文件名一起改，就直接敲完整路径。")
        grid.addWidget(self.ed_cust_out, 2, 1, 1, 3)
        btn_out = QPushButton("浏览…")
        btn_out.setToolTip("选择存放的文件夹（文件名保持 <店铺名>_客户.xlsx）")
        btn_out.clicked.connect(
            lambda: self._pick_output_dir(self.ed_cust_out, DEFAULT_CUSTOMER_OUT))
        grid.addWidget(btn_out, 2, 4)

        grid.addWidget(QLabel("请求间隔(秒)："), 3, 0)
        self.sp_interval = QDoubleSpinBox()
        self.sp_interval.setRange(MIN_CUSTOMER_INTERVAL, 60.0)
        self.sp_interval.setSingleStep(0.5)
        self.sp_interval.setValue(MIN_CUSTOMER_INTERVAL)
        self.sp_interval.setToolTip(f"每条之间的等待，防限流；最小 {MIN_CUSTOMER_INTERVAL}s")
        self.sp_interval.setMinimumWidth(100)
        grid.addWidget(self.sp_interval, 3, 1)
        grid.addWidget(QLabel("失败重试："), 3, 2)
        self.sp_retry = QSpinBox()
        self.sp_retry.setRange(0, 10)
        self.sp_retry.setValue(2)
        self.sp_retry.setMinimumWidth(100)
        grid.addWidget(self.sp_retry, 3, 3)

        self.chk_skip_ip = QCheckBox("跳过 IP 检测")
        grid.addWidget(self.chk_skip_ip, 3, 4)

        bar = QHBoxLayout()
        self.btn_customers = QPushButton("抓取客户信息")
        self.btn_customers.clicked.connect(lambda: self._start_pipeline(do_orders=False,
                                                                       do_customers=True))
        bar.addWidget(self.btn_customers)
        bar.addStretch(1)
        grid.addLayout(bar, 4, 0, 1, 5)

        grid.setColumnStretch(1, 1)
        return box

    def _build_rating_box(self) -> QGroupBox:
        """⑤ 商品评分 —— 走评分页自己的「导出」按钮，和订单那套是两条路。

        评分页**没有**订单页那种「导出范围 / 格式」面板，只有一颗导出按钮；
        点完可能立刻下载，也可能进导出历史异步生成，两种都实现了（见
        ``tiktok_export.RatingExporter``）。
        """
        box = QGroupBox("⑤ 导出商品评分（/product/rating 页面的「导出」按钮）")
        grid = QGridLayout(box)
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(6)

        grid.addWidget(QLabel("评分 Excel："), 0, 0)
        self.ed_rating_out = QLineEdit(str(DEFAULT_RATING_OUT))
        self.ed_rating_out.setMinimumWidth(120)
        self.ed_rating_out.setToolTip(
            "文件名默认按店铺命名（选中店铺后自动填入）。\n"
            "点「浏览…」在目录里**选中**一个 xlsx / csv 文件（导出就写它，已存在则覆盖）；\n"
            "想指定一个还不存在的路径，直接在这个框里敲完整路径。")
        grid.addWidget(self.ed_rating_out, 0, 1, 1, 3)
        btn_out = QPushButton("浏览…")
        btn_out.setToolTip("浏览选择评分要写进的 xlsx / csv 文件（已存在则覆盖）")
        btn_out.clicked.connect(
            lambda: self._pick_output_file(self.ed_rating_out, DEFAULT_RATING_OUT))
        grid.addWidget(btn_out, 0, 4)

        # ⑥ 差评客户：优先用评分文件自带的订单号；没有才靠昵称回订单表
        grid.addWidget(QLabel("客户差评 Excel："), 1, 0)
        self.ed_bad_out = QLineEdit(str(DEFAULT_BAD_CUSTOMER_OUT))
        self.ed_bad_out.setMinimumWidth(120)
        self.ed_bad_out.setToolTip(
            "⑥ 抓取 1~3 星评价对应客户的收货信息，文件名默认按店铺命名。\n"
            "点「浏览…」选择**文件夹**，文件就存进这个文件夹（文件名不变）；\n"
            "想连文件名一起改，就直接敲完整路径。\n"
            "评分文件自带订单号列（真机导出的 CSV 就有）时不需要订单表；\n"
            "只有老版「仅昵称」的评分文件才需要③的订单表做关联。")
        grid.addWidget(self.ed_bad_out, 1, 1, 1, 3)
        btn_bad = QPushButton("浏览…")
        btn_bad.setToolTip("选择存放的文件夹（文件名保持 <店铺名>_客户差评信息.xlsx）")
        btn_bad.clicked.connect(
            lambda: self._pick_output_dir(self.ed_bad_out, DEFAULT_BAD_CUSTOMER_OUT))
        grid.addWidget(btn_bad, 1, 4)

        # ⑤ 和 ⑥ 是两颗**独立**的按钮：可以只导出评分，也可以拿一份已经导好的
        # 评分表直接抓差评客户（不用重新开评分页导一次，省一两分钟）。
        bar = QHBoxLayout()
        self.btn_rating = QPushButton("导出商品评分")
        self.btn_rating.clicked.connect(lambda: self._start_pipeline(
            do_orders=False, do_customers=False, do_rating=True))
        self.btn_rating.setToolTip("走评分页的「导出」按钮导出，筛出 1~3 星后覆盖保存；"
                                   "完整版留在 _评分原始.xlsx")
        bar.addWidget(self.btn_rating)

        self.btn_bad = QPushButton("抓取差评客户（1~3 星）")
        self.btn_bad.clicked.connect(lambda: self._start_pipeline(
            do_orders=False, do_customers=False, do_rating=False, do_bad=True))
        self.btn_bad.setToolTip(
            "读「评分 Excel」框里那份文件 → 筛 1~3 星覆盖保存 → 按昵称回订单表换出\n"
            "订单号 → order/get 抓客户 → 写「客户差评 Excel」。\n"
            "评分还没导出（或里面没有 1~3 星）时会照实提示，不会写出空表。")
        bar.addWidget(self.btn_bad)
        bar.addStretch(1)
        grid.addLayout(bar, 2, 0, 1, 5)

        grid.setColumnStretch(1, 1)
        return box

    def _build_run_bar(self) -> QHBoxLayout:
        bar = QHBoxLayout()
        self.btn_run_all = QPushButton(
            "▶ 一键完成：导出订单 → 抓客户信息 → 导出商品评分 → 抓差评客户")
        self.btn_run_all.setMinimumHeight(36)
        self.btn_run_all.clicked.connect(lambda: self._start_pipeline(
            do_orders=True, do_customers=True,
            do_rating=self.chk_rating.isChecked(),
            do_bad=self.chk_bad.isChecked()))
        bar.addWidget(self.btn_run_all, 2)

        self.btn_stop = QPushButton("■ 停止")
        self.btn_stop.setMinimumHeight(36)
        self.btn_stop.clicked.connect(self._stop)
        bar.addWidget(self.btn_stop)

        # 商品评分要额外开一次页面、等一次导出，不是每次都需要 —— 但默认跟着跑
        self.chk_rating = QCheckBox("含商品评分")
        self.chk_rating.setChecked(True)
        self.chk_rating.setToolTip(
            "一键完成时是否也导出商品评分（会多花一两分钟）。\n"
            "只想跑订单 + 客户就取消勾选，或单独用 ⑤ 区那个按钮。")
        bar.addWidget(self.chk_rating)

        # ⑥ 依赖⑤的评分文件：一键完成时跟着勾选即可；评分不跑就自动跳过
        self.chk_bad = QCheckBox("含差评客户")
        self.chk_bad.setChecked(True)
        self.chk_bad.setToolTip("⑤导出后筛出 1~3 星 → 覆盖保存评分表 → 再抓这些订单的客户")
        bar.addWidget(self.chk_bad)

        self.lbl_stat = QLabel("等待开始")
        bar.addWidget(self.lbl_stat, 1)

        self.bar_progress = QProgressBar()
        self.bar_progress.setMinimumWidth(220)
        self.bar_progress.setRange(0, 100)
        self.bar_progress.setValue(0)
        bar.addWidget(self.bar_progress, 2)
        return bar

    def _build_log_dock(self) -> None:
        dock = QDockWidget("日志", self)
        dock.setAllowedAreas(Qt.DockWidgetArea.BottomDockWidgetArea)
        self.log_box = QPlainTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setFont(QFont("Menlo", 11))
        self.log_box.setMaximumBlockCount(4000)
        dock.setWidget(self.log_box)
        self.addDockWidget(Qt.DockWidgetArea.BottomDockWidgetArea, dock)
        self.resizeDocks([dock], [220], Qt.Orientation.Vertical)
        self._start_log_file()

    def _build_menu(self) -> None:
        m = self.menuBar().addMenu("工具")
        act_cfg = QAction("紫鸟配置…", self)
        act_cfg.triggered.connect(self._edit_config)
        m.addAction(act_cfg)
        act_json = QAction("查看选中店铺原始数据", self)
        act_json.triggered.connect(lambda: self._show_store_json(self._current_store()))
        m.addAction(act_json)
        act_clear = QAction("清空日志", self)
        act_clear.triggered.connect(self.log_box.clear)
        m.addAction(act_clear)
        act_open = QAction("打开输出目录", self)
        act_open.triggered.connect(self._open_output_dir)
        m.addAction(act_open)
        act_log = QAction("打开日志文件", self)
        act_log.triggered.connect(self._open_log_file)
        m.addAction(act_log)
        act_log_save = QAction("日志另存为…", self)
        act_log_save.triggered.connect(self._save_log_as)
        m.addAction(act_log_save)
        act_check = QAction("环境自检", self)
        act_check.triggered.connect(self._self_check)
        m.addAction(act_check)

    # -- 日志 ---------------------------------------------------------------
    def log_line(self, msg: str) -> None:
        """线程安全的日志入口（唯一写日志的地方）。"""
        app = QApplication.instance()
        if app is None or QThread.currentThread() is app.thread():
            self._append_log(msg)
        else:
            self.log_bus.message.emit(str(msg))

    def _pick_log_path(self) -> Path | None:
        """日志文件放哪：优先程序（exe）旁边，不可写就退到用户主目录。"""
        for base in (zb.app_dir(), Path.home()):
            try:
                base.mkdir(parents=True, exist_ok=True)
                probe = base / ".__log_probe"
                probe.write_text("", encoding="utf-8")
                probe.unlink()
                return base / "运行日志.txt"
            except Exception:
                continue
        return None

    def _start_log_file(self) -> None:
        """把日志**同时**写一份到文件。

        为什么必须落盘：exe 是窗口程序，日志只在界面里（而且只留最近 4000 行）。
        真机出问题时用户往往只能截图，而「页面停在哪个区域站」「每种写法各返回
        什么 code」这些关键行一屏根本截不全 —— 2026-10-14 排查 98001008 就卡在
        这里。现在每次运行都写「运行日志.txt」，出问题一键把整份现场发出来即可。

        只在交互运行时写：``--check`` 是打包后跑的自检，没人看日志，写了只会
        在发布目录里留一个**含账号名与本机路径**的文件（差点误提交，见 .gitignore）。
        """
        if "--check" in sys.argv:
            self._log_path = None
            self._log_handle = None
            return
        self._log_path = self._pick_log_path()
        self._log_handle = None
        if self._log_path is None:
            return
        try:
            with open(self._log_path, "w", encoding="utf-8", errors="replace") as fh:
                fh.write("===== tiktok商家 运行日志 =====\n")
                fh.write(f"时间     : {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
                fh.write(f"平台     : {sys.platform}\n")
                fh.write(f"Python   : {sys.version.split()[0]}\n")
                fh.write(f"冻结打包 : {bool(getattr(sys, 'frozen', False))}\n")
                fh.write(f"程序目录 : {zb.app_dir()}\n\n")
            self._log_handle = open(self._log_path, "a", encoding="utf-8",
                                    errors="replace")
        except Exception:
            self._log_handle = None
        if self._log_handle is not None:
            self.log_line(f"日志文件：{self._log_path}")
            self.log_line("（出问题时「工具 → 打开日志文件」，把整份发来即可）")

    def _open_log_file(self) -> None:
        p = getattr(self, "_log_path", None)
        if not p:
            QMessageBox.information(self, "日志", "这份运行没能创建日志文件。")
            return
        reveal_path(p)

    def _save_log_as(self) -> None:
        out, _ = QFileDialog.getSaveFileName(
            self, "日志另存为", str(self._log_path or (Path.home() / "运行日志.txt")),
            "文本文件 (*.txt)")
        if not out:
            return
        try:
            Path(out).write_text(self.log_box.toPlainText(), encoding="utf-8")
            self.log_line(f"日志已另存为：{out}")
        except Exception as exc:
            QMessageBox.critical(self, "保存失败", f"写入失败：\n{exc}")

    def _append_log(self, msg: str) -> None:
        if not hasattr(self, "log_box"):
            return
        self.log_box.appendPlainText(str(msg))
        fh = getattr(self, "_log_handle", None)
        if fh is not None:
            try:
                fh.write(str(msg) + "\n")
                fh.flush()          # 崩溃/强杀时也别丢最后几行
            except Exception:
                self._log_handle = None

    # -- 状态 ---------------------------------------------------------------
    def _set_busy(self, state: str) -> None:
        busy = state in ("env", "list", "pipe")
        for w in (self.btn_env, self.btn_list, self.btn_exit, self.btn_cfg,
                  self.btn_orders, self.btn_customers, self.btn_rating,
                  self.btn_bad, self.btn_run_all, self.ed_root, self.chk_kill,
                  self.chk_rating, self.chk_bad, self.cb_shop):
            w.setEnabled(not busy)
        self.btn_stop.setEnabled(state == "pipe")
        if state == "env":
            self.lbl_stat.setText("正在准备环境…")
        elif state == "list":
            self.lbl_stat.setText("正在获取店铺列表…")
        elif state == "pipe":
            self.lbl_stat.setText("流水线运行中…")
        else:
            self.lbl_stat.setText("就绪")

    # -- 配置 ---------------------------------------------------------------
    def _load_config(self, initial: bool = False) -> None:
        root_text = self.ed_root.text().strip()
        try:
            root = zb.locate_ziniao_root(root_text or None)
        except Exception as exc:
            # 冻结成 exe 后 DEFAULT_ZINIAO_ROOT 指向临时解包目录，没意义；
            # 用「当前形态下最可能的位置」给用户提示。
            self.ed_root.setText(str(zb.preferred_ziniao_root()))
            self.lbl_account.setText("账号：—（未找到紫鸟项目）")
            self.ed_cfgpath.setText("—")
            self.cfg_error = f"{type(exc).__name__}: {exc}"
            self.log_line(f"未找到紫鸟项目：{exc}")
            return
        try:
            self.cfg = zb.load_config(root)
        except Exception as exc:
            self.cfg_error = f"{type(exc).__name__}: {exc}"
            self.log_line(f"读取紫鸟配置失败：{exc}")
            return
        self.cfg_error = ""
        self.ed_root.setText(str(root))
        self.ed_cfgpath.setText(str(root / "config" / "config.yaml"))
        missing = self.cfg.missing()
        self.lbl_account.setText(
            f"账号：{self.cfg.company or '(未填)'} / {self.cfg.username or '(未填)'}"
            f"　端口：{self.cfg.socket_port}"
            + (f"　[!] 缺：{', '.join(missing)}" if missing else "")
        )
        self._refresh_client_label()
        if not initial:
            self.log_line(f"已加载紫鸟项目：{root}")

    def _refresh_client_label(self) -> None:
        """把解析到的客户端 exe 显示在 ① 区，找不到就红字提示怎么填。"""
        cfg = self.cfg
        raw = (getattr(cfg, "client_path", "") or "") if cfg else ""
        try:
            exe = zb.find_client_executable(raw, root=self.ed_root.text().strip() or None)
        except Exception as exc:
            self.lbl_client.setText(f"—（解析失败：{exc}）")
            self.lbl_client.setStyleSheet("color:#c62828;")
            return
        if exe:
            prefix = "自动找到：" if not raw.strip() else ""
            self.lbl_client.setText(prefix + exe)
            self.lbl_client.setToolTip(exe)
            self.lbl_client.setStyleSheet("")
            return
        # 找不到：把平台相关的完整诊断（填法示例 / 已找过的位置）塞进 tooltip，
        # 界面上只留一行短提示，别把 ① 区撑爆。
        hint = zb.client_hint(cfg) if cfg is not None else ""
        self.lbl_client.setText(
            "没找到客户端程序 —— 点「编辑配置…」→「自动检测」，"
            "或把「客户端路径」留空（悬停看具体填法）")
        self.lbl_client.setToolTip(hint or "「客户端路径」留空即可自动查找；Windows 也可直接填 "
                                          "exe 全路径（如 D:\\紫鸟\\ziniao\\ziniao.exe）")
        self.lbl_client.setStyleSheet("color:#c62828;")

    def _edit_config(self) -> None:
        if self.cfg is None:
            QMessageBox.warning(self, "未加载", "还没找到紫鸟项目，无法编辑配置。")
            return
        dlg = ZiniaoConfigDialog(self.cfg, self, root=self.ed_root.text().strip() or None)
        if dlg.exec():
            try:
                path = zb.save_config(self.cfg, self.ed_root.text().strip() or None)
                self.log_line(f"紫鸟配置已保存：{path}")
            except Exception as exc:
                QMessageBox.critical(self, "保存失败", f"写入 config.yaml 失败：\n{exc}")
                return
            self._load_config()
            self.env_ready = False
            # 保存完立刻复检一次客户端能不能解析出来 —— 这一栏填错的概率很高
            # （真机上就有人把 exe 填进了隔壁的「数据目录」），早点说比等启动失败好。
            self._warn_if_client_unresolved()

    def _warn_if_client_unresolved(self) -> None:
        """配置刚存完，若仍解析不到客户端就当场提示（不再等到「启动环境」才炸）。"""
        if self.cfg is None:
            return
        root_text = self.ed_root.text().strip() or None
        try:
            exe = zb.find_client_executable(self.cfg.client_path, root=root_text, force=True)
        except Exception as exc:
            QMessageBox.warning(self, "客户端解析失败", f"解析客户端路径时出错：\n{exc}")
            return
        self._refresh_client_label()
        if exe:
            self.log_line(f"客户端程序：{exe}")
            return
        QMessageBox.warning(self, "还没找到紫鸟客户端", zb.client_hint(self.cfg))

    # -- ① 环境 -------------------------------------------------------------
    def _start_env(self) -> None:
        self._load_config()
        if self.cfg is None:
            QMessageBox.warning(self, "未就绪", "找不到紫鸟项目，请先设置「紫鸟项目」路径。")
            return
        missing = self.cfg.missing()
        if missing:
            QMessageBox.warning(self, "配置不全", f"config.yaml 缺少：{', '.join(missing)}")
            return
        # Windows 客户端没有固定安装位置，启动前先强制探一次并缓存：
        # 之后 Flow 会先杀掉客户端进程，那时再按「运行中的进程」就找不到了。
        root_text = self.ed_root.text().strip() or None
        try:
            exe = zb.find_client_executable(self.cfg.client_path, root=root_text, force=True)
        except Exception:
            exe = ""
        self._refresh_client_label()
        if not exe:
            QMessageBox.warning(
                self, "找不到紫鸟客户端",
                "没找到紫鸟浏览器的客户端程序，无法启动环境。\n\n"
                + zb.client_hint(self.cfg))
            return

        self._set_busy("env")
        self.log_line("=" * 60)
        self.worker = EnvThread(self.cfg, self.ed_root.text().strip(),
                                self.chk_kill.isChecked(), self)
        self.worker.log.connect(self.log_line)
        self.worker.done.connect(self._on_env_done)
        self.worker.start()

    def _on_env_done(self, ok: bool, message: str) -> None:
        self.env_ready = ok
        self._set_busy("idle")
        self.log_line(f"环境：{message}")
        if not ok:
            QMessageBox.critical(self, "启动环境失败", message)

    def _exit_client(self) -> None:
        try:
            client = zb.make_client(self.cfg, log=self.log_line,
                                    root=self.ed_root.text().strip() or None)
            client.get_exit()
            self.env_ready = False
            self.log_line("已请求关闭紫鸟客户端")
        except Exception as exc:
            self.log_line(f"关闭客户端失败：{exc}")

    # -- ② 店铺 -------------------------------------------------------------
    def _load_stores(self) -> None:
        self._load_config()
        if self.cfg is None:
            QMessageBox.warning(self, "未就绪", "找不到紫鸟项目，请先设置「紫鸟项目」路径。")
            return
        self._set_busy("list")
        self.worker = ListThread(self.cfg, self.ed_root.text().strip(), self)
        self.worker.log.connect(self.log_line)
        self.worker.done.connect(self._on_stores)
        self.worker.start()

    def _on_stores(self, ok: bool, stores: object, message: str) -> None:
        self._set_busy("idle")
        self.log_line(message)
        if not ok:
            QMessageBox.critical(self, "获取店铺列表失败", message)
            return
        self.stores = list(stores or [])  # type: ignore[arg-type]
        self._compute_store_fields()
        self._fill_store_table()
        if not self.stores:
            self.log_line("店铺列表为空；若刚启动环境，请稍等几秒后重试。")

    def _compute_store_fields(self) -> None:
        fields: list[tuple[str, str]] = []
        for key, label in PREFERRED_FIELDS:
            if any(key in s for s in self.stores):
                fields.append((key, label))
        if not fields:
            fields = [("browserName", "店铺名称"), ("browserOauth", "店铺 ID")]
        self._fields = fields
        self.tbl_stores.setColumnCount(len(fields))
        self.tbl_stores.setHorizontalHeaderLabels([label for _, label in fields])
        header = self.tbl_stores.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        for i in range(1, len(fields)):
            header.setSectionResizeMode(i, QHeaderView.ResizeMode.Interactive)

    def _filtered_stores(self) -> list[dict[str, Any]]:
        kw = self.ed_filter.text().strip().lower()
        if not kw:
            return list(self.stores)
        hit = []
        for s in self.stores:
            blob = " ".join(_cell(s.get(k, "")) for k, _ in self._fields).lower()
            if kw in blob:
                hit.append(s)
        return hit

    def _fill_store_table(self) -> None:
        if not hasattr(self, "tbl_stores"):
            return
        rows = self._filtered_stores()
        self.tbl_stores.setRowCount(len(rows))
        for r, store in enumerate(rows):
            for c, (key, _label) in enumerate(self._fields):
                item = QTableWidgetItem(_cell(store.get(key, "")))
                if c == 0:
                    item.setData(Qt.ItemDataRole.UserRole, r)
                self.tbl_stores.setItem(r, c, item)
        self.lbl_store_count.setText(
            f"共 {len(rows)} 个" + (f" / 全部 {len(self.stores)}" if len(rows) != len(self.stores)
                                   else ""))

    def _current_store(self) -> dict[str, Any] | None:
        item = self.tbl_stores.currentItem()
        if item is None:
            return None
        idx = item.row()
        rows = self._filtered_stores()
        if 0 <= idx < len(rows):
            return rows[idx]
        return None

    def _on_store_selected(self) -> None:
        store = self._current_store()
        if store is None:
            self.lbl_selected.setText("已选店铺：—")
            return
        self.lbl_selected.setText(
            f"已选店铺：{store.get('browserName') or '?'}"
            f"　ID={store.get('browserOauth') or store.get('browserId') or '?'}"
            f"　siteId={store.get('siteId', '?')}")
        self._apply_store_output_names(store)

    def _apply_store_output_names(self, store: dict[str, Any] | None) -> None:
        """选中店铺后，把输出文件名自动改成「<店铺名>_订单.xlsx」「<店铺名>_客户.xlsx」。

        只在用户**没有手动改过**这些框时才覆盖 —— 即框里仍是默认名、或仍是
        上一次自动填进去的名字。用户「另存为」过、或自己敲过路径，就原样保留。
        顺带把④的「订单 Excel」输入框也指向该店的订单文件（若它原来也是默认值）。
        """
        if not store:
            return
        order_path, cust_path, rating_path, bad_path = store_output_paths(store)
        for attr, default, target in (
            ("ed_order_out", DEFAULT_ORDER_OUT, order_path),
            ("ed_cust_out", DEFAULT_CUSTOMER_OUT, cust_path),
            ("ed_cust_in", DEFAULT_ORDER_OUT, order_path),
            ("ed_rating_out", DEFAULT_RATING_OUT, rating_path),
            ("ed_bad_out", DEFAULT_BAD_CUSTOMER_OUT, bad_path),
        ):
            line = getattr(self, attr, None)
            if line is None:
                continue
            cur = line.text().strip()
            last_auto = getattr(self, f"_auto_{attr}", None)
            if cur in ("", str(default)) or (last_auto and cur == last_auto):
                line.setText(str(target))
                setattr(self, f"_auto_{attr}", str(target))

    def _show_store_json(self, store: dict[str, Any] | None) -> None:
        if not store:
            QMessageBox.information(self, "未选择", "请先在店铺列表里选中一行。")
            return
        StoreJsonDialog(store, self).exec()

    # -- ③④⑤⑥ 文件选择 ------------------------------------------------------
    def _pick_output_file(self, line: QLineEdit, default: Path) -> None:
        """「浏览…」：**浏览并选中一个文件**（③订单 / ⑤评分用）。

        这两个框存的是「导出文件本身」，用户要的是**在目录树里点中那个
        `.xlsx` / `.csv`** —— 所以弹的是「打开」式的浏览框，选中即填进框；
        导出时就写这个文件（已存在则覆盖）。

        不是「另存为」对话框：那个要用户现起一个文件名，还得应付覆盖确认。

        起始位置沿用**上次选过的目录**，并尽量把光标落在框里当前那个文件上。
        想指定一个还不存在的文件（新目录 / 新名字），直接在框里敲完整路径 ——
        框本来就可编辑。
        """
        cur = line.text().strip()
        start = dir_of(cur or default)
        if self._last_out_dir and Path(self._last_out_dir).is_dir():
            start = Path(self._last_out_dir)
        initial = str(Path(start) / Path(cur).name) if cur else str(start)
        # 过滤器只放**我们真能写**的两种：`.xls`（老格式）和 `.xlsm`（带宏）都写不了，
        # 放进来只会让用户选中一个然后拿到个「格式不对」的文件。
        chosen, _sel = QFileDialog.getOpenFileName(
            self, "选择文件", initial,
            "Excel / CSV 文件 (*.xlsx *.csv);;所有文件 (*)")
        if not chosen:
            return
        p = Path(chosen)
        self._last_out_dir = str(p.parent)
        line.setText(str(p))
        self.log_line(f"输出位置：{p}")

    def _pick_output_dir(self, line: QLineEdit, default: Path) -> None:
        """「浏览…」：给**目录**型输出框挑存放目录（④客户 / ⑥差评用这个）。

        ④⑥ 的文件名一律按店铺名自动生成（`<店铺名>_客户.xlsx` /
        `<店铺名>_客户差评信息.xlsx`），用户要决定的只是「放进哪个文件夹」——
        所以只弹目录选择框：**文件名原样保留**，只换它前面的目录部分。
        想连文件名一起改，直接在输入框里敲完整路径即可（框本来就可编辑）。

        ③订单 / ⑤评分不一样：那两个存的是「导出文件本身」，用
        ``_pick_output_file``（弹保存对话框，能改名、能换 csv）。

        起始目录优先用**上次选过的那个**，所以想把几个文件都收进同一个文件夹时，
        第一下选好，后面几个点开就是对的目录，直接确定。
        """
        cur = Path(line.text().strip() or str(default))
        # 注意：`Path("")` 是 `.`、`is_dir()` 为真 —— 空串必须单独判，否则起始目录
        # 会变成一个点，Qt 拿不到有效位置。
        start = dir_of(line.text().strip() or default)
        if self._last_out_dir and Path(self._last_out_dir).is_dir():
            start = Path(self._last_out_dir)
        chosen = QFileDialog.getExistingDirectory(self, "选择存放目录", str(start))
        if not chosen:
            return
        self._last_out_dir = chosen
        new_path = Path(chosen) / cur.name
        line.setText(str(new_path))
        self.log_line(f"输出位置：{cur.name} → {new_path}")

    def _pick_customer_input(self) -> None:
        start = self.ed_cust_in.text().strip() or str(BASE_DIR)
        path, _ = QFileDialog.getOpenFileName(
            self, "选择订单 Excel", start, "Excel 文件 (*.xlsx *.xlsm *.xls)")
        if path:
            self.ed_cust_in.setText(path)

    def _refresh_columns(self) -> None:
        path = Path(self.ed_cust_in.text().strip())
        if not path.exists():
            return
        try:
            sheets = tc.list_sheets(path)
        except Exception as exc:
            self.log_line(f"读取工作表失败：{exc}")
            return
        self.cb_sheet.blockSignals(True)
        self.cb_sheet.clear()
        self.cb_sheet.addItems(sheets)
        self.cb_sheet.blockSignals(False)

        sheet = self.cb_sheet.currentText() or None
        try:
            cols = tc.read_header(path, sheet, row=1)
        except Exception as exc:
            self.log_line(f"读取列名失败：{exc}")
            return
        if not cols:
            return
        current = self.cb_column.currentText()
        self.cb_column.blockSignals(True)
        self.cb_column.clear()
        self.cb_column.addItems(cols)
        self.cb_column.setCurrentText(
            current if current in cols else ("Order ID" if "Order ID" in cols else cols[0]))
        self.cb_column.blockSignals(False)

    # -- ③④ 流水线 ----------------------------------------------------------
    def _start_pipeline(self, do_orders: bool, do_customers: bool,
                        do_rating: bool = False, do_bad: bool = False) -> None:
        self._load_config()
        if self.cfg is None:
            QMessageBox.warning(self, "未就绪", "找不到紫鸟项目，请先设置「紫鸟项目」路径。")
            return
        store = self._current_store()
        if store is None:
            QMessageBox.warning(self, "未选店铺", "请先在「店铺列表」里选中一个店铺。")
            return

        # 兜底：保证输出文件名带上店铺名（用户中途清空过框 / 没触发过选店信号时）
        self._apply_store_output_names(store)

        order_out = Path(self.ed_order_out.text().strip() or str(DEFAULT_ORDER_OUT))
        customer_out = Path(self.ed_cust_out.text().strip() or str(DEFAULT_CUSTOMER_OUT))
        rating_out = Path(self.ed_rating_out.text().strip() or str(DEFAULT_RATING_OUT))
        bad_out = Path(self.ed_bad_out.text().strip() or str(DEFAULT_BAD_CUSTOMER_OUT))
        cust_in_text = self.ed_cust_in.text().strip()
        cust_in = Path(cust_in_text) if cust_in_text else None

        if do_orders and order_out.suffix.lower() not in (".xlsx", ".xlsm"):
            order_out = order_out.with_suffix(".xlsx")
            self.ed_order_out.setText(str(order_out))
        if do_rating and rating_out.suffix.lower() not in (".xlsx", ".xlsm"):
            rating_out = rating_out.with_suffix(".xlsx")
            self.ed_rating_out.setText(str(rating_out))
        if do_bad and bad_out.suffix.lower() not in (".xlsx", ".xlsm"):
            bad_out = bad_out.with_suffix(".xlsx")
            self.ed_bad_out.setText(str(bad_out))
        if do_bad and not do_rating:
            # 单独跑⑥：必须有⑤导出的评分表（筛过的 out 或完整版 _评分原始，CSV 也认）
            if not any(p.exists() for p in rating_file_candidates(rating_out)):
                QMessageBox.warning(
                    self, "找不到评分表",
                    f"单独抓差评客户需要先有商品评分表：\n{rating_out}\n\n"
                    "请先点「导出商品评分」，或在「评分 Excel」框里指向已有文件。")
                return
            # ★ 订单表**只在「评分文件里没有订单号列」时才必需**。
            #   真机⑤导出的是带 order_id 的 CSV → ⑥ 拿订单号直接 order/get，
            #   压根不碰订单表。旧版在这里无条件拦「找不到订单表」，用户明明
            #   有评分表却被挡住（而且弹窗指着订单 Excel，看着像⑥读错了文件）。
            if not rating_order_column(rating_out) \
                    and not (Path(cust_in_text).exists() if cust_in_text else False) \
                    and not order_out.exists():
                QMessageBox.warning(
                    self, "找不到订单表",
                    f"评分文件里没有订单号列，只能靠买家昵称回订单表换出 Order ID，\n"
                    f"但订单表不存在：\n{cust_in_text or order_out}\n\n"
                    "请先执行「导出订单」，或在「订单 Excel」框里指向已有文件。")
                return
        if do_customers:
            if customer_out.suffix.lower() not in (".xlsx", ".xlsm"):
                customer_out = customer_out.with_suffix(".xlsx")
                self.ed_cust_out.setText(str(customer_out))
            if not do_orders:
                if cust_in is None or not cust_in.exists():
                    QMessageBox.warning(self, "路径无效",
                                        f"找不到订单 Excel：\n{cust_in}\n"
                                        "可先执行「导出订单」，或点「浏览…」选择文件。")
                    return

        opts = {
            "do_orders": do_orders,
            "do_customers": do_customers,
            "do_rating": do_rating,
            "do_bad": do_bad,     # ⑥ 可以单独跑（读已有的评分 Excel）
            "order_out": str(order_out),
            "customer_out": str(customer_out),
            "rating_out": str(rating_out),
            "bad_out": str(bad_out),
            "customer_input": str(cust_in) if cust_in else "",
            "customer_sheet": self.cb_sheet.currentText().strip(),
            "order_column": self.cb_column.currentText().strip() or "Order ID",
            # ③④ 的这几个参数界面上已不再暴露（③只留「订单 Excel」框），一律走默认
            "export_scope": DEFAULT_EXPORT_SCOPE,
            "export_format": DEFAULT_EXPORT_FORMAT,
            "interval": float(self.sp_interval.value()),
            "retries": int(self.sp_retry.value()),
            "api_base": tc.DEFAULT_API_BASE,
            "page_origins": "",
            "seller_id": "",
            "target_shop": self.cb_shop.currentText().strip() or DEFAULT_TARGET_SHOP,
            "transport_mode": DEFAULT_TRANSPORT_MODE,
            "skip_ip_check": self.chk_skip_ip.isChecked(),
        }

        self._set_busy("pipe")
        self.bar_progress.setRange(0, 0)
        self.bar_progress.setValue(0)
        if do_orders:
            self.tbl_orders.setRowCount(0)
        if do_customers:
            self.tbl_customers.setRowCount(0)
        if do_rating:
            self.tbl_rating.setRowCount(0)
        if do_bad:
            self.tbl_bad.setRowCount(0)
        steps = [s for s, on in (("导出订单", do_orders), ("抓客户信息", do_customers),
                                 ("导出商品评分", do_rating),
                                 ("抓差评客户", do_bad)) if on]
        self.log_line("=" * 60)
        self.log_line(f"开始流水线：{' + '.join(steps) or '(无步骤)'}"
                      f"　店铺={store.get('browserName') or '?'}")

        self.worker = PipelineThread(self.cfg, self.ed_root.text().strip(), store, opts, self)
        self.worker.log.connect(self.log_line)
        self.worker.stage.connect(self._on_stage)
        self.worker.order_progress.connect(self._on_order_progress)
        self.worker.customer_progress.connect(self._on_customer_progress)
        self.worker.rating_progress.connect(self._on_rating_progress)
        self.worker.orders_ready.connect(self._on_orders_ready)
        self.worker.customer_row.connect(self._on_customer_row)
        self.worker.rating_ready.connect(self._on_rating_ready)
        self.worker.bad_progress.connect(self._on_bad_progress)
        self.worker.bad_ready.connect(self._on_bad_ready)
        self.worker.done.connect(self._on_pipeline_done)
        self.worker.failed.connect(self._on_pipeline_failed)
        self.worker.start()

    def _on_stage(self, stage: str) -> None:
        self.lbl_stat.setText(f"阶段：{stage}")

    def _on_order_progress(self, got: int, total: int) -> None:
        self.bar_progress.setRange(0, 0)  # 未知总数时用忙碌态
        if total:
            self.bar_progress.setRange(0, total)
            self.bar_progress.setValue(min(got, total))
        self.lbl_stat.setText(f"订单：已拉取 {got}" + (f" / {total}" if total else ""))

    def _on_customer_progress(self, done: int, total: int) -> None:
        self.bar_progress.setRange(0, max(1, total))
        self.bar_progress.setValue(done)
        self.lbl_stat.setText(f"客户信息：{done}/{total}")

    def _on_rating_progress(self, got: int, total: int) -> None:
        self.bar_progress.setRange(0, 0)
        if total:
            self.bar_progress.setRange(0, total)
            self.bar_progress.setValue(min(got, total))
        self.lbl_stat.setText(f"商品评分：{got}" + (f" / {total}" if total else ""))

    def _on_bad_progress(self, got: int, total: int) -> None:
        self.bar_progress.setRange(0, max(1, total))
        self.bar_progress.setValue(got)
        self.lbl_stat.setText(f"差评客户：{got}/{total}")

    def _on_bad_ready(self, rows: list) -> None:
        """⑥ 差评客户表：客户列 + 评分 + 商品（列固定，不随导出文件变）。"""
        if not rows:
            self._fill_table(self.tbl_bad, [dict.fromkeys(BAD_CUSTOMER_COLUMNS, "")])
            self.log_line("没有可写入的差评客户")
            return
        shown = self._fill_table(self.tbl_bad, rows)
        self.tabs.setCurrentWidget(self.tbl_bad)
        if shown:
            self.log_line(f"差评客户表格只显示前 {shown} 行，"
                          f"完整 {len(rows)} 行见导出的 Excel")

    def _on_orders_ready(self, rows: list) -> None:
        """订单表按导出文件自己的表头重建列（导出文件 60+ 列，与内置列名不同）。"""
        if not rows:        # 空结果也要建出列，不然表格还是上一批的旧列
            rows = []
        shown = self._fill_table(self.tbl_orders, rows)
        if not rows:
            self._fill_table(self.tbl_orders, [dict.fromkeys(tc.ORDER_COLUMNS, "")])
        if shown:
            self.log_line(f"表格只显示前 {shown} 行，完整 {len(rows)} 行见导出的 Excel")
        self.tabs.setCurrentWidget(self.tbl_orders)

    def _on_rating_ready(self, rows: list) -> None:
        """商品评分表按导出文件自己的表头建列（列数与订单完全不同）。"""
        shown = self._fill_table(self.tbl_rating, rows)
        if rows:
            self.tabs.setCurrentWidget(self.tbl_rating)
        if shown:
            self.log_line(f"商品评分表格只显示前 {shown} 行，完整 {len(rows)} 行见导出的 Excel")

    @staticmethod
    def _fill_table(table: Any, rows: list) -> None:
        cols: list[str] = list(rows[0].keys()) if rows else []
        shown = rows[:MAX_TABLE_ROWS]
        table.setUpdatesEnabled(False)
        try:
            table.clear()
            table.setColumnCount(len(cols))
            table.setHorizontalHeaderLabels(cols)
            table.setRowCount(len(shown))
            for r, row in enumerate(shown):
                for c, col in enumerate(cols):
                    table.setItem(r, c, QTableWidgetItem(str(row.get(col, ""))))
        finally:
            table.setUpdatesEnabled(True)
        if len(rows) > len(shown):
            return len(shown)
        return None

    def _on_customer_row(self, idx: int, row: dict) -> None:
        if idx < 0:
            return
        if self.tbl_customers.rowCount() <= idx:
            self.tbl_customers.setRowCount(idx + 1)
        for c, col in enumerate(tc.CUSTOMER_COLUMNS):
            self.tbl_customers.setItem(idx, c, QTableWidgetItem(str(row.get(col, ""))))

    def _on_pipeline_done(self, summary: dict) -> None:
        self._set_busy("idle")
        self.bar_progress.setRange(0, 100)
        self.bar_progress.setValue(100)
        parts = []
        if summary.get("order_out"):
            parts.append(f"订单 {summary.get('order_count', 0)} 条 → {summary['order_out']}")
        if summary.get("customer_out"):
            parts.append(f"客户 {summary.get('customer_count', 0)} 条 → {summary['customer_out']}")
        if summary.get("rating_out"):
            parts.append(f"商品评分 {summary.get('rating_count', 0)} 条 → {summary['rating_out']}")
        if summary.get("bad_out"):
            parts.append(f"差评客户 {summary.get('bad_count', 0)} 条 → {summary['bad_out']}")
        text = "；".join(parts) or "无输出"
        self.log_line(f"流水线完成：{text}")
        if summary.get("errors"):
            self.log_line(f"其中失败 {summary['errors']} 条")
        if summary.get("stopped"):
            self.log_line("（用户中途停止）")
        self.lbl_stat.setText("完成")
        if summary.get("bad_out"):
            self.tabs.setCurrentWidget(self.tbl_bad)
        elif summary.get("rating_out"):
            self.tabs.setCurrentWidget(self.tbl_rating)
        elif summary.get("customer_out"):
            self.tabs.setCurrentWidget(self.tbl_customers)
        self.worker = None

    def _on_pipeline_failed(self, message: str) -> None:
        self._set_busy("idle")
        self.bar_progress.setRange(0, 100)
        self.bar_progress.setValue(0)
        self.lbl_stat.setText("失败")
        self.log_line(f"流水线失败：{message}")
        QMessageBox.critical(self, "流水线失败", message)
        self.worker = None

    def _stop(self) -> None:
        if self.worker is not None and self.worker.isRunning() and hasattr(self.worker, "stop"):
            self.worker.stop()
            self.btn_stop.setEnabled(False)
            self.log_line("已发送停止信号，等待当前请求返回…")

    # -- 其它 ---------------------------------------------------------------
    def _open_output_dir(self) -> None:
        for text in (self.ed_cust_out.text().strip(), self.ed_order_out.text().strip(),
                     self.ed_rating_out.text().strip(), self.ed_bad_out.text().strip()):
            p = Path(text)
            if p.parent.exists():
                reveal_path(p.parent)
                return
        reveal_path(BASE_DIR)

    def _self_check(self) -> None:
        lines = [
            f"Python         : {sys.version.split()[0]}",
            f"工作目录       : {BASE_DIR}",
            f"紫鸟项目       : {self.ed_root.text().strip() or '(未设置)'}",
            f"配置文件       : {self.ed_cfgpath.text().strip() or '(未加载)'}",
            f"账号           : {self.cfg.company if self.cfg else '-'}"
            f" / {self.cfg.username if self.cfg else '-'}",
            f"接口主机       : {tc.DEFAULT_API_BASE}",
            f"订单输出       : {self.ed_order_out.text().strip()}",
            f"客户输出       : {self.ed_cust_out.text().strip()}",
            f"商品评分输出   : {self.ed_rating_out.text().strip()}",
            f"差评客户输出   : {self.ed_bad_out.text().strip()}",
            f"评分页         : {te.rating_page_suffix('https://seller.us.')}",
            "⑤/⑥           : 两颗独立按钮，可分开单独执行",
            f"后台店铺       : {self.cb_shop.currentText().strip() or '(未设置)'}"
            f"  →  {te.shop_home_url(self.cb_shop.currentText()) or '(认不出，不会导航)'}",
            f"订单列/客户列/差评列: {len(tc.ORDER_COLUMNS)} / "
            f"{len(tc.CUSTOMER_COLUMNS)} / {len(BAD_CUSTOMER_COLUMNS)}",
        ]
        QMessageBox.information(self, "环境自检", "\n".join(lines))
        for line in lines:
            self.log_line(line)

    def closeEvent(self, event) -> None:  # noqa: N802
        if self.worker is not None and self.worker.isRunning():
            if hasattr(self.worker, "stop"):
                self.worker.stop()
            self.worker.wait(5000)
        # 关掉日志句柄：Windows 上文件被占用时用户连删都删不掉
        fh = getattr(self, "_log_handle", None)
        if fh is not None:
            try:
                fh.close()
            except Exception:
                pass
            self._log_handle = None
        event.accept()


# ----------------------------------------------------------------------------
# 入口
# ----------------------------------------------------------------------------
def main(argv: list[str] | None = None) -> int:
    argv = list(sys.argv if argv is None else argv)
    if "--check" in argv:
        app = QApplication.instance() or QApplication(argv)
        win = MainWindow()
        win.show()
        app.processEvents()
        # 客户端解析：这一项在 Windows 上是**最容易出问题**的地方（安装目录
        # 用户自选、没有固定位置），CI 又在 Windows runner 上跑，所以顺手把它
        # 也打进报告 —— 既是给用户看的诊断，也是「Windows 分支没崩」的回归哨兵。
        try:
            cli = zb.find_client_executable(
                (getattr(win.cfg, "client_path", "") or "") if win.cfg else "",
                root=win.ed_root.text().strip() or None, force=True)
        except Exception as exc:  # noqa: BLE001
            cli = f"(解析异常：{type(exc).__name__}: {exc})"
        lines = [
            "GUI 构建成功",
            f"冻结打包 : {bool(getattr(sys, 'frozen', False))}",
            f"程序目录 : {zb.app_dir()}",
            f"紫鸟项目 : {win.ed_root.text().strip() or '(未找到)'}",
            f"配置文件 : {win.ed_cfgpath.text().strip() or '(未加载)'}",
            f"配置缺项 : {win.cfg.missing() if win.cfg else ['未加载紫鸟配置']}",
            f"装载错误 : {win.cfg_error or '(无)'}",
            f"客户端程序 : {cli or '(本机没装紫鸟客户端，自动探测已跑完未命中)'}",
            f"订单列   : {len(tc.ORDER_COLUMNS)} ｜客户列: {len(tc.CUSTOMER_COLUMNS)}",
        ]
        if win.cfg is None or win.cfg_error:
            lines.append("候选目录 :")
            lines.extend(f"  - {c}" for c in zb.ziniao_root_candidates()[:6])
        report = "\n".join(lines)
        try:
            print(report)
        except Exception:
            pass  # 冻结成窗口程序后 stdout 可能不可用
        # exe（console=False）看不到 print，CI 靠这个文件判定打包有没有漏依赖
        if "--check-out" in argv:
            i = argv.index("--check-out")
            if i + 1 < len(argv):
                try:
                    Path(argv[i + 1]).write_text(report, encoding="utf-8")
                except Exception as exc:
                    print(f"写自检报告失败：{exc}")
        win.close()
        return 0

    app = QApplication(argv)
    win = MainWindow()
    win.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
