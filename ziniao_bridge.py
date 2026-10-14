"""紫鸟浏览器桥接层。

TikTok 商家后台的订单/客户接口都必须带**店铺自己的登录态**。手工粘贴 cookie
会随会话过期而失效，所以这里改成：用紫鸟打开店铺 → 直接复用店铺浏览器的
登录态和出口 IP 去调接口。

协议客户端不在这里重复实现，而是复用已经调通的
`ziniao_webdriver_demo/util/ziniao_client.py`（那边解决了 v6 启动的一堆坑：
`--user-data-dir` 必填、`ELECTRON_RUN_AS_NODE` 污染、`open -a` 丢参数、
`killall` 杀不干净、本地 IPC 被系统代理拦截等）。

那份客户端是**运行时按文件路径加载**的（`spec_from_file_location`，不是 import），
所以它必须**以 .py 源码形式存在于磁盘上** —— 打包 exe 时不能编译进包里，只能当
数据文件带着。项目里内联了一份在 `vendor/ziniao_webdriver_demo/`，
`locate_ziniao_root()` 会按「exe 旁挂 → 仓库内 vendor → 开发态兄弟目录」依次找。

本模块提供四件事：
  1. locate_ziniao_root() / load_ziniao()：定位并加载那个项目
  2. BrowserSession：按 debuggingPort 用**原生 CDP**（websocket-client）接管店铺浏览器，
     Page.navigate 打断连接后自动重连重放，保证「当前页面」始终可用
  3. open_store / check_ip / grab_cookies / close_store：店铺浏览器操作
  4. BrowserTransport：在店铺页面里直接 fetch TikTok 接口（走店铺代理 IP）
     另有 fetch_bytes / intercept_url / pump / add_init_script，用于「点页面按钮导出订单
     并把文件取回来」——详见 tiktok_export.py

实测结论（2026-10-13）：
  · ``order/get``（订单详情）本地 requests 直连就能通（code=0），
    所以它是浏览器内通道的可靠兜底；
  · ``order/list``（订单列表）本地直连**恒定**返回裸 ``code=10000``
    （与 aid / X-Tts-Oec-Bsid 无关，试过 4068 / 6556 / 无 aid 都一样）；
    页面真实用的是同源 ``GET /api/fulfillment/na/order/list``，缺签名重放 404。
    结论：**订单别再走接口了，改用后台「导出订单」按钮**（Orders 导出 Excel 含全部
    客户字段），所以本模块额外提供了 Fetch 拦截 + 页面内取字节的能力。
  · ``seller_id`` 必须从**该店铺自己的 cookie** 推导
    （``oec_seller_id_unified_seller_env``）。不同店铺的 seller_id 完全不同，
    硬编码成别的店铺会拿到 ``code=98001004``；
  · 页面内 fetch 要从店铺后台页面发起：从 about:blank / 扩展页发起时，
    跨域 + ``content-type: application/json`` 会触发 CORS 预检失败，
    浏览器只报 ``TypeError: Failed to fetch``。
    所以先 navigate 到后台首页，再 fetch；BrowserTransport 仍保留候选 origin 切换。
  · **紫鸟内核拿不到下载落盘路径**：``Browser.setDownloadBehavior``（``allow`` 与
    DrissionPage 用的 ``allowAndName`` 都试过）显示 completed 但文件凭空消失；
    且导出下载 URL 带 ``use-disposable=1``（只能用一次），让浏览器真下载会消耗掉它。
    → 用 ``intercept_url()`` 把请求拦在发送前，再用 ``fetch_bytes()`` 在页面里取回。
"""

from __future__ import annotations

import json
import shutil
import sys
import time
import types
import urllib.parse
from pathlib import Path
from typing import Any, Callable, Iterable

BASE_DIR = Path(__file__).resolve().parent

# 开发态默认位置：<工作区>/紫鸟浏览器/ziniao_webdriver_demo
# （这份是带 config.yaml 的「真身」，所以优先级要排在仓库内 vendor 之前）
DEFAULT_ZINIAO_ROOT = BASE_DIR.parent / "紫鸟浏览器" / "ziniao_webdriver_demo"

# 店铺后台首页（跨境店）。实际优先用 startBrowser 返回的 launcherPage。
TIKTOK_SELLER_HOME = "https://seller.tiktokshopglobalselling.com/"

# 页面来源候选：fetch 必须从这些 origin 之一发起才不会被 CORS 拦掉。
# 顺序即尝试顺序：先按真实后台页面（服务端 Origin 白名单里就有它），
# 最后才用接口自身的 origin（同源，绕开 CORS，但可能被服务端校验 Origin）。
TIKTOK_PAGE_ORIGINS: list[str] = [
    "https://seller.tiktokshopglobalselling.com",
    "https://seller.eu.tiktokshopglobalselling.com",
    "https://seller.us.tiktokshopglobalselling.com",
]

TIKTOK_DOMAIN_KEY = "tiktokshopglobalselling"

# 店铺身份类错误码：都不是「重试就有用」的错，而是「这次请求发得不对」，
# 换个 host / 换种写法可能就通了（实测 98001008 就是这么来的）。
PERMISSION_CODES: frozenset[str] = frozenset({
    "98001008",   # 需要权限 —— cookie / host 不是这批订单所属的店
    "98001004",   # 部分信息填写错误 —— seller_id 上下文对不上
    "98001001",   # 商家账号相关
    "13009001",   # 此商家账号不存在
})

# 照抄抓包模板时**不**带过来的查询参数：这些是列表的分页/排序/筛选，
# 详情接口带上会变成另一种查询语义。
TEMPLATE_DROP_QUERY: frozenset[str] = frozenset({
    "page_size", "pageSize", "page_token", "pageToken", "cursor", "offset",
    "limit", "page", "page_number", "pageNumber", "sort", "sort_by",
    "sort_field", "sort_order", "sort_type", "tab", "selected_sort",
    "search_type", "create_time_ge", "create_time_lt", "order_status",
})

# capture_order_request.py 抓到的真实请求模板落在哪
CAPTURE_FILE = BASE_DIR / "devtools" / "captured_order_request.json"


def load_captured_template(path: str | Path | None = None) -> dict[str, Any] | None:
    """读抓包模板里最像「订单接口」的那条请求，供 BrowserTransport 照抄。

    返回 ``{"url": ..., "method": ..., "headers": {...}}``；没有/读不到返回 None。
    """
    p = Path(path) if path else CAPTURE_FILE
    if not p.is_file():
        return None
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return None
    cands = [r for r in (data.get("network") or [])
             if isinstance(r, dict) and "fulfillment" in str(r.get("url") or "")]
    if not cands:
        hook = [r for r in (data.get("page_hook") or [])
                if isinstance(r, dict) and "fulfillment" in str(r.get("url") or "")]
        if not hook:
            return None
        return {"url": hook[0].get("url") or "", "method": hook[0].get("method") or "POST",
                "headers": hook[0].get("headers") or {}}
    # 优先 POST（两个订单接口都是 POST）
    cands.sort(key=lambda r: 0 if str(r.get("method")).upper() == "POST" else 1)
    best = cands[0]
    return {"url": str(best.get("url") or ""),
            "method": str(best.get("method") or "POST"),
            "headers": dict(best.get("headers") or {})}

LogFn = Callable[[str], None]


def _noop(_msg: str) -> None:  # pragma: no cover - 占位
    pass


def origin_of(url: str) -> str:
    """取 URL 的 origin（scheme://host）。"""
    parts = urllib.parse.urlsplit(str(url or ""))
    if not parts.scheme or not parts.netloc:
        return ""
    return f"{parts.scheme}://{parts.netloc}"


class ZiniaoBridgeError(RuntimeError):
    """桥接层自身的错误（找不到项目、连不上浏览器等）。"""


# ----------------------------------------------------------------------------
# DrissionPage 异常分类：掉线要重连，JS 自身报错不能重连
# ----------------------------------------------------------------------------
_FATAL_TYPES: tuple[type, ...] | None = None


def _fatal_types() -> tuple[type, ...]:
    global _FATAL_TYPES
    if _FATAL_TYPES is None:
        found: list[type] = []
        try:
            from DrissionPage import errors as E  # noqa: N812

            for name in ("PageDisconnectedError", "ContextLostError",
                         "BrowserConnectError", "TargetNotFoundError"):
                cls = getattr(E, name, None)
                if isinstance(cls, type):
                    found.append(cls)
        except Exception:  # pragma: no cover - DrissionPage 缺失时退化
            pass
        _FATAL_TYPES = tuple(found)
    return _FATAL_TYPES


def is_disconnect(exc: BaseException) -> bool:
    """判断异常是否属于「CDP 会话掉线」——只有这类才值得重连。"""
    # 原生 websocket-client / socket 层的断线
    if isinstance(exc, (ConnectionResetError, BrokenPipeError, ConnectionError,
                        ConnectionAbortedError)):
        return True
    types_ = _fatal_types()
    if types_ and isinstance(exc, types_):
        return True
    text = str(exc).lower()
    return any(k in text for k in ("disconnected", "has been closed", "target closed",
                                   "connection lost", "session closed",
                                   "connection reset", "broken pipe"))


# ----------------------------------------------------------------------------
# 1. 定位 / 加载紫鸟项目
# ----------------------------------------------------------------------------
def app_dir() -> Path:
    """「程序」所在目录 —— 冻结成 exe 后是 exe 旁边的目录，开发态是项目根。

    为什么不能用 ``__file__``：PyInstaller onefile 下 ``__file__`` 指向
    ``sys._MEIPASS``（每次启动新建、退出即删的临时解包目录），拿它去推
    ``../紫鸟浏览器/...`` 必然扑空。所以冻结时改按 ``sys.executable`` 定位。
    """
    if getattr(sys, "frozen", False):
        exe = getattr(sys, "executable", "") or ""
        if exe:
            return Path(exe).resolve().parent
    return BASE_DIR


def _bundle_dir() -> Path | None:
    """PyInstaller 的 ``--add-data`` 解包目录；非冻结时为 None。"""
    mp = getattr(sys, "_MEIPASS", None)
    return Path(mp) if mp else None


# 相对「程序所在目录」的候选子路径，按优先级排列（发布包布局 / 临时摆放 / 仓库内 vendor）
_ZINIAO_REL_PATHS: tuple[Path, ...] = (
    Path("紫鸟浏览器") / "ziniao_webdriver_demo",
    Path("ziniao_webdriver_demo"),
    Path("vendor") / "ziniao_webdriver_demo",
)


def ziniao_root_candidates(override: str | Path | None = None) -> list[Path]:
    """按优先级列出所有会被尝试的紫鸟项目目录（只列，不校验存在性）。

    冻结态（exe）顺序：
      1. 显式指定（界面上的「紫鸟项目路径」）
      2. 程序旁的 紫鸟浏览器/ziniao_webdriver_demo   ← 发布包布局（exe 旁挂）
      3. 程序旁的 ziniao_webdriver_demo
      4. 程序旁的 vendor/ziniao_webdriver_demo
      5. exe 内置的 vendor（万一哪天改成 --add-data 打进 exe）
      6. 从 exe 目录向上最多 4 级 glob ``*ziniao_webdriver_demo``

    开发态顺序（兄弟目录提到最前）：
      1. 显式指定
      2. ../紫鸟浏览器/ziniao_webdriver_demo
         —— 必须排在 vendor 之前：开发时**只有它带 config.yaml**，
            先命中 vendor 会让界面读不到账号、起不来浏览器。
      3. 项目旁的 紫鸟浏览器/ziniao_webdriver_demo / ziniao_webdriver_demo
      4. 项目内的 vendor/ziniao_webdriver_demo（兄弟目录不在时才用）
      5. 向上最多 4 级 glob
    """
    cands: list[Path] = []
    if override:
        cands.append(Path(override).expanduser())

    here = app_dir()
    rels = [here / rel for rel in _ZINIAO_REL_PATHS]
    bundle = _bundle_dir()
    bundle_rels = [bundle / rel for rel in _ZINIAO_REL_PATHS] if bundle else []

    if getattr(sys, "frozen", False):
        cands.extend(rels)
        cands.append(DEFAULT_ZINIAO_ROOT)
    else:
        cands.append(DEFAULT_ZINIAO_ROOT)
        cands.extend(rels)

    cands.extend(bundle_rels)

    # 冻结态下 BASE_DIR 是临时解包目录，向上搜没意义，改用 exe 所在目录
    up = app_dir()
    for _ in range(4):
        up = up.parent
        cands.extend(sorted(up.glob("*ziniao_webdriver_demo")))
        cands.extend(sorted(up.glob("*/*ziniao_webdriver_demo")))

    seen: set[str] = set()
    uniq: list[Path] = []
    for c in cands:
        key = str(c)
        if key not in seen:
            seen.add(key)
            uniq.append(c)
    return uniq


def preferred_ziniao_root() -> Path:
    """界面上「紫鸟项目路径」的推荐值：发布态看 exe 旁，开发态看兄弟目录。"""
    if getattr(sys, "frozen", False):
        return app_dir() / _ZINIAO_REL_PATHS[0]
    return DEFAULT_ZINIAO_ROOT


def locate_ziniao_root(override: str | Path | None = None) -> Path:
    """找到 ziniao_webdriver_demo 目录（必须含 util/ziniao_client.py）。

    找不到抛 ZiniaoBridgeError。
    """
    candidates = ziniao_root_candidates(override)

    for cand in candidates:
        if (cand / "util" / "ziniao_client.py").is_file():
            return cand.resolve()

    raise ZiniaoBridgeError(
        "找不到紫鸟项目 ziniao_webdriver_demo（需要其中的 util/ziniao_client.py）。\n"
        f"已尝试：\n  " + "\n  ".join(str(c) for c in candidates[:8]) + "\n"
        "请在界面上把「紫鸟项目路径」改成该目录的绝对路径。"
    )


_MODS: dict[str, Any] | None = None
_LOADED_ROOT: Path | None = None


def _load_file(name: str, path: Path):
    """按文件路径加载模块并登记到 sys.modules（避免依赖 sys.path 顺序）。"""
    import importlib.util

    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise ZiniaoBridgeError(f"无法加载模块：{path}")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


def load_ziniao(root: str | Path | None = None, force: bool = False) -> dict[str, Any]:
    """加载紫鸟协议客户端，返回 {"client", "dp", "config_path", "root"}。

    注意：`ziniao_client.py` 内部写死了 `from util.dp import DpBrowser`，
    所以这里必须把那个项目的 util 目录注册成名为 `util` 的顶级包。
    """
    global _MODS, _LOADED_ROOT

    real_root = locate_ziniao_root(root)
    if _MODS is not None and not force and _LOADED_ROOT == real_root:
        return _MODS

    util_dir = real_root / "util"

    if "util" in sys.modules:
        existing = getattr(sys.modules["util"], "__path__", None)
        same = existing and str(util_dir) in [str(p) for p in existing]
        if not same:
            raise ZiniaoBridgeError(
                "sys.modules 里已经有另一个 `util` 包，与紫鸟项目的 util 冲突："
                f"{existing}。请改名后重试。"
            )
    else:
        pkg = types.ModuleType("util")
        pkg.__path__ = [str(util_dir)]  # type: ignore[attr-defined]
        pkg.__file__ = str(util_dir / "__init__.py")
        sys.modules["util"] = pkg

    dp = _load_file("util.dp", util_dir / "dp.py")
    zc = _load_file("util.ziniao_client", util_dir / "ziniao_client.py")
    zc.__package__ = "util"

    if str(real_root) not in sys.path:
        sys.path.append(str(real_root))

    _MODS = {
        "client": zc,
        "dp": dp,
        "root": real_root,
        "config_path": real_root / "config" / "config.yaml",
    }
    _LOADED_ROOT = real_root
    return _MODS


def _seed_config(cfg_path: Path) -> bool:
    """config.yaml 不存在时，从同目录的 config.example.yaml 播种一份。

    为什么需要：发布用的 zip 里**故意不带 config.yaml**（否则用户解压升级时会把
    已填好的企业账号盖掉）。全新解压出来的目录只有 config.example.yaml，得有人
    把它变成真配置文件，否则界面一上来就是「配置缺项」。
    只读目录等失败情况一律咽掉，让上层用默认值继续跑。
    """
    example = cfg_path.with_name("config.example.yaml")
    try:
        cfg_path.parent.mkdir(parents=True, exist_ok=True)
        if example.is_file():
            shutil.copyfile(example, cfg_path)
            return True
    except OSError:
        pass
    return False


def load_config(root: str | Path | None = None, force: bool = False):
    """读紫鸟的 config.yaml，返回 ZiniaoConfig（缺失时先播种一份）。"""
    mods = load_ziniao(root, force=force)
    cfg_path: Path = mods["config_path"]
    if not cfg_path.exists():
        _seed_config(cfg_path)
    return mods["client"].ZiniaoConfig.from_yaml(cfg_path)


def make_client(cfg, log: LogFn | None = None, root: str | Path | None = None):
    """按配置创建 ZiniaoClient。"""
    mods = load_ziniao(root)
    return mods["client"].ZiniaoClient(cfg, log=log or _noop)


def save_config(cfg, root: str | Path | None = None) -> Path:
    mods = load_ziniao(root)
    return cfg.save_yaml(mods["config_path"])


def find_client_executable(client_path: str = "", root: str | Path | None = None,
                           force: bool = False) -> str:
    """按平台解析（必要时自动探测）紫鸟客户端可执行文件，给界面的「自动检测」用。

    - Windows 没有固定安装目录，``client_path`` 可以是 exe 全路径、所在目录或留空
    - macOS 支持 ``ziniao`` / ``xxx.app`` / bundle 内的可执行文件
    ``force=True`` 时忽略上次探测的缓存重新扫一遍（点「自动检测」、或启动环境前用它，
    因为启动会先杀掉客户端进程，必须在杀之前把路径探到并缓存住）。
    找不到返回空串。
    """
    mods = load_ziniao(root)
    client_mod = mods["client"]
    if getattr(client_mod, "IS_WINDOWS", False):
        return client_mod.resolve_windows_client_executable(client_path, force=force)
    if getattr(client_mod, "IS_MAC", False):
        # 解析不出就回落到「ziniao」（去 /Applications 里找）——和 Windows 分支
        # 一致：配置可能是从别的机器抄来的，写死了别的平台路径不该让整个流程卡死。
        return (client_mod.resolve_mac_app_executable(client_path)
                or client_mod.resolve_mac_app_executable("ziniao"))
    return client_path or "/opt/ziniao/ziniaobrowser"


def client_hint(cfg: Any, root: str | Path | None = None) -> str:
    """找不到客户端时的诊断文本（按平台给具体填法 + 已探过的位置）。

    直接转发给紫鸟客户端的 ``ZiniaoConfig.client_path_hint()``；万一分支没实现
    或抛错，也给一句兜底，绝不让「提示本身」把界面搞崩。
    """
    fn = getattr(cfg, "client_path_hint", None)
    if callable(fn):
        try:
            return str(fn())
        except Exception as exc:  # pragma: no cover - 只在客户端分支异常时走到
            return f"生成客户端诊断信息失败：{exc}"
    return ("没找到紫鸟客户端程序。「客户端路径」留空即可自动查找，"
            r"Windows 也可直接填 exe 全路径（如 D:\紫鸟\ziniao\ziniao.exe）。")


# ----------------------------------------------------------------------------
# 2. 店铺浏览器会话（原生 CDP + 断线自愈）
# ----------------------------------------------------------------------------
# 为什么不用 DrissionPage（2026-10-13 真机实测，店铺「嘉恒业」）：
#   · DrissionPage 接管 ziniao 店铺浏览器后，页面一旦发生导航，它的 CDP WebSocket
#     就被服务端 reset，且**不会自愈** —— 之后所有页面操作恒定
#     PageDisconnectedError（连 new_tab 都失败），reconnect 也救不回来；
#   · 换成原生 websocket-client 直连同一个 webSocketDebuggerUrl 则完全正常：
#     发 Page.navigate 会 reset 一次连接，**重连后页面已正常导航完成**，
#     Runtime.evaluate / Storage.getCookies 全部可用，order/list 直接 code=0。
# 所以这里直接说 CDP：自己管连接生命周期（断线重连 + 请求重放），不再走
# DrissionPage 的页面抽象。
_DISCONNECT_HINTS = ("disconnected", "has been closed", "target closed",
                     "connection lost", "session closed", "connection reset",
                     "broken pipe", "websocket")


def _is_ws_error(exc: BaseException) -> bool:
    """CDP 连接层错误（可重连重试），区别于 JS 自身报错（直接上抛）。"""
    if isinstance(exc, (ConnectionResetError, BrokenPipeError, ConnectionError,
                        ConnectionAbortedError, OSError)):
        # 注意：OSError 涵盖 socket 层，但 WebSocketTimeoutException 也是 OSError 子类，
        # 超时同样意味着连接不可用，重连重试是可接受的（两个接口都是读操作）。
        return True
    try:
        import websocket as _ws

        for name in ("WebSocketException", "WebSocketConnectionClosedException",
                     "WebSocketTimeoutException", "WebSocketBadStatusException"):
            cls = getattr(_ws, name, None)
            if isinstance(cls, type) and isinstance(exc, cls):
                return True
    except Exception:
        pass
    return is_disconnect(exc)


def _host_tail(url: str) -> str:
    """取「站点尾巴」——主机名的最后两段，用于容忍 region 子域差异。"""
    host = urllib.parse.urlsplit(str(url or "")).netloc.rsplit("@", 1)[-1]
    host = host.split(":")[0].lower()
    labels = [p for p in host.split(".") if p]
    return ".".join(labels[-2:]) if len(labels) >= 2 else host


def _same_site(a: str, b: str) -> bool:
    """两个 URL 是否同一站点（seller.us.x.com 与 seller.x.com 视为同一站）。"""
    ta, tb = _host_tail(a), _host_tail(b)
    return bool(ta) and ta == tb


def _host_of(url: str) -> str:
    host = urllib.parse.urlsplit(str(url or "")).netloc.rsplit("@", 1)[-1]
    return host.split(":")[0].lower()


def _same_host(a: str, b: str) -> bool:
    """两个 URL 是否**同一个 host**（``seller.us.x.com`` ≠ ``seller.eu.x.com``）。

    和 ``_same_site`` 的区别很关键：站点级比较会把各个区域站混为一谈，
    于是 ``ensure_on`` 从不真正换站 —— 实测踩过，见 ``_ensure_origin``。
    """
    ha, hb = _host_of(a), _host_of(b)
    return bool(ha) and ha == hb


def _is_api_host(url: str) -> bool:
    """是不是接口主机（``api16-*`` / ``api-*``）—— 它不能当页面来源。"""
    return _host_of(url).startswith("api")


class BrowserSession:
    """按 debuggingPort 接管紫鸟店铺浏览器（原生 CDP），掉线自动重连。

    对外接口与旧 DrissionPage 版保持一致：``navigate`` / ``ensure_on`` /
    ``current_url`` / ``run_js`` / ``cookies`` / ``close``，另有 ``reconnect`` /
    ``tabs`` / ``eval`` / ``call`` 供调试。
    """

    def __init__(self, port: Any, log: LogFn | None = None,
                 home_urls: Iterable[str] | None = None,
                 retry: int = 20, interval: float = 1.0):
        self.port = int(port)
        self.log = log or _noop
        self.home_urls = [u for u in (home_urls or []) if u]
        self._retry = max(1, int(retry))
        self._interval = max(0.2, float(interval))
        self._ws: Any = None
        self._ws_url = ""
        self._target_id = ""
        self._req_id = 0
        self.connect()

    # -- 设备发现（HTTP /json） --------------------------------------------
    def _http_json(self, path: str) -> Any:
        import requests

        s = requests.Session()
        s.trust_env = False  # 本地回环别被系统代理（Clash 等）拦掉
        r = s.get(f"http://127.0.0.1:{self.port}{path}", timeout=5,
                  headers={"Connection": "close"})
        return r.json()

    def targets(self) -> list[dict[str, Any]]:
        try:
            data = self._http_json("/json/list")
        except Exception:
            return []
        if not isinstance(data, list):
            return []
        return [t for t in data if isinstance(t, dict)]

    def pages(self) -> list[dict[str, Any]]:
        return [t for t in self.targets() if t.get("type") == "page"]

    def new_page(self, url: str = "about:blank") -> dict[str, Any]:
        """新建一个标签页。

        实测坑：从紫鸟的 ``chrome-extension://`` 检测页发 ``Page.navigate`` 时，
        该 target 可能直接被销毁，接着 ``reconnect()`` 会发现「一个 page 都没有」
        而彻底卡死。所以这里留一条救急通道：直接用浏览器自带的 ``/json/new``。
        """
        import requests

        endpoint = (f"http://127.0.0.1:{self.port}/json/new?"
                    + urllib.parse.quote(url, safe=":/?&=%#"))
        s = requests.Session()
        s.trust_env = False
        last: Exception | None = None
        for method in ("put", "get"):     # 新版 Chrome 只认 PUT
            try:
                r = getattr(s, method)(endpoint, timeout=10,
                                       headers={"Connection": "close"})
                if r.status_code < 400 and r.text.strip().startswith("{"):
                    data = r.json()
                    self.log(f"新建标签页：{str(data.get('id'))[:8]}")
                    return data
                last = ZiniaoBridgeError(f"/json/new HTTP {r.status_code}")
            except Exception as exc:
                last = exc
        raise ZiniaoBridgeError(f"新建标签页失败：{last}")

    def _choose(self, pages: list[dict[str, Any]]) -> dict[str, Any]:
        """挑工作页面：同站 > 任意 http 页 > 最后一个（通常 about:blank）。"""
        hosts = [urllib.parse.urlsplit(u).netloc for u in self.home_urls]
        for p in pages:
            url = str(p.get("url") or "")
            if any(h and h in url for h in hosts):
                return p
        for p in pages:
            if str(p.get("url") or "").startswith("http"):
                return p
        return pages[-1]

    # -- CDP 连接 ----------------------------------------------------------
    def _open_ws(self, target: dict[str, Any]) -> None:
        try:
            import websocket
        except Exception as exc:  # pragma: no cover - 依赖缺失时给明确提示
            raise ZiniaoBridgeError(
                "缺少 websocket-client，请先执行 `pip install websocket-client`"
            ) from exc
        ws_url = target.get("webSocketDebuggerUrl")
        if not ws_url:
            raise ZiniaoBridgeError(f"target 缺少 webSocketDebuggerUrl：{target.get('id')}")
        self._close_ws()
        self._ws = websocket.create_connection(ws_url, suppress_origin=True, timeout=15)
        self._ws_url = ws_url
        self._target_id = str(target.get("id") or "")

    def _close_ws(self) -> None:
        ws, self._ws = self._ws, None
        if ws is not None:
            try:
                ws.close()
            except Exception:
                pass

    def connect(self) -> bool:
        """等到调试端口出现 page target 并建好 CDP 连接。"""
        last: Exception | None = None
        deadline = time.time() + self._retry * self._interval
        while time.time() < deadline:
            pages = self.pages()
            if pages:
                try:
                    self._open_ws(self._choose(pages))
                    self.log(f"浏览器内核就绪（调试端口 {self.port}，标签页 {len(pages)} 个）")
                    return True
                except Exception as exc:
                    last = exc
            time.sleep(self._interval)
        raise ZiniaoBridgeError(f"等待调试端口 {self.port} 就绪超时：{last}")

    def reconnect(self) -> bool:
        """掉线自愈：重新发现 page target 并重建 CDP 连接。"""
        self.log("CDP 会话掉线，正在重连…")
        pages: list[dict[str, Any]] = []
        for _ in range(max(1, self._retry)):
            pages = self.pages()
            if pages:
                break
            time.sleep(self._interval)
        if not pages:
            # 一个 page 都没有：紫鸟内核把旧标签页回收了，自己造一个
            try:
                self.new_page()
                pages = self.pages()
            except Exception as exc:
                self.log(f"新建标签页也失败：{exc}")
        if not pages:
            raise ZiniaoBridgeError("重连失败：调试端口没有可用页面")
        t = self._choose(pages)
        self._open_ws(t)
        self.log(f"CDP 会话已恢复（target {str(t.get('id'))[:8]}）")
        return True

    def close(self) -> None:
        """只断开 CDP 连接，**不**关闭浏览器（紫鸟的 stopBrowser 负责关）。"""
        self._close_ws()

    def alive(self) -> bool:
        try:
            return any(str(t.get("id")) == self._target_id for t in self.pages())
        except Exception:
            return False

    # -- CDP 调用 ----------------------------------------------------------
    def call(self, method: str, params: dict[str, Any] | None = None,
             timeout: float | None = None, retry: bool = True) -> Any:
        """发一条 CDP 命令并等它的响应（自动跳过事件消息）。

        连接层异常时重连并重放一次（两个业务接口都是读操作，重放安全）。
        """
        ws = self._ws
        if ws is None:
            self.reconnect()
            ws = self._ws
        self._req_id += 1
        mid = self._req_id
        prev = ws.gettimeout()
        if timeout:
            ws.settimeout(float(timeout))
        try:
            ws.send(json.dumps({"id": mid, "method": method, "params": params or {}}))
            while True:
                msg = json.loads(ws.recv())
                if msg.get("id") == mid:
                    if "error" in msg:
                        raise RuntimeError(f"{method}: {msg['error']}")
                    return msg.get("result")
        except Exception as exc:
            if retry and _is_ws_error(exc):
                self.log(f"CDP 连接中断（{type(exc).__name__}），重连后重试 {method}")
                self.reconnect()
                return self.call(method, params, timeout, retry=False)
            raise
        finally:
            try:
                ws.settimeout(prev)
            except Exception:
                pass

    def eval(self, script: str, await_promise: bool = True,
             timeout: float = 120) -> Any:
        """执行 JS 并取回值。JS 自身报错 → ZiniaoBridgeError（不重连）。"""
        res = self.call("Runtime.evaluate",
                        {"expression": script, "returnByValue": True,
                         "awaitPromise": bool(await_promise)},
                        timeout=timeout) or {}
        out = res.get("result") or {}
        if out.get("subtype") == "error" or res.get("exceptionDetails"):
            desc = out.get("description") or str(res.get("exceptionDetails"))[:300]
            raise ZiniaoBridgeError(f"页面 JS 报错：{desc}")
        return out.get("value")

    # -- 页面 --------------------------------------------------------------
    def current_url(self) -> str:
        try:
            return str(self.eval("location.href", await_promise=False, timeout=15) or "")
        except Exception:
            for t in self.pages():
                if str(t.get("id")) == self._target_id:
                    return str(t.get("url") or "")
            return ""

    def _wait_url(self, url: str, timeout: float = 60.0) -> bool:
        deadline = time.time() + max(1.0, float(timeout))
        while time.time() < deadline:
            cur = self.current_url()
            if cur.startswith("http") and _same_site(cur, url):
                return True
            time.sleep(1)
        return False

    def navigate(self, url: str, retry: int = 5, interval: float = 2.0,
                 wait: float = 60.0) -> "BrowserSession":
        """打开目标页面。

        ``Page.navigate`` 会把 CDP WebSocket 打断（该内核的既定行为），
        所以这里把「发送导航」和「等页面就位」拆开：断线由 ``call`` 内部重连，
        之后再轮询 ``location.href`` 确认真的到了。
        """
        last: Exception | None = None
        for i in range(1, max(1, retry) + 1):
            try:
                self.call("Page.navigate", {"url": url}, timeout=30)
            except Exception as exc:
                last = exc
                self.log(f"[{i}/{retry}] 导航请求失败（{type(exc).__name__}）：{exc}")
                try:
                    self.reconnect()
                except Exception as exc2:
                    last = exc2
            if self._wait_url(url, timeout=wait):
                self.log(f"已打开：{url}")
                return self
            last = last or ZiniaoBridgeError("页面未到达目标地址")
            self.log(f"[{i}/{retry}] 等待页面到达 {url} 超时")
            time.sleep(interval)
        raise ZiniaoBridgeError(f"无法打开页面：{url}（{last}）")

    def ensure_on(self, url: str) -> "BrowserSession":
        """确保当前页面位于 url（同站即视为已在，避免多余跳转）。"""
        if _same_site(self.current_url(), url):
            return self
        return self.navigate(url)

    def run_js(self, script: str, timeout: float = 120, retry: int = 2) -> Any:
        """在页面里执行 JS。

        只在**连接层**出错时重连重试；JS 自身抛错（例如跨域 ``Failed to fetch``）
        直接上抛，交给上层换页面来源或换通道，避免把真问题重试掉。
        """
        last: Exception | None = None
        for _ in range(max(1, retry)):
            try:
                return self.eval(script, await_promise=True, timeout=timeout)
            except Exception as exc:
                last = exc
                if not _is_ws_error(exc):
                    break
        raise ZiniaoBridgeError(f"页面 JS 执行失败：{last}")

    def read_json(self, expr: str, default: Any = None) -> Any:
        """执行 JS 表达式，并把返回的 JSON 字符串解析成 Python 对象。"""
        raw = self.eval(expr, await_promise=False, timeout=60)
        text = str(raw or "").strip()
        if not text or text in ("null", "undefined"):
            return default
        try:
            return json.loads(text)
        except Exception:
            return default

    # -- 事件流 / 原始命令 --------------------------------------------------
    def send(self, method: str, params: dict[str, Any] | None = None) -> int:
        """裸发一条 CDP 命令，不等响应。

        用于必须在**事件回调内**立刻回复的场景（Fetch 拦截里放行/中止请求），
        这时不能进 ``call()`` 去读响应，否则会把指令流搅乱。
        """
        ws = self._ws
        if ws is None:
            self.reconnect()
            ws = self._ws
        self._req_id += 1
        ws.send(json.dumps({"id": self._req_id, "method": method,
                            "params": params or {}}))
        return self._req_id

    def pump(self, seconds: float,
             on_event: Callable[[dict[str, Any]], None] | None = None,
             stop: Callable[[list[dict[str, Any]]], bool] | None = None) -> list[dict[str, Any]]:
        """手动读取 CDP 事件流 ``seconds`` 秒（命令响应会被跳过）。

        :param on_event: 每条事件回调，可在里面 ``send()`` 回命令
        :param stop: 收到事件列表后返回 True 即提前结束
        """
        ws = self._ws
        out: list[dict[str, Any]] = []
        if ws is None:
            return out
        end = time.time() + max(0.0, float(seconds))
        while time.time() < end:
            ws.settimeout(max(0.3, min(2.0, end - time.time())))
            try:
                msg = json.loads(ws.recv())
            except Exception:
                continue
            if not msg.get("method"):
                continue
            out.append(msg)
            if on_event:
                try:
                    on_event(msg)
                except Exception as exc:  # 回调里的异常不该打断收事件
                    self.log(f"事件处理异常：{exc}")
            if stop and stop(out):
                break
        return out

    def add_init_script(self, source: str) -> str:
        """注入「每次新文档加载都执行」的脚本（导航后依然生效）。

        必须在导航**之前**调用；页面已加载后再注入只对当前文档有效。
        """
        self.call("Page.enable", {}, timeout=20)
        res = self.call("Page.addScriptToEvaluateOnNewDocument",
                        {"source": source}, timeout=20) or {}
        return str(res.get("identifier") or "")

    # -- 页面内取二进制 -----------------------------------------------------
    def fetch_bytes(self, url: str, timeout: float = 180.0,
                    chunk: int = 120_000, tag: str = "__ziniao_dl") -> bytes:
        """在**页面里** fetch 目标 URL，把响应体取回 Python。

        为什么不让浏览器自己下载：实测 ziniao 内核会忽略
        ``Browser.setDownloadBehavior`` 的 ``downloadPath``（``allow`` /
        ``allowAndName`` 都试过），文件下完就"消失"，拿不到落盘路径。
        这里改成页面内 fetch + base64 分块回传：既走店铺出口 IP 和登录态，
        又能自己决定存到哪个路径。
        """
        import base64 as _b64

        start = (
            "(()=>{const url=%s;window[%r]={b64:'',meta:null};"
            "fetch(url,{credentials:'include'}).then(async r=>{"
            "const b=await r.arrayBuffer();const u8=new Uint8Array(b);"
            "let s='';const CH=0x8000;"
            "for(let i=0;i<u8.length;i+=CH){"
            "s+=String.fromCharCode.apply(null,u8.subarray(i,i+CH));}"
            "window[%r]={b64:btoa(s),meta:{status:r.status,len:u8.length,"
            "type:String(r.headers.get('content-type')||'')}};"
            "}).catch(e=>{window[%r]={b64:'',meta:{err:String(e)}};});"
            "return 'started';})()"
        ) % (json.dumps(url), tag, tag, tag)
        self.eval(start, await_promise=False, timeout=30)

        deadline = time.time() + max(5.0, float(timeout))
        meta: dict[str, Any] = {}
        while time.time() < deadline:
            time.sleep(1)
            got = self.read_json(f"JSON.stringify(window[{tag!r}]&&window[{tag!r}].meta)",
                                 default=None)
            if isinstance(got, dict):
                meta = got
                break
        if not meta:
            raise ZiniaoBridgeError(f"页面内取字节超时：{url[:120]}")
        if meta.get("err"):
            raise ZiniaoBridgeError(f"页面内 fetch 失败：{meta['err']}")
        if int(meta.get("len") or 0) <= 0:
            raise ZiniaoBridgeError(
                f"页面内 fetch 返回空（HTTP {meta.get('status')}，{meta.get('type')}）")

        b64_len = int(self.eval(
            f"(window[{tag!r}]&&window[{tag!r}].b64||'').length",
            await_promise=False, timeout=60) or 0)
        parts: list[str] = []
        offset = 0
        while offset < b64_len:
            piece = str(self.eval(
                f"(window[{tag!r}]&&window[{tag!r}].b64||'').slice({offset},{offset + chunk})",
                await_promise=False, timeout=90) or "")
            if not piece:
                break
            parts.append(piece)
            offset += chunk
        data = _b64.b64decode("".join(parts))
        if len(data) != int(meta.get("len") or 0):
            raise ZiniaoBridgeError(
                f"取回字节数不符：期望 {meta.get('len')}，实际 {len(data)}")
        self.log(f"页面内取回 {len(data)} 字节（HTTP {meta.get('status')}，"
                 f"{str(meta.get('type'))[:60]}）")
        return data

    def intercept_url(self, url_pattern: str, click_js: str,
                      timeout: float = 25.0, marker: str | None = None) -> list[str]:
        """点 ``click_js``，用 Fetch 域**在请求发出前**拦下匹配的请求并中止。

        为什么要中止：导出文件的下载 URL 带 ``use-disposable=1``（一次性），
        让浏览器真下载会把它消耗掉，之后任何重取都返回
        ``[CheckDisposableURL] out of use times``。拦在发送前，服务器根本没看到，
        URL 仍然可用 —— 我们再自己在页面里 fetch 它。
        """
        key = marker if marker is not None else url_pattern.strip("*")
        try:
            self.call("Fetch.enable",
                      {"patterns": [{"urlPattern": url_pattern,
                                     "requestStage": "Request"}]}, timeout=20)
        except Exception as exc:
            raise ZiniaoBridgeError(f"Fetch.enable 失败（内核可能不支持拦截）：{exc}") from exc

        caught: list[str] = []

        def on_event(msg: dict[str, Any]) -> None:
            if msg.get("method") != "Fetch.requestPaused":
                return
            params = msg.get("params") or {}
            url = str((params.get("request") or {}).get("url") or "")
            rid = str(params.get("requestId") or "")
            if not rid:
                return
            if key and key not in url:
                self.send("Fetch.continueRequest", {"requestId": rid})
                return
            caught.append(url)
            self.log(f"拦下下载请求（未发给服务器，URL 未消耗）：{url[:130]}")
            self.send("Fetch.failRequest", {"requestId": rid,
                                            "errorReason": "Aborted"})

        try:
            self.run_js(click_js)
            self.pump(timeout, on_event=on_event, stop=lambda evs: bool(caught))
        finally:
            try:
                self.call("Fetch.disable", {}, timeout=10)
            except Exception:
                pass
        return caught

    # -- cookie ------------------------------------------------------------
    def cookies_verbose(self, domain_keyword: str | None = TIKTOK_DOMAIN_KEY,
                        url: str | None = None) -> list[dict[str, Any]]:
        """原始 cookie 列表（**不合并**），可按 url 精确取。

        为什么优先 ``Network.getCookies(urls=[url])``：
        ``Storage.getCookies`` 返回的是**整个 profile 的全部 cookie**。统一商家
        环境下同一个账号可能同时有多个区域的店铺（``seller.us.*`` / ``seller.eu.*``），
        于是 ``oec_seller_id_unified_seller_env`` 这类同名 cookie 会**同时存在多个
        值**；合并成 dict 时「后写覆盖前写」→ 可能取到**另一个店铺**的 id
        （实测就用另一个店铺的身份去请求，全批 98001008 需要权限）。
        带 url 的 ``Network.getCookies`` 返回的是「浏览器真的会发给这个 url 的
        那一份」，按域名/路径精确匹配，不会串味。
        """
        items: list[dict[str, Any]] = []
        if url:
            try:
                self.call("Network.enable", {}, timeout=10)
            except Exception:
                pass  # 有些内核不 enable 也能 getCookies
            try:
                res = self.call("Network.getCookies", {"urls": [str(url)]},
                                timeout=20) or {}
                items = [dict(c) for c in (res.get("cookies") or [])]
                self.log(f"按域名精确取 cookie（{_host_tail(url) or url}）：{len(items)} 条")
            except Exception as exc:
                self.log(f"Network.getCookies 失败，退回 Storage.getCookies：{exc}")
                items = []
        if not items:
            try:
                res = self.call("Storage.getCookies", {}, timeout=20) or {}
                items = [dict(c) for c in (res.get("cookies") or [])]
            except Exception as exc:
                self.log(f"取 cookie 失败：{exc}")
                return []
        if not items:
            return []
        if domain_keyword:
            filtered = [c for c in items if domain_keyword in str(c.get("domain") or "")]
            if filtered:
                items = filtered
            else:
                self.log(f"没有域名包含 {domain_keyword!r} 的 cookie（共 {len(items)} 条），"
                         "可能店铺未登录该平台")
                return []
        return items

    def _merge_cookies(self, items: list[dict[str, Any]]) -> dict[str, str]:
        """原始 cookie 列表 → name→value；同名不同值会**显式告警**。"""
        out: dict[str, str] = {}
        dups: list[tuple[str, str, str, str]] = []
        for c in items:
            name = c.get("name")
            if not name:
                continue
            name = str(name)
            val = "" if c.get("value") is None else str(c.get("value"))
            if name in out and out[name] != val:
                dups.append((name, out[name], val, str(c.get("domain") or "")))
            out[name] = val
        if dups:
            self.log(f"[警告] {len(dups)} 个 cookie 同名不同值被覆盖（多店铺串味），"
                     "现在取到的是最后一条：")
            for name, old, new, dom in dups[:8]:
                self.log(f"        {name}: {old} → {new} @ {dom}")
            self.log("        → 店铺身份可能取错！请用 cookies(url=店铺后台地址) 精确取。")
        return out

    def cookies(self, domain_keyword: str | None = TIKTOK_DOMAIN_KEY,
                url: str | None = None) -> dict[str, str]:
        """取 cookie（含 HttpOnly）→ name→value。

        传了 ``url`` 就按「浏览器真会发给该 url 的那一份」取（推荐）；
        否则退化为 ``Storage.getCookies`` 全量合并，并对同名覆盖显式告警。
        """
        items = self.cookies_verbose(domain_keyword=domain_keyword, url=url)
        if not items:
            self.log("没取到任何 cookie（可能店铺未登录该平台）")
            return {}
        return self._merge_cookies(items)

    def store_context(self, url: str, expect: str = "",
                      timeout: float = 25.0, interval: float = 1.5) -> dict[str, Any]:
        """等「当前店铺」在 cookie 里稳定下来，并回报读到的全部候选值。

        店铺浏览器刚导航到后台时，前端还没把 ``oec_seller_id_unified_seller_env``
        刷成新店铺 —— 取太早会拿到**上一次**的店铺 id。这里轮询到「连续两次读到
        同一个非空值」或「命中 expect」才返回。

        返回 ``{"cookies": {...}, "seller_id": str, "stable": bool,
        "seen": [...], "candidates": [{"name","value","domain","path"}]}``
        """
        seen: list[str] = []
        prev = ""
        items: list[dict[str, Any]] = []
        cookies: dict[str, str] = {}
        deadline = time.time() + max(1.0, float(timeout))
        while True:
            items = self.cookies_verbose(url=url)
            cookies = self._merge_cookies(items) if items else {}
            cur = seller_id_from_cookies(cookies)
            if cur and cur not in seen:
                seen.append(cur)
                self.log(f"读到店铺候选 seller_id={cur}"
                         f"（{store_name_of(cur) or '未登记'}）")
            if cur and (cur == expect or cur == prev):
                return {"cookies": cookies, "seller_id": cur, "stable": True,
                        "seen": seen, "candidates": seller_candidates(items)}
            if time.time() >= deadline:
                self.log(f"[警告] 等店铺 cookie 稳定超时（{timeout}s）；"
                         f"已见候选：{seen or ['(空)']}")
                return {"cookies": cookies, "seller_id": cur, "stable": False,
                        "seen": seen, "candidates": seller_candidates(items)}
            prev = cur
            time.sleep(interval)

    # -- 兼容旧 DrissionPage 版的少量接口 -----------------------------------
    def tabs(self) -> list[dict[str, Any]]:
        return self.pages()

    def _working_tab(self) -> "BrowserSession":
        return self
def open_session(port: Any, log: LogFn | None = None,
                 home_urls: Iterable[str] | None = None,
                 retry: int = 20, interval: float = 1.0) -> BrowserSession:
    """接管店铺浏览器，返回可自愈的 BrowserSession。"""
    return BrowserSession(port, log=log, home_urls=home_urls,
                          retry=retry, interval=interval)


def connect_browser(port: Any, log: LogFn | None = None,
                    retry: int = 20, interval: float = 1.0,
                    home_urls: Iterable[str] | None = None) -> BrowserSession:
    """open_session 的旧名（保持向后兼容）。"""
    return open_session(port, log=log, home_urls=home_urls, retry=retry, interval=interval)


def _as_session(browser: Any) -> Any:
    """把入参归一成「有 navigate/tabs」的对象。"""
    return browser


def goto(browser: Any, url: str, log: LogFn | None = None,
         retry: int = 5, interval: float = 2.0) -> Any:
    """稳定地打开目标页面并返回该标签页对象（兼容 session / 裸 Chromium）。"""
    nav = getattr(browser, "navigate", None)
    if callable(nav):
        return nav(url, retry=retry, interval=interval)

    log = log or _noop
    last: Exception | None = None
    for i in range(1, max(1, retry) + 1):
        try:
            tabs = browser.get_tabs()
        except Exception as exc:
            tabs = []
            last = exc
        for t in tabs:
            try:
                if str(getattr(t, "url", "") or "").startswith("http"):
                    t.get(url, retry=1, interval=1, timeout=30)
                    return t
                break
            except Exception as exc:
                last = exc
                continue
        try:
            t = browser.new_tab(url)
            if t is not None:
                return t
        except Exception as exc:
            last = exc
            log(f"[{i}/{retry}] 打开页面失败：{exc}")
        time.sleep(interval)
    raise ZiniaoBridgeError(f"无法打开页面：{url}（{last}）")


def check_ip(browser: Any, ip_url: str, log: LogFn | None = None,
             timeout: float = 30.0) -> bool:
    """打开 ip 检测页，等「检测成功」按钮出现。

    这一步**失败不该中断流程**：检测页是紫鸟自己的扩展页，取不到按钮有很多原因
    （扩展页不允许 CDP 导航、页面改版、代理慢），此时仍然可以尝试后面的接口。
    """
    log = log or _noop
    if not ip_url:
        log("ipDetectionPage 为空，跳过 IP 检测")
        return False
    if not str(ip_url).startswith("http"):
        # 实测 ipDetectionPage 是 chrome-extension://…/index.html，
        # Page.navigate 到扩展页会被内核拒绝，白白等一轮超时，所以直接跳过。
        scheme = urllib.parse.urlsplit(str(ip_url)).scheme or "未知"
        log(f"IP 检测页不是 http 页面（{scheme}），跳过检测（不影响后续）")
        return False
    nav = getattr(browser, "navigate", None)
    if not callable(nav):
        log("当前浏览器对象不支持导航，跳过 IP 检测")
        return False
    try:
        nav(ip_url, retry=2, interval=2.0, wait=min(20.0, max(5.0, timeout)))
    except Exception as exc:
        log(f"IP 检测页打开失败（忽略，继续）：{exc}")
        return False

    js = ("(function(){const b=document.querySelector('.styles_btn--success');"
          "return b ? (b.textContent || 'ok') : '';})()")
    deadline = time.time() + max(1.0, float(timeout))
    while time.time() < deadline:
        try:
            if browser.eval(js, await_promise=False, timeout=10):
                log("IP 检测通过")
                return True
        except Exception:
            pass
        time.sleep(1)
    log("IP 检测未通过：未出现成功按钮（代理可能未就绪）")
    return False


def grab_cookies(browser: Any, domain_keyword: str = TIKTOK_DOMAIN_KEY,
                 log: LogFn | None = None, url: str | None = None) -> dict[str, str]:
    """抓 cookie（含 HttpOnly），返回 name→value。

    **强烈建议传 ``url`` = 店铺后台地址**：那样只取「浏览器真会发给该地址的那一份」，
    避免统一商家环境下多店铺同名 cookie 相互覆盖（实测会取到别的店铺的 seller_id）。
    """
    log = log or _noop
    fn = getattr(browser, "cookies", None)
    if callable(fn):
        if url:
            try:
                return fn(domain_keyword, url=url)
            except TypeError:
                pass
        try:
            return fn(domain_keyword)
        except TypeError:
            return fn()
    if isinstance(browser, dict):
        return {}

    items: list[dict[str, Any]] = []
    try:
        items = [dict(c) for c in browser.cookies(all_info=True)]
    except Exception as exc:
        log(f"浏览器级取 cookie 失败，退回逐标签页取：{exc}")
        for t in browser.get_tabs():
            try:
                items += [dict(c) for c in t.cookies(all_info=True)]
            except Exception:
                continue

    if not items:
        return {}
    if domain_keyword:
        filtered = [c for c in items if domain_keyword in str(c.get("domain") or "")]
        if filtered:
            items = filtered
    out: dict[str, str] = {}
    for c in items:
        name = c.get("name")
        if name:
            out[str(name)] = "" if c.get("value") is None else str(c.get("value"))
    return out


#: 「店铺 id 类」cookie 的名字，按优先级排列
SELLER_COOKIE_KEYS: tuple[str, ...] = (
    "oec_seller_id_unified_seller_env",
    "global_seller_id_unified_seller_env",
    "oec_seller_id",
    "seller_id",
)

#: 已知店铺 id → 店名（写日志/报错时把 id 翻译成人话）
KNOWN_STORES: dict[str, str] = {
    "7494876782645708274": "汉维2-圣诞树",
    "7496000161034505075": "嘉恒业",
}


def store_name_of(seller_id: Any) -> str:
    """已知店铺 id → 店名，未知返回空串。"""
    return KNOWN_STORES.get(str(seller_id or "").strip(), "")


def seller_id_from_cookies(cookies: dict[str, str]) -> str:
    """从 cookie 推导 seller_id / oec_seller_id。"""
    for key in SELLER_COOKIE_KEYS:
        val = cookies.get(key)
        if val and str(val).strip():
            return str(val).strip()
    return ""


def seller_candidates(items: Iterable[dict[str, Any]] | None) -> list[dict[str, str]]:
    """列出所有「店铺 id 类」cookie（含域名/路径），用于排查多店铺串味。

    返回形如 ``[{"name":..., "value":..., "domain":..., "path":...}]``。
    同一个店铺 id 出现多条属于正常（多域名共享）；**出现多个不同 value 就是串味**。
    """
    out: list[dict[str, str]] = []
    for c in items or []:
        name = str(c.get("name") or "")
        if name not in SELLER_COOKIE_KEYS:
            continue
        val = "" if c.get("value") is None else str(c.get("value")).strip()
        if not val:
            continue
        out.append({"name": name, "value": val,
                    "domain": str(c.get("domain") or ""),
                    "path": str(c.get("path") or "/")})
    return out


# ----------------------------------------------------------------------------
# 3. 店铺操作
# ----------------------------------------------------------------------------
def browser_key(store: Any) -> str:
    """取店铺的启动/关闭凭据（browserOauth 优先）。"""
    if isinstance(store, dict):
        return str(store.get("browserOauth") or store.get("browserId") or "")
    return str(store or "")


def store_label(store: dict[str, Any]) -> str:
    return str(store.get("browserName") or browser_key(store) or "未知店铺")


def close_store(client: Any, key: str, log: LogFn | None = None) -> None:
    log = log or _noop
    try:
        client.close_store(key)
        log(f"已关闭店铺：{key}")
    except Exception as exc:
        log(f"关闭店铺失败（忽略）：{exc}")


# ----------------------------------------------------------------------------
# 4. 浏览器内 fetch
# ----------------------------------------------------------------------------
class BrowserTransport:
    """在店铺浏览器页面里发起请求，返回解析后的 JSON。

    为什么不用本地 requests：店铺会话绑在店铺代理出口 IP 上，且 order/list 需要
    浏览器侧签名，本机直连恒定 code=10000。

    为什么需要「页面来源」：页面内 fetch 是跨域的，`content-type: application/json`
    会触发 CORS 预检；若当前页面不是接口白名单里的 origin（例如停在 ziniao 的
    ip 检测页、about:blank），浏览器只报 `TypeError: Failed to fetch`。
    因此这里按候选 origin 依次切换工作页面，并把**真正打通的那个组合**记下来。

    实测（2026-10-14 真机日志）——「页面来源」必须换到**正确的区域站**，
    光换 host 没用，路径还得跟站对上：

        seller.tiktokshopglobalselling.com + /api/fulfillment/order/get → 98001008
        seller.eu.tiktokshopglobalselling.com + 同一路径 → HTTP 200 "No matching route"
        seller.us.tiktokshopglobalselling.com + 同一路径 → code=0 ✅

    所以「同源」只在页面**真的停在那站**时成立 —— 命中后要把
    ``(写法, 页面实际 origin)`` 一起记（``winner_label`` / ``winner_origin``）；
    换站必须走 ``navigate`` 而不是 ``ensure_on``（后者把各区域站视作同站，不换）。

    调用约定与 tiktok_client 的 transport 一致：``fn(url, params, body) -> dict``。
    """

    def __init__(self, browser: Any, log: LogFn | None = None,
                 origins: Iterable[str] | None = None,
                 blocked_codes: Iterable[Any] = ("10000",),
                 use_page_origin: bool = True,
                 template: dict[str, Any] | None = None):
        self._browser = browser
        self.log = log or _noop
        self.origins = [str(o).rstrip("/") for o in (origins or []) if o]
        self.blocked_codes = {str(c) for c in blocked_codes}
        self.calls = 0
        # 实测教训：后台自己发的是**同源** ``seller.<region>.../api/...``，
        # 而不是抓包留档里的 ``api16-normal-no1a...``。用绝对 host 去页面内 fetch
        # 等于换了个域发请求，cookie / Origin / Referer 全都对不上 →
        # ``98001008``。所以默认把请求改写成「当前工作页面的 origin + 原路径」。
        self.use_page_origin = bool(use_page_origin)
        # 更进一步：用 capture_order_request.py 抓到的**真实请求**当模板，
        # 照抄它的 host、路径前缀（如北美的 /na/）和查询参数。
        self.template = template or None
        self._po_cache = ""
        self._last_url = ""
        # 页面**实际**所在的 origin（每次导航后实测）。可能和我们要去的地址不同 ——
        # 服务端会把 eu 地址重定向到 us，这时「同源」必须按实际值来拼。
        self.page_origin: str = ""
        # 命中过的组合：写法 **和** 当时页面实际所在的 origin，两个一起记才成立。
        # 实测：us 站同源 /api/fulfillment/order/get → code=0；
        # 换成 eu 站同源同一路径 → HTTP 200 "No matching route"。
        # 只记「同源」不记站 = 下次页面在 eu 就直接崩（踩过）。
        self.winner_label: str | None = None
        self.winner_origin: str = ""
        # 「来源 × 写法」全扫过一遍仍没打通 → 之后单发
        self.swept = False
        # 这些码说明「这次请求发得不对」—— 换写法再试（页面来源换过了就换 host）
        self.retry_codes = {str(c) for c in blocked_codes} | PERMISSION_CODES

    @property
    def tab(self):
        getter = getattr(self._browser, "_working_tab", None)
        if callable(getter):
            return getter()
        return getattr(self._browser, "latest_tab", self._browser)

    # -- 候选顺序 -----------------------------------------------------------
    def _order(self) -> list[str | None]:
        """候选页面来源：**已打通的站 → 当前站 → 其余**（接口主机不能当页面）。"""
        cands = [o for o in self.origins if not _is_api_host(o)]
        first = self.winner_origin or self.page_origin
        if not first or _is_api_host(first):
            first = ""
        rest = [o for o in cands if o != first]
        return (([first] if first else []) + rest) or [None]

    def _ask_page_origin(self) -> str:
        """问浏览器：工作页面**现在**在哪个 origin。"""
        try:
            cur = self._browser.current_url()
        except Exception:
            cur = ""
        return origin_of(cur)

    def _ensure_origin(self, origin: str | None) -> bool:
        """把工作页面切到指定 origin，返回「现在确实在这个站」。

        两个必须踩过的坑：

        1. **不能用 ``ensure_on``** —— 它按站点尾巴比较，会把
           ``seller.us.x.com`` / ``seller.eu.x.com`` / ``seller.x.com`` 当成同站，
           于是**从不真正换站**（页面停在通用站，我们却以为切过去了）。
           这里用 ``navigate``，并且按 host 严格复核。
        2. ``api16-*`` 是接口主机，不是页面 —— 拿它当页面来源只会拿到一堆
           JS 报错，直接跳过。
        """
        if origin and _is_api_host(origin):
            return False
        if origin is None:                      # 没有候选站：就在当前页面试
            return True
        cur = self.page_origin or self._ask_page_origin()
        if origin and cur and _same_host(cur, origin):
            self.page_origin = cur          # 已经在目标站，别多跳（导航会掉线）
            self._po_cache = ""
            return True
        if origin:
            nav = getattr(self._browser, "navigate", None)
            if callable(nav):
                nav(origin)
            else:
                fn = getattr(self._browser, "ensure_on", None)
                if callable(fn):
                    fn(origin)
            # 关键：记**实际**落地的站，服务端可能重定向（eu 地址 → us 页面）
            self.page_origin = self._ask_page_origin() or self.page_origin
            self._po_cache = ""
            return bool(self.page_origin and _same_host(self.page_origin, origin))
        return bool(cur)

    # -- 目标 URL：同源 / 照抄模板 ------------------------------------------
    def _page_origin(self) -> str:
        """当前工作页面的 origin（同源改写用）—— 始终以**实测**为准。"""
        if not self.page_origin:
            self.page_origin = self._ask_page_origin()
        self._po_cache = self.page_origin
        return self.page_origin

    def _from_template(self, url: str) -> tuple[str, dict[str, str]]:
        """按抓到的真实请求改写 host 与路径前缀，并带上它的查询参数。

        例：模板 ``seller.us.../api/fulfillment/na/order/list?...`` +
        请求 ``api16-normal-no1a..../api/fulfillment/order/get``
        → ``seller.us.../api/fulfillment/na/order/get?...``
        """
        tmpl = self.template or {}
        turl = str(tmpl.get("url") or "")
        if not turl:
            return "", {}
        tu = urllib.parse.urlsplit(turl)
        if not tu.scheme or not tu.netloc:
            return "", {}
        pu = urllib.parse.urlsplit(url)
        tail = pu.path.rstrip("/").rsplit("/", 1)[-1]
        tdir = tu.path.rstrip("/").rsplit("/", 1)[0]
        new_path = (tdir + "/" + tail) if tail else tu.path
        extra = {k: v for k, v in urllib.parse.parse_qsl(tu.query, keep_blank_values=True)
                 if k not in TEMPLATE_DROP_QUERY}
        return f"{tu.scheme}://{tu.netloc}{new_path}", extra

    def _variants(self, url: str,
                  params: dict[str, Any] | None = None) -> list[tuple[str, str]]:
        """列出这次要试的写法（按优先级），每条都带拼好的完整 URL。

          · ``模板`` —— 照抄 capture_order_request.py 抓到的真实请求（最可信）
          · ``同源`` —— 当前页面 origin + 原路径（页面自己就是这么发的）
          · ``绝对`` —— 抓包留档里的 api16-* host（某些店只有它通）

        一旦某种写法成功过，``self.winner`` 生效，后续只用它 —— 不会每次
        都白试一轮。
        """
        base: dict[str, Any] = dict(params or {})
        out: list[tuple[str, str]] = []

        def add(label: str, endpoint: str, extra: dict[str, Any] | None = None) -> None:
            q = dict(base)
            for k, v in (extra or {}).items():
                q.setdefault(k, v)
            qs = urllib.parse.urlencode({k: v for k, v in q.items() if v not in (None, "")})
            full = f"{endpoint}?{qs}" if qs else endpoint
            if all(full != f for _l, f in out):
                out.append((label, full))

        page = self._page_origin() if self.use_page_origin else ""
        pu = urllib.parse.urlsplit(url)
        turl, extra = self._from_template(url)

        order = [self.winner_label] if self.winner_label else ["模板", "同源", "绝对"]
        for label in order:
            if label == "模板" and turl:
                add("模板", turl, extra)
            elif label == "同源" and page and pu.path:
                add("同源", page.rstrip("/") + pu.path)
            elif label == "绝对":
                add("绝对", url)
        if not out:
            add("绝对", url)
        return out

    # -- JS ----------------------------------------------------------------
    def _js(self, full: str, body: dict[str, Any] | None) -> str:
        self._last_url = full
        payload = json.dumps(body or {}, ensure_ascii=False)
        # json.dumps 产出的双引号字符串是合法的 JS 字符串字面量
        return (
            "(async()=>{"
            f"const r=await fetch({json.dumps(full)},"
            "{method:'POST',headers:{'content-type':'application/json'},"
            f"credentials:'include',body:{json.dumps(payload)}}});"
            "const t=await r.text();"
            "return JSON.stringify({status:r.status,text:t});"
            "})()"
        )

    def _run(self, script: str) -> Any:
        runner = getattr(self._browser, "run_js", None)
        if callable(runner):
            return runner(script, timeout=120)
        return self.tab.run_js(script, as_expr=True, timeout=120)

    # -- 单次请求 -----------------------------------------------------------
    def _once(self, full: str, body: dict[str, Any] | None) -> dict[str, Any]:
        raw = self._run(self._js(full, body))
        self.calls += 1
        if isinstance(raw, str):
            try:
                raw = json.loads(raw)
            except Exception as exc:
                raise ZiniaoBridgeError(f"页面返回无法解析：{raw[:200]}（{exc}）")
        if not isinstance(raw, dict) or "text" not in raw:
            raise ZiniaoBridgeError(f"页面 fetch 返回异常：{str(raw)[:200]}")

        status = raw.get("status")
        text = str(raw.get("text") or "")
        preview = text if len(text) <= 400 else text[:200] + " … " + text[-120:]
        self.log(f"[浏览器内] POST {full} → HTTP {status} {preview}")

        if not text.strip():
            raise ZiniaoBridgeError(f"页面 fetch 空响应（HTTP {status}）")
        try:
            payload = json.loads(text)
        except Exception as exc:
            raise ZiniaoBridgeError(
                f"接口返回非 JSON（HTTP {status}）：{text[:200]}（{exc}）"
            )
        if not isinstance(payload, dict):
            raise ZiniaoBridgeError(f"接口返回结构异常：{text[:200]}")
        if not isinstance(status, int) or not (200 <= status < 300):
            # 传输层问题（页面未登录 / 被拦 / 跨域）才抛，交给链上的下一通道
            raise ZiniaoBridgeError(f"页面 fetch HTTP {status}：{text[:200]}")
        return payload

    def __call__(self, url: str, params: dict[str, Any] | None = None,
                 body: dict[str, Any] | None = None) -> dict[str, Any]:
        """依次尝试「页面来源 × URL 写法」。

        换页面来源解决 CORS / 页面站不对；换写法解决 host 与接口前缀不对
        （模板 / 同源 / 绝对 api16-*）。命中一次就记住，后续不再白试。
        """
        last: Exception | None = None
        last_payload: dict[str, Any] | None = None
        tried: set[str] = set()          # 同一次调用里，同一个 URL 不重复打
        cands = self._order()
        if self.swept:
            # 已经「来源 × 写法」全扫过一遍都没打通 → 之后单发一次就返回，
            # 免得 600 条订单把同一个错误重试上千次。
            cands = cands[:1]
        for i, origin in enumerate(cands):
            if origin and _is_api_host(origin):
                self.log(f"页面来源 {origin} 是接口主机（不是页面），跳过")
                continue
            try:
                self._ensure_origin(origin)
            except Exception as exc:
                last = exc
                self.log(f"切到页面来源 {origin} 失败：{exc}")
                continue
            # 注意：**不因「没到达指定站」而跳过**。这些区域站会互相重定向
            # （实测：打开 eu 地址，页面实际落在 us），此时「同源」按实测值拼
            # 反而正好打对 —— 真机就是靠这一步通的。

            variants = self._variants(url, params)
            if self.swept:
                variants = variants[:1]
            for j, (label, full) in enumerate(variants):
                if full in tried:
                    continue
                tried.add(full)
                more = j < len(variants) - 1
                try:
                    payload = self._once(full, body)
                except Exception as exc:      # 跨域被拦 / No matching route / 掉线
                    last = exc
                    self.log(f"[{label}] 请求失败{'，换写法' if more else ''}：{exc}")
                    continue

                code = payload.get("code")
                bad = code is not None and str(code) in self.retry_codes
                if bad:
                    # 这个写法没打通：先换写法，写法试完再换页面来源
                    last_payload = payload
                    last = ZiniaoBridgeError(f"[{label}] 返回 code={code}")
                    if more:
                        self.log(f"[{label}] 返回 code={code}（这个写法不对，换一个试）")
                    continue
                if (label != self.winner_label
                        or self.page_origin != self.winner_origin):
                    # 只有真正拿到业务数据才有资格被记住。
                    # 「写法」和「当时页面实际所在站」必须一起记 —— 同源写法
                    # 只在那个站的 host 下成立（us 通、eu 直接 No matching route）。
                    self.winner_label = label
                    self.winner_origin = self.page_origin or ""
                    self.log(f"✅ 锁定可用通道：{label} @ {self.page_origin or '当前页面'}")
                return payload

            if i < len(cands) - 1:
                self.log(f"在 {origin or '当前页面'} 上没打通，换页面来源 {cands[i + 1]}")

        if last_payload is not None:
            # 所有「来源 × 写法」都只拿到同一个业务码（例如 98001008）——
            # 原样返回，让上层抛带人话提示的 ApiError，而不是变成「传输失败」。
            self.swept = True
            return last_payload
        raise ZiniaoBridgeError(f"浏览器内 fetch 全部失败：{last}")


class RequestsTransport:
    """本地直连（requests），作为浏览器内 fetch 的兜底通道。

    实测：``order/get`` 本地直连可用（code=0）；``order/list`` 本地恒定
    code=10000，所以它只能救详情/客户信息那一段。
    """

    def __init__(self, cookies: dict[str, str], headers: dict[str, str],
                 log: LogFn | None = None, timeout: float = 30.0):
        import requests

        self.log = log or _noop
        self.timeout = timeout
        self._headers = dict(headers)
        self._session = requests.Session()
        self._session.trust_env = False  # 本地直连不要走系统代理
        self._session.cookies.update(cookies or {})

    def __call__(self, url: str, params: dict[str, Any] | None = None,
                 body: dict[str, Any] | None = None) -> dict[str, Any]:
        resp = self._session.post(url, params=params or {}, headers=self._headers,
                                  json=body or {}, timeout=self.timeout)
        text = resp.text
        preview = text if len(text) <= 400 else text[:200] + " … " + text[-120:]
        self.log(f"[本地直连] POST {url.split('/api/')[-1]} → HTTP {resp.status_code} {preview}")
        if not resp.ok:
            raise ZiniaoBridgeError(f"本地直连 HTTP {resp.status_code}：{text[:200]}")
        try:
            payload = resp.json()
        except Exception as exc:
            raise ZiniaoBridgeError(
                f"接口返回非 JSON（HTTP {resp.status_code}）：{text[:200]}（{exc}）"
            )
        if not isinstance(payload, dict):
            raise ZiniaoBridgeError(f"接口返回结构异常：{text[:200]}")
        return payload


class TransportChain:
    """按顺序尝试多个通道。

    切换规则：
      · **传输层失败**（异常 / 非 2xx / 跨域被拦）→ 换下一通道；
      · 业务响应里的 ``code`` 落在 ``fallthrough_codes``（默认 10000 = 被风控/限流）
        也换下一通道 —— 因为同一个请求换条路可能就通了；
      · 其它业务码（含 code=0）原样返回，不重复请求；
      · 如果所有通道都只拿到 fallthrough 的 code，就把最后一个响应返回，让调用方
        （``OrderClient._check``）抛出带人话提示的 ApiError。
    """

    def __init__(self, items: Iterable[tuple[str, Callable]], log: LogFn | None = None,
                 fallthrough_codes: Iterable[Any] = ("10000",)):
        self.items = list(items)
        self.log = log or _noop
        self.fallthrough = {str(c) for c in fallthrough_codes}
        self.used: str | None = None

    def __call__(self, url: str, params: dict[str, Any] | None = None,
                 body: dict[str, Any] | None = None) -> dict[str, Any]:
        last_exc: Exception | None = None
        last_payload: dict[str, Any] | None = None

        for idx, (name, fn) in enumerate(self.items):
            nxt = self.items[idx + 1][0] if idx + 1 < len(self.items) else None
            try:
                payload = fn(url, params, body)
            except Exception as exc:
                last_exc = exc
                self.log(f"通道「{name}」失败" + (f"，改用「{nxt}」" if nxt else "")
                         + f"：{exc}")
                continue

            code = payload.get("code") if isinstance(payload, dict) else None
            last_payload = payload if isinstance(payload, dict) else None
            if (nxt and code is not None and self.fallthrough
                    and str(code) in self.fallthrough):
                self.log(f"通道「{name}」返回 code={code}（疑似风控/限流），改用「{nxt}」")
                continue

            self.used = name
            return payload

        if last_payload is not None:
            # 结构上是成功的响应，交给 _check 抛带提示的 ApiError
            return last_payload
        raise ZiniaoBridgeError(f"所有通道都失败：{last_exc}")
