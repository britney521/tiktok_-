"""紫鸟浏览器 HTTP IPC 协议客户端 —— 跨平台 + DrissionPage。

把官方 demo 里散落在 `__main__` 中的协议实现抽出来，做成可复用的类：
- 去掉 `input()` 阻塞确认与 `exit()`，改成返回值 + 可注入的日志回调
- 修掉 macOS/Linux 下 download_path 被拼成反斜杠的 bug
- 进程名 / 启动命令按平台自动判断
- **浏览器接管改用 DrissionPage**：不走 WebDriver 协议、不需要 chromedriver，
  直接连 startBrowser 返回的 debuggingPort

典型用法：
    cfg = load_config()                    # config/config.yaml
    client = ZiniaoClient(cfg, log=print)
    client.kill_process()                  # 可选
    client.start_browser()
    client.update_core()
    stores = client.get_browser_list()
    ret = client.open_store(stores[0])
    driver = client.get_driver(ret)        # DpBrowser
    ...
"""

from __future__ import annotations

import hashlib
import json
import os
import platform
import shutil
import socket
import subprocess
import time
import traceback
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Literal

import requests

IS_WINDOWS = platform.system() == "Windows"
IS_MAC = platform.system() == "Darwin"
IS_LINUX = platform.system() == "Linux"

PROJECT_ROOT = Path(__file__).resolve().parent.parent
CONFIG_PATH = PROJECT_ROOT / "config" / "config.yaml"
LOG_DIR = PROJECT_ROOT / "logs"
LAUNCH_LOG = LOG_DIR / "ziniao_launch.log"

# 紫鸟 v6 的 HTTP-IPC 服务端口就绪检测默认超时（秒）
READY_TIMEOUT = 90.0

# Electron 应用被这两个环境变量“劫持”时会退化成纯 Node 进程，
# 从而把 --run_type 之类的参数当成 node 的参数直接报 bad option。
CHILD_ENV_BLOCKLIST = ("ELECTRON_RUN_AS_NODE", "NODE_OPTIONS", "ELECTRON_NO_ATTACH_CONSOLE")


def resolve_mac_app_executable(client_path: str) -> str:
    """把 macOS 上的客户端写法解析成真正的可执行文件路径。

    支持三种写法：
      - 已经是可执行文件        /Applications/ziniao.app/Contents/MacOS/ziniao
      - bundle 目录             /Applications/ziniao.app
      - 应用名                  ziniao
    解析失败返回空串。
    """
    raw = (client_path or "").strip()
    if not raw:
        return ""

    p = Path(raw).expanduser()
    if p.is_file():
        return str(p)

    # /Applications/ziniao.app 形式
    if raw.endswith(".app"):
        candidate = p / "Contents" / "MacOS" / p.stem
        return str(candidate) if candidate.is_file() else ""

    # 应用名形式 → 去常见目录里找 <name>.app
    name = Path(raw).name
    for base in (Path("/Applications"), Path.home() / "Applications"):
        candidate = base / f"{name}.app" / "Contents" / "MacOS" / name
        if candidate.is_file():
            return str(candidate)
    return ""


def default_user_data_dir() -> Path:
    """给 web_driver 实例一个独立的 user-data-dir。

    为什么必需：紫鸟 v6 在启动时判断
        `if (hasUserDataDir || cliServiceMode) { 继续启动 } else { relaunching GUI instance }`
    不带 `--user-data-dir` 时它只会转发给 GUI 实例，**永远不会创建 HTTP IPC 端口**，
    于是 updateCore / getBrowserList 全部拿不到响应。
    独立目录同时保证自动化实例与手动打开的客户端互不干扰。
    """
    if IS_WINDOWS:
        base = Path(os.getenv("LOCALAPPDATA") or (Path.home() / "AppData" / "Local"))
        base = base / "ziniaobrowser"
    elif IS_MAC:
        base = Path.home() / "Library" / "Application Support" / "ziniaobrowser"
    else:
        base = Path(os.getenv("XDG_CONFIG_HOME") or (Path.home() / ".config")) / "ziniaobrowser"
    return base / "instances" / "ziniao_webdriver"


LogFn = Callable[[str], None]


def _print_log(msg: str) -> None:
    print(msg, flush=True)


# ----------------------------------------------------------------------------
# 运行时配置
# ----------------------------------------------------------------------------
@dataclass
class ZiniaoConfig:
    """config.yaml 的强类型视图。"""

    version: str = "v6"
    client_path: str = "ziniao"
    webdriver_path: str = ""
    socket_port: int = 16851
    company: str = ""
    username: str = ""
    password: str = ""
    # web_driver 实例的独立数据目录（紫鸟 v6 必填，留空则自动推导）
    user_data_dir: str = ""

    @classmethod
    def from_dict(cls, cfg: dict[str, Any]) -> "ZiniaoConfig":
        browser = (cfg.get("ziniao") or {}).get("browser") or {}
        user = (cfg.get("ziniao") or {}).get("user_info") or {}
        return cls(
            version=str(browser.get("version") or "v6"),
            client_path=str(browser.get("client_path") or ""),
            webdriver_path=str(browser.get("webdriver_path") or ""),
            socket_port=int(browser.get("socket_port") or 16851),
            company=str(user.get("company") or ""),
            username=str(user.get("username") or ""),
            password=str(user.get("password") or ""),
            user_data_dir=str(browser.get("user_data_dir") or ""),
        )

    @classmethod
    def from_yaml(cls, path: Path | str = CONFIG_PATH) -> "ZiniaoConfig":
        import yaml

        p = Path(path)
        if not p.exists():
            return cls()
        with p.open("r", encoding="utf-8") as f:
            return cls.from_dict(yaml.safe_load(f) or {})

    def to_dict(self) -> dict[str, Any]:
        return {
            "ziniao": {
                "browser": {
                    "version": self.version,
                    "client_path": self.client_path,
                    "webdriver_path": self.webdriver_path,
                    "socket_port": self.socket_port,
                    "user_data_dir": self.user_data_dir,
                },
                "user_info": {
                    "company": self.company,
                    "username": self.username,
                    "password": self.password,
                },
            }
        }

    def save_yaml(self, path: Path | str = CONFIG_PATH) -> Path:
        """把当前配置写回 config.yaml（保留 task 段）。"""
        import yaml

        p = Path(path)
        existing: dict[str, Any] = {}
        if p.exists():
            with p.open("r", encoding="utf-8") as f:
                existing = yaml.safe_load(f) or {}
        merged = self.to_dict()
        if "task" in existing:
            merged["task"] = existing["task"]
        p.parent.mkdir(parents=True, exist_ok=True)
        with p.open("w", encoding="utf-8") as f:
            yaml.safe_dump(merged, f, allow_unicode=True, sort_keys=False)
        return p

    def default_client_path(self) -> str:
        """按平台给出客户端路径占位。"""
        if self.client_path:
            return self.client_path
        if IS_MAC:
            return "ziniao"
        if IS_WINDOWS:
            return r"D:\ziniao\ziniao.exe"
        return "/opt/ziniao/ziniaobrowser"

    def client_executable(self) -> str:
        """真正要执行的文件路径（macOS 会把 .app / 应用名解析成内层可执行文件）。"""
        raw = self.default_client_path()
        if IS_MAC:
            return resolve_mac_app_executable(raw)
        return raw if Path(raw).is_file() else ""

    def resolved_user_data_dir(self) -> Path:
        """web_driver 实例的数据目录（配置留空则自动推导），并确保已创建。"""
        p = Path(self.user_data_dir).expanduser() if self.user_data_dir else default_user_data_dir()
        p.mkdir(parents=True, exist_ok=True)
        return p

    def missing(self) -> list[str]:
        return [k for k in ("company", "username", "password") if not getattr(self, k)]


class ZiniaoError(RuntimeError):
    """紫鸟 IPC 返回非 0 statusCode 时抛出。"""

    def __init__(self, action: str, status_code: Any, raw: Any = None):
        self.action = action
        self.status_code = status_code
        self.raw = raw
        detail = ""
        if isinstance(raw, dict):
            detail = raw.get("message") or raw.get("msg") or ""
        super().__init__(f"[{action}] statusCode={status_code} {detail}".strip())


# ----------------------------------------------------------------------------
# 客户端
# ----------------------------------------------------------------------------
class ZiniaoClient:
    """紫鸟浏览器 HTTP-IPC 协议客户端。

    driver_factory 可注入，便于无客户端环境下做桩测试。
    """

    def __init__(
        self,
        config: ZiniaoConfig,
        log: LogFn | None = None,
        driver_factory: Callable[[dict[str, Any]], Any] | None = None,
        http_timeout: float = 120.0,
    ):
        self.config = config
        self.log: LogFn = log or _print_log
        self.http_timeout = http_timeout
        self.socket_port = config.socket_port
        self.driver_factory = driver_factory
        self._session = requests.Session()
        # 127.0.0.1 的本地 IPC 绝不能走系统/环境代理。
        # mac 上开着 Clash/Surge（系统代理 HTTPEnable=1）时，
        # requests 默认会把 http://127.0.0.1:16851 交给代理，代理返回 502，
        # 症状就是「启动环境卡住 / 获取店铺列表失败」。
        self._session.trust_env = False

    # -- 基础 ---------------------------------------------------------------
    @property
    def user_info(self) -> dict[str, str]:
        return {
            "company": self.config.company,
            "username": self.config.username,
            "password": self.config.password,
        }

    def send_http(self, data: dict[str, Any]) -> dict[str, Any] | None:
        """向本地客户端 IPC 端口 POST 一条指令，失败返回 None（与原 demo 一致）。"""
        try:
            url = f"http://127.0.0.1:{self.socket_port}"
            resp = self._session.post(
                url, data=json.dumps(data).encode("utf-8"), timeout=self.http_timeout,
            )
            return json.loads(resp.text)
        except Exception as err:
            self.log(f"IPC 通讯异常：{err}")
            return None

    def _call(self, action: str, extra: dict[str, Any] | None = None, fatal: bool = False) -> dict[str, Any]:
        """发一条带 user_info 的指令，校验 statusCode。"""
        data: dict[str, Any] = {"action": action, "requestId": str(uuid.uuid4())}
        if extra:
            data.update(extra)
        data.update(self.user_info)

        result = self.send_http(data)
        if result is None:
            raise ZiniaoError(action, "no-response", None)
        code = result.get("statusCode")
        if str(code) == "0":
            return result
        if str(code) == "-10003":
            raise ZiniaoError(action, code, result)  # 登录信息错误
        if not fatal:
            self.log(f"{action} 返回异常：{json.dumps(result, ensure_ascii=False)}")
        raise ZiniaoError(action, code, result)

    def _try(self, action: str, extra: dict[str, Any] | None = None) -> dict[str, Any] | None:
        """容错版 _call：失败只记日志，返回 None。"""
        try:
            return self._call(action, extra)
        except ZiniaoError as exc:
            self.log(f"{action} 失败：{exc}")
            return None

    # -- 进程与启动 ---------------------------------------------------------
    @staticmethod
    def process_name(version: str) -> str:
        if IS_WINDOWS:
            return "SuperBrowser.exe" if version == "v5" else "ziniao.exe"
        return "ziniaobrowser" if IS_LINUX else "ziniao"

    def port_open(self, port: int | None = None) -> bool:
        port = port or self.socket_port
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(0.5)
            return s.connect_ex(("127.0.0.1", port)) == 0

    def wait_port(self, port: int | None = None, timeout: float = READY_TIMEOUT,
                  interval: float = 0.5) -> bool:
        """轮询等待本地 IPC 端口可连接。"""
        end = time.time() + timeout
        while time.time() < end:
            if self.port_open(port):
                return True
            time.sleep(interval)
        return False

    def kill_process(self, version: str | None = None, wait_release: float = 20.0) -> bool:
        """终止紫鸟客户端及其子进程，并等待 IPC 端口释放。

        注意：macOS 下 `killall ziniao` 只能干掉主进程，Electron 的 Helper 进程和
        `env-kit/agentic/node` 子进程会残留成孤儿，继续占着端口。所以这里按命令行
        整棵树清理，否则新实例起不来。
        """
        version = version or self.config.version
        name = self.process_name(version)
        try:
            if IS_WINDOWS:
                os.system("taskkill /f /t /im " + name + " >nul 2>&1")
            else:
                os.system("pkill -f 'ziniao.app/Contents/MacOS/ziniao' >/dev/null 2>&1")
                os.system("pkill -f 'ziniaobrowser' >/dev/null 2>&1")
                os.system(f"pkill -x '{name.replace('.exe', '')}' >/dev/null 2>&1")

            # 等端口真正释放，避免刚杀完就启动导致端口冲突
            end = time.time() + wait_release
            while time.time() < end and self.port_open():
                time.sleep(0.5)
            self.log(f"已终止紫鸟客户端进程：{name}")
            return True
        except Exception:
            self.log("终止进程失败：" + traceback.format_exc())
            return False

    @staticmethod
    def _child_env() -> dict[str, str]:
        """构造传给客户端子进程的干净环境。"""
        env = dict(os.environ)
        # 如果 ELECTRON_RUN_AS_NODE 被继承进来，Electron 会退化成纯 Node，
        # 把 --run_type/--ipc_type 当成 node 自己的参数，直接 `bad option` 退出。
        for key in CHILD_ENV_BLOCKLIST:
            env.pop(key, None)
        return env

    def launch_command(self) -> list[str]:
        """拼出完整的启动命令（便于日志与自检）。"""
        exe = self.config.client_executable()
        args = ["--run_type=web_driver", "--ipc_type=http", f"--port={self.socket_port}"]
        if self.config.version == "v6":
            args.append(f"--user-data-dir={self.config.resolved_user_data_dir()}")
        if IS_LINUX:
            args.insert(0, "--no-sandbox")
        return [exe or "<未找到客户端可执行文件>", *args]

    def start_browser(self) -> bool:
        """以 web_driver + http IPC 模式启动客户端，并等到端口就绪。

        走的是直接执行 bundle 内可执行文件，而不是 `open -a`：
        `open` 在客户端刚退出/仍在运行时经常只做 activate，把 `--args` 丢掉，
        结果参数没生效、端口也不开。
        """
        exe = self.config.client_executable()
        if not exe:
            self.log(
                f"找不到紫鸟客户端可执行文件（client_path={self.config.client_path!r}）。\n"
                "macOS 可填：ziniao 或 /Applications/ziniao.app 或 "
                "/Applications/ziniao.app/Contents/MacOS/ziniao"
            )
            return False

        cmd = self.launch_command()
        try:
            LOG_DIR.mkdir(parents=True, exist_ok=True)
            with LAUNCH_LOG.open("a", encoding="utf-8") as out:
                out.write(f"\n===== {time.strftime('%F %T')} launch =====\n")
                out.write(" ".join(cmd) + "\n")
                out.flush()
                subprocess.Popen(  # noqa: S603
                    cmd,
                    env=self._child_env(),
                    stdout=out,
                    stderr=subprocess.STDOUT,
                    start_new_session=True,
                )
        except Exception:
            self.log("启动客户端失败：" + traceback.format_exc())
            return False

        self.log(f"已启动客户端，等待 IPC 端口 {self.socket_port}…")
        if self.wait_port(self.socket_port, timeout=READY_TIMEOUT):
            self.log(f"IPC 端口 {self.socket_port} 已就绪")
            return True

        self.log(
            f"等待端口 {self.socket_port} 超时，客户端没能进入 web_driver 模式。可能原因：\n"
            f"  1) v6 的 --user-data-dir 被改坏（当前：{self.config.resolved_user_data_dir()}）\n"
            "  2) 客户端版本低于 5.285.7，不支持 http IPC\n"
            f"  3) 端口 {self.socket_port} 被别的程序占用，换个 socket_port\n"
            f"  4) 启动详情见 {LAUNCH_LOG}"
        )
        return False

    def update_core(self, on_progress: LogFn | None = None,
                    timeout: float = READY_TIMEOUT + 180) -> bool:
        """下载/更新所有内核。HTTP 有超时，需循环调用直到成功（这里带总超时兜底）。"""
        cb = on_progress or self.log
        deadline = time.time() + timeout
        while True:
            result = self.send_http({"action": "updateCore", "requestId": str(uuid.uuid4()),
                                     **self.user_info})
            if result is None:
                if time.time() > deadline:
                    cb("更新内核超时：客户端 IPC 始终无响应（端口未就绪 / 版本过旧 / 代理拦截）")
                    return False
                cb("等待客户端启动…")
                time.sleep(2)
                continue
            code = result.get("statusCode")
            if code is None or str(code) == "-10003":
                cb("当前版本不支持此接口，请升级客户端")
                return False
            if str(code) == "0":
                cb("更新内核完成")
                return True
            cb(f"等待更新内核：{json.dumps(result, ensure_ascii=False)}")
            time.sleep(2)

    def get_exit(self) -> None:
        """关闭客户端。"""
        self._try("exit")

    # -- 店铺 ---------------------------------------------------------------
    def get_browser_list(self) -> list[dict[str, Any]]:
        result = self._call("getBrowserList")
        return result.get("browserList") or []

    def open_store(self, store: str | dict[str, Any], is_webdriver_read_only: int = 0,
                   is_privacy: int = 0, is_headless: int = 0,
                   cookie_type_save: int = 0, js_info: str = "") -> dict[str, Any]:
        """打开店铺，返回 startBrowser 的原始响应。store 可传 ID 或整条店铺对象。"""
        if isinstance(store, dict):
            store_key = store.get("browserOauth") or store.get("browserId") or ""
        else:
            store_key = str(store)
        if not store_key:
            raise ZiniaoError("startBrowser", "empty-store", {"store": store})

        extra: dict[str, Any] = {
            "isWaitPluginUpdate": 0,
            "isHeadless": is_headless,
            "isWebDriverReadOnlyMode": is_webdriver_read_only,
            "cookieTypeLoad": 0,
            "cookieTypeSave": cookie_type_save,
            "runMode": "1",
            "isLoadUserPlugin": False,
            "pluginIdType": 1,
            "privacyMode": is_privacy,
            "notPromptForDownload": 1,
        }
        if str(store_key).isdigit():
            extra["browserId"] = str(store_key)
        else:
            extra["browserOauth"] = str(store_key)
        if len(str(js_info)) > 2:
            extra["injectJsInfo"] = json.dumps(js_info)

        return self._call("startBrowser", extra)

    def close_store(self, browser_oauth: str) -> None:
        self._try("stopBrowser", {"duplicate": 0, "browserOauth": browser_oauth})

    # -- 驱动 ---------------------------------------------------------------
    def download_driver(self) -> bool:
        """无需 chromedriver。

        DrissionPage 直接接管紫鸟已启动 Chromium 的调试端口，不走 WebDriver 协议，
        所以这里不再下载任何驱动。保留该方法只是为了让旧调用点不用改。
        """
        self.log("使用 DrissionPage 接管浏览器，无需下载 chromedriver，跳过")
        return True

    @staticmethod
    def _sha1(path: Path) -> str:
        with open(path, "rb") as f:
            return hashlib.new("sha1", f.read()).hexdigest()

    def _download_file(self, url: str, save_path: Path) -> None:
        resp = requests.get(url, stream=True, timeout=120)
        if resp.status_code != 200:
            self.log(f"下载失败，状态码：{resp.status_code}")
            return
        with open(save_path, "wb") as f:
            for chunk in resp.iter_content(chunk_size=1024):
                if chunk:
                    f.write(chunk)
        self.log(f"已下载并保存：{save_path}")

    def get_browser(self, port: Any) -> Any:
        """连接紫鸟已启动浏览器的调试端口，返回 DpBrowser。"""
        from DrissionPage import Chromium

        from util.dp import DpBrowser

        chromium = Chromium(int(port))
        return DpBrowser(chromium)

    def get_driver(self, open_ret: dict[str, Any]) -> Any:
        """根据 startBrowser 响应接管浏览器。

        DrissionPage 不需要 chromedriver/chromedriver 路径，只要 debuggingPort。
        """
        if self.driver_factory is not None:
            return self.driver_factory(open_ret)

        port = open_ret.get("debuggingPort")
        core_type = open_ret.get("core_type")
        if core_type not in (None, "Chromium", 0):
            self.log(f"注意：当前内核为 {core_type}，DrissionPage 仅支持 Chromium 内核")

        if not port:
            self.log("startBrowser 未返回 debuggingPort，无法接管浏览器")
            return None

        self.log(f"接管浏览器调试端口：127.0.0.1:{port}")
        try:
            return self.get_browser(port)
        except Exception:
            self.log("接管浏览器失败：" + traceback.format_exc())
            return None

    # -- 页面辅助 -----------------------------------------------------------
    @staticmethod
    def _as_tab(browser: Any) -> Any:
        """拿到可执行页面操作的标签页对象（兼容 DpBrowser / Chromium / Tab）。"""
        if hasattr(browser, "latest_tab"):
            return browser.latest_tab
        return browser

    def open_ip_check(self, browser: Any, ip_check_url: str, timeout: float = 60) -> bool:
        """打开 ip 检测页，等待「检测成功」按钮出现。"""
        try:
            tab = self._as_tab(browser)
            tab.get(ip_check_url)
            ele = tab.ele(
                'xpath://button[contains(@class, "styles_btn--success")]',
                timeout=timeout,
            )
            if ele:
                self.log("ip 检测通过")
                return True
            self.log("ip 检测超时：未出现成功按钮")
            return False
        except Exception:
            self.log("ip 检测异常：" + traceback.format_exc())
            return False

    def open_launcher_page(self, browser: Any, launcher_page: str, wait: float = 6.0) -> None:
        tab = self._as_tab(browser)
        tab.get(launcher_page)
        time.sleep(wait)

    # -- 其它 ---------------------------------------------------------------
    @staticmethod
    def normalize_download_path(path: str) -> str:
        """跨平台修正下载目录结尾分隔符（原 demo 在 mac/linux 上会拼成反斜杠）。"""
        if not path:
            return str(PROJECT_ROOT / "downloads")
        if not (path.endswith("/") or path.endswith("\\")):
            path += os.sep
        return path

    @staticmethod
    def delete_all_cache() -> None:
        """删除所有店铺缓存（非必要，仅磁盘紧张时用）。"""
        if not IS_WINDOWS:
            return
        local_appdata = os.getenv("LOCALAPPDATA")
        if not local_appdata:
            return
        cache_path = os.path.join(local_appdata, "SuperBrowser")
        if os.path.exists(cache_path):
            shutil.rmtree(cache_path)
