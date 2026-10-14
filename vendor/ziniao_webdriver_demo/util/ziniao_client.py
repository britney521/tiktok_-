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
import ntpath  # Windows 路径解析；跨平台可用，便于在 macOS 上离线测试 Windows 分支
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


# ----------------------------------------------------------------------------
# Windows 客户端定位
# ----------------------------------------------------------------------------
# 为什么要这么麻烦：紫鸟 Windows 版的安装目录是**安装时用户自选**的，没有固定
# 位置。真机见到的就有
#     D:\紫鸟\ziniao\ziniao.exe
#     C:\Program Files\ZiniaoBrowser\ZiniaoBrowser.exe
#     C:\Program Files\ziniao\ziniao.exe
# 而官方 demo 的示例配置写的是 `D:\ziniao\ziniao.exe` —— 照抄必然找不到，报
# 「找不到紫鸟客户端可执行文件」。macOS 上一句 `killall ziniao` 就够，Windows
# 上必须按下面几级去探。
#
# 探测顺序（可信度递减）：
#   1. 环境变量 ZINIAO_EXE_PATH（影刀等 RPA 平台的通行做法）
#   2. 正在运行的客户端进程的真实 exe 路径（用户一般已经开着紫鸟）
#   3. 注册表卸载信息 / App Paths
#   4. 常见安装目录 + 各盘根浅层扫 *ziniao* / *紫鸟*
# 找到即缓存，避免每次读配置都扫一遍盘。
_WIN_EXE_NAMES = ("ziniao.exe", "ZiniaoBrowser.exe", "SuperBrowser.exe")
_WIN_ENV_VARS = ("ZINIAO_EXE_PATH", "ZINIAO_CLIENT_PATH", "ZINIAO_PATH")
# 通配兜底时要排掉的名字（卸载器 / 更新器 / 安装包 / Electron 子进程）
_WIN_EXE_SKIP = ("unins", "update", "upgrade", "setup", "install", "helper",
                 "crashpad", "uninstall", "elevate")
# 目录名同时匹配 "ziniao" 与中文「紫鸟」两种写法
_WIN_DIR_KEYWORDS = ("ziniao", "紫鸟")
_WIN_DIR_HINTS = (
    r"C:\Program Files\ziniao",
    r"C:\Program Files (x86)\ziniao",
    r"C:\Program Files\ZiniaoBrowser",
    r"C:\Program Files (x86)\ZiniaoBrowser",
    r"C:\Program Files\Ziniao Browser",
    r"C:\Program Files (x86)\Ziniao Browser",
    r"C:\Program Files\紫鸟浏览器",
    r"C:\Program Files (x86)\紫鸟浏览器",
    r"D:\紫鸟\ziniao",
    r"D:\紫鸟",
    r"D:\ziniao",
    r"D:\紫鸟浏览器\ziniao",
    r"C:\紫鸟\ziniao",
    r"C:\ziniao",
)

# 文件系统钩子抽成模块级变量：便于在非 Windows 上做离线测试（注入假目录树）。
_FS_ISFILE = os.path.isfile
_FS_ISDIR = os.path.isdir


def _listdir(path: str) -> list[str]:
    try:
        return sorted(os.listdir(path))
    except OSError:
        return []


def _name_matches(name: str) -> bool:
    low = name.lower()
    return any(k in low for k in _WIN_DIR_KEYWORDS)


def _win_exe_in_dir(directory: str) -> str:
    """在目录里找客户端 exe（先按已知文件名，再按名字含 ziniao/紫鸟 的 exe 兜底）。"""
    if not directory or not _FS_ISDIR(directory):
        return ""
    for name in _WIN_EXE_NAMES:
        candidate = ntpath.join(directory, name)
        if _FS_ISFILE(candidate):
            return candidate
    for name in _listdir(directory):
        low = name.lower()
        # 排除卸载/更新/安装包这类同前缀的 exe，别把 uninstaller 当客户端
        if not low.endswith(".exe") or any(k in low for k in _WIN_EXE_SKIP):
            continue
        if _name_matches(name):
            return ntpath.join(directory, name)
    return ""


def _win_running_client_paths() -> list[str]:
    """问系统要正在运行的客户端 exe 路径（PowerShell 优先，wmic 兜底）。"""
    out: list[str] = []
    create_no_window = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    for proc in ("ziniao", "ZiniaoBrowser", "SuperBrowser"):
        ps = (
            f"Get-Process -Name {proc} -ErrorAction SilentlyContinue | "
            "Where-Object { $_.Path } | Select-Object -First 1 -ExpandProperty Path"
        )
        try:
            r = subprocess.run(  # noqa: S603
                ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
                capture_output=True, text=True, timeout=15,
                creationflags=create_no_window,
            )
            for line in (r.stdout or "").splitlines():
                p = line.strip()
                if p and _FS_ISFILE(p):
                    out.append(p)
        except Exception:
            continue
        if out:
            break
    return out


def _reg_str(key, name: str) -> str:  # noqa: ANN001
    try:
        import winreg  # noqa: PLC0415

        value, _ = winreg.QueryValueEx(key, name)
        return str(value or "").strip()
    except Exception:
        return ""


def _win_registry_client_paths() -> list[str]:
    """从注册表卸载信息与 App Paths 里挖客户端路径。"""
    try:
        import winreg  # noqa: PLC0415
    except ImportError:
        return []
    out: list[str] = []
    uninstall = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    for hive, sub in (
        (winreg.HKEY_LOCAL_MACHINE, uninstall),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_CURRENT_USER, uninstall),
    ):
        try:
            with winreg.OpenKey(hive, sub) as key:
                count = winreg.QueryInfoKey(key)[0]
                for i in range(count):
                    try:
                        name = winreg.EnumKey(key, i)
                        with winreg.OpenKey(key, name) as sk:
                            display = _reg_str(sk, "DisplayName")
                            if not display or not _name_matches(display):
                                continue
                            for field in ("InstallLocation", "DisplayIcon"):
                                val = _reg_str(sk, field)
                                # DisplayIcon 常是 "D:\...\ziniao.exe,0"，去掉图标序号
                                if val:
                                    out.append(val.strip('"').split(",")[0].strip())
                    except OSError:
                        continue
        except OSError:
            continue

    app_paths = r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\ziniao.exe"
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            with winreg.OpenKey(hive, app_paths) as key:
                out.append(_reg_str(key, ""))
        except OSError:
            pass
    return [p for p in out if p]


def _win_dir_hints() -> list[str]:
    """常见安装目录 + 按环境变量推出来的用户级目录。"""
    hints = list(_WIN_DIR_HINTS)
    for env in ("LOCALAPPDATA", "APPDATA", "USERPROFILE", "ProgramData"):
        base = (os.getenv(env) or "").strip()
        if not base:
            continue
        for leaf in (("ziniao",), ("ZiniaoBrowser",), ("Programs", "ziniao"),
                     ("Programs", "ZiniaoBrowser"), ("紫鸟", "ziniao")):
            hints.append(ntpath.join(base, *leaf))
    return hints


def _win_drive_scan() -> list[str]:
    """盘根浅扫：只看根目录下名字含 ziniao/紫鸟 的目录（及其 zinia? 子目录）。"""
    for letter in "CDEFGH":
        root = f"{letter}:\\"
        if not _FS_ISDIR(root):
            continue
        hits: list[str] = []
        for name in _listdir(root):
            if not _name_matches(name):
                continue
            base = ntpath.join(root, name)
            found = _win_exe_in_dir(base)
            if not found:
                for sub in _listdir(base):
                    if _name_matches(sub):
                        found = _win_exe_in_dir(ntpath.join(base, sub))
                        if found:
                            break
            if found:
                hits.append(found)
        if hits:
            return hits
    return []


_win_client_cache: dict[str, Any] = {"done": False, "exe": "", "tried": []}


def windows_client_candidates(force: bool = False) -> list[tuple[str, str]]:
    """返回 ``(候选路径, 来源说明)``，按可信度排序（Windows 之外返回空）。"""
    if not IS_WINDOWS:
        return []
    cands: list[tuple[str, str]] = []

    def add(path: str, why: str) -> None:
        path = (path or "").strip().strip('"')
        if path:
            cands.append((path, why))

    for key in _WIN_ENV_VARS:
        add(os.getenv(key, ""), f"环境变量 {key}")
    for path in _win_running_client_paths():
        add(path, "正在运行的客户端进程")
    for path in _win_registry_client_paths():
        add(path, "注册表")
    for path in _win_dir_hints():
        add(path, "常见安装目录")
    for path in _win_drive_scan():
        add(path, "盘根扫描")
    return cands


def find_windows_client(force: bool = False) -> str:
    """在 Windows 上自动定位紫鸟客户端 exe；找不到返回空串。"""
    if not IS_WINDOWS:
        return ""
    if _win_client_cache["done"] and not force:
        return str(_win_client_cache["exe"])

    exe, tried = "", []
    for path, why in windows_client_candidates(force=force):
        tried.append(f"{path}  ({why})")
        hit = _win_exe_in_dir(path) or (path if _FS_ISFILE(path) else "")
        if hit:
            exe = hit
            break

    _win_client_cache.update(done=True, exe=exe, tried=tried)
    return exe


def windows_client_tried() -> list[str]:
    """上次探测尝试过的位置（用于报错时列给用户看）。"""
    return list(_win_client_cache.get("tried") or [])


def resolve_windows_client_executable(client_path: str = "", force: bool = False) -> str:
    """把配置里的客户端写法解析成 Windows 上真正可执行的 exe 路径。

    支持这些写法（全部实测过）：

      ==============================  ===================================
      写法                            解析结果
      ==============================  ===================================
      ``D:\\紫鸟\\ziniao\\ziniao.exe``   原样
      ``D:\\紫鸟\\ziniao``               目录 → 里面的 ziniao.exe
      ``D:\\紫鸟``                       上一层目录 → 再往里走一层
      ``ziniao`` / ``ziniao.exe``      裸名 → 自动探测
      ``"D:\\...\\ziniao.exe"``          带引号（从属性面板复制的样子）
      ``d:/紫鸟/ziniao/ziniao.exe``      正斜杠 / 大小写随意
      ``D:\\紫鸟\\ziniao``（其实该目录下
      只有一个 ``ziniao.exe``）          见下面「省略 .exe」那条
      留空                             自动探测
      ==============================  ===================================

    ★ 「省略 .exe」为什么要单独处理 ★
    Windows 资源管理器**默认隐藏已知扩展名**，用户点开桌面快捷方式的「属性 → 目标」
    抄下来的是 ``D:\\紫鸟\\ziniao`` —— 看着像目录/无扩展名文件，实际那个文件叫
    ``D:\\紫鸟\\ziniao.exe``。照抄进「客户端路径」后，``isfile`` 和 ``isdir`` 都是
    False，于是被当成「路径写错」落到自动探测；如果这个目录又不在常见位置表里
    （比如装在 ``E:\\soft\\ziniao.exe``），自动探测也找不到 → 报「找不到客户端」。
    真机上就是这么卡的。所以这里补一步 ``raw + ".exe"``。

    明确写了路径却不存在时，**回落到自动探测**而不是报错 —— 配置常常是从别的
    机器抄过来的，换个安装位置就失效，不该让整个流程卡死。
    """
    raw = (client_path or "").strip().strip('"').strip()
    if raw:
        raw = os.path.expandvars(raw)

    if raw:
        # ① 已经是 exe 全路径。同时容忍「省略 .exe」的写法（见 docstring）。
        for cand in (raw, raw + ".exe"):
            if _FS_ISFILE(cand):
                return cand

        # ② 是目录 → 找里面的 zinia? / 紫鸟*.exe
        hit = _win_exe_in_dir(raw)
        if hit:
            return hit

        # ③ D:\紫鸟 这种「上一层目录」，往里再走一层找 zinia?
        if _FS_ISDIR(raw):
            for name in _listdir(raw):
                if _name_matches(name):
                    hit = _win_exe_in_dir(ntpath.join(raw, name))
                    if hit:
                        return hit

    # 刻意**不查 PATH**：紫鸟客户端从不往 PATH 里注册，而 which 会把任何同名的
    # 无关程序当成客户端（真机上就撞到过 ~/.local/bin/ziniao 这种同名 shim）。
    # 下面这几级探测已经覆盖了实际安装方式。
    return find_windows_client(force=force)


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
            # Windows 没有固定安装位置（安装时用户自选），留空 = 走自动探测。
            return ""
        return "/opt/ziniao/ziniaobrowser"

    def client_executable(self) -> str:
        """真正要执行的文件路径。

        macOS 会把 .app / 应用名解析成内层可执行文件；Windows 支持填目录、exe
        全路径、裸程序名或留空，详见 ``resolve_windows_client_executable``。
        """
        raw = self.default_client_path()
        if IS_MAC:
            return resolve_mac_app_executable(raw)
        if IS_WINDOWS:
            return resolve_windows_client_executable(raw)
        return raw if Path(raw).is_file() else ""

    def client_path_hint(self) -> str:
        """找不到客户端时给用户看的诊断文本（按平台给具体做法）。"""
        lines = [
            f"找不到紫鸟客户端可执行文件。当前「客户端路径」填的是：{self.client_path!r}",
        ]
        if IS_WINDOWS:
            lines += [
                "",
                "【怎么修 · 二选一】",
                "  方法 A（推荐）：把「客户端路径」清空 → 程序自己去常见位置、",
                "                 正在运行的进程、注册表里找。",
                "  方法 B：填对路径。下面三种写法都认：",
                r"      D:\紫鸟\ziniao\ziniao.exe   ← 最保险，一直到 .exe",
                r"      D:\紫鸟\ziniao              ← 目录，自动找里面的 ziniao.exe",
                r"      D:\紫鸟                     ← 上一层目录也行",
                "",
                "【路径从哪抄】右键桌面紫鸟快捷方式 → 属性 → 「目标」那一栏。",
                r"⚠ 注意 Windows 默认隐藏扩展名：属性里看着是 D:\紫鸟\ziniao，",
                r"   那其实是个 exe，抄成 D:\紫鸟\ziniao.exe 最稳。",
                r"⚠ 别填到隔壁的「数据目录」框里 —— 那是另一个东西（文件夹），",
                r"   填错会被自动忽略，然后就是这个报错。",
                "也可以设环境变量 ZINIAO_EXE_PATH 指向那个 exe。",
                r"界面上改：① 紫鸟环境 →「编辑配置…」→「客户端路径」。",
            ]
        elif IS_MAC:
            lines += [
                "可填：ziniao 或 /Applications/ziniao.app 或 "
                "/Applications/ziniao.app/Contents/MacOS/ziniao",
            ]
        else:
            lines += ["Linux 可填 /opt/ziniao/ziniaobrowser"]
        if IS_WINDOWS:
            if not windows_client_tried():
                find_windows_client(force=True)
            tried = windows_client_tried()
            if tried:
                lines.append("")
                lines.append("已自动找过这些位置（都没命中）：")
                lines += [f"  - {t}" for t in tried[:8]]
            else:
                lines.append("")
                lines.append("自动探测这一轮没生成任何候选位置（环境变量都为空、"
                             "也没有运行中的客户端进程）。")
        return "\n".join(lines)

    def resolved_user_data_dir(self) -> Path:
        """web_driver 实例的数据目录（配置留空则自动推导），并确保已创建。

        防御：界面上「客户端路径」和「数据目录」是挨着的两个框，很容易把客户端
        exe 填到这里来。真发生时要**静默回退到默认目录**，而不是拿一个文件路径
        去 mkdir（那会在启动客户端前抛 FileExistsError，把真正的失败原因盖掉）。
        """
        raw = (self.user_data_dir or "").strip()
        if raw.lower().endswith(".exe"):
            raw = ""
        if raw and not Path(raw).expanduser().is_absolute():
            # 相对路径会被解到「exe 所在目录」下（onefile 时那是只读的解包临时目录），
            # 不可靠 → 回退到默认位置。
            raw = ""
        p = Path(raw).expanduser() if raw else default_user_data_dir()
        if p.exists() and not p.is_dir():
            p = default_user_data_dir()
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
            self.log(self.config.client_path_hint())
            return False

        self.log(f"客户端可执行文件：{exe}")
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
