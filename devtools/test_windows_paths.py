"""离线验证 Windows 客户端定位逻辑（在 macOS 上注入假 Windows 目录树）。

为什么要这个：`vendor/ziniao_webdriver_demo/util/ziniao_client.py` 里把文件系统
探针抽成了模块级变量（`_FS_ISFILE` / `_FS_ISDIR` / `_listdir`），就是为了能在
非 Windows 上把 Windows 分支跑一遍 —— 真机验证要等用户那边，太慢。

跑法：
    python devtools/test_windows_paths.py
"""

from __future__ import annotations

import importlib.util
import ntpath
import sys
import tempfile
from pathlib import Path

# 这个脚本全程 print 中文（含 ⚠ ★）。Windows 控制台默认 cp1252，不切编码就
# UnicodeEncodeError —— CI 里 workflow 已设 PYTHONUTF8/PYTHONIOENCODING，
# 这层是防「有人在本机 Windows 上直接跑」时踩坑。
for _name in ("stdout", "stderr"):
    _s = getattr(sys, _name, None)
    if _s is not None:
        try:
            _s.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

ROOT = Path(__file__).resolve().parent.parent
ZC_PATH = ROOT / "vendor" / "ziniao_webdriver_demo" / "util" / "ziniao_client.py"


def load_module():
    name = "zc_win_test"
    spec = importlib.util.spec_from_file_location(name, ZC_PATH)
    mod = importlib.util.module_from_spec(spec)
    # 必须先登记进 sys.modules：dataclass 处理时要用 cls.__module__ 反查模块命名空间
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


class FakeWinFS:
    """用一个明确的文件集合冒充 Windows 文件系统。

    `files` 里放完整路径，目录由文件路径自动推导（含中间层）。
    """

    def __init__(self, files: list[str], dirs: list[str] | None = None):
        self.files = {f.lower() for f in files}
        self.dirs = {d.lower() for d in (dirs or [])}
        for f in files:
            parent = ntpath.dirname(f)
            while parent and parent not in ("", ntpath.splitdrive(parent)[0] + "\\"):
                self.dirs.add(parent.lower())
                parent = ntpath.dirname(parent)
        self.dirs |= {ntpath.splitdrive(d)[0] + "\\" for d in list(self.dirs) if ntpath.splitdrive(d)[0]}
        self.dirs = {d.rstrip("\\").lower() or d.lower() for d in self.dirs}
        self.listings: dict[str, list[str]] = {}
        for d in self.dirs:
            names = []
            for f in self.files:
                if ntpath.dirname(f).lower() == d:
                    names.append(ntpath.basename(f))
            for sub in self.dirs:
                if sub == d:
                    continue
                if ntpath.dirname(sub + "\\").rstrip("\\").lower() == d:
                    names.append(ntpath.basename(sub))
            if names:
                self.listings[d] = sorted(names)

    def isfile(self, p: str) -> bool:
        return str(p).lower() in self.files

    def isdir(self, p: str) -> bool:
        return str(p).rstrip("\\").lower() in self.dirs

    def listdir(self, p: str) -> list[str]:
        return list(self.listings.get(str(p).rstrip("\\").lower(), []))


def bind(mod, fs: FakeWinFS) -> None:
    mod.IS_WINDOWS = True
    mod._FS_ISFILE = fs.isfile
    mod._FS_ISDIR = fs.isdir
    mod._listdir = fs.listdir
    mod._win_client_cache.clear()
    mod._win_client_cache.update(done=False, exe="", tried=[])
    # 让「运行中的进程 / 注册表」两级必定落空，把测试聚焦在目录探测上
    mod._win_running_client_paths = lambda: []
    mod._win_registry_client_paths = lambda: []


# 场景：(名字, 假文件集合, (输入 client_path, 期望结果))
LAYOUTS = [
    (
        "D:\\紫鸟\\ziniao\\ziniao.exe（真机最常见）",
        ["D:\\紫鸟\\ziniao\\ziniao.exe"],
        [
            ("", r"D:\紫鸟\ziniao\ziniao.exe"),                # 留空 → 自动探测
            ("ziniao", r"D:\紫鸟\ziniao\ziniao.exe"),           # 裸名 → 自动探测
            (r"D:\紫鸟\ziniao", r"D:\紫鸟\ziniao\ziniao.exe"),  # 目录
            (r"D:\紫鸟\ziniao\ziniao.exe", r"D:\紫鸟\ziniao\ziniao.exe"),  # exe 全路径
            (r"D:\紫鸟", r"D:\紫鸟\ziniao\ziniao.exe"),         # 上一层目录
        ],
    ),
    (
        "D:\\紫鸟\\ziniao.exe（exe 直接在父目录）",
        ["D:\\紫鸟\\ziniao.exe"],
        [
            ("", r"D:\紫鸟\ziniao.exe"),
            (r"D:\紫鸟", r"D:\紫鸟\ziniao.exe"),
            (r"D:\紫鸟\ziniao", r"D:\紫鸟\ziniao.exe"),
        ],
    ),
    (
        "C:\\Program Files\\ZiniaoBrowser\\ZiniaoBrowser.exe",
        ["C:\\Program Files\\ZiniaoBrowser\\ZiniaoBrowser.exe"],
        [("", "C:\\Program Files\\ZiniaoBrowser\\ZiniaoBrowser.exe")],
    ),
    (
        "中文 exe 名：D:\\紫鸟\\紫鸟浏览器.exe",
        ["D:\\紫鸟\\紫鸟浏览器.exe"],
        [("", "D:\\紫鸟\\紫鸟浏览器.exe")],
    ),
    (
        "旧配置抄来的错路径 → 回落到自动探测",
        ["D:\\紫鸟\\ziniao\\ziniao.exe"],
        [(r"D:\ziniao\ziniao.exe", r"D:\紫鸟\ziniao\ziniao.exe")],
    ),
    (
        "带引号 / 大小写混杂 / 正斜杠的写法",
        ["D:\\紫鸟\\ziniao\\ziniao.exe"],
        [
            ('"D:\\紫鸟\\ziniao\\ziniao.exe"', r"D:\紫鸟\ziniao\ziniao.exe"),
            ("d:/紫鸟/ziniao/ziniao.exe", r"D:\紫鸟\ziniao\ziniao.exe"),
            ("D:\\紫鸟\\ziniao\\", r"D:\紫鸟\ziniao\ziniao.exe"),
        ],
    ),
    (
        "目录里混着卸载器，别误判成客户端",
        ["D:\\ziniao\\ziniao.exe", "D:\\ziniao\\unins000.exe",
         "D:\\ziniao\\Update.exe"],
        [("", "D:\\ziniao\\ziniao.exe")],
    ),
    (
        "★ 隐藏扩展名：用户照抄成 D:\\紫鸟\\ziniao（真身是 .exe）",
        ["D:\\紫鸟\\ziniao.exe"],
        [(r"D:\紫鸟\ziniao", r"D:\紫鸟\ziniao.exe")],
    ),
    (
        "★ 隐藏扩展名 + 目录不在常见位置表里（E 盘）",
        ["E:\\soft\\ziniao.exe"],
        [
            (r"E:\soft\ziniao", r"E:\soft\ziniao.exe"),
            ("", ""),   # 自动探测不认识 E 盘 → 只能靠用户填对
        ],
    ),
]

# 明确应当**找不到**的情况
NEGATIVE = [
    ("空文件系统", [], "", ""),
    ("只有卸载器，没有客户端",
     ["D:\\ziniao\\unins000.exe", "D:\\ziniao\\Update.exe"], "", ""),
]


def main() -> int:
    mod = load_module()
    ok = fail = 0
    for title, files, cases in LAYOUTS:
        print(f"\n== {title} ==")
        for given, want in cases:
            bind(mod, FakeWinFS(files))
            got = mod.resolve_windows_client_executable(given, force=True)
            good = str(got) == want
            ok, fail = (ok + 1, fail) if good else (ok, fail + 1)
            print(f"  [{'OK ' if good else 'FAIL'}] client_path={given!r}"
                  f" → {got!r}" + ("" if good else f"  期望 {want!r}"))

    for title, files, given, want in NEGATIVE:
        print(f"\n== {title} ==")
        bind(mod, FakeWinFS(files))
        got = mod.resolve_windows_client_executable(given, force=True)
        good = str(got) == want
        ok, fail = (ok + 1, fail) if good else (ok, fail + 1)
        print(f"  [{'OK ' if good else 'FAIL'}] → {got!r}"
              + ("" if good else f"  期望 {want!r}"))

    # 顺带验一下「数据目录误填 exe」的防错。
    # ★ 必须把默认目录换到临时目录：resolved_user_data_dir() 会 mkdir，
    #   不换就会在跑测试的机器上凭空建出 ~/AppData/Local/... 那种垃圾目录。
    print("\n== 数据目录误填客户端 exe 时的防错 ==")
    bind(mod, FakeWinFS(["D:\\紫鸟\\ziniao\\ziniao.exe"]))
    with tempfile.TemporaryDirectory() as tmp:
        fake_default = Path(tmp) / "ziniaobrowser" / "instances" / "ziniao_webdriver"
        mod.default_user_data_dir = lambda: fake_default
        cfg = mod.ZiniaoConfig(user_data_dir=r"D:\紫鸟\ziniao\ziniao.exe")
        try:
            p = mod.ZiniaoConfig.resolved_user_data_dir(cfg)
            good = "ziniao_webdriver" in str(p) and Path(tmp) in p.parents
            ok, fail = (ok + 1, fail) if good else (ok, fail + 1)
            print(f"  [{'OK ' if good else 'FAIL'}] 误填 exe → 回退默认目录 {p}")
        except Exception as exc:  # noqa: BLE001
            fail += 1
            print(f"  [FAIL] 抛异常了：{type(exc).__name__}: {exc}")

        # 正常填一个目录时应当原样采用
        real_dir = Path(tmp) / "my_data"
        cfg2 = mod.ZiniaoConfig(user_data_dir=str(real_dir))
        try:
            p2 = mod.ZiniaoConfig.resolved_user_data_dir(cfg2)
            good = p2 == real_dir and p2.is_dir()
            ok, fail = (ok + 1, fail) if good else (ok, fail + 1)
            print(f"  [{'OK ' if good else 'FAIL'}] 正常填目录 → {p2}")
        except Exception as exc:  # noqa: BLE001
            fail += 1
            print(f"  [FAIL] 抛异常了：{type(exc).__name__}: {exc}")

    print(f"\n结果：{ok} 通过 / {fail} 失败")
    return 1 if fail else 0


if __name__ == "__main__":
    sys.exit(main())
