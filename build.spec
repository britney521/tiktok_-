# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置 —— tiktok商家 / 抓客户信息

本地构建（必须在目标平台上构建，PyInstaller 不能交叉编译）：

    pip install -r requirements.txt pyinstaller
    pyinstaller build.spec --noconfirm

产出：

    dist/tiktok商家.exe          （ONEFILE = True）

发布方式：把 exe 和 `紫鸟浏览器/ziniao_webdriver_demo/` 放在**同一个目录**里。
    <发布目录>/
    ├── tiktok商家.exe
    └── 紫鸟浏览器/ziniao_webdriver_demo/{util,config}/

  程序会自己按 exe 位置去找那份协议客户端（见 ziniao_bridge.locate_ziniao_root）。

为什么紫鸟项目要「旁挂」而不是打进 exe：
    `ziniao_bridge.load_ziniao()` 用 `spec_from_file_location()` **按文件路径**
    加载 `util/ziniao_client.py`，所以它必须是磁盘上真实存在的 .py 源码。
    打进 exe 会变成解包临时目录里的东西，路径一改就失效；旁挂则最稳、也方便
    单独替换那份客户端。发布包里由 CI 从 `vendor/ziniao_webdriver_demo/` 复制出来。
"""

import ast
import sys
from pathlib import Path


def _say(msg: str) -> None:
    """打一行进度日志；打不出来就咽掉，绝不让它影响构建。

    ★ 为什么要包一层 ★
    spec 是被 PyInstaller 用 ``exec()`` 执行的 —— 里面任何未捕获异常都会把
    整个构建带崩。而 Windows 的控制台默认是 cp1252 编码，中文一 print 就是::

        UnicodeEncodeError: 'charmap' codec can't encode characters ...

    实测（GitHub 的 windows runner，2026-10-14）：构建步骤 **1 秒内** 失败，
    整份日志就只有这一行 traceback，PyInstaller 连依赖分析都没开始 ——
    为了打一行进度日志而炸掉整条流水线，实在不划算。

    注：workflow 里已经把 PYTHONIOENCODING 设成 utf-8，正常情况下走不到兜底分支；
    这层是防「有人在本机 Windows 上直接跑 pyinstaller」时再踩同一个坑。
    """
    for stream in (sys.stdout, sys.stderr):
        if stream is None:
            continue
        try:
            stream.write(msg + "\n")
            return
        except Exception:
            continue


# ---------------------------------------------------------------------------
# 可调项
# ---------------------------------------------------------------------------
# True  = 单文件 exe（好分发，但每次启动要解包，冷启动慢几秒）
# False = onedir（一个文件夹，启动快，但文件多）
ONEFILE = True

NAME = "tiktok商家"
ENTRY = "customer_gui.py"

ROOT = Path(SPECPATH)  # noqa: F821  # PyInstaller 在 exec spec 时注入

VENDOR_UTIL = ROOT / "vendor" / "ziniao_webdriver_demo" / "util"

# 扫描到这些名字时跳过（要么是本项目自己在管，要么是明确排除的）
_SKIP_NAMES = {
    "__future__",
    "util",          # ziniao_bridge 自己注册的包，不能由 PyInstaller 接管
    "DrissionPage",  # 已弃用，见 excludes
    "selenium",      # 已弃用，见 excludes
}


def vendored_imports() -> list[str]:
    """解析紫鸟客户端源码，收集它真正 import 的**所有**模块名。

    ★ 为什么非这么做不可 ★
    那份客户端是运行时用 `spec_from_file_location()` 动态加载的，PyInstaller
    的静态分析**完全看不见它的 import 图**。后果不只是第三方库会漏 ——
    **连标准库都会漏**。实测就栽在这上面：

        ModuleNotFoundError: No module named 'uuid'

    `ziniao_client.py` 顶层写着 `import uuid`，但本项目没有任何地方静态引用
    uuid，PyInstaller 就认定用不到、没打进包；等运行到那一步才炸。
    这种漏法连 warn-build.txt 都不会提醒（静态图上它压根不存在）。

    所以这里直接读源码喂 hiddenimports：比手写清单可靠，
    以后替换那份客户端也不用改 spec。
    """
    if not VENDOR_UTIL.is_dir():
        raise SystemExit(
            f"找不到内联的紫鸟客户端：{VENDOR_UTIL}\n"
            "（发布包要靠它，仓库里缺了就没法打包）"
        )

    names: set[str] = set()
    for py in sorted(VENDOR_UTIL.glob("*.py")):
        tree = ast.parse(py.read_text(encoding="utf-8"), filename=str(py))
        for node in ast.walk(tree):  # 连函数体内的延迟导入一起收
            if isinstance(node, ast.Import):
                names.update(a.name.split(".")[0] for a in node.names)
            elif isinstance(node, ast.ImportFrom):
                if node.level or not node.module:
                    continue  # 相对导入
                names.add(node.module.split(".")[0])

    found = sorted(n for n in names if n not in _SKIP_NAMES)
    _say(f"[build.spec] 从紫鸟客户端扫描到 {len(found)} 个依赖：{found}")
    return found


# ---------------------------------------------------------------------------
# 依赖
# ---------------------------------------------------------------------------
hiddenimports = [
    # ★★★ 紫鸟客户端的依赖，自动扫描得出（含标准库，如 uuid）★★★
    *vendored_imports(),
    # websocket-client 子模块较多，显式带上更保险
    "websocket",
    "websocket._abnf",
    "websocket._core",
    "websocket._http",
    "websocket._url",
    "websocket._utils",
    # 本项目里是延迟 import（写 Excel 时才用）
    "openpyxl",
    "openpyxl.styles",
    "openpyxl.utils",
]

excludes = [
    # 用不到，而且目标机器上根本没装：不排除的话 PyInstaller 只会刷警告
    "DrissionPage",
    "selenium",
    # 体积大户，本项目零依赖
    "numpy",
    "pandas",
    "matplotlib",
    "scipy",
    "PIL",
    "IPython",
    "jupyter",
    "notebook",
    # 标准库里的边角料
    "tkinter",
    "test",
    "unittest",
    "pydoc_data",
    "lib2to3",
]

a = Analysis(
    [str(ROOT / ENTRY)],
    pathex=[str(ROOT)],
    binaries=[],
    datas=[],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

_common_exe_kwargs = dict(
    name=NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,          # 紫鸟/杀软对 UPX 壳比较敏感，关掉
    console=False,      # GUI 程序：不要黑框
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon=str(ROOT / "docs" / "app.ico"),  # 有图标就打开这行
)

if ONEFILE:
    exe = EXE(pyz, a.scripts, a.binaries, a.datas, [], **_common_exe_kwargs)
else:
    exe = EXE(
        pyz, a.scripts, [],
        exclude_binaries=True,
        **_common_exe_kwargs,
    )
    coll = COLLECT(
        exe, a.binaries, a.datas,
        strip=False,
        upx=False,
        name=NAME,
    )
