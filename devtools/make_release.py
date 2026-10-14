#!/usr/bin/env python3
"""把 PyInstaller 产物 + 内联的紫鸟客户端，组装成可直接分发的发布目录。

发布布局（exe 旁挂紫鸟项目）：

    release/
    ├── tiktok商家.exe
    ├── 使用说明.txt
    └── 紫鸟浏览器/
        └── ziniao_webdriver_demo/
            ├── util/{__init__.py, ziniao_client.py, dp.py}
            └── config/{__init__.py, config.yaml}

`ziniao_bridge.locate_ziniao_root()` 会在冻结态下**优先**找 exe 旁边的
`紫鸟浏览器/ziniao_webdriver_demo`，所以这个布局开箱即用。

用法：
    python devtools/make_release.py              # 组装到 release/
    python devtools/make_release.py --zip        # 顺便打成 tiktok商家.zip
    python devtools/make_release.py --exe dist/tiktok商家.exe --out /tmp/rel

为什么用 Python 而不是 PowerShell 组装：Windows runner 上中文路径经
PowerShell 容易因编码问题变成乱码，Python 处理 Unicode 路径是可靠的。
"""

from __future__ import annotations

import argparse
import shutil
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
VENDOR = ROOT / "vendor" / "ziniao_webdriver_demo"

# 发布目录里紫鸟项目放在哪（相对发布根）
SIDE_REL = Path("紫鸟浏览器") / "ziniao_webdriver_demo"

APP_NAME = "tiktok商家"

README_TXT = """tiktok商家 —— 订单导出 / 抓客户信息
================================================

【第一次使用】
1. 先把「紫鸟浏览器」客户端装好，并确认能正常打开店铺。
2. 双击 tiktok商家.exe。
3. 界面上「紫鸟项目路径」应该已经自动指向本目录下的
   紫鸟浏览器\\ziniao_webdriver_demo\。
   如果显示「未找到紫鸟项目」，就手动把它改成该目录的绝对路径。
4. 点「编辑配置」，填上你的紫鸟企业账号（company / username / password），
   保存后即写入 紫鸟浏览器\\ziniao_webdriver_demo\\config\\config.yaml。
   ⚠ 该文件含明文密码，别外传、别提交到仓库。

【流水线四步】
① 启动环境  ② 选店铺  ③ 导出订单  ④ 抓客户信息

【目录说明】
tiktok商家.exe
紫鸟浏览器\\ziniao_webdriver_demo\\   ← 紫鸟协议客户端（必须和 exe 放一起）
     util\\ziniao_client.py           ← 真正干活的那份
     util\\dp.py
     config\\config.yaml              ← 你的账号配置（首次运行时自己填）

【常见问题】
· 提示「找不到紫鸟项目」：确认 紫鸟浏览器\\ziniao_webdriver_demo\\util\\ziniao_client.py
  存在，且和 exe 在同一级目录下。
· 报 code=98001008：当前登录态看不到这批订单。先确认店铺选对了，
  再确认订单确实是这家店的。
· 抓客户很慢是正常的：每条之间有最小 1.5 秒间隔，是为了避开平台风控。
"""


def find_exe(explicit: str | None) -> tuple[Path, bool]:
    """返回 (exe 路径, 是否 onedir)。

    onefile  -> dist/tiktok商家.exe   / dist/tiktok商家
    onedir   -> dist/tiktok商家/tiktok商家.exe（连同整个目录一起分发）
    """
    if explicit:
        p = Path(explicit)
        if not p.exists():
            sys.exit(f"指定的 exe 不存在：{p}")
        return p, p.is_dir()

    candidates = [
        ROOT / "dist" / f"{APP_NAME}.exe",
        ROOT / "dist" / APP_NAME,
        ROOT / "dist" / APP_NAME / f"{APP_NAME}.exe",
        ROOT / "dist" / APP_NAME / APP_NAME,
    ]
    for c in candidates:
        if c.exists():
            return c, c.is_dir()
    sys.exit(
        "在 dist/ 下找不到打包产物。先跑：\n"
        "    pyinstaller build.spec --noconfirm"
    )


def copy_ziniao(dst_root: Path) -> Path:
    """把内联的紫鸟客户端铺成发布布局，并生成 config.yaml。"""
    if not (VENDOR / "util" / "ziniao_client.py").is_file():
        sys.exit(f"内联的紫鸟客户端不完整：{VENDOR}")

    dst = dst_root / SIDE_REL
    (dst / "util").mkdir(parents=True, exist_ok=True)
    (dst / "config").mkdir(parents=True, exist_ok=True)

    for name in ("__init__.py", "ziniao_client.py", "dp.py"):
        shutil.copy2(VENDOR / "util" / name, dst / "util" / name)

    shutil.copy2(VENDOR / "config" / "__init__.py", dst / "config" / "__init__.py")

    # 配置模板 → 真实文件名；已存在就不覆盖（别把用户填好的账号冲掉）
    target_cfg = dst / "config" / "config.yaml"
    if target_cfg.exists():
        print(f"  保留已有配置：{target_cfg.relative_to(dst_root)}")
    else:
        shutil.copy2(
            VENDOR / "config" / "config.example.yaml",
            target_cfg,
        )
        print(f"  生成配置模板：{target_cfg.relative_to(dst_root)}（账号待填）")

    return dst


def main() -> int:
    ap = argparse.ArgumentParser(description="组装发布目录")
    ap.add_argument("--exe", help="exe 路径（默认自动在 dist/ 里找）")
    ap.add_argument("--out", default="release", help="发布目录（默认 release/）")
    ap.add_argument("--zip", action="store_true", help="顺便打成 zip")
    args = ap.parse_args()

    exe, onedir = find_exe(args.exe)
    out = Path(args.out)
    if not out.is_absolute():
        out = ROOT / out

    if out.exists():
        shutil.rmtree(out)
    out.mkdir(parents=True)

    print(f"打包产物 : {exe}")
    if onedir:
        # onedir：整目录搬过去
        shutil.copytree(exe.parent, out / exe.parent.name, dirs_exist_ok=True)
        print(f"  已复制整个目录 {exe.parent.name}/")
    else:
        shutil.copy2(exe, out / exe.name)
        print(f"  已复制 {exe.name}")

    ziniao_dst = copy_ziniao(out)

    (out / "使用说明.txt").write_text(README_TXT, encoding="utf-8")

    print()
    print("发布目录内容：")
    for p in sorted(out.rglob("*")):
        if p.is_file():
            rel = p.relative_to(out)
            print(f"  {rel}  ({p.stat().st_size:,} B)")

    if args.zip:
        # zip 直接以程序名命名，放在发布目录旁边（CI 按这个名字找产物）
        zip_path = out.parent / f"{APP_NAME}.zip"
        if zip_path.exists():
            zip_path.unlink()
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for p in sorted(out.rglob("*")):
                if p.is_file():
                    zf.write(p, p.relative_to(out))
        print()
        print(f"已打 zip：{zip_path}  ({zip_path.stat().st_size:,} B)")

    # 自检：发布布局必须是「exe 和紫鸟项目同级」
    assert (out / SIDE_REL / "util" / "ziniao_client.py").is_file(), "旁挂布局不对"
    print()
    print("✔ 旁挂布局自检通过（exe + 紫鸟浏览器/ziniao_webdriver_demo 同级）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
