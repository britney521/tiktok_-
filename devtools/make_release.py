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


def _force_utf8_stdio() -> None:
    """把 stdout/stderr 切到 UTF-8 输出。

    Windows 的控制台默认编码是 cp1252，而本脚本从头到尾都在 print 中文
    （发布目录清单、使用说明里的中文路径…），一 print 就是::

        UnicodeEncodeError: 'charmap' codec can't encode characters ...

    实测这种报错会让 CI 的「组装发布目录」步骤直接挂掉。
    ``errors="replace"`` 是最后兜底：万一某个环境不允许改编码，
    也只会显示成 ?，不会把构建打断。
    """
    for name in ("stdout", "stderr"):
        stream = getattr(sys, name, None)
        if stream is None:
            continue
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


_force_utf8_stdio()

ROOT = Path(__file__).resolve().parent.parent
VENDOR = ROOT / "vendor" / "ziniao_webdriver_demo"

# 发布目录里紫鸟项目放在哪（相对发布根）
SIDE_REL = Path("紫鸟浏览器") / "ziniao_webdriver_demo"

APP_NAME = "tiktok商家"

# ★ 用原始字符串（r"""..."""）★
# 这段文案里全是 Windows 路径（D:\紫鸟\ziniao、util\ziniao_client.py …）。
# 普通三引号字符串下 `\u` / `\z` 这类会被当成转义序列 —— 其中 `\util` 会直接
# 变成 **SyntaxError: truncated \uXXXX escape**，把整个脚本、进而把 CI 打挂。
# 实测就踩过一次。r 前缀一劳永逸。
README_TXT = r"""tiktok商家 —— 订单导出 / 抓客户信息
================================================

【第一次使用】
1. 先把「紫鸟浏览器」客户端装好，并确认能正常打开店铺。
2. 双击 tiktok商家.exe。
3. 界面上「紫鸟项目路径」应该已经自动指向本目录下的
   紫鸟浏览器\ziniao_webdriver_demo\。
   如果显示「未找到紫鸟项目」，就手动把它改成该目录的绝对路径。
4. 点「编辑配置」，填上你的紫鸟企业账号（company / username / password），
   保存后即写入 紫鸟浏览器\ziniao_webdriver_demo\config\config.yaml。
   ⚠ 该文件含明文密码，别外传、别提交到仓库。
5.「客户端路径」**留空就行**，程序会自己去找紫鸟客户端装在哪。

【流水线四步】
① 启动环境  ② 选店铺  ③ 导出订单  ④ 抓客户信息

【目录说明】
tiktok商家.exe
紫鸟浏览器\ziniao_webdriver_demo\   ← 紫鸟协议客户端（必须和 exe 放一起）
     util\ziniao_client.py           ← 真正干活的那份
     util\dp.py
     config\config.yaml              ← 你的账号配置（首次运行自动生成，填上账号即可）
     config\config.example.yaml      ← 配置模板，别删

【常见问题】
· 提示「找不到紫鸟客户端」：先确认紫鸟浏览器自己装好了、能打开店铺。
  然后把「编辑配置」里的「客户端路径」**留空**（程序会自动去找）。
  也可以手动填，下面三种写法都认：
      D:\紫鸟\ziniao\ziniao.exe    ← 最保险，一直到 .exe
      D:\紫鸟\ziniao               ← 目录，自动找里面的 ziniao.exe
      D:\紫鸟                      ← 上一层目录也行
  路径从哪抄：右键桌面紫鸟快捷方式 → 属性 →「目标」那一栏。
  ⚠ Windows 默认隐藏扩展名：属性里看着是 D:\紫鸟\ziniao，那其实是个 exe。
  ⚠ 别填到「数据目录」那一栏里 —— 那是另一样东西（文件夹），填错会被忽略，
    然后就是这个报错（两个框挨着，容易看错）。
  ⚠ 如果报错里出现 client_path='ziniao'，说明你用的是旧版 exe，换新版即可。
· 提示「找不到紫鸟项目」：确认 紫鸟浏览器\ziniao_webdriver_demo\util\ziniao_client.py
  存在，且和 exe 在同一级目录下。
· 报 code=98001008：当前登录态看不到这批订单。先确认店铺选对了，
  再确认订单确实是这家店的。
· 抓客户很慢是正常的：每条之间有最小 1.5 秒间隔，是为了避开平台风控。

【升级】
直接解压覆盖即可 —— 发布包里**不含** config\config.yaml，你填好的账号不会被冲掉
（新版会自动沿用原有的那份配置）。
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
    # ★ 模板必须带上 ★ zip 里**故意不含**真的 config.yaml（免得解压升级覆盖用户填好的
    # 账号），程序靠这份 example 在首次运行时播种出 config.yaml
    # （ziniao_bridge.load_config）。少了它，新解压出来的目录就永远没配置。
    shutil.copy2(VENDOR / "config" / "config.example.yaml", dst / "config" / "config.example.yaml")

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
        # ★ 不把 config/config.yaml 打进 zip ★
        # 它只是从 config.example.yaml 复制出来的空模板。用户解压新版覆盖时，
        # 若把这个文件一起盖回去，已经填好的企业账号 / 密码就丢了。
        # 程序在 config.yaml 缺失时会自动从 config.example.yaml 播种一份
        # （ziniao_bridge.load_config），所以发布包里不带它也完全能跑。
        user_cfg = str(SIDE_REL / "config" / "config.yaml").replace("\\", "/")
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for p in sorted(out.rglob("*")):
                if not p.is_file():
                    continue
                arc = str(p.relative_to(out)).replace("\\", "/")
                if arc == user_cfg:
                    print(f"  跳过（保留用户自己的配置，不打进 zip）：{arc}")
                    continue
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
