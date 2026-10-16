"""离线验证「强制关掉已开启的紫鸟客户端」（`zb.kill_all_clients`）。

背景（2026-10-16 用户定案）：本来开着紫鸟、再让程序启动一个，两个实例会串到
同一个进程上 —— 新实例的 `--run_type=web_driver` / `--port` 不生效，startBrowser
返回的 `debuggingPort` 是旧实例的，CDP 连上去就出各种怪问题。所以「启动环境」
一律先清干净再起。

这里**不碰本机进程**：把 `zb._run_quiet`（真正发 taskkill / tasklist / pgrep 的
那个函数）和 `zb._IS_WINDOWS` 换掉，就能在 macOS 上把 Windows 分支跑一遍，
断言命令行的**内容与顺序**。

跑法：
    python devtools/test_client_kill.py
"""

from __future__ import annotations

import contextlib
import sys
import types
from pathlib import Path

# 全程 print 中文（含 ⚠ ★）。Windows 控制台默认 cp1252，不切编码就 UnicodeEncodeError。
for _name in ("stdout", "stderr"):
    _s = getattr(sys, _name, None)
    if _s is not None:
        try:
            _s.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import ziniao_bridge as zb  # noqa: E402

OK = 0
FAIL = 0


def check(name: str, good: bool, detail: str = "") -> None:
    global OK, FAIL
    if good:
        OK += 1
        print(f"  [OK ] {name}")
    else:
        FAIL += 1
        print(f"  [FAIL] {name}  ← {detail}")


class FakeRunner:
    """假的外部命令执行器。

    · ``pids``：{进程名: [pid, ...]}，用来模拟「现在有哪些紫鸟进程在跑」
    · ``stick``：True = taskkill **杀不掉**（模拟杀不干净，看会不会再清一轮）
    · 命令一律记进 ``cmds``，断言就靠它
    """

    def __init__(self, pids: dict[str, list[int]] | None = None,
                 stick: bool = False, explode: bool = False,
                 timeline: list[str] | None = None) -> None:
        self.pids = {k: list(v) for k, v in (pids or {}).items()}
        self.stick = stick
        self.explode = explode
        self.timeline = timeline if timeline is not None else []
        self.cmds: list[list[str]] = []
        self.kills: list[str] = []          # 收到的 taskkill 目标名
        self.queries: list[str] = []        # 收到的 tasklist 查询名
        self.pgrep: list[str] = []          # 收到的 pgrep 模式

    def __call__(self, cmd, timeout: float = 20.0):
        self.cmds.append(list(cmd))
        self.timeline.append(str(cmd[0]))
        if self.explode:
            raise OSError("命令跑不起来")
        head = cmd[0]
        if head == "tasklist":
            name = str(cmd[2]).replace("imagename eq ", "").strip()
            self.queries.append(name)
            lines = [f'"{name}","{p}","Console","1","132,608 K"'
                     for p in self.pids.get(name, [])]
            if not lines:
                # 没有匹配时 tasklist 退化成分号开头的提示行（退出码 1）
                lines = ["信息: 没有运行的任务匹配指定标准。"]
            return types.SimpleNamespace(stdout="\n".join(lines) + "\n", returncode=0)
        if head == "taskkill":
            name = str(cmd[-1])
            self.kills.append(name)
            if not self.stick:
                self.pids[name] = []
            return types.SimpleNamespace(stdout="", returncode=0)
        if head == "pgrep":
            name = str(cmd[-1])
            self.pgrep.append(name)
            return types.SimpleNamespace(
                stdout="\n".join(str(p) for p in self.pids.get(name, [])) + "\n",
                returncode=0 if self.pids.get(name) else 1)
        return types.SimpleNamespace(stdout="", returncode=0)


class FakeClient:
    """假的紫鸟客户端：只关心 kill_process 被调用了几次 / 抛不抛。"""

    def __init__(self, boom: bool = False, timeline: list[str] | None = None) -> None:
        self.calls = 0
        self.boom = boom
        self.timeline = timeline if timeline is not None else []

    def kill_process(self, version=None, wait_release: float = 20.0) -> bool:  # noqa: ARG002
        self.calls += 1
        self.timeline.append("kill_process")
        if self.boom:
            raise RuntimeError("taskkill 炸了")
        return True


@contextlib.contextmanager
def fake_env(windows: bool, runner):
    """把 `_IS_WINDOWS` / `_run_quiet` 换成假的，退出时还原。"""
    old_win, old_run = zb._IS_WINDOWS, zb._run_quiet
    zb._IS_WINDOWS, zb._run_quiet = windows, runner
    try:
        yield
    finally:
        zb._IS_WINDOWS, zb._run_quiet = old_win, old_run


def main() -> int:
    print("== 1. running_clients()：数得准、不炸 ==")
    with fake_env(True, FakeRunner({"ziniao.exe": [111],
                                    "ZiniaoBrowser.exe": [222, 222]})):
        pids = zb.running_clients()
        check("多个名字的进程都数到", pids == [111, 222], str(pids))
        check("PID 去重 + 排序", len(pids) == len(set(pids)) and pids == sorted(pids),
              str(pids))
    with fake_env(True, FakeRunner()):
        check("一个都没有 → 空列表（不是报错）", zb.running_clients() == [],
              str(zb.running_clients()))
    with fake_env(True, FakeRunner(explode=True)):
        check("★ 命令跑不起来 → 返回空列表，绝不抛异常",
              zb.running_clients() == [], str(zb.running_clients()))
    with fake_env(False, FakeRunner({"ziniao": [333], "ziniaobrowser": [444]})):
        check("非 Windows 走 pgrep，也能数到", zb.running_clients() == [333, 444],
              str(zb.running_clients()))

    print("== 2. kill_all_clients()：Windows 下把三种名字全杀一遍 ==")
    tl: list[str] = []
    r = FakeRunner({"ziniao.exe": [111], "ZiniaoBrowser.exe": [222]}, timeline=tl)
    c = FakeClient(timeline=tl)
    logs: list[str] = []
    with fake_env(True, r):
        n = zb.kill_all_clients(c, log=logs.append)
    check("返回「动手前数到几个」", n == 2, str(n))
    check("★ 三种可能的进程名都发了 taskkill（只杀一种会漏）",
          sorted(r.kills) == sorted(zb._WIN_CLIENT_EXE_NAMES), str(r.kills))
    check("taskkill 都带 /f /t（强制 + 连子进程树）",
          all(cmd[:3] == ["taskkill", "/f", "/t"] for cmd in r.cmds
              if cmd[0] == "taskkill"), str(r.cmds[:1]))
    check("还调了 client.kill_process()（兜底 + 等端口释放）", c.calls == 1,
          str(c.calls))
    check("★ 顺序：先数 → 再 taskkill → 最后 kill_process",
          tl.index("tasklist") < tl.index("taskkill")
          and tl.index("taskkill") < tl.index("kill_process"), str(tl))
    check("日志里有「检测到 2 个紫鸟进程」",
          any("检测到 2 个紫鸟进程" in m for m in logs), str(logs[:2]))
    check("杀干净就不喊「再清一轮」",
          not any("再清一轮" in m for m in logs), str(logs))

    print("== 3. 杀不干净 → 再清一轮，两轮后如实报出来 ==")
    r = FakeRunner({"ziniao.exe": [111]}, stick=True)
    c = FakeClient()
    logs = []
    with fake_env(True, r):
        zb.kill_all_clients(c, log=logs.append)
    check("杀不掉时再清一轮（taskkill 共 2 轮 × 3 个名字）", len(r.kills) == 6,
          str(len(r.kills)))
    check("kill_process 也跟着来第二轮", c.calls == 2, str(c.calls))
    check("日志写「再清一轮」", any("再清一轮" in m for m in logs), str(logs))
    check("★ 两轮还有残留 → 提示手动结束（不是静默放过）",
          any("仍检测到紫鸟进程" in m and "任务管理器" in m for m in logs),
          str(logs[-2:]))

    print("== 4. 异常都不外抛（杀不干净也得让「启动环境」继续走）==")
    r = FakeRunner({"ziniao.exe": [111]})
    c = FakeClient(boom=True)
    logs = []
    with fake_env(True, r):
        try:
            n = zb.kill_all_clients(c, log=logs.append)
            err = None
        except Exception as exc:  # noqa: BLE001
            n, err = -1, exc
    check("kill_process 抛异常 → 不往外抛", err is None, repr(err))
    check("异常记进日志（有据可查）",
          any("终止客户端进程异常" in m for m in logs), str(logs))
    check("返回值照旧是「之前有几个」", n == 1, str(n))

    r = FakeRunner({"ziniao.exe": [111]})

    def boom_log(_msg: str) -> None:
        raise RuntimeError("日志回调也炸")

    with fake_env(True, r):
        try:
            zb.kill_all_clients(FakeClient(), log=boom_log)
            err2 = None
        except Exception as exc:  # noqa: BLE001
            err2 = exc
    check("日志回调自己抛异常 → 也不影响主流程", err2 is None, repr(err2))

    print("== 5. 非 Windows：不发 taskkill，交给 kill_process ==")
    # 注：假 client 不会真的把进程「杀掉」，所以这里用「本来就没有进程在跑」的
    # 场景 —— 否则复查永远有残留、要跑满两轮（那是测试桩的局限，不是实现的问题）。
    r = FakeRunner()
    c = FakeClient()
    logs = []
    with fake_env(False, r):
        n = zb.kill_all_clients(c, log=logs.append)
    check("不发任何 taskkill", r.kills == [], str(r.kills))
    check("走 vendor 的 kill_process（macOS 是 pkill 三连）", c.calls == 1,
          str(c.calls))
    check("数进程走 pgrep，两个名字都查",
          set(r.pgrep) == set(zb._POSIX_CLIENT_NAMES), str(r.pgrep))
    check("返回值仍是数到的个数", n == 0, str(n))
    check("一个都没数到也照样清一遍残留（不提前退出）",
          any("仍清一次残留" in m for m in logs), str(logs[:1]))

    print("== 6. 默认（真机）语义：先数、杀、再数 ==")
    r = FakeRunner({"ziniao.exe": [111]})
    with fake_env(True, r):
        zb.kill_all_clients(FakeClient())
    heads = [cmd[0] for cmd in r.cmds]
    check("命令序列 = 查 → 杀 → 查（复查残留）",
          heads[0] == "tasklist" and heads[-1] == "tasklist"
          and "taskkill" in heads, str(heads))
    check("查询覆盖全部进程名（不只看当前配置里那一个）",
          set(r.queries) == set(zb._WIN_CLIENT_EXE_NAMES), str(sorted(set(r.queries))))

    print(f"\n结果：{OK} 通过 / {FAIL} 失败")
    return 1 if FAIL else 0


if __name__ == "__main__":
    sys.exit(main())
