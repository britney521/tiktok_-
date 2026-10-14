#!/usr/bin/env python3
"""真机验证：用「导出订单」按钮导出并下载订单文件（不调 order/list 接口）。

用法：
    python devtools/export_check.py [店铺序号，默认 1]
"""
from __future__ import annotations

import sys
import time
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import requests  # noqa: E402
import ziniao_bridge as zb  # noqa: E402
import tiktok_export as te  # noqa: E402


def log(*a):
    print(*a, flush=True)


def ensure_env(client) -> None:
    s = requests.Session()
    s.trust_env = False
    try:
        s.get("http://127.0.0.1:16851/", timeout=2)
        log("紫鸟 IPC 已在运行")
        return
    except Exception:
        pass
    log("启动紫鸟客户端…")
    client.start_browser()
    try:
        client.update_core(on_progress=lambda m: None)
    except Exception:
        pass


def main(idx: int = 1) -> int:
    cfg = zb.load_config()
    client = zb.make_client(cfg, log=log)
    try:
        client.http_timeout = 300.0
    except Exception:
        pass
    ensure_env(client)

    stores = client.get_browser_list()
    store = stores[idx]
    log(f"店铺：{store.get('browserName')}（index={idx}）")

    ret = None
    for i in range(1, 4):
        try:
            ret = client.open_store(store)
            break
        except Exception as exc:
            log(f"[open_store {i}/3] {exc}")
            time.sleep(5)
    if not ret:
        return 1
    port = ret.get("debuggingPort")
    key = ret.get("browserOauth") or ret.get("browserId")

    sess = None
    try:
        sess = zb.open_session(port, log=log,
                               home_urls=[zb.TIKTOK_SELLER_HOME,
                                          "tiktokshopglobalselling.com"])
        exporter = te.OrderExporter(sess, log=log)
        t0 = time.time()
        result = exporter.export(out_path=ROOT / "_导出原始.xlsx",
                                 scope=te.SCOPE_FILTERED, fmt="excel",
                                 timeout=300, poll=6)
        log(f"\n导出耗时 {time.time() - t0:.0f}s")

        header, rows = result["header"], result["rows"]
        log(f"文件名：{result['file_name']}")
        log(f"列数：{len(header)}；数据行：{len(rows)}")
        log("前 30 列：" + str(header[:30]))

        if rows:
            out = te.write_rows(header, rows, ROOT / "订单.xlsx")
            log(f"已写干净版：{out}")
            id_col = header.index("Order ID") if "Order ID" in header else 0
            ids = [str(r[id_col]) for r in rows]
            log(f"Order ID 示例：{ids[:3]} … {ids[-3:]}（共 {len(ids)}）")

            def col(name: str) -> str:
                if name not in header:
                    return "(无该列)"
                i = header.index(name)
                vals = [str(r[i]) for r in rows[:3]]
                return str(vals)

            for name in ("Buyer Nickname", "Recipient", "Phone #", "Country",
                         "State", "City", "Address Line 1"):
                log(f"  {name}: {col(name)}")
        return 0
    except Exception:
        log(traceback.format_exc())
        return 1
    finally:
        if sess is not None:
            try:
                sess.close()
            except Exception:
                pass
        log("\n=== 关闭店铺 ===")
        try:
            client.close_store(key)
            log("已关闭：", str(key)[:48])
        except Exception as exc:
            log("关闭失败：", exc)


if __name__ == "__main__":
    sys.exit(main(int(sys.argv[1]) if len(sys.argv) > 1 else 1))
