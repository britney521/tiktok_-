"""获取客户工具 · 离屏冒烟测试（零网络）。

覆盖:
  [1] 界面构建 / 列数
  [2] cookie 文本解析（两种格式）
  [3] 订单 Excel 读 Order ID（去重 / 过滤杂质）
  [4] fetch_customers 核心流程（桩化 HTTP）
  [5] GUI 端到端：合成订单表 → 抓取 → 客户 Excel 回读
  [6] 失败路径：接口回非 0 code → 占位行 + Fetch Status
  [7] 无权限路径：code=98001008 → probe 提示 + 只发 1 次请求即中止（其余 SKIPPED）

运行:
    QT_QPA_PLATFORM=offscreen python devtools/smoke_gui.py
"""

from __future__ import annotations

import json
import os
import sys
import tempfile
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
PKG = HERE.parent
sys.path.insert(0, str(PKG))

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import openpyxl  # noqa: E402
from PyQt6.QtWidgets import QApplication  # noqa: E402

import customer_api as api  # noqa: E402
import get_customers_gui as gui  # noqa: E402

ORDER_IDS = ["576966020592278290", "576966020592278291", "576966020592278292"]


def _order_get_payload(oid: str) -> dict:
    return {
        "code": 0,
        "data": {
            "main_order": [{
                "main_order_id": oid,
                "buyer_info_module": {
                    "buyer_nickname": f"buyer_{oid[-3:]}",
                    "shipping_address": {
                        "address_id": "addr-1",
                        "region": {"name": "United Kingdom", "code": "GB"},
                        "districts": [{"name": "England"}, {"name": "London"},
                                      {"name": "Camden"}],
                        "items": [
                            {"key": "name", "value": "Jane Doe"},
                            {"key": "phone", "value": "+44 20 0000 0000"},
                            {"key": "email", "value": "jane@example.com"},
                            {"key": "address", "value": "1 High Street"},
                            {"key": "address_detail", "value": "Flat 2"},
                            {"key": "zipcode", "value": "NW1 1AA"},
                        ],
                    },
                },
            }],
        },
    }


def make_order_excel(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Orders"
    ws.append(["Order ID", "Order Status", "Buyer Nickname"])
    ws.append([ORDER_IDS[0], "COMPLETED", "x"])
    ws.append([ORDER_IDS[1], "COMPLETED", "x"])
    ws.append([ORDER_IDS[2], "COMPLETED", "x"])
    ws.append([ORDER_IDS[2], "DUPLICATE", "x"])   # 重复应被去掉
    ws.append(["说明行不是 ID", "", ""])           # 杂质应被过滤
    ws.append([None, "", ""])
    wb.save(path)
    wb.close()


def pump(app: QApplication, cond, timeout: float = 30.0) -> bool:
    end = time.time() + timeout
    while time.time() < end:
        app.processEvents()
        if cond():
            return True
        time.sleep(0.02)
    return False


def main() -> int:
    app = QApplication.instance() or QApplication(sys.argv)

    print("[1] 界面构建")
    win = gui.MainWindow()
    win.show()
    app.processEvents()
    assert win.tbl.columnCount() == len(api.PREVIEW_COLUMNS), win.tbl.columnCount()
    assert len(api.CUSTOMER_COLUMNS) == 17, len(api.CUSTOMER_COLUMNS)
    assert win.ed_api.text().strip() == api.DEFAULT_API_BASE
    print(f"    列数 OK（预览 {win.tbl.columnCount()} / 导出 {len(api.CUSTOMER_COLUMNS)}）")

    print("[2] cookie 文本解析")
    assert api.parse_cookie_text("a=1; b=2") == {"a": "1", "b": "2"}
    assert api.parse_cookie_text('{"x": "y"}') == {"x": "y"}
    assert api.parse_cookie_text("sessionid=zzz\nmsToken=q")["sessionid"] == "zzz"
    print("    两种格式均 OK")

    with tempfile.TemporaryDirectory() as td:
        tdp = Path(td)
        src = tdp / "orders.xlsx"
        out = tdp / "客户.xlsx"
        make_order_excel(src)

        print("[3] 读 Order ID")
        ids = api.read_order_ids(src)
        assert ids == ORDER_IDS, ids
        print(f"    去重 + 过滤杂质 OK（{len(ids)} 个）")

        print("[4] fetch_customers 核心流程（桩化 HTTP）")
        calls: list[str] = []

        def fake_post(_self, url, params, body):
            oid = body["main_order_id"][0]
            calls.append(oid)
            return _order_get_payload(oid)

        real = api.CustomerClient._post_requests
        api.CustomerClient._post_requests = fake_post
        try:
            client = api.CustomerClient.from_cookies(dict(api.DEFAULT_COOKIES))
            res = api.fetch_customers(client, ids, out_path=out, interval=0.0)
            assert res["ok"] == 3 and res["fail"] == 0, res
            assert len(calls) == 3, calls
            rows = api.iter_sheet_rows(out)
            assert rows[0][0] == "Order ID"
            assert len(rows) == 4, len(rows)          # 表头 + 3 行
            assert rows[1][2] == "Jane Doe"           # Recipient
            assert rows[1][8] == "United Kingdom"     # Country
            assert rows[1][9] == "GB"                 # Country Code
            assert rows[1][10] == "England"           # Province
            print(f"    抓到 {res['ok']} 条，落盘 {out.name}，字段映射 OK")

            print("[5] GUI 端到端")
            win.ed_in.setText(str(src))
            app.processEvents()
            win.ed_out.setText(str(out))
            win.cb_column.setCurrentText("Order ID")
            win._start()
            assert pump(app, lambda: win.worker is None, 30), "GUI 抓取超时"
            app.processEvents()
            assert win.tbl.rowCount() == 3, win.tbl.rowCount()
            assert win.tbl.item(0, 1) is not None and win.tbl.item(0, 1).text().startswith("buyer_")
            assert win.btn_run.isEnabled()
            print(f"    表格 {win.tbl.rowCount()} 行；按钮状态复位 OK")

            print("[6] 失败路径（接口回 code=1000003）")
            def bad_post(_self, url, params, body):
                return {"code": 1000003, "message": "session expired"}

            api.CustomerClient._post_requests = bad_post
            out2 = tdp / "客户_失败.xlsx"
            client2 = api.CustomerClient.from_cookies(dict(api.DEFAULT_COOKIES))
            res2 = api.fetch_customers(client2, ids[:1], out_path=out2, interval=0.0)
            assert res2["fail"] == 1 and res2["ok"] == 0, res2
            rows2 = api.iter_sheet_rows(out2)
            assert rows2[1][0] == ids[0] and rows2[1][16] == "API_ERROR", rows2[1]
            print("    占位行 + Fetch Status=API_ERROR OK")

            print("[7] 无权限路径（cookie 非本店铺 → code=98001008）")
            def perm_post(_self, url, params, body):
                return {"code": 98001008, "data": None, "message": "需要权限"}

            api.CustomerClient._post_requests = perm_post
            client3 = api.CustomerClient.from_cookies(dict(api.DEFAULT_COOKIES))
            # 7a 预检：应明确报出「不是这家店」
            probe = client3.probe(ids[0])
            assert probe["ok"] is False and probe["code"] == 98001008, probe
            assert "cookie" in probe["hint"], probe["hint"]
            # 7b 整批：第 1 条失败后立刻中止，后面不再发请求
            sent: list[str] = []
            api.CustomerClient._post_requests = lambda s, u, p, b: (
                sent.append(b["main_order_id"][0]) or perm_post(s, u, p, b))
            out3 = tdp / "客户_无权限.xlsx"
            res3 = api.fetch_customers(client3, ids, out_path=out3, interval=0.0)
            assert len(sent) == 1, f"应只发 1 次请求，实际 {sent}"
            assert res3["fatal"] and res3["ok"] == 0 and res3["fail"] == 1, res3
            assert [r["Fetch Status"] for r in res3["rows"]] == \
                ["API_ERROR", "SKIPPED", "SKIPPED"], res3["rows"]
            print(f"    只发 {len(sent)} 次请求即中止；其余标 SKIPPED，fatal 已上报 OK")
        finally:
            api.CustomerClient._post_requests = real

    win.close()
    print("\n全部通过 ✅")
    return 0


if __name__ == "__main__":
    sys.exit(main())
