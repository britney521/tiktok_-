"""⑥ 前置检查的一次性验证（用**真实文件**，但不联网、不起真流水线）。

背景：用户只跑过⑤，磁盘上有 `秋图2_商品评分.xlsx`（自带 order_id）、**没有**订单表。
旧代码在「单独跑⑥」时无条件要求订单表，于是弹窗拦人，还指着 `<店铺>_订单.xlsx`，
看着像⑥把订单表当成了评分表。

本脚本把守卫单独拎出来跑：`QMessageBox.warning` 收进列表（不弹模态），
`PipelineThread` 换成不干活的桩（只记录被创建 → 说明守卫放行了）。

用法：``QT_QPA_PLATFORM=offscreen python devtools/check_bad_guard.py``
"""
from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from PyQt6.QtWidgets import QApplication  # noqa: E402

import customer_gui as cg  # noqa: E402
import tiktok_export as te  # noqa: E402

GONE = ROOT / "这个订单表不存在.xlsx"


class Sink:
    def __init__(self) -> None:
        self.msgs: list[str] = []

    def connect(self, *_a, **_k) -> None:
        pass

    def emit(self, *_a, **_k) -> None:
        pass


class FakeWorker:
    """替换 PipelineThread：只证明「守卫放行了」，不真跑。"""

    created: list[dict] = []

    def __init__(self, cfg, root, store, opts, parent=None) -> None:  # noqa: ARG002
        FakeWorker.created.append(dict(opts))
        self.signals = {}

    def __getattr__(self, name):
        return Sink()          # log/stage/done/failed… 一律吃掉

    def start(self) -> None:
        pass


def main() -> int:
    app = QApplication(sys.argv)  # noqa: F841

    warns: list[str] = []
    cg.QMessageBox.warning = lambda *a, **k: warns.append(str(a[2]) if len(a) > 2 else str(a))
    orig_worker = cg.PipelineThread
    cg.PipelineThread = FakeWorker

    win = cg.MainWindow()
    store = {"browserName": "秋图2", "browserOauth": "x", "siteId": 2}
    win._on_stores(True, [store], "共 1 个店铺")
    win.tbl_stores.selectRow(0)
    app.processEvents()

    rating = ROOT / "秋图2_商品评分.xlsx"
    fails = 0

    def check(name: str, got: bool, detail: str = "") -> None:
        nonlocal fails
        print(f"  [{'OK  ' if got else 'FAIL'}] {name}" + (f"  ← {detail}" if detail else ""))
        fails += 0 if got else 1

    print("① 只有评分表、没有订单表 → 不该拦（用户的实际情形）")
    win.ed_rating_out.setText(str(rating))
    win.ed_bad_out.setText(str(ROOT / "秋图2_客户差评信息.xlsx"))
    win.ed_order_out.setText(str(GONE))       # 故意不存在的订单表
    win.ed_cust_in.setText(str(GONE))
    warns.clear()
    FakeWorker.created.clear()
    win._start_pipeline(do_orders=False, do_customers=False, do_bad=True)
    app.processEvents()
    check("没有弹窗", not warns, str(warns))
    check("流水线起来了（守卫放行）", len(FakeWorker.created) == 1,
          str(len(FakeWorker.created)))
    check("读的就是评分表（不是订单表）",
          bool(FakeWorker.created) and FakeWorker.created[0]["rating_out"] == str(rating),
          str(FakeWorker.created[0].get("rating_out") if FakeWorker.created else None))
    check("评分表里认得出 order_id（⑥ 靠它直连）",
          cg.rating_order_column(rating) == "order_id", cg.rating_order_column(rating))

    print("② 老版「仅昵称」评分表 + 没有订单表 → 才该拦，且文案说清原因")
    # 临时文件一律放临时目录：这个脚本**不在用户仓库里留任何东西**
    legacy = Path(tempfile.mkdtemp(prefix="guard_legacy_")) / "老版评分.xlsx"
    te.write_rows(["Product Name", "Rating", "Buyer Nickname"],
                  [["圣诞树", "2", "bob"]], legacy, sheet="商品评分")
    try:
        win.ed_rating_out.setText(str(legacy))
        warns.clear()
        FakeWorker.created.clear()
        win._start_pipeline(do_orders=False, do_customers=False, do_bad=True)
        app.processEvents()
        check("弹窗拦下了", len(warns) == 1, str(warns))
        check("文案说的是「订单号列」而不是赖评分表",
              bool(warns) and "订单号列" in warns[0], str(warns))
        check("没起流水线", not FakeWorker.created, str(len(FakeWorker.created)))

        print("③ 连锁都没有 → 该拦，文案指向评分表")
        # 注意：必须放在**空目录**里 —— 项目根目录下真有一份 _评分原始.csv，
        # 守卫认出它会（正确地）放行，那样就测不出「什么都没有」这一支了。
        empty = Path(tempfile.mkdtemp(prefix="guard_empty_"))
        win.ed_rating_out.setText(str(empty / "根本没有.xlsx"))
        warns.clear()
        FakeWorker.created.clear()
        win._start_pipeline(do_orders=False, do_customers=False, do_bad=True)
        app.processEvents()
        check("弹窗拦下了", len(warns) == 1, str(warns))
        check("文案指向「评分表」", bool(warns) and "评分表" in warns[0], str(warns))
        check("没起流水线", not FakeWorker.created, str(len(FakeWorker.created)))
    finally:
        legacy.unlink(missing_ok=True)
        cg.PipelineThread = orig_worker

    print(f"\n{'全部通过' if not fails else f'{fails} 项失败'}")
    return 1 if fails else 0


if __name__ == "__main__":
    raise SystemExit(main())
