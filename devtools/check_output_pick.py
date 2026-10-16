"""只验证「浏览…」这一块：③订单 / ⑤评分 = 浏览选中一个文件，④客户 / ⑥差评 = 选文件夹。

为什么单独一个脚本：改这块 UI 很频繁，而整套冒烟要跑三四分钟。这里 20 来条断言
秒级出结果，专门盯「弹的是哪个对话框、框里被改成什么、起始路径落在哪」。
"""
import os
import sys
import tempfile
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "devtools"))

from PyQt6.QtWidgets import QApplication, QPushButton  # noqa: E402

import customer_gui as cg  # noqa: E402
import smoke_gui as sm  # noqa: E402
import tiktok_export as te  # noqa: E402

OK = FAIL = 0


def check(name, cond, extra=""):
    global OK, FAIL
    if cond:
        OK += 1
        print(f"  [OK  ] {name}")
    else:
        FAIL += 1
        print(f"  [FAIL] {name}  {extra}")


app = QApplication(sys.argv)
sm.install_stubs()
win = cg.MainWindow()
app.processEvents()

print("== 四个「浏览…」按钮的分工 ==")
btns = [b for b in win.findChildren(QPushButton) if b.text() == "浏览…"]
check("还是 5 个「浏览…」（④输入 1 + 四个输出）", len(btns) == 5, str(len(btns)))
tips = {b.toolTip() for b in btns}
check("③⑤ 的提示是「浏览选择文件」",
      any("浏览选择" in t for t in tips), str(tips))
check("③⑤ 提示里写明已存在则覆盖",
      any("已存在则覆盖" in t for t in tips), str(tips))
check("④⑥ 的提示是「选文件夹」",
      any("存放的文件夹" in t for t in tips), str(tips))
check("④⑥ 提示里写明文件名不变",
      any("文件名保持" in t for t in tips), str(tips))

tmp = Path(tempfile.mkdtemp(prefix="pick_"))
old_dir = tmp / "旧的"
old_dir.mkdir()
other = tmp / "另一个文件夹"
other.mkdir()
calls = []
picked = {"v": ""}


def fake_open(parent=None, title="", path="", flt="", *a, **k):
    calls.append((str(title), str(path), str(flt)))
    return picked["v"], ""


def fake_dir(parent=None, title="", directory="", *a, **k):
    calls.append((str(title), str(directory)))
    return picked["v"]


cg.QFileDialog.getOpenFileName = staticmethod(fake_open)
cg.QFileDialog.getExistingDirectory = staticmethod(fake_dir)

print("\n== ③ 订单 / ⑤ 评分：浏览选中一个文件 ==")
win._last_out_dir = ""
probe = old_dir / "某店_商品评分.xlsx"
probe.write_text("x")
win.ed_rating_out.setText(str(probe))
target = other / "手工选的评分.csv"
target.write_text("x")
picked["v"] = str(target)
win._pick_output_file(win.ed_rating_out, cg.DEFAULT_RATING_OUT)
check("⑤ 选中的文件原样写回框里", win.ed_rating_out.text() == str(target),
      win.ed_rating_out.text())
check("★ 弹的是「浏览选择」框（标题「选择文件」），不是「另存为」",
      calls[-1][0] == "选择文件", calls[-1][0])
check("过滤器只放真能写的两种（xlsx + csv，不给 .xls/.xlsm）",
      "*.xlsx" in calls[-1][2] and "*.csv" in calls[-1][2]
      and "*.xls;" not in calls[-1][2] and "*.xlsm" not in calls[-1][2],
      calls[-1][2])
check("起始路径 = 原目录 + 原文件名（光标直接落在那个文件上）",
      calls[-1][1] == str(old_dir / "某店_商品评分.xlsx"), calls[-1][1])
check("起始位置落在原文件所在目录", Path(calls[-1][1]).parent == old_dir, calls[-1][1])

picked["v"] = str(old_dir / "另一个文件.xlsx")
win._pick_output_file(win.ed_rating_out, cg.DEFAULT_RATING_OUT)
check("换一个文件 → 也原样写回",
      win.ed_rating_out.text() == str(old_dir / "另一个文件.xlsx"),
      win.ed_rating_out.text())
check("起始目录已跟着上一次的选择走（other）",
      calls[-1][1] == str(other / "手工选的评分.csv"), calls[-1][1])

win._last_out_dir = ""
win.ed_order_out.setText(str(old_dir / "某店_订单.xlsx"))
picked["v"] = str(other / "订单.xlsx")
win._pick_output_file(win.ed_order_out, cg.DEFAULT_ORDER_OUT)
check("③ 订单是同一套（浏览选中即填）",
      win.ed_order_out.text() == str(other / "订单.xlsx"), win.ed_order_out.text())
check("③ 起始路径同样是原目录+原文件名",
      calls[-1][1] == str(old_dir / "某店_订单.xlsx"), calls[-1][1])

print("\n== 选中 .csv → 真的写成 csv ==")
picked["v"] = str(other / "评分.csv")
win._pick_output_file(win.ed_rating_out, cg.DEFAULT_RATING_OUT)
te.write_rows(["rating", "order_id"], [["1", "2"]], Path(win.ed_rating_out.text()))
data = Path(win.ed_rating_out.text()).read_bytes()
check("★ 不是把 xlsx 塞进 .csv（没有 PK 头）", data[:2] != b"PK", repr(data[:2]))
check("csv 带 BOM（Excel 双击不乱码）", data[:3] == b"\xef\xbb\xbf", repr(data[:3]))

print("\n== ④ 客户 / ⑥ 差评：仍然只选文件夹，文件名不变 ==")
win._last_out_dir = ""
win.ed_cust_out.setText(str(old_dir / "某店_客户.xlsx"))
picked["v"] = str(tmp)
win._pick_output_dir(win.ed_cust_out, cg.DEFAULT_CUSTOMER_OUT)
check("客户框：只换存放目录", win.ed_cust_out.text() == str(tmp / "某店_客户.xlsx"),
      win.ed_cust_out.text())
check("走的是目录对话框（标题「选择存放目录」）", calls[-1][0] == "选择存放目录", calls[-1][0])
win.ed_bad_out.setText(str(old_dir / "某店_客户差评信息.xlsx"))
win._pick_output_dir(win.ed_bad_out, cg.DEFAULT_BAD_CUSTOMER_OUT)
check("差评框：只换存放目录",
      win.ed_bad_out.text() == str(tmp / "某店_客户差评信息.xlsx"),
      win.ed_bad_out.text())

print("\n== 取消 ==")
picked["v"] = ""
before = win.ed_rating_out.text()
win._pick_output_file(win.ed_rating_out, cg.DEFAULT_RATING_OUT)
check("③⑤ 取消 → 一个字都不动", win.ed_rating_out.text() == before,
      win.ed_rating_out.text())
before = win.ed_cust_out.text()
win._pick_output_dir(win.ed_cust_out, cg.DEFAULT_CUSTOMER_OUT)
check("④⑥ 取消 → 一个字都不动", win.ed_cust_out.text() == before,
      win.ed_cust_out.text())

print(f"\n合计 {OK} 项通过，{FAIL} 项失败")
sys.exit(1 if FAIL else 0)
