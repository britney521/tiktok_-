"""差评链路（⑤→⑥）—— 离线回归测试（不联网、不碰 Qt、不需要紫鸟）。

真机⑤导出的「商品评分」其实是 **CSV，自带 order_id**（2026-10-15 实测），所以
1~3 星要变成客户信息，主链条是**订单号直连**：

    评分表(1~3 星，含 order_id) → order_id → order/get → 客户差评表

老版「只有昵称」的评分文件才走兜底：

    评分表(1~3 星) → 昵称(+商品) → 订单表换出 Order ID → order/get → 客户差评表

要守住的三段：

    1. 评分识别：列名别名（``Rating``/``评分``…）、值形态（``5``/``5.0``/``2星``/``★3``）、
       认不出评分列时**不筛选**（绝不把文件写成空的）
    2. 筛选：1~3 留下，4~5 走掉；筛选结果才落盘
    3. 订单号：自带列优先直连；认不出才退回「昵称 + 商品」两级匹配
       （脱敏昵称 ``K*****`` 匹配不上要如实报告）

用法：``python devtools/test_bad_review.py``
"""

from __future__ import annotations

import sys
import types
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import openpyxl  # noqa: E402
import tiktok_client as tc  # noqa: E402
import tiktok_export as te  # noqa: E402
from customer_gui import (  # noqa: E402
    BAD_CUSTOMER_COLUMNS, PipelineThread,
    rating_file_candidates, rating_order_column,
)

OK = 0
FAIL = 0
TMP = Path(tempfile.mkdtemp(prefix="bad_test_"))


def check(name: str, got: bool, detail: str = "") -> None:
    global OK, FAIL
    if got:
        OK += 1
        print(f"  [OK  ] {name}")
    else:
        FAIL += 1
        print(f"  [FAIL] {name}" + (f"  ← {detail}" if detail else ""))


class Stub:
    """冒充 PipelineThread 的 self（⑥ 的这几个方法只用到静态逻辑）。

    不实例化真的线程：QThread 一构造就要 Qt 运行环境，而本测试要保持零依赖。
    """

    opts: dict = {}
    _sku_map_cache: dict = {}
    _norm = staticmethod(PipelineThread._norm)
    _cols_of = staticmethod(PipelineThread._cols_of)


# 评分文件（真机形状：表头 + 说明行 + 数据；没有订单号列）
RATING_HEADER = ["Product Name", "Rating", "Buyer Nickname", "Review Content"]
RATING_ROWS = [
    ["圣诞树 1.8m", 5, "alice", "好评"],
    ["圣诞树 2.1m", 2, "bob", "一般"],
    ["圣诞树 2.1m", 1, "carol", "很差"],
    ["花环 24in", 3, "dave", "凑合"],
    ["花环 24in", 4, "erin", "还行"],
]

# 订单表（③ 导出的形状）
ORDER_ROWS = [
    {"Order ID": "577616376050126944", "Buyer Nickname": "Alice",
     "Product Name": "圣诞树 1.8m"},
    {"Order ID": "577616320746459381", "Buyer Nickname": "Bob",
     "Product Name": "圣诞树 2.1m"},
    {"Order ID": "577616282657263723", "Buyer Nickname": "Carol",
     "Product Name": "花环 24in"},
]


def main() -> int:
    print("1) 评分值解析")
    for raw, want in ((5, 5), ("5", 5), ("5.0", 5), ("2星", 2), ("★3", 3),
                      ("1 分", 1), ("4.6", 5), (None, 0), ("", 0),
                      ("没评分", 0), ("abc", 0)):
        check(f"{raw!r} → {want}", te.rating_value(raw) == want,
              str(te.rating_value(raw)))

    print("2) 列名识别（别名 / 空格 / 大小写）")
    check("Rating", te.pick_column(["Rating"], te.RATING_SCORE_COLUMNS) == "Rating")
    check("评分 中文", te.pick_column(["评分"], te.RATING_SCORE_COLUMNS) == "评分")
    check("Rating Score", te.pick_column(["Rating Score"], te.RATING_SCORE_COLUMNS)
          == "Rating Score")
    check("带前导空格", te.pick_column([" Rating "], te.RATING_SCORE_COLUMNS)
          == " Rating ")
    check("昵称 Buyer Nickname",
          te.pick_column(["Buyer Nickname"], te.RATING_BUYER_COLUMNS)
          == "Buyer Nickname")
    check("昵称 买家昵称",
          te.pick_column(["买家昵称"], te.RATING_BUYER_COLUMNS) == "买家昵称")
    check("不认识就给空串", te.pick_column(["SKU ID"], te.RATING_SCORE_COLUMNS) == "")

    print("3) 筛选 1~3 星")
    col, kept, scores = te.filter_low_ratings(RATING_HEADER, RATING_ROWS)
    check("认到评分列", col == "Rating", col)
    check("留下 3 条", len(kept) == 3, str(len(kept)))
    check("分数对得上", scores == [2, 1, 3], str(scores))
    check("4~5 星被剔掉",
          all(te.rating_value(r[1]) >= 4 for r in RATING_ROWS
              if r not in kept), "4/5 星漏进去了")
    # ★ 认不出评分列 → 原样返回，绝不把文件写成空的
    col2, kept2, _ = te.filter_low_ratings(["Product", "评论内容"], RATING_ROWS)
    check("认不出时不筛选", col2 == "" and len(kept2) == len(RATING_ROWS),
          f"{col2!r}/{len(kept2)}")

    print("4) 订单表 → 昵称索引")
    index = PipelineThread._order_index(Stub(), ORDER_ROWS)
    check("3 个买家", len(index) == 3, str(sorted(index)))
    check("昵称进了索引", "bob" in index, str(sorted(index)))
    check("同一单的商品也挂上了",
          PipelineThread._match_orders(Stub(), index, "bob", "圣诞树 2.1m")
          == ["577616320746459381"])
    check("只给昵称也能命中",
          PipelineThread._match_orders(Stub(), index, "carol", "")
          == ["577616282657263723"])
    check("昵称对不上 → 空",
          PipelineThread._match_orders(Stub(), index, "nobody", "") == [])
    # 平台脱敏成 K***** 时对不上 —— 如实返回空，不伪装成匹配
    check("脱敏昵称匹配不上",
          PipelineThread._match_orders(Stub(), index, "b***", "圣诞树 2.1m") == [])

    print("5) 没有订单行时回读订单 Excel")
    path = TMP / "orders.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Order ID", "Buyer Nickname", "Product Name"])
    ws.append(["Platform unique order ID.", "Buyer nickname.", "Product."])
    for row in ORDER_ROWS:
        ws.append([row["Order ID"], row["Buyer Nickname"], row["Product Name"]])
    wb.save(path)
    stub = Stub()
    stub.opts = {"customer_input": str(path)}
    from_file = PipelineThread._order_index(stub, [])
    check("回读成功且有 3 个买家", len(from_file) == 3, str(sorted(from_file)))
    check("回读的值与内存一致",
          PipelineThread._match_orders(stub, from_file, "alice", ""), str(from_file))

    print("6) 差评 → 订单（两级匹配）")
    _, kept3, _ = te.filter_low_ratings(RATING_HEADER, RATING_ROWS)
    idx = PipelineThread._order_index(Stub(), ORDER_ROWS)
    picked: list[str] = []
    seen: set[str] = set()
    missed = 0
    for row in kept3:
        nick = PipelineThread._norm(row[2])
        prod = PipelineThread._norm(row[0])
        ids = PipelineThread._match_orders(Stub(), idx, nick, prod)
        if not ids:
            missed += 1
            continue
        for oid in ids:
            if oid not in seen:
                seen.add(oid)
                picked.append(oid)
    check("3 条差评 → 2 单", len(picked) == 2, str(picked))
    check("没匹配到的算 1 条（dave 没下过单）", missed == 1, str(missed))
    check("命中顺序 = 差评顺序",
          picked == ["577616320746459381", "577616282657263723"], str(picked))

    print("7) 差评客户表列")
    check("= 客户列 + 评分 + 商品",
          BAD_CUSTOMER_COLUMNS[-2:] == ["评分", "商品"], str(BAD_CUSTOMER_COLUMNS[-3:]))
    check("跟④一样带「SKU ID」+「商家 SKU」两列（紧挨着）",
          BAD_CUSTOMER_COLUMNS[1:3] == ["SKU ID", "商家 SKU"],
          str(BAD_CUSTOMER_COLUMNS[:4]))
    check("前 N 列就是客户表本身（列定义只加不改）",
          BAD_CUSTOMER_COLUMNS[:-2] == list(tc.CUSTOMER_COLUMNS),
          str(BAD_CUSTOMER_COLUMNS[:4]))

    print("8) ⑥ 单独跑：读已有评分表（⑤ 不重跑）")

    class Sink:
        def __init__(self) -> None:
            self.msgs: list[str] = []

        def emit(self, msg=None, *_a, **_k) -> None:
            self.msgs.append(str(msg))

    class Runner(Stub):
        """把 ⑥ 单独跑那几个方法绑过来（log/stage 换成纯收集器，不碰 Qt）。"""

        _apply_rating_filter = PipelineThread._apply_rating_filter
        _load_rating_bad = PipelineThread._load_rating_bad
        # SKU 兜底这两条路（两张表）在 ⑥ 主循环里也会被调到
        _load_sku_map = PipelineThread._load_sku_map
        _sku_map_from_rows = PipelineThread._sku_map_from_rows
        _sku_col_of = staticmethod(PipelineThread._sku_col_of)
        _sku_pair = PipelineThread._sku_pair

        def __init__(self, opts: dict) -> None:
            self.opts = opts
            self.log = Sink()
            self.stage = Sink()
            self.rating_progress = Sink()
            self.rating_ready = Sink()
            self._sku_map_cache: dict = {}
            self._seller_sku_cache: dict = {}
            self._rating_header: list = []
            self._rating_cols: dict = {"score": "", "buyer": "", "product": ""}
            self._rating_bad: list = []
            self._rating_all: list = []
            self._bad_hit = False

    # (a) 评分表不存在 → 返回 False 且日志里说清楚，绝不静默跑下去
    a = Runner({"rating_out": str(TMP / "不存在的评分.xlsx")})
    check("文件不存在 → False", a._load_rating_bad() is False)
    check("日志里写清找不到", any("找不到评分文件" in m for m in a.log.msgs),
          str(a.log.msgs))

    # (b) 读完整版（_评分原始.xlsx）→ 筛成 1~3 星写进 out
    out_b = TMP / "评分_b.xlsx"
    raw_b = out_b.with_name("_评分原始.xlsx")
    te.write_rows(RATING_HEADER, RATING_ROWS, raw_b, sheet="商品评分")
    b = Runner({"rating_out": str(out_b)})
    check("从完整版读到 → True", b._load_rating_bad() is True)
    hb, rb = te.parse_rating(out_b)
    check("out 里只有 1~3 星", len(rb) == 3, str(len(rb)))
    check("完整版没被盖", len(te.parse_rating(raw_b)[1]) == len(RATING_ROWS))
    check("记住了评分列/昵称列",
          b._rating_cols["score"] == "Rating" and b._rating_cols["buyer"] == "Buyer Nickname",
          str(b._rating_cols))

    # (c) 再跑一次（读已筛过的 out）→ 幂等，行数不变
    c = Runner({"rating_out": str(out_b)})
    check("读已筛过的 → 仍然 True", c._load_rating_bad() is True)
    check("幂等：还是 3 行", len(te.parse_rating(out_b)[1]) == 3)

    # (d) 读的是「完整版还没备份过」的 out → 先把完整版另存一份
    out_d = TMP / "评分_d.xlsx"
    te.write_rows(RATING_HEADER, RATING_ROWS, out_d, sheet="商品评分")
    d = Runner({"rating_out": str(out_d)})
    check("未筛过的 out → True", d._load_rating_bad() is True)
    raw_d = out_d.with_name("_评分原始.xlsx")
    check("完整版被另存下来", raw_d.exists() and
          len(te.parse_rating(raw_d)[1]) == len(RATING_ROWS))
    check("out 变成只有差评", len(te.parse_rating(out_d)[1]) == 3)

    print("9) 真机 CSV：评分文件**自带 order_id** → 不再靠昵称反查")
    REAL_HEADER = ["rating", "review_text", "review_date", "order_id",
                   "product_id", "sku_id"]
    REAL_ROWS = [
        ["1", "", "2026-10-15 07:14:13", "577589364214960772",
         "1732172928080450348", "1732172945808462636"],
        ["3", "", "2026-10-13 09:14:49", "577586254751765136",
         "1732172928080450348", "1732302878131589932"],
        ["5", "", "2026-10-12 09:23:22", "577588884391367032",
         "1732172928080450348", "1732172945808462636"],
    ]
    check("认出 order_id 列",
          te.pick_column(REAL_HEADER, te.RATING_ORDER_COLUMNS) == "order_id")
    check("这份文件里确实没有买家昵称列（老路走不通）",
          te.pick_column(REAL_HEADER, te.RATING_BUYER_COLUMNS) == "")
    check("评分照样认得出", te.pick_column(REAL_HEADER, te.RATING_SCORE_COLUMNS)
          == "rating")
    check("订单号归一：去 # 与空白",
          PipelineThread._order_id(" #577589364214960772 ") == "577589364214960772")
    check("末尾 .0 去掉", PipelineThread._order_id("577589364214960772.0")
          == "577589364214960772")
    check("科学计数串也认",
          PipelineThread._order_id("5.7758936421496e+13") == "57758936421496")
    check("数字单元格（float）也认",
          PipelineThread._order_id(57758936421496.0) == "57758936421496")
    check("位数不够不当订单号", PipelineThread._order_id("123") == "")
    check("空 → 空串", PipelineThread._order_id(None) == "")

    print("10) ⑥ 走订单号列（真机形状）：不查订单表也能抓客户")

    class Client:
        def __init__(self) -> None:
            self.asked: list[str] = []

        def fetch(self, ids, chunk_size: int = 1):
            self.asked.extend(list(ids))
            return {"code": 0, "data": {"main_order": [
                {"main_order_id": ids[0],
                 "buyer_info_module": {"buyer_nickname": "buyer_" + ids[0][-4:]},
                 "sku_module": [{"sku_id": "1732172945808462636",
                                 "seller_sku_name": "X004RIMEUV"}]}]}}

    class Runner2(Runner):
        """⑥ 主循环的桩：把会碰浏览器/网络的几处换掉。"""

        _fetch_bad_customers = PipelineThread._fetch_bad_customers
        _bad_row = staticmethod(PipelineThread._bad_row)
        _flush_bad_customers = staticmethod(PipelineThread._flush_bad_customers)
        _order_id = staticmethod(PipelineThread._order_id)

        def __init__(self, opts: dict) -> None:
            super().__init__(opts)
            self.stop_event = types.SimpleNamespace(is_set=lambda: False)
            self.bad_progress = Sink()
            self._cust_cache = {}
            self._sku_map_cache = {"x": "y"}     # 非空 → 不去读订单文件
            self._seller_sku_cache = {"x": "y"}  # 同上（两张兜底表都得置上）
            self.asked: list[str] = []

        def _ensure_order_page(self, browser) -> None:
            self.asked.append("ensure_order_page")

    _, kept_real, _ = te.filter_low_ratings(REAL_HEADER, REAL_ROWS)
    check("筛出 1/3 星 2 条", len(kept_real) == 2, str(len(kept_real)))
    bad_rows_real = [dict(zip(REAL_HEADER, r)) for r in kept_real]
    r2 = Runner2({"bad_out": str(TMP / "客户差评.xlsx")})
    r2._rating_header = list(REAL_HEADER)
    r2._rating_cols = {"score": "rating", "buyer": "",
                       "product": "product_id", "order": "order_id"}
    cli = Client()
    got_c = r2._fetch_bad_customers(browser=None, client_obj=cli,
                                    bad_rows=bad_rows_real, origins=None,
                                    order_rows=[])
    check("两条差评各抓了一单",
          cli.asked == ["577589364214960772", "577586254751765136"], str(cli.asked))
    check("日志写明走的是「自带订单号列」",
          any("自带订单号列" in m for m in r2.log.msgs), str(r2.log.msgs[:4]))
    check("先回到订单页（站对页也要对）", "ensure_order_page" in r2.asked)
    check("成功 2 条", got_c.get("ok") == 2, str(got_c.get("ok")))
    check("差评客户表落盘", Path(got_c["out"]).exists(), str(got_c.get("out")))
    hdr_c, rows_c = te.parse_rating(got_c["out"])
    check("列 = 客户列 + 评分 + 商品", hdr_c == BAD_CUSTOMER_COLUMNS, str(hdr_c))
    check("评分跟着带进表", [r[list(hdr_c).index("评分")] for r in rows_c] == ["1", "3"],
          str(rows_c))
    check("商品跟着带进表（product_id 兜底）",
          [r[list(hdr_c).index("商品")] for r in rows_c]
          == ["1732172928080450348", "1732172928080450348"], str(rows_c))
    # ★ 商家 SKU 也在差评表里：差评是哪个 SKU 吃的，一眼能看出来 ★
    check("★ 商家 SKU 跟着带进表（响应里的 seller_sku_name）",
          [r[list(hdr_c).index("商家 SKU")] for r in rows_c]
          == ["X004RIMEUV", "X004RIMEUV"], str(rows_c))
    check("SKU ID 那一列不受影响",
          [r[list(hdr_c).index("SKU ID")] for r in rows_c]
          == ["1732172945808462636", "1732172945808462636"], str(rows_c))

    print("10b) ⑥ 单独跑：自己按订单表补**两张**兜底表")
    # 这次响应里没有 seller_sku_name → 商家 SKU 只能来自订单表兜底
    class Client3:
        def __init__(self) -> None:
            self.asked: list[str] = []

        def fetch(self, ids, chunk_size: int = 1):
            self.asked.extend(list(ids))
            return {"code": 0, "data": {"main_order": [
                {"main_order_id": ids[0],
                 "buyer_info_module": {"buyer_nickname": "b"},
                 "sku_module": [{"sku_id": "1732172945808462636"}]}]}}

    order_rows = [
        {"Order ID": "577589364214960772", "SKU ID": "1732172945808462636",
         "Seller SKU": "X004RIMEUV"},
        {"Order ID": "577586254751765136", "SKU ID": "1732302878131589932",
         "Seller SKU": "WT-1"},
    ]
    r4 = Runner2({"bad_out": str(TMP / "客户差评3.xlsx")})
    r4._sku_map_cache = {}          # 模拟「④ 没跑过」
    r4._seller_sku_cache = {}
    r4._rating_header = list(REAL_HEADER)
    r4._rating_cols = {"score": "rating", "buyer": "",
                       "product": "product_id", "order": "order_id"}
    got_d = r4._fetch_bad_customers(browser=None, client_obj=Client3(),
                                    bad_rows=bad_rows_real, origins=None,
                                    order_rows=order_rows)
    check("两张兜底表都从订单表建起来了",
          bool(r4._sku_map_cache) and bool(r4._seller_sku_cache),
          f"{r4._sku_map_cache} / {r4._seller_sku_cache}")
    check("日志里两张表各报一行",
          any("SKU ID 兜底表已就绪" in m for m in r4.log.msgs)
          and any("商家 SKU 兜底表已就绪" in m for m in r4.log.msgs),
          str([m for m in r4.log.msgs if "兜底表" in m]))
    hdr_d, rows_d = te.parse_rating(got_d["out"])
    check("★ 商家 SKU 落的是订单表里的那一列（响应没有时）",
          [r[list(hdr_d).index("商家 SKU")] for r in rows_d]
          == ["X004RIMEUV", "WT-1"],
          str([r[list(hdr_d).index("商家 SKU")] for r in rows_d]))
    # 这次 Client3 的响应里**有 sku_id、没有 seller_sku_name** → 正好证明两列
    # 是**各自**择源的：SKU ID 用响应（第 2 单的订单表值 1732302878131589932 被忽略），
    # 商家 SKU 用订单表兜底。任何一列取错源都会在这里露出来。
    check("★ 同一行两列各自择源：SKU ID 用响应 / 商家 SKU 用订单表",
          [r[list(hdr_d).index("SKU ID")] for r in rows_d]
          == ["1732172945808462636", "1732172945808462636"],
          str([r[list(hdr_d).index("SKU ID")] for r in rows_d]))

    print("11) ⑥ 没有订单号列、也没有昵称列 → 明确跳过不硬跑")
    r3 = Runner2({"bad_out": str(TMP / "客户差评2.xlsx")})
    r3._rating_header = ["Product Name", "Rating"]
    r3._rating_cols = {"score": "Rating", "buyer": "", "product": "Product Name",
                       "order": ""}
    got_x = r3._fetch_bad_customers(browser=None, client_obj=Client(),
                                    bad_rows=[{"Product Name": "圣诞树", "Rating": "2"}],
                                    origins=None, order_rows=[])
    check("返回空结果（不写空表）",
          got_x.get("ok") == 0 and got_x.get("out") == "", str(got_x))
    check("日志说明两种列都没认出来",
          any("既没订单号列" in m for m in r3.log.msgs), str(r3.log.msgs[:3]))

    print("12) ⑥ 的依赖判断：评分文件自带订单号 → **不需要**订单表")
    # 真机⑤导出的完整版是 CSV（自带 order_id），落成 _评分原始.csv
    sub_a = TMP / "r12a"
    sub_a.mkdir(exist_ok=True)
    csv_path = sub_a / "_评分原始.csv"
    csv_path.write_text(
        "rating,review_text,review_date,order_id,product_id,sku_id\n"
        "1,很差,2026-10-01,577589364214960772,1732172928080450348,111\n"
        "3,一般,2026-10-02,577586254751765136,1732172928080450348,222\n",
        encoding="utf-8")
    out_x = sub_a / "秋图2_商品评分.xlsx"        # 筛过的那份（本例故意不存在）
    cands = rating_file_candidates(out_x)
    check("候选里认 _评分原始.csv（真机完整版就是 CSV）",
          csv_path in cands, str([p.name for p in cands]))
    check("候选无重复", len(cands) == len({str(p) for p in cands}),
          str([p.name for p in cands]))
    check("只有 CSV 完整版也算「有评分表」→ 守卫不该拦",
          any(p.exists() for p in cands), str([p.name for p in cands]))
    check("认出自带的订单号列 order_id",
          rating_order_column(out_x) == "order_id",
          repr(rating_order_column(out_x)))

    # 筛过的那份 xlsx 自带 order_id 也认（用户就是这种情况：有评分表、没订单表）
    sub_b = TMP / "r12b"
    sub_b.mkdir(exist_ok=True)
    rated = sub_b / "秋图2_商品评分.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["rating", "review_text", "order_id", "product_id"])
    ws.append([2, "一般", "577589364214960772", "1732172928080450348"])
    wb.save(rated)
    check("筛过的 xlsx 也认 order_id（→ 不必再要订单表）",
          rating_order_column(rated) == "order_id",
          repr(rating_order_column(rated)))

    # 老版「仅昵称」文件 → 认不出订单号列 → 这时才需要订单表
    legacy = TMP / "老版评分.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Product Name", "Rating", "Buyer Nickname"])
    ws.append(["圣诞树", 1, "bob"])
    wb.save(legacy)
    check("老版「仅昵称」认不出订单号列（→ 才要订单表）",
          rating_order_column(legacy) == "",
          repr(rating_order_column(legacy)))

    # 既没文件、文件又坏 → 空串且不抛异常（守卫据此回落到「要订单表」）
    check("文件不存在不抛异常",
          rating_order_column(TMP / "根本没有.xlsx") == "")
    sub_c = TMP / "r12c"
    sub_c.mkdir(exist_ok=True)
    (sub_c / "_评分原始.xlsx").write_bytes(b"")      # 空字节 = 坏 xlsx
    check("坏/空评分文件不抛异常",
          rating_order_column(sub_c / "评分.xlsx") == "")

    print(f"\n通过 {OK} 项，失败 {FAIL} 项（临时目录 {TMP}）")
    return 1 if FAIL else 0


if __name__ == "__main__":
    raise SystemExit(main())
