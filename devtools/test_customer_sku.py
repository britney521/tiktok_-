"""客户表的两列 SKU —— 离线回归测试（不联网、不碰 Qt、不需要紫鸟）。

客户表紧挨着放两列，是同一件商品 SKU 的**两个 ID**（对齐订单表那两列）：

    · 「SKU ID」  = 商品 SKU 的**平台 ID**（``sku_module[].sku_id``）
    · 「商家 SKU」= 商家自己填的 **Seller SKU**（``sku_module[].seller_sku_name``），
      如 ``X004RIMEUV`` / ``WT-1``

两列的主源都随请求一起回来，订单表只做兜底（``SKU ID`` / ``Seller SKU`` 列），
所以两边都要守住：

    1. 客户表确实有这两列，且紧挨着（Order ID / SKU ID / 商家 SKU）
    2. 从响应里取值：多 SKU 合并去重、**两列互不串味**
    3. 接口优先于订单表（订单表只兜底）
    4. 订单表兜底：字段说明行要跳过、一个订单多 SKU 要合并
    5. 列名变体（``SKU ID`` / ``商品 SKU`` / ``Seller SKU`` / 大小写）要分得清，
       ★ 「商品 SKU」属于**平台 ID** 那组别名，不能被当成商家 SKU
    6. 表里没有 SKU 列 → 返回空 dict，而不是抛错中断流水线
    7. 失败占位行的两列都不能丢

用法：``python devtools/test_customer_sku.py``
"""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import openpyxl  # noqa: E402
import tiktok_client as tc  # noqa: E402
from customer_gui import PipelineThread  # noqa: E402

OK = 0
FAIL = 0
TMP = Path(tempfile.mkdtemp(prefix="sku_test_"))


def check(name: str, got: bool, detail: str = "") -> None:
    global OK, FAIL
    if got:
        OK += 1
        print(f"  [OK  ] {name}")
    else:
        FAIL += 1
        print(f"  [FAIL] {name}" + (f"  ← {detail}" if detail else ""))


def write_xlsx(path: Path, rows: list[list]) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def order(sku_module: object = None, **extra) -> dict:
    node: dict = {"main_order_id": "577616376050126944", "buyer_info_module": {}}
    if sku_module is None:
        node["sku_module"] = [{"sku_id": "1732505384254411635"}]
    elif sku_module == "":              # 显式表示「没有这个节点」
        pass
    else:
        node["sku_module"] = sku_module
    node.update(extra)
    return node


# 平台导出文件的形状：表头 + 一行字段说明 + 数据行（含一个订单两个 SKU）
SAMPLE = [
    ["Order ID", "SKU ID", "Seller SKU", "Buyer Nickname"],
    ["Platform unique order ID.", "Platform SKU ID.",
     "Seller sku input by the seller.", None],
    ["577616376050126944", "1732505384254411635", "130-F24", "Alice"],
    ["577616320746459381", "1732505384254411636", "130-F24", "Bob"],
    ["577616320746459381", "1732505384254411636", "130-F24", "Bob"],   # 重复行 → 去重
    ["577616320746459381", "1732505384254411637", "130-F25", "Bob"],   # 同单第二个 SKU
    ["577616282657263723", None, None, "Cindy"],
    [None, "Ghost-ID", None, None],                     # 没有订单号 → 跳过
    ["", "", "", ""],
    ["订单说明文字", "x", "y", "z"],                      # 说明行变体 → 跳过
]


def main() -> int:
    print("1) 客户表要有「SKU ID」+「商家 SKU」，紧挨着")
    check("SKU ID 列存在", "SKU ID" in tc.CUSTOMER_COLUMNS)
    check("SKU ID 紧跟 Order ID", tc.CUSTOMER_COLUMNS[1] == "SKU ID",
          str(tc.CUSTOMER_COLUMNS[:4]))
    check("SKU ID 只出现一次", tc.CUSTOMER_COLUMNS.count("SKU ID") == 1)
    check("★ 商家 SKU 紧跟 SKU ID（第 3 列）",
          tc.CUSTOMER_COLUMNS[2] == "商家 SKU", str(tc.CUSTOMER_COLUMNS[:4]))
    check("商家 SKU 只出现一次", tc.CUSTOMER_COLUMNS.count("商家 SKU") == 1)
    check("两列之间不夹别的列",
          tc.CUSTOMER_COLUMNS[1:3] == ["SKU ID", "商家 SKU"])

    print("2) 主源：响应里的 sku_id")
    check("单 SKU", tc.parse_customer(order())["SKU ID"] == "1732505384254411635",
          repr(tc.parse_customer(order())["SKU ID"]))
    check("多 SKU 合并去重",
          tc.parse_customer(order([{"sku_id": "1"}, {"sku_id": "1"}, {"sku_id": "2"}]))
          ["SKU ID"] == "1 | 2")
    check("sku_list 也算", tc.parse_customer(
        {"main_order_id": "1", "sku_list": [{"sku_id": "9"}]})["SKU ID"] == "9")
    # 同一个 sku 节点里既有 seller_sku 又有 sku_id 时，必须取后者
    mixed = order([{"sku_id": "1732505384254411635", "seller_sku_name": "130-F24",
                    "seller_sku": "130-F24"}])
    check("不拿 seller_sku 顶上", tc.parse_customer(mixed)["SKU ID"]
          == "1732505384254411635", repr(tc.parse_customer(mixed)["SKU ID"]))
    check("只有 seller_sku 时不瞎填",
          tc.parse_customer(order([{"seller_sku_name": "130-F24"}]))["SKU ID"] == "")

    print("2b) 商家 SKU：主源是响应里的 seller_sku_name")
    check("单 SKU", tc.parse_customer(order([
        {"sku_id": "1732505384254411635", "seller_sku_name": "X004RIMEUV"}]))
        ["商家 SKU"] == "X004RIMEUV")
    check("多 SKU 合并去重", tc.parse_customer(order([
        {"seller_sku_name": "A"}, {"seller_sku_name": "A"}, {"seller_sku_name": "B"}]))
        ["商家 SKU"] == "A | B")
    check("sku_list 也算", tc.parse_customer(
        {"main_order_id": "1",
         "sku_list": [{"seller_sku_name": "WT-1"}]})["商家 SKU"] == "WT-1")
    check("seller_sku（不带 _name）也认", tc.parse_customer(order([
        {"seller_sku": "B-2"}]))["商家 SKU"] == "B-2")
    # ★ 关键：同一个节点里两个 ID 都在，必须各取各的，不能串 ★
    check("★ 两列互不串味",
          (tc.parse_customer(mixed)["SKU ID"] == "1732505384254411635"
           and tc.parse_customer(mixed)["商家 SKU"] == "130-F24"),
          f"SKU ID={tc.parse_customer(mixed)['SKU ID']!r} "
          f"商家 SKU={tc.parse_customer(mixed)['商家 SKU']!r}")
    check("只有平台 ID 时商家 SKU 不瞎填（不拿 sku_id 顶上）",
          tc.parse_customer(order([{"sku_id": "1"}]))["商家 SKU"] == "")
    check("值为 None 的商家 SKU 当没有", tc.parse_customer(order([
        {"sku_id": "1", "seller_sku_name": None}]))["商家 SKU"] == "")

    print("3) 接口优先，订单表只兜底")
    check("响应有值时不覆盖",
          tc.parse_customer(order(), "999")["SKU ID"] == "1732505384254411635")
    check("响应没 sku 节点 → 用兜底", tc.parse_customer(order(""), "999")["SKU ID"] == "999")
    check("两边都没有 → 留空", tc.parse_customer(order(""))["SKU ID"] == "")
    check("商家 SKU 同样：响应没有 → 用兜底",
          tc.parse_customer(order(), "9", "FALLBACK")["商家 SKU"] == "FALLBACK")
    check("商家 SKU 同样：响应有 → 兜底不覆盖",
          tc.parse_customer(order([{"sku_id": "1", "seller_sku_name": "REAL"}]),
                            "9", "FALLBACK")["商家 SKU"] == "REAL")
    check("商家 SKU 两边都没有 → 留空",
          tc.parse_customer(order(""))["商家 SKU"] == "")
    check("两个兜底值一起传，各进各的列",
          (tc.parse_customer(order(""), "ID-1", "SKU-1")["SKU ID"] == "ID-1"
           and tc.parse_customer(order(""), "ID-1", "SKU-1")["商家 SKU"] == "SKU-1"))

    print("4) 订单表兜底映射")
    xlsx = write_xlsx(TMP / "orders.xlsx", SAMPLE)
    got = tc.read_sku_map(xlsx)
    check("读到 3 个订单", len(got) == 3, str(sorted(got)))
    check("单 SKU", got.get("577616376050126944") == "1732505384254411635",
          repr(got.get("577616376050126944")))
    check("多 SKU 合并去重",
          got.get("577616320746459381") == "1732505384254411636 | 1732505384254411637",
          repr(got.get("577616320746459381")))
    check("值为空的订单仍在表里", got.get("577616282657263723") == "")
    check("说明行没被当成订单", "订单说明文字" not in got)
    check("认的是 SKU ID 列而不是 Seller SKU 列",
          "130-F24" not in " ".join(got.values()), str(got))

    print("4b) 商家 SKU 也从订单表兜底（换一组列名）")
    got_s = tc.read_sku_map(xlsx, sku_columns=tc.SELLER_SKU_SOURCE_COLUMNS)
    check("读到 3 个订单", len(got_s) == 3, str(sorted(got_s)))
    check("单 SKU", got_s.get("577616376050126944") == "130-F24",
          repr(got_s.get("577616376050126944")))
    check("同一单多个商家 SKU 合并去重",
          got_s.get("577616320746459381") == "130-F24 | 130-F25",
          repr(got_s.get("577616320746459381")))
    check("★ 没把平台 SKU ID 混进来",
          "1732505384254411635" not in " ".join(got_s.values()), str(got_s))
    check("说明行同样跳过", "Seller sku" not in got_s)

    print("5) 列名变体")
    check("商品 SKU 中文名（平台 ID 那组）",
          tc.read_sku_map(write_xlsx(TMP / "cn.xlsx", [
              ["Order ID", "商品 SKU"], ["577616376050126944", "CN-01"]]))
          == {"577616376050126944": "CN-01"})
    check("Sku ID 大小写",
          tc.read_sku_map(write_xlsx(TMP / "case.xlsx", [
              ["Order ID", "Sku ID"], ["577616376050126944", "ID-02"]]))
          == {"577616376050126944": "ID-02"})
    check("商家 SKU 中文名",
          tc.read_sku_map(write_xlsx(TMP / "s_cn.xlsx", [
              ["Order ID", "商家 SKU"], ["577616376050126944", "CN-9"]]),
              sku_columns=tc.SELLER_SKU_SOURCE_COLUMNS)
          == {"577616376050126944": "CN-9"})
    check("Seller 大小写不敏感",
          tc.read_sku_map(write_xlsx(TMP / "s_case.xlsx", [
              ["Order ID", "seller sku"], ["577616376050126944", "S-03"]]),
              sku_columns=tc.SELLER_SKU_SOURCE_COLUMNS)
          == {"577616376050126944": "S-03"})
    # ★ 「商品 SKU」是平台 ID 的中文别名 —— 找商家 SKU 时不能把它当成命中 ★
    check("★ 找商家 SKU 时「商品 SKU」列不算命中（否则两边取到同一列）",
          tc.read_sku_map(write_xlsx(TMP / "s_mix.xlsx", [
              ["Order ID", "商品 SKU"], ["577616376050126944", "1732"]]),
              sku_columns=tc.SELLER_SKU_SOURCE_COLUMNS) == {})
    check("两组别名不重叠",
          not (set(map(tc.norm_col, tc.SKU_SOURCE_COLUMNS))
               & set(map(tc.norm_col, tc.SELLER_SKU_SOURCE_COLUMNS))),
          f"{tc.SKU_SOURCE_COLUMNS} ∩ {tc.SELLER_SKU_SOURCE_COLUMNS}")

    print("6) 没有 SKU 列不能炸")
    check("缺列返回空 dict",
          tc.read_sku_map(write_xlsx(TMP / "nocol.xlsx", [["Order ID", "Recipient"]]))
          == {})
    check("找商家 SKU 缺列也返回空 dict",
          tc.read_sku_map(write_xlsx(TMP / "nocol2.xlsx", [["Order ID", "Recipient"]]),
                          sku_columns=tc.SELLER_SKU_SOURCE_COLUMNS) == {})
    check("不存在的文件才抛", isinstance(
        _raises(lambda: tc.read_sku_map(TMP / "nope.xlsx")), Exception))

    print("7) 占位行")
    check("占位行带 SKU ID", tc.blank_customer("577616376050126944", "SKIPPED", "1 | 2")
          ["SKU ID"] == "1 | 2")
    check("占位行带商家 SKU",
          tc.blank_customer("577616376050126944", "SKIPPED", "1 | 2", "WT-1")
          ["商家 SKU"] == "WT-1")
    check("占位行不带时两列都为空",
          (tc.blank_customer("1", "EMPTY")["SKU ID"] == ""
           and tc.blank_customer("1", "EMPTY")["商家 SKU"] == ""))
    check("占位行仍按表顺序输出",
          list(tc.blank_customer("1", "EMPTY").keys()) == list(tc.CUSTOMER_COLUMNS))

    print("8) 流水线：③导出的行直接折叠")
    rows = [
        {"Order ID": "577616376050126944", "SKU ID": "1732505384254411635"},
        {"Order ID": "577616320746459381", "SKU ID": "1732505384254411636"},
        {"Order ID": "577616320746459381", "SKU ID": "1732505384254411637"},
    ]
    check("多行同一单合并", PipelineThread._sku_map_from_rows(rows)
          == {"577616376050126944": "1732505384254411635",
              "577616320746459381": "1732505384254411636 | 1732505384254411637"})
    check("没有 SKU 列返回空", PipelineThread._sku_map_from_rows(
        [{"Order ID": "1", "Recipient": "x"}]) == {})
    check("空 rows 返回空", PipelineThread._sku_map_from_rows([]) == {})
    # 真机订单表那一行同时有 SKU ID 和 Seller SKU：两组别名必须各认各的
    both = [{"Order ID": "577616376050126944", "SKU ID": "1732505384254411635",
             "Seller SKU": "X004RIMEUV"}]
    check("同一行按默认列名折叠 → 取平台 ID",
          PipelineThread._sku_map_from_rows(both)
          == {"577616376050126944": "1732505384254411635"})
    check("★ 同一行按商家列名折叠 → 取 Seller SKU",
          PipelineThread._sku_map_from_rows(both, tc.SELLER_SKU_SOURCE_COLUMNS)
          == {"577616376050126944": "X004RIMEUV"})

    print("9) 写表：列数对齐")
    out = tc.write_customers([tc.parse_customer(order()),
                              tc.blank_customer("2", "EMPTY", "B-2", "SKU-2")],
                             TMP / "customers.xlsx")
    head = tc.read_header(out)
    check("表头含 SKU ID", "SKU ID" in head, str(head))
    check("表头含商家 SKU", "商家 SKU" in head, str(head))
    check("表头长度 == CUSTOMER_COLUMNS", len(head) == len(tc.CUSTOMER_COLUMNS))
    check("商家 SKU 落在第 3 列", head[2] == "商家 SKU", str(head[:4]))

    print(f"\n通过 {OK} 项，失败 {FAIL} 项（临时目录 {TMP}）")
    return 1 if FAIL else 0


def _raises(fn) -> BaseException | None:
    try:
        fn()
    except Exception as exc:      # noqa: BLE001 - 测试里就是要接住
        return exc
    return None


if __name__ == "__main__":
    raise SystemExit(main())
