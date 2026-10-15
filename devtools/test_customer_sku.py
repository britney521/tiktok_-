"""客户表「SKU ID」列 —— 离线回归测试（不联网、不碰 Qt、不需要紫鸟）。

要的是**商品 SKU 的平台 ID**（``sku_module[].sku_id``），不是商家自定义的 Seller SKU。
主源随请求一起回来，订单表只做兜底，所以两边都要守住：

    1. 客户表确实有这一列，且紧跟 Order ID
    2. 从响应里取 sku_id：多 SKU 合并去重、**不拿 seller_sku_name 顶上**
    3. 接口优先于订单表（订单表只兜底）
    4. 订单表兜底：字段说明行要跳过、一个订单多 SKU 要合并
    5. 列名变体（``SKU ID`` / ``商品 SKU`` / 大小写）都能认
    6. 表里没有 SKU 列 → 返回空 dict，而不是抛错中断流水线
    7. 失败占位行的 SKU ID 不能丢

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
    print("1) 客户表要有「SKU ID」列")
    check("列存在", "SKU ID" in tc.CUSTOMER_COLUMNS)
    check("紧跟 Order ID", tc.CUSTOMER_COLUMNS[1] == "SKU ID", str(tc.CUSTOMER_COLUMNS[:3]))
    check("列顺序里只出现一次", tc.CUSTOMER_COLUMNS.count("SKU ID") == 1)
    check("没有误加商家 SKU 列",
          not any("seller" in c.lower() for c in tc.CUSTOMER_COLUMNS))

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

    print("3) 接口优先，订单表只兜底")
    check("响应有值时不覆盖",
          tc.parse_customer(order(), "999")["SKU ID"] == "1732505384254411635")
    check("响应没 sku 节点 → 用兜底", tc.parse_customer(order(""), "999")["SKU ID"] == "999")
    check("两边都没有 → 留空", tc.parse_customer(order(""))["SKU ID"] == "")

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

    print("5) 列名变体")
    check("商品 SKU 中文名",
          tc.read_sku_map(write_xlsx(TMP / "cn.xlsx", [
              ["Order ID", "商品 SKU"], ["577616376050126944", "CN-01"]]))
          == {"577616376050126944": "CN-01"})
    check("Sku ID 大小写",
          tc.read_sku_map(write_xlsx(TMP / "case.xlsx", [
              ["Order ID", "Sku ID"], ["577616376050126944", "ID-02"]]))
          == {"577616376050126944": "ID-02"})

    print("6) 没有 SKU 列不能炸")
    check("缺列返回空 dict",
          tc.read_sku_map(write_xlsx(TMP / "nocol.xlsx", [["Order ID", "Recipient"]]))
          == {})
    check("不存在的文件才抛", isinstance(
        _raises(lambda: tc.read_sku_map(TMP / "nope.xlsx")), Exception))

    print("7) 占位行")
    check("占位行带 SKU ID", tc.blank_customer("577616376050126944", "SKIPPED", "1 | 2")
          ["SKU ID"] == "1 | 2")
    check("占位行不带时为空", tc.blank_customer("1", "EMPTY")["SKU ID"] == "")
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

    print("9) 写表：列数对齐")
    out = tc.write_customers([tc.parse_customer(order()),
                              tc.blank_customer("2", "EMPTY", "B-2")],
                             TMP / "customers.xlsx")
    head = tc.read_header(out)
    check("表头含 SKU ID", "SKU ID" in head, str(head))
    check("表头长度 == CUSTOMER_COLUMNS", len(head) == len(tc.CUSTOMER_COLUMNS))

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
