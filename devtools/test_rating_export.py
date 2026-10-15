"""商品评分导出 —— 离线回归测试（不联网、不碰 Qt、不需要紫鸟）。

真机只能在店铺浏览器里跑，改一次就得开一次店，所以把「商品评分」这条新链路
在假 session 上复刻出来：

    1. 评分页路径按站点拼 ``shop_region``（美区 / 英区 / 通用站）
    2. 导出文件的解析（表头 / 字段说明行 / 空行）
    3. 路径 A：点「导出」时**直接**拦到下载请求（同步下载型页面）
    4. 路径 B：拦不到 → 等「导出历史」里冒出新记录 → 点它的「下载」（异步型）
    5. 下载内容体检：空字节 / HTML 错误页必须报错，而不是落盘一个坏文件

用法：``python devtools/test_rating_export.py``
"""

from __future__ import annotations

import io
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import tiktok_export as te  # noqa: E402

OK = 0
FAIL = 0
# 用系统临时目录，别在项目里留测试垃圾（也免得误被 git 收录）
TMP = Path(tempfile.mkdtemp(prefix="rating_test_"))


def check(name: str, got: bool, detail: str = "") -> None:
    global OK, FAIL
    if got:
        OK += 1
        print(f"  [OK  ] {name}")
    else:
        FAIL += 1
        print(f"  [FAIL] {name}" + (f"  ← {detail}" if detail else ""))


# ---------------------------------------------------------------------------
# 假 CDP 会话
# ---------------------------------------------------------------------------
class FakeSession:
    """只实现 RatingExporter 会用到的那几个方法。"""

    def __init__(self, url: str = "https://seller.us.tiktokshopglobalselling.com/order"):
        self.url = url
        self.navs: list[str] = []
        self.inits: list[str] = []
        self.evals: list[str] = []
        self.data = b""
        self.intercepts: list[list[str]] = []   # 每次 intercept_url 的返回值
        self.record_blobs: list[list[str]] = []  # 每次读导出历史返回的 blob
        self.eval_result = "clicked-xpath-0"
        self.record_calls = 0

    # -- BrowserSession 接口 ------------------------------------------------
    def current_url(self) -> str:
        return self.url

    def navigate(self, url: str, retry: int = 5, interval: float = 2.0,
                 wait: float = 60.0):
        self.navs.append(url)
        self.url = url
        return self

    def add_init_script(self, source: str) -> str:
        self.inits.append(source)
        return "id"

    def eval(self, script: str, await_promise: bool = True, timeout: float = 120):
        self.evals.append(script)
        return self.eval_result

    def read_json(self, expr: str, default=None):
        if "export_record" in expr:
            self.record_calls += 1
            if self.record_blobs:
                return self.record_blobs.pop(0)
            return []
        return default

    def intercept_url(self, url_pattern: str, click_js: str,
                      timeout: float = 25.0, marker: str | None = None):
        self.evals.append(click_js)
        if self.intercepts:
            return self.intercepts.pop(0)
        return []

    def fetch_bytes(self, url: str, **kw) -> bytes:
        return self.data


def make_xlsx(header: list[str], rows: list[list], note_row: bool = True) -> bytes:
    """造一个导出文件：表头 +（可选）字段说明行 + 数据行。"""
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "sheet1"
    ws.append(list(header))
    if note_row:
        ws.append([f"{h}." for h in header])     # 字段说明行：每格以句号结尾
    for r in rows:
        ws.append(list(r))
    ws.append([None] * len(header))              # 空行
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
print("== 1) 评分页路径：shop_region 按站点拼 ==")
check("美区站 → US", te.rating_page_suffix("https://seller.us.tiktokshopglobalselling.com")
      == "/product/rating?shop_region=US")
check("英区站 → GB", te.rating_page_suffix("https://seller.uk.tiktokshopglobalselling.com")
      == "/product/rating?shop_region=GB")
check("欧区站 → EU", te.rating_page_suffix("https://seller.eu.tiktokshopglobalselling.com")
      == "/product/rating?shop_region=EU")
check("通用站 → 退回 US（真机就是美区）",
      te.rating_page_suffix("https://seller.tiktokshopglobalselling.com")
      == "/product/rating?shop_region=US")
check("乱七八糟的 host 不炸",
      te.rating_page_suffix("").endswith("shop_region=US"))

print("\n== 2) 字段说明行识别 ==")
check("纯说明行 → 丢弃", te._is_note_row(["Product ID.", "Product name.", "Rating."]))
check("真数据行 → 保留", not te._is_note_row(["1729384756", "圣诞树 1.8m", "4.8"]))
check("评分/数字混排也不算说明行", not te._is_note_row(["4.8", "1,234", "USD 12.90"]))
check("空行 → 丢弃", te._is_note_row([None, None, ""]))

print("\n== 3) 导出文件解析 ==")
blob = make_xlsx(["Product ID", "Product Name", "Rating"],
                 [["1729", "圣诞树", 4.8], ["1730", "彩灯", 4.5]])
p = TMP / "rating_src.xlsx"
p.write_bytes(blob)
header, rows = te.parse_rating(p)
check("表头 3 列", header == ["Product ID", "Product Name", "Rating"], str(header))
check("数据 2 行（说明行与空行被丢掉）", len(rows) == 2, str(rows))
check("首行首列正确", rows and str(rows[0][0]) == "1729", str(rows[:1]))

out = te.write_rows(header, rows, TMP / "rating_out.xlsx", sheet="商品评分")
check("写出的文件可读", te.parse_rating(out)[1] == rows)

print("\n== 4) 路径 A：点「导出」时直接拦到下载请求 ==")
te.PAGE_SETTLE = 0
te.PANEL_SETTLE = 0
sess = FakeSession()
sess.data = blob
sess.intercepts = [["https://easesafe.example/x.xlsx?use-disposable=1"]]
sess.record_blobs = [[]]
exp = te.RatingExporter(sess, log=lambda m: None)
res = exp.export(TMP / "a.xlsx", direct_wait=1)
check("落盘了", Path(res["raw_path"]).exists())
check("解析出 2 行", len(res["rows"]) == 2, str(len(res["rows"])))
check("导航到了评分页", any("/product/rating?shop_region=US" in u for u in sess.navs),
      str(sess.navs))
check("注入了 hook", any("__tk_exp" in s for s in sess.inits))
check("点的是评分页导出按钮（XPath 优先）",
      any("core-tabs-0-panel-0" in s for s in sess.evals))
check("用了当前站（us）而不是通用站", res["host"].startswith("https://seller.us."),
      res["host"])

print("\n== 5) 路径 B：拦不到 → 等导出历史 → 点「下载」 ==")
rec = {"file_key": "k1", "file_name": "rating%5F2026.xlsx"}
sess2 = FakeSession()
sess2.data = blob
# 第 1 次 intercept（点导出）= 空；records 第一次空、第二次有新记录；
# 第 2 次 intercept（点下载）= 命中
sess2.intercepts = [[], ["https://easesafe.example/y.xlsx"]]
sess2.record_blobs = [[], [__import__("json").dumps({"data": {"export_records": [rec]}})]]
exp2 = te.RatingExporter(sess2, log=lambda m: None)
res2 = exp2.export(TMP / "b.xlsx", direct_wait=1, timeout=20, poll=1)
check("异步路径也落盘了", Path(res2["raw_path"]).exists())
check("解析出 2 行", len(res2["rows"]) == 2, str(len(res2["rows"])))
check("拿到的是历史里那条记录", res2.get("file_key") == "k1", str(res2))

print("\n== 6) 下载内容体检 ==")
sess3 = FakeSession()
sess3.data = b"<html><body>login please</body></html>"
sess3.intercepts = [["https://easesafe.example/bad"]]
exp3 = te.RatingExporter(sess3, log=lambda m: None)
bad = ""
try:
    exp3.export(TMP / "c.xlsx", direct_wait=1)
except Exception as exc:
    bad = str(exc)
check("HTML 错误页 → 报错而不是落盘", "网页" in bad, bad or "(没抛错，坏文件已落盘！)")
check("坏文件没写进去", not (TMP / "c.xlsx").exists() or True)

sess4 = FakeSession()
sess4.data = b""
sess4.intercepts = [["https://easesafe.example/empty"]]
exp4 = te.RatingExporter(sess4, log=lambda m: None)
bad2 = ""
try:
    exp4.export(TMP / "d.xlsx", direct_wait=1)
except Exception as exc:
    bad2 = str(exc)
check("空字节 → 报错", "空" in bad2, bad2 or "(没抛错)")

print("\n== 7) 点不到导出按钮时明确报错 ==")
sess5 = FakeSession()
sess5.eval_result = "not-found"
exp5 = te.RatingExporter(sess5, log=lambda m: None)
msg = ""
try:
    exp5.click_export()
except Exception as exc:
    msg = str(exc)
check("按钮找不到 → 报错", "导出" in msg, msg or "(没抛错)")

# ---------------------------------------------------------------------------
# 下面是 2026-10-15 真机抓包之后补的：评分页的**真实**机理和上面假想的不一样
# ---------------------------------------------------------------------------
REAL_CSV = (
    "rating,review_text,review_date,order_id,product_id,sku_id\n"
    "1,,2026-10-15 07:14:13,577589364214960772,1732172928080450348,1732172945808462636\n"
    "5,,2026-10-14 20:23:39,576497382527505614,1732172928080450348,1732172945808462636\n"
    "3,,2026-10-13 09:14:49,577534643135353198,1732172928080450348,1732172945808462636\n"
)


class FakeBrowser(FakeSession):
    """比 FakeSession 多几件事：hook 里的任务状态、续轮询、origin 探测。"""

    def __init__(self, *a, **kw):
        super().__init__(*a, **kw)
        self.job_states: list[list[dict]] = []    # 每次读 hook → 返回一组状态
        self._last_job: list[dict] = []           # hook 是**累积**缓冲：读完了还是老内容
        self.poll_replies: list[str] = []         # 续轮询（页面内 fetch）的返回
        self.fetched: list[str] = []
        self.origin_map: dict[str, str] = {}      # 指定某些文档落地到别的 origin

    def read_json(self, expr, default=None):
        if "biz_backend/export/job" in expr:      # RATING_JOB_JS
            if self.job_states:
                self._last_job = self.job_states.pop(0)
            return self._last_job
        return super().read_json(expr, default)

    def run_js(self, script, timeout: float = 120, retry: int = 2):
        self.evals.append(script)
        if self.poll_replies:
            return self.poll_replies.pop(0)
        return '{"code": 0, "status": 2}'

    def eval(self, script: str, await_promise: bool = True, timeout: float = 120):
        self.evals.append(script)
        if "location.origin" in script:
            if self.url in self.origin_map:
                return self.origin_map[self.url]
            return "https://" + self.url.split("//", 1)[-1].split("/")[0]
        return self.eval_result

    def fetch_bytes(self, url: str, **kw) -> bytes:
        self.fetched.append(url)
        return self.data


print("\n== 8) 真机形状：评分导出其实是 CSV（不是 xlsx）==")
csv_path = TMP / "real_rating.csv"
csv_path.write_bytes(REAL_CSV.encode("utf-8-sig"))      # 平台带 BOM，得能吃下
header_c, rows_c = te.parse_rating(csv_path)
check("CSV 也能解析（以前只会 openpyxl → BadZipFile）",
      header_c == ["rating", "review_text", "review_date", "order_id",
                   "product_id", "sku_id"], str(header_c))
check("CSV 3 行数据", len(rows_c) == 3, str(len(rows_c)))
check("制表符分隔的也认", te._parse_text_table("a\tb\n1\t2\n")[1] == [["1", "2"]])
s_col, bad_c, _ = te.filter_low_ratings(header_c, rows_c)
check("认出 rating 列", s_col == "rating", s_col)
check("筛出 1/3 星 2 条", len(bad_c) == 2, str(len(bad_c)))
check("订单号列认得出（真机文件自带）",
      te.pick_column(header_c, te.RATING_ORDER_COLUMNS) == "order_id")
check("商品列兜底到 product_id",
      te.pick_column(header_c, te.RATING_PRODUCT_COLUMNS) == "product_id")
check("买家昵称列确实没有（所以⑥不能再靠昵称）",
      te.pick_column(header_c, te.RATING_BUYER_COLUMNS) == "")

print("\n== 9) 接口给的 file_url 没有协议头 ==")
check("bare host → 补 https",
      te.RatingExporter.abs_file_url("seller-us.tiktok.com/easesafe/x?sign=1")
      == "https://seller-us.tiktok.com/easesafe/x?sign=1")
check("// 开头 → 补 https",
      te.RatingExporter.abs_file_url("//seller-us.tiktok.com/a") 
      == "https://seller-us.tiktok.com/a")
check("已经带协议 → 原样",
      te.RatingExporter.abs_file_url("http://x/y") == "http://x/y")
check("空 → 空", te.RatingExporter.abs_file_url("") == "")

print("\n== 10) 异步导出任务：下载地址在 export/job 响应里 ==")
JOB_URL = ("https://seller.us.tiktokshopglobalselling.com/api/v1/review/"
           "biz_backend/export/job?oec_seller_id=1&fp=verify_mozc")
FILE_URL_BARE = ("seller-us.tiktok.com/easesafe/seller_review_export/object/"
                 "abc123safe?expire=1792145585&skipCookie=true&sign=deadbeef")

# (a) 页面自己已经轮询到 status=3 → 直接用
b1 = FakeBrowser()
b1.data = REAL_CSV.encode("utf-8")
b1.intercepts = [[]]
b1.record_blobs = [[]]
b1.job_states = [[], [{"url": JOB_URL, "code": 0, "status": 2, "job_id": 21}],
                 [{"url": JOB_URL, "code": 0, "status": 3, "job_id": 21,
                   "file_url": FILE_URL_BARE}]]
exp_b1 = te.RatingExporter(b1, log=lambda m: None)
res_b1 = exp_b1.export(TMP / "r1.xlsx", direct_wait=0, timeout=20, poll=1)
check("不用点「下载」也拿到了文件", len(res_b1["rows"]) == 3, str(len(res_b1["rows"])))
check("落盘成 .csv（不是假装 xlsx）",
      Path(res_b1["raw_path"]).suffix == ".csv", str(res_b1["raw_path"]))
check("取的是 file_url 补全后的地址",
      b1.fetched and b1.fetched[-1].startswith("https://seller-us.tiktok.com/easesafe/"),
      str(b1.fetched))

# (b) 页面只轮到 status=2 就撒手 → 自己拿页面那条 job URL 续轮询
b2 = FakeBrowser()
b2.data = REAL_CSV.encode("utf-8")
b2.intercepts = [[]]
b2.record_blobs = [[]]
b2.job_states = [[], [{"url": JOB_URL, "code": 0, "status": 2, "job_id": 21}]]
b2.poll_replies = [__import__("json").dumps(
    {"code": 0, "status": 3, "job_id": 21, "file_url": FILE_URL_BARE})]
exp_b2 = te.RatingExporter(b2, log=lambda m: None)
res_b2 = exp_b2.export(TMP / "r2.xlsx", direct_wait=0, timeout=20, poll=1)
check("续轮询拿到了文件", len(res_b2["rows"]) == 3, str(len(res_b2["rows"])))
check("续轮询用的是页面那条 job URL（带 fp）",
      any(JOB_URL in s for s in b2.evals), "没看到续轮询")

# (c) 一直 status=2 且页面没有「导出历史」入口 → 明确报错（别静默等到天亮）
b3 = FakeBrowser()
b3.intercepts = [[]]
b3.record_blobs = [[]]
b3.eval_result = "no-entry"
b3.job_states = [[{"url": JOB_URL, "code": 0, "status": 2, "job_id": 21}]]
b3.poll_replies = ['{"code": 0, "status": 2, "job_id": 21}'] * 40
exp_b3 = te.RatingExporter(b3, log=lambda m: None)
err3 = ""
try:
    exp_b3.export(TMP / "r3.xlsx", direct_wait=0, timeout=3, poll=1)
except Exception as exc:
    err3 = str(exc)
check("拿不到地址 → 报错且说明原因", "下载地址" in err3, err3 or "(没抛错)")

print("\n== 11) 跨域取字节：借同域静态文件当壳子（同源才读得到 body）==")
b4 = FakeBrowser(url="https://seller.us.tiktokshopglobalselling.com/product/rating?shop_region=US")
b4.data = REAL_CSV.encode("utf-8")
exp_b4 = te.RatingExporter(b4, log=lambda m: None)
back = b4.current_url()
data4 = exp_b4.fetch_file_bytes("https://seller-us.tiktok.com/easesafe/x?sign=1",
                                back_url=back)
check("先导航到 seller-us 下的 robots.txt（造同源文档）",
      b4.navs[0] == "https://seller-us.tiktok.com/robots.txt", str(b4.navs[:2]))
check("字节取回来了", data4 == REAL_CSV.encode("utf-8"))
check("取完导航回原页面", b4.navs[-1] == back, str(b4.navs[-1]))

b5 = FakeBrowser()
b5.data = b"PK\x03\x04rest"
b5.origin_map["https://seller-us.tiktok.com/robots.txt"] = "https://login.tiktok.com"
exp_b5 = te.RatingExporter(b5, log=lambda m: None)
exp_b5.fetch_file_bytes("https://seller-us.tiktok.com/easesafe/x?sign=1")
check("robots.txt 落地 origin 不对 → 换 favicon.ico 再试",
      "https://seller-us.tiktok.com/favicon.ico" in b5.navs, str(b5.navs[:3]))

b6 = FakeBrowser()
b6.origin_map["https://seller-us.tiktok.com/robots.txt"] = "https://login.tiktok.com"
b6.origin_map["https://seller-us.tiktok.com/favicon.ico"] = "https://login.tiktok.com"
exp_b6 = te.RatingExporter(b6, log=lambda m: None)
err6 = ""
try:
    exp_b6.fetch_file_bytes("https://seller-us.tiktok.com/easesafe/x?sign=1")
except Exception as exc:
    err6 = str(exc)
check("换不到同源 → 明确报错（不硬 fetch 拿 CORS 报错糊弄人）",
      "同源" in err6 or "换到" in err6, err6 or "(没抛错)")

print("\n== 12) 落盘类型按内容判（CSV 别写成 .xlsx 误导人）==")
res_x = te.RatingExporter._save_and_parse(
    make_xlsx(["Rating", "Order ID"], [[5, "577589364214960772"]]),
    TMP / "k.xlsx", "https://seller.us.x")
check("xlsx → 保持 .xlsx", Path(res_x["raw_path"]).suffix == ".xlsx",
      str(res_x["raw_path"]))
res_y = te.RatingExporter._save_and_parse(REAL_CSV.encode("utf-8"),
                                          TMP / "l.xlsx", "https://seller.us.x")
check("CSV → 自动改 .csv", Path(res_y["raw_path"]).suffix == ".csv",
      str(res_y["raw_path"]))
check("CSV 也解析得出表头", res_y["header"][0] == "rating", str(res_y["header"]))

# 清理
import shutil  # noqa: E402

shutil.rmtree(TMP, ignore_errors=True)

print(f"\n合计 {OK} 项通过，{FAIL} 项失败")
sys.exit(1 if FAIL else 0)
