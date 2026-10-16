# TikTok 订单 / 客户 / 商品评分导出 GUI（紫鸟店铺浏览器版）

六步流水线：**① 启动紫鸟环境 → ② 获取店铺列表 → ③ 选店铺导出订单 →
④ 读订单 Excel 抓客户信息 → ⑤ 导出商品评分（筛出 1~3 星后覆盖保存）→
⑥ 抓差评客户**。

⑤ 和 ⑥ 是**两颗独立的按钮**，可以分开跑：

- **「导出商品评分」**（`商品评分.xlsx`，默认 `<店铺名>_商品评分.xlsx`）：走评分页
  自己的导出按钮，导出后只留 1~3 星并覆盖保存；
- **「抓取差评客户（1~3 星）」**：**不开评分页**，直接读「评分 Excel」框里那份文件
  （筛过的是幂等的，没筛过的就地筛一遍再覆盖保存），然后抓差评客户；
- 「一键完成」可以两个都带上（旁边「含商品评分」「含差评客户」默认都勾）。

⑤⑥ **放在最后**：评分页会把浏览器带离订单页，而 ④ 抓客户要求页面停在订单页
（见「98001008 四层」那节），顺序反了就要多导航一次。单独跑 ⑥ 时也会先把页面
停回订单页再发请求。

## ③ 订单为什么改成「点后台的导出按钮」

接口那条路实测走不通，所以订单导出改走**页面上的「导出订单」按钮**：

- `order/list` 本地 requests 直连 **恒定** 裸 `code=10000`（aid 4068 / 6556 / 不带都试过）；
- 后台页面真实用的是同源 `GET /api/fulfillment/na/order/list`，带签名参数，
  去掉签名重放直接 `404 page not found`，没法在页面外调。

而订单页右上角有个导出图标按钮（`data-log_click_for="export_entry"`），点开面板可
选导出范围与格式，点「导出」后**异步生成**，完成后出现在「导出历史记录」里，
点那一行的「下载」即可下载一份完整订单 Excel —— **实测 63 列 × 616 行，
`Buyer Nickname` / `Recipient` / `Phone #` / `Country` / `State` / `City` /
`Address Line 1` 等客户字段都在里面**，一次导出拿全。

### 下载环节的两个坑（都已解决）

1. **ziniao 内核忽略 CDP 的下载目录**：`Browser.setDownloadBehavior` 用
   `allow` / `allowAndName` 都试过（DrissionPage 用的是 `allowAndName`），下载在 CDP 里
   显示 `completed`，但文件不落指定目录，全盘也搜不到 —— 内核把文件吞了。
2. **下载 URL 是一次性的**：`.../easesafe/oec_trade_order_export/object/<file_key>
   ?timeStamp=&sign=&use-disposable=1&attachment`，让浏览器真下载会消耗掉它，
   之后重取返回 `[CheckDisposableURL] out of use times`。

→ 解法：用 CDP **`Fetch` 域把这次请求拦在发送前**（服务器根本看不到，URL 未消耗），
再**在页面里 `fetch` 它**取回二进制（走店铺出口 IP + 登录态），自己决定存哪。
详见 `tiktok_export.py`，以及 `ziniao_bridge.BrowserSession` 的
`fetch_bytes()` / `intercept_url()` / `pump()` / `add_init_script()`。
每次重试都会重新点一次「下载」拿到新签名 URL，所以失败可安全重试。

## ⑤ 商品评分

页面 `<站>/product/rating?shop_region=US`（区域按站点自动推：us→US / uk→GB / eu→EU，
通用站退回 US）。和订单是**两条独立的路** —— 2026-10-15 连店铺浏览器抓包之后才看
清楚，之前按订单那套猜的是错的：

| | 订单 ③ | 商品评分 ⑤ |
|---|---|---|
| 入口 | 右上角导出图标 → 面板（可选范围 / 格式） | 页面里那颗「导出」按钮，**没有面板** |
| 点完之后 | 生成记录进「导出历史」→ 点「下载」**触发 easesafe 下载请求** | **没有任何浏览下载请求**：异步任务，下载地址在接口响应里 |
| 怎么拿到文件 | Fetch 域拦下一次性 URL → 页面内取字节 | 读 `export/job` 的 `file_url` → 页面内取字节（跨域，见下） |
| 文件格式 | xlsx | **CSV**（`rating,review_text,review_date,order_id,product_id,sku_id`）|
| 定位按钮 | `data-log_click_for="export_entry"` + 结构 XPath | 结构 XPath（`core-tabs-0-panel-0`…，**真机已失效**）+ **文案兜底**（含「导出 / Export / 导出报表」就点） |

真机抓到的完整链路（改 `RatingExporter` 前先看这段）：

```
点「导出」  POST /api/v1/review/biz_backend/export/start      → {"code":0}
页面轮询    GET  /api/v1/review/biz_backend/export/job        → {"status":2,"job_id":21}
生成完成    GET  .../export/job                               → {"status":3,
                "file_url":"seller-us.tiktok.com/easesafe/seller_review_export/
                            object/<id>safe?expire=…&skipCookie=true&sign=…"}
收尾        POST /api/v1/review/biz_backend/export/update      → {"code":0}
```

三个真机坑，都在代码注释里写死了：

1. **hook 关键词**：评分接口在 `/api/v1/review/` 下，多一层 `v1` —— 写成 `/api/review/`
   一条响应都收不到（`__tk_exp` 永远空）。
2. **`file_url` 没有协议头**（`seller-us.tiktok.com/…`）→ 统一补 `https:`。
3. **跨域**：`file_url` 在 `seller-us.tiktok.com`，从店铺页(`seller.us.…`) fetch 它是
   跨域，`cors` 模式一律 `TypeError: Failed to fetch`（`no-cors` 只能拿到 opaque，
   读不了 body）。解法：先把当前文档导航到**同域**的静态文件
   （`https://seller-us.tiktok.com/robots.txt`），文档 origin 变成它之后就是**同源**
   请求，CORS 不参与 → body 随便读，读完再导航回评分页。

下载回来的字节会先体检 —— **空 / HTML 登录页直接报错**，不会落一个坏文件。

### ⑤ 之后：只留 1~3 星，**直接覆盖保存**

导出完立刻筛掉 4~5 星，把 1~3 星**覆盖**写回 `<店铺名>_商品评分.xlsx`（真机导出的是
CSV，完整版按真实格式落盘成 `_评分原始.csv`，不会伪装成 xlsx）。没筛到差评时
**保留完整表**——宁可多存一份完整数据，也绝不往用户的文件里写一个空表。

## ⑥ 差评客户（1~3 星 → 客户信息）

**可以单独执行**（「抓取差评客户（1~3 星）」那颗按钮），不用重新导一次评分：
它读「评分 Excel」框里早已导好的那份文件。找不到评分表时会先弹窗提示，
而不是静默跑出个空表。

**⑥ 的依赖只有一样：评分表本身。** 评分文件自带 `order_id` 时**不需要订单表**
（真机导出的 CSV、以及⑤筛过保存的 xlsx 都有这一列）；只有老版「仅昵称」的评分
文件才需要订单表做关联。所以单独跑⑥的前提是「有⑤导出的评分表」，不是
「有③导出的订单表」—— 旧版这里无条件要求订单表，明明有评分表也被拦下，
弹窗还指着 `<店铺>_订单.xlsx`，看起来像⑥读错了文件（`_start_pipeline` 的守卫已修）。

真机的评分文件**自带 `order_id`**（2026-10-15 实测），所以链条很直：

```
评分表(1~3 星) → 文件里的 order_id → order/get → <店铺名>_客户差评信息.xlsx
```

- 输出列 = 客户表那 18 列 **+ 「评分」+「商品」**；
- 评分文件**没有 `order_id` 列**（老版本后台）时才退回老路：昵称(+商品) →
  订单表换出 Order ID（两级匹配：先「昵称+商品」精确锁定，锁不到再退回只用昵称）；
  这时候才需要订单表；
- ④ 抓过的订单**直接复用缓存**，⑥ 不重复发请求；
- 两种列都没认出来（既无订单号、又无昵称）→ 日志写明原因并跳过，**绝不写一张空表**；
  昵称脱敏成 `K*****` 对不上订单时照实报「N/M 条没匹配到订单」。

## ④ 客户信息

从订单 Excel 的 `Order ID` 列逐条调 `order/get`（该接口本地直连就通），写 `<店铺名>_客户.xlsx`（18 列）。

**「SKU ID」列（第 2 列）= 商品 SKU 的**平台 ID**，不是商家自定义的 Seller SKU。**

- **主源**：`order/get` 响应里的 `sku_module[].sku_id`（请求本身就带着，最准）；
- **兜底**：响应里没有 sku 节点、或是请求失败/中止的占位行 → 按 `Order ID` 从订单表
  的 `SKU ID` 列补（也认 `商品 SKU` / `Sku ID` 等别名，单独跑④时回读订单 Excel）；
- 一个订单多个 SKU → 合并成一格（` | ` 分隔、去重）；
- **注意别拿 `Seller SKU` 列顶上**：那是商家自己填的编码，和平台 SKU ID 是两回事。

**cookie 必须与订单同属一个店铺**（2026-10-13 实测）：一个店铺的登录态只能看它自己的订单。
汉维2-圣诞树 cookie + 汉维订单 → `code=0`；换成嘉恒业订单 → `code=98001008 需要权限`
（把 cookie 里的 seller_id 字段硬改也无效，服务端 `98001004` 拦下）。所以报「需要权限」
基本都不是 cookie 过期，而是 **cookie 拿错店**。

**店铺身份怎么取才算对**（2026-10-14 实测踩坑）：

| 取法 | 结果 |
| --- | --- |
| `Storage.getCookies` 全量 → 合并成 name→value | ❌ 会串味：统一商家环境下同一 profile 可能同时有多个区域店铺（`seller.us.*` / `seller.eu.*`）的同名 cookie，**后写覆盖前写**，可能拿到**另一个店铺**的 id |
| `Network.getCookies(urls=[店铺后台地址])` | ✅ 只取「浏览器真会发给这个店铺地址的那一份」，按域名/路径精确匹配 |
| 页面**自己发出的**请求里的 `seller_id`（`tiktok_export.page_seller_id()`） | ✅ ✅ 最可信 —— 订单就是它导出的 |

所以 ③ 现在按「店铺后台地址」精确取 cookie，并**等店铺身份稳定**再往下走；
导出后还会拿**页面自己用的 seller_id** 复核，不一致就以页面为准并大声告警；
④ 动手前先**预检前 3 条订单**，一旦是 `98001008/98001004` 这类店铺身份错误就**整批中止**
（其余标 `SKIPPED`），不再对着同一个错误重复请求几百次。

**④ 的「换站」机制（2026-10-14 真机日志后定稿）**：

店铺 cookie 是**按站隔离**的 —— `seller.tiktokshopglobalselling.com` / `seller.us.*` /
`seller.eu.*` 各带一份 `oec_seller_id_unified_seller_env`；但光换 cookie **没用** ——
同一账号在这些站里的 `seller_id` 常常**相同**，真正决定能不能看这批订单的是
**页面停在哪个区域站**。真机实测（订单表 619 条**全是美国订单**）：

| 页面站 + 同源路径 `/api/fulfillment/order/get` | 结果 |
|---|---|
| `seller.tiktokshopglobalselling.com` | `98001008` 需要权限 |
| `seller.eu.tiktokshopglobalselling.com` | HTTP 200 **"No matching route"** |
| `seller.us.tiktokshopglobalselling.com` | **`code=0`** ✅ |

所以 ④ 现在按 **「区域站 × 写法」** 逐个试，**命中一次就把「写法 + 该站」一起锁定**
（`winner_label` / `winner_origin`），之后每条订单只发 1 次请求、不再导航：

```
[浏览器内] POST …/seller.tiktokshopglobalselling.com/api/fulfillment/order/get → 98001008
[同源] 返回 code=98001008（这个写法不对，换一个试）
[绝对] 请求失败：页面 JS 报错：Cannot read properties of undefined (reading 'headers')
在 https://seller.tiktokshopglobalselling.com 上没打通，换页面来源 https://seller.eu…
[浏览器内] POST https://seller.us.tiktokshopglobalselling.com/api/fulfillment/order/get → {"code":0,…} ✅
✅ 锁定可用通道：同源 @ https://seller.us.tiktokshopglobalselling.com
```

两个曾经把它搞崩的实现细节（都已修）：

- **换站不能用 `ensure_on`**：它按「站点尾巴」比较，把 `seller.us.*` / `seller.eu.*` /
  `seller.*` 当成**同一站**，于是从不真正换站。现在用 `navigate` 并**按 host 严格复核**。
- **必须记「实际落地的站」**：区域站之间会互相重定向（打开 eu 地址，页面其实落在 us）。
  以前把「我们请求的地址」当成「页面所在的地址」记下来，之后每次都用错站发 → 全批失败。
  现在每次导航后都回读 `location`，以**实测值**为准。

主循环里还有一道保险：预检过了但首条就撞凭据错误时，会自动**重扫区域站并复检**，
复检通过就重试本条继续（只自救一次），失败才整批中止（其余标 `SKIPPED`）。

单独用 `获取客户/` 工具时才需要自己保证 cookie 对得上
（那里有「🔍 凭据预检」+ `sync_cookie.py`）。

## 运行

```bash
python customer_gui.py            # 打开界面
python customer_gui.py --check    # 离屏自检（构建界面，不联网）
python devtools/test_customer_sku.py   # 客户表「SKU ID」回归（30 条断言，零依赖）
python devtools/test_bad_review.py     # 差评链路回归（75 条断言：筛选 / order_id 直连 / 昵称反查兜底 / ⑥单独跑）
python devtools/test_shop_switch.py    # 后台店铺回归（43 条断言：名字→URL 解析 / 已是目标不动 / 落错站如实报，零依赖）
python devtools/test_rating_export.py  # 商品评分导出回归（52 条断言：异步任务 / CSV / 同源取字节）
QT_QPA_PLATFORM=offscreen python devtools/smoke_gui.py   # 冒烟测试（桩化紫鸟，14 段，零网络，约 15s）
python devtools/export_check.py 1  # 真机验证：用第 2 个店铺跑一遍完整导出
python devtools/diagnose_store.py --store 2   # 报 98001008 时：把「页面/cookie/订单」三者身份一次问清
```

## 打包成 exe

### 本地

```bash
pip install -r requirements.txt pyinstaller
pyinstaller build.spec --noconfirm     # → dist/tiktok商家.exe
python devtools/make_release.py --zip  # → release/ + tiktok商家.zip
```

⚠ PyInstaller **不能交叉编译**：在 macOS 上只能得到 Mach-O，要 Windows 的 `.exe`
就得在 Windows 上构建 —— 这正是用 GitHub Actions 的意义。

### CI

`.github/workflows/build.yml`（`:windows-latest`）：

```
检出 → 装依赖 → 源码自检 → PyInstaller → 组装发布目录 → 冻结产物自检 → 上传 artifact
                                                              └ 打 tag 时自动发 Release
```

触发：push 到 main/master 只出 artifact；push `v*` tag 会额外建 Release；也能手动跑。

### 发布布局（exe 旁挂紫鸟项目）

```
tiktok商家.exe
使用说明.txt
紫鸟浏览器/ziniao_webdriver_demo/
    util/{__init__.py, ziniao_client.py, dp.py}
    config/{__init__.py, config.yaml}      ← 首次运行由用户填账号
```

`ziniao_bridge.locate_ziniao_root()` 在冻结态下**优先**找 exe 旁边的
`紫鸟浏览器/ziniao_webdriver_demo`，所以这个布局开箱即用。

### 打包时必须注意的三件事

| 事项 | 说明 |
| --- | --- |
| **紫鸟客户端要旁挂** | `load_ziniao()` 用 `spec_from_file_location()` **按文件路径**加载 `util/ziniao_client.py`，它必须是磁盘上真实存在的 .py 源码 —— 打进 exe 会落在临时解包目录里，路径一变就失效。所以只在仓库里内联一份（`vendor/`），由 `make_release.py` 复制到 exe 旁边。 |
| **动态加载文件的依赖会全漏** | PyInstaller 静态分析看不见「运行时才按路径加载」的文件的 import 图，**连标准库都会漏**（实测栽在 `uuid` 上：`No module named 'uuid'`）。`build.spec` 因此**自动用 `ast` 扫 `vendor/.../util/*.py`**，把扫到的依赖全喂进 `hiddenimports`（含 `yaml`）。换新版客户端不用改 spec。 |
| **`--check` 要能写报告** | exe 是窗口程序（`console=False`），看不到 `print`。所以 `--check` 支持 `--check-out <文件>`，CI 靠这个文件断言打包有没有漏依赖 —— 这是唯一能在流水线里拦住上面那类问题的手段（`warn-build.txt` 不会提醒，因为静态图上那些模块压根不存在）。 |

### ⚠ 别把凭据推上仓库

`.gitignore` 已经挡住这些（**仓库是公开的，务必别 `git add -f`**）：

- `**/ziniao_webdriver_demo/config/config.yaml` —— 紫鸟企业账号**明文密码**
  （仓库里只保留 `config.example.yaml` 模板；`紫鸟浏览器/` 那份的 `.gitignore` 也补了）
- `cookies.json` —— 店铺登录态（`sid_tt` / `d_ticket` / `passport_csrf_token`）
- `*_订单.xlsx` / `*_客户.xlsx` / `_导出原始.xlsx` —— 客户姓名、地址、电话

### 报 `98001008` 时怎么定位（真机，1 次跑完）

```bash
python devtools/diagnose_store.py --store 2 --orders 订单.xlsx --sample 3
```

它会打印：当前页面落在哪个站 → 每个站的 cookie `seller_id` → **页面自己请求里用的 `seller_id`**
→ 然后用「页面身份 / cookie 身份 / 不带参数」三种姿势各打一次 `order/get`。按结果对号入座：

- **页面身份 `code=0`** → 订单属于页面这家店，抓客户就该用页面的值（主 GUI 已自动改用）。
- **页面报 `98001008` 但 cookie 的 id 不同** → cookie 串味，用 `获取客户/sync_cookie.py` 重新按站同步。
- **三种姿势全 `98001008`** → 这批订单不在当前登录账号名下的任何店里 —— 导出这批订单用的是
  **另一个账号/店铺**，必须换那个店铺的登录态（紫鸟里换店/换账号），改参数无解。

### 抓「后台自己发的请求」，照抄它的写法（2026-10-14 新增）

猜参数行不通 —— 页面到底怎么发的，只有页面自己知道（同源 host、`/na/` 这类区域前缀、
aid、签名字段、必要请求头）。一次性抓清：

```bash
# 接管第 2 个店铺 → 打开订单列表 → 收 30s 事件（期间可手动点开一个订单）
python devtools/capture_order_request.py --store 2 --wait 30

# 浏览器已在跑：不导航，直接抓当前页面
python devtools/capture_order_request.py --port 16896 --no-nav --wait 45

# 抓完顺手用订单表第 1 条，在页面内重放若干候选 order/get，看哪个 code=0
python devtools/capture_order_request.py --store 2 --probe-get 订单.xlsx
```

它输出四段：**[1]** CDP `Network` 抓到的真实请求（完整 URL/query/headers/body ＋
`initiator` 调用栈）→ **[2]** 页面内 fetch/XHR hook 记录（URL/body/调用栈）→
**[3]** 顺调用栈去那个 JS 文件里搜接口与参数构造片段（**这就是「发送请求位置」**）→
**[4]** 候选 `order/get` 重放结果。抓到的模板写到 `devtools/captured_order_request.json`。

**抓到模板后，主 GUI 会自动照抄它**：④ 的浏览器内请求会依次尝试
**「模板 → 同源 → 绝对 api16-*」**三种写法，并且把「写法 + 当时页面实际所在的区域站」
**一起锁定**，后续只发这一种（不会每条订单都白试一轮）。
所以典型流程是：先跑一次 `capture_order_request.py` → 再正常跑 `customer_gui.py`。

> 为什么是三种写法：后台真实接口是**同源** `seller.<region>.tiktokshopglobalselling.com/api/...`
> （带区域前缀，如北美的 `/api/fulfillment/na/...`），而抓包留档里的 `api16-normal-no1a...`
> 是另一个域 —— 直接拿它去页面内 fetch，cookie/Origin/Referer 都对不上，正是 `98001008` 的常见来源。
> 即使没抓模板，现在的「区域站 × 写法」自动扫描也够用（真机已通）。

依赖见 `requirements.txt`（另需同工作区的 `紫鸟浏览器/ziniao_webdriver_demo`，会自动定位）。

## 使用顺序

1. **启动环境**：定位紫鸟客户端程序（Windows 上自动探测安装位置，见「排错对照」）→
   读紫鸟 `config/config.yaml`（企业账号 / 端口 / `--user-data-dir`）→
   终止旧进程 → 启动客户端（web_driver + http IPC）→ 更新内核
2. **获取店铺列表**：表格列按真实返回字段自动推导，支持关键词过滤，双击看原始 JSON
   - **账号下挂了多个后台店铺**时（如「秋图2」= Spain + United States），按「后台店铺」
     框里填的店铺**直接导航**到它的后台首页，**不模拟点击**右上角切换器：
     `https://seller.<区域>.tiktokshopglobalselling.com/homepage?shop_region=<代码>`，
     其中 `shop_region=US` 就代表 United States（用户 2026-10-15 定案）。
     已经在这家店就不动页面；导航后等落地 URL 的 `shop_region` 变成目标才往下走
     （区域站之间会重定向，真机见过「开 eu 落在 us」，所以按**实际落地 URL** 继续）。
     切错店铺的后果是 cookie 里 seller_id 是另一家的 → 后面整批 98001008，
     所以日志里会写清导航到了哪个 URL、页面右上角显示哪家店
   - 「后台店铺」框可以填：显示名（`United States`，不分大小写）/ 中文名（`美国`）/
     区域码（`US`）/ 完整后台 URL。认不出的名字**只告警、不乱猜区域码、不乱跑**，
     继续用当前页面（后面身份核对还有兜底）
   - 为什么不用点击：真机实测点开菜单要等渲染、点中后要轮询等胶囊重渲染（6~14s），
     菜单还会浮在右上角挡住导出按钮；URL 一步到位，落错站一眼就能从 URL 看出来。
     右上角切换器只保留**只读**读取（把店铺名写进日志，读不到不影响流程）
3. **选中店铺 → 导出订单**：打开订单页 → 点「导出」按钮 → 等生成 → 拦截 URL 取回文件，
   写 `<店铺名>_订单.xlsx`（列名 = 导出文件表头，全量保留），原始文件另存 `_导出原始.xlsx`
4. **抓客户信息**：从订单 Excel 读 `Order ID` 列，逐条 `order/get`，写 `<店铺名>_客户.xlsx`（18 列，含取自响应 `sku_module` 的「SKU ID」）
5. **导出商品评分**：⑤ 区「导出商品评分」按钮，走 `/product/rating` 页自己的「导出」
   按钮；导出后**只留 1~3 星**并覆盖保存（完整版留 `_评分原始.xlsx`）。
   **只导评分、先不抓客户**就点这一颗
6. **抓差评客户**：⑤ 区「抓取差评客户（1~3 星）」按钮，**可单独点**——读已有的评分
   表，取 1~3 星那批的**订单号**（真机评分文件自带 `order_id`），逐条 `order/get`，
   写 `<店铺名>_客户差评信息.xlsx`（客户 18 列 + 「评分」「商品」）。
   只有老版「仅昵称」的评分文件才需要订单表来换订单号

> 输出文件名默认带店铺名（选中店铺后自动填入），一个店一份，不同店铺互不覆盖。
> 想换存放位置就点框旁边的「浏览…」——**只换目录，文件名不变**；选过一次之后，
> 其它几个框再点「浏览…」会直接落在同一个目录（记着上次选的）。
> 想连文件名一起改，直接在框里敲完整路径（手填过的不会被自动改名覆盖）。

「一键完成」会把 ③④⑤⑥ 连起来跑；中途可「停止」，已抓部分照常落盘（每 10 条写一次）。
⑥ 正常情况下**不需要订单表**（评分文件自带 `order_id`）；单独点它时只要「评分 Excel」
指向该店的评分文件即可。只有老版「仅昵称」的评分文件，才要求④的「订单 Excel」
指向该店的订单文件来做关联。

## 关键参数

③ 区只保留「订单 Excel」一个文件框，其余参数都已固定为默认值、不再出现在界面上。

| 参数 | 说明 |
| --- | --- |
| 订单 Excel | ③ 唯一的框：导出文件名，默认 `<店铺名>_订单.xlsx`（选中店铺后自动填入）；「浏览…」只换存放目录 |
| 导出范围 | **固定**「筛选出的订单」（= 页面当前筛选结果） |
| 导出格式 | **固定** Excel |
| 传输通道 | **固定**自动（浏览器内优先，失败转本地直连） |
| 接口主机 | **固定** `api16-normal-no1a.tiktokshopglobalselling.com`（本账号实测唯一可用） |
| 页面来源 | **固定**留空自动（后台首页 → 内置候选 → 接口 origin） |
| seller_id | **固定**从店铺 cookie `oec_seller_id_unified_seller_env` 自动取 |
| 请求间隔 | ④ 抓客户每条之间的等待，**最小 1.5s**（默认 1.5s，防限流） |
| 跳过 IP 检测 | 默认会等紫鸟的 ip 检测页通过，确认代理生效（失败不中断） |
| 客户端路径 | ① 区「编辑配置…」里：紫鸟**客户端程序**的位置（不是数据目录）。**留空即可**，会自动探测。Windows 上填 exe 所在**目录**（如 `D:\紫鸟\ziniao`）或全路径都行；macOS 填 `ziniao` |
| 数据目录 | ① 区「编辑配置…」里：web_driver 实例的数据**文件夹**。留空自动推导。误填成 exe 会被拦下并提示搬走 |

## 排错对照

| 现象 | 原因 | 处理 |
| --- | --- | --- |
| `PageDisconnectedError` / `connection to the page has been disconnected` | DrissionPage 在该内核上只要页面一导航，CDP WebSocket 就被服务端 reset 且不会自愈 | **已换成原生 CDP**（`websocket-client`）：导航打断后自动重连重放。日志里看到「CDP 连接中断…重连后重试 Page.navigate」是**正常现象** |
| 日志「新建标签页：xxxxxxxx」 | 从紫鸟的 `chrome-extension://` 检测页导航会把 target 弄丢 | 已内置兜底：没有可用 page 时用 `/json/new` 新建一个 |
| 「导出面板没打开（按钮结构可能变了）」 | 后台改版，导出入口按钮的文案/结构变了 | 调整 `tiktok_export.EXPORT_ENTRY_XPATH` 或 `data-log_click_for` |
| 「等待导出生成超时」 | 订单量大时后端生成慢 | 稍后重试；可在 opts 里调大 `export_timeout` |
| 「没拦到下载请求」 | 导出还在生成中（那一行没有「下载」按钮），或文件名不匹配 | 会自动重试 3 次；仍失败把日志发我 |
| 「页面内 fetch 失败：TypeError: Failed to fetch」 | 拦截 abort 后首次重取偶发失败 | **会自动重试**（重新点「下载」拿新签名 URL），实测 1 次重试即成功 |
| 订单 0 条 | 导出范围固定为「筛选出的订单」，而当前页面的筛选结果是空的 | 先在订单后台清掉筛选（让它显示全部订单）再导出 |
| `[code=98001004]` | `seller_id` 与当前店铺不匹配（典型：手动填了别的店铺的 id） | seller_id 现已固定为「从 cookie 自动取」，若仍报错说明③打开的店铺不是订单所属的那家 |
| 客户信息全 `API_ERROR` | cookie 失效 | 重跑；或把 ④ 的传输通道切到「本地直连」（`order/get` 实测可通） |
| `[code=98001008] 需要权限` | 三种可能：① **cookie 不是导出这张订单表的那个店铺的**（不是过期）—— 实测「汉维2-圣诞树 cookie + 嘉恒业订单」必报此码；② **页面没停在正确的区域站**（同一个店铺，`seller.*` 通用站看不到 `seller.us.*` 的订单）；③ **页面没停在订单页**（站对、页错同样报这个码，2026-10-14 真机定案） | ①换成本店铺的 cookie：`获取客户/sync_cookie.py --store N` 从紫鸟当前店铺同步，或在上级 GUI 里跑（自动取当前店铺 cookie）；②③正常跑主 GUI 即可 —— ④ 会先导航到**该站的订单页**，再按「区域站 × 写法」逐个试并锁定能通的那个 |
| 同一份 exe：Mac 上能抓客户，拷到 Windows 上全批 `98001008` | **页面停的页面/站不一样**，不是「用了本地 cookie」—— 抓客户全程**不读任何本地 cookie 文件**，身份来自紫鸟店铺浏览器里那份实时登录态（CDP `Network.getCookies`，按店铺后台域名精确取）。真机定案就是「必须导航到订单页」 | 已修：④ 开始前 `_ensure_order_page()` 显式停回订单页；换站时也直接开「该站的订单页」。日志里有 `页面已在订单页：…` / `页面已停在：…` 可直接确认。出问题请「工具 → 打开日志文件」把整份发来 |
| `98001008` 时想「改 `params` 里的 `seller_id` 修好」 | **改不动**（2026-10-14 实测）：同一份 cookie 下把 `seller_id`/`oec_seller_id` 换成别的店 → `98001004`；把这两个参数**整段删掉** → 照样 `98001008`；改 `aid` / 加 `X-Tts-Oec-Bsid` / 换 `seller.us` 域名 host → 全是 `98001008`。服务端是**按会话（cookie）判定**你能看哪家店，参数不是权限开关 | 别在参数上试了，先跑 `python devtools/diagnose_store.py` 把三者身份打出来（见下） |
| 抓客户全批 `98001008`，但订单确实是这个店导出的 | **店铺身份取错**：`Storage.getCookies` 把多个区域店铺的同名 cookie 混在一起，合并时覆盖成了另一个店（实测：cookie 读到「13 单」那家的 id，而页面实际开着「619 单」那家） | 已修：③ 改按店铺后台地址精确取 cookie（`Network.getCookies`）+ 等身份稳定，导出后用**页面自己请求里的 seller_id** 复核，不一致以页面为准。升级后重跑即可；日志会打印 `⚠ 店铺身份不一致` 与页面实际用的 id |
| `[code=98001004] 部分信息填写错误` | 参数里的 `seller_id` 与 cookie 里的店铺对不上 | 清空界面上的 seller_id；注意硬改 cookie 里的 seller_id 字段也无效（服务端会拦），必须换真的那份 cookie |
| 换了正确店铺仍全批 `98001008`，页面内请求也一起被拒 | **页面没停在正确的「区域站」上**：同一个接口路径在 `seller.*`（通用站）回 `98001008`、`seller.eu.*` 回 `No matching route`、**只有 `seller.us.*` 通**（2026-10-14 真机日志）。其次才是 host / 路径前缀（`/na/` 之类） | 已自动兜底：④ 按 **「区域站 × 写法（模板/同源/绝对）」** 逐个试，命中即**锁定「写法 + 该站」组合**，之后每条只发 1 次。要做实可先跑 `python devtools/capture_order_request.py --store N --wait 30` 抓一次真实请求，再正常跑 `customer_gui.py` |
| 日志里明明有一次 `code=0`，紧接着又全是 `98001008` / `No matching route` | 曾经的实现缺陷（已修）：只记住「写法」没记住「站」，于是下一条跑到 eu 站还用同源路径；换站用的是 `ensure_on`（按站点尾巴比较，各区域站被视为同站 → **从不换站**）；且记录的是**请求地址**而非**页面实际落地的址**（区域站会互相重定向） | 已修：`_ensure_origin` 改用 `navigate` + 按 host 严格复核；每次导航后回读 `location` 取**实测站**；`winner_label` + `winner_origin` 一起锁定 |
| 启动环境报「找不到紫鸟客户端可执行文件（client_path='ziniao'）」 | Windows 上紫鸟的**安装目录是安装时自选的**，没有固定位置（真机见过 `D:\紫鸟\ziniao`、`C:\Program Files\ZiniaoBrowser` 等）；官方 demo 示例里的 `D:\ziniao\ziniao.exe` 只是其中一种，照抄多半不存在 | **留空即可**，新版会自动找（环境变量 `ZINIAO_EXE_PATH` → 运行中的进程 → 注册表 → 常见安装目录 → 盘根浅扫）。也可手动填 exe 所在**目录**（如 `D:\紫鸟\ziniao`）或全路径，或点配置里的「自动检测」。① 区「客户端」一行会直接显示解析结果，**红字就是没找到**。填错路径不会直接失败，会自动回落到探测。⚠ 如果报错里出现 `client_path='ziniao'`，说明你跑的是**旧版 exe**（旧版没有自动探测、Windows 上恒定失败）—— 换新版即可 |
| 从快捷方式属性抄下来的路径填进去还是不认 | Windows 资源管理器**默认隐藏已知扩展名**：属性里显示的目标是 `D:\紫鸟\ziniao`，其实那个文件叫 `D:\紫鸟\ziniao.exe`；照抄后 `isfile`/`isdir` 都是假，又不在常见位置表里（例如装在 `E:\soft\ziniao.exe`）→ 自动探测也找不到 | 已修：解析时会自动补一次 `.exe`（`D:\紫鸟\ziniao` → 认到 `D:\紫鸟\ziniao.exe`）。最保险还是抄到 `.exe` 那一位 |
| 把客户端 exe 填到了「数据目录」 | ① 区配置里「客户端路径」和「数据目录」是挨着的两个框，容易填反 | 保存时会弹窗拦住并**问你搬不搬到「客户端路径」**；就算硬填进去，运行时也会忽略它、回退到默认数据目录（不会崩）。**报错文案里也会点出这个坑** |
| 「紫鸟启动后没有界面」 | **这是自动化模式的既定行为，不是 bug**：`--run_type=web_driver` 的官方语义就是「让超级浏览器以**无界面状态**运行主进程」，去掉它就没有 HTTP IPC 端口、整条链路都跑不起来 | 不用管它，看日志即可（界面日志 + `运行日志.txt`）。顺带：「强制重启客户端」默认**不勾选**，客户端已在运行就直接复用，不会把你打开的实例关掉 |
| ⑤ 报「没找到商品评分页的『导出』按钮」 | 评分页改版，XPath 失效（那个老 XPath 在 2026-10-15 真机上**已经命中不了**） | 代码里还有**文案兜底**（按钮文字含「导出 / Export / 导出报表」就点），一般能自愈。真换了地方就改 `tiktok_export.RATING_EXPORT_XPATHS`，或把新 XPath 加进去（是个元组，可写多个，按顺序试） |
| ⑤ 报「没认出评分列」 | 后台改了列名（现在认 `Rating` / `Rating Score` / `评分` / `星级` 等一串别名） | **不会把文件写成空的**：认不出就保留完整表。把文件里的真实列名发我，加进 `tiktok_export.RATING_SCORE_COLUMNS`；日志里会打印评分文件的实际列名 |
| ⑤ 日志停在「点击后 20s 内没有浏览器下载请求」 | **正常**：评分导出是异步任务，本来就不发下载请求（2026-10-15 真机抓包确认） | 接着看下一行「读 export/job 的 file_url」。若一直 `status=2` 说明服务端还在生成，加大「导出超时」（默认 300s）即可 |
| ⑤ 报「页面内 fetch 失败：TypeError: Failed to fetch」 | `file_url` 在**另一个域名**（`seller-us.tiktok.com`），从店铺页取它是跨域 → CORS（`no-cors` 也读不了 body） | 已自动处理：先把文档导航到同域的静态文件当壳子再取（日志会写「文档同源就位」）。若报「没法把页面换到 … 下」＝那个域下的静态文件都打不开，用 `python devtools/live_rating_probe.py --api` 看真实响应 |
| ⑥ 日志说「评分文件自带订单号列…」 | **正常且最好**：真机 CSV 自带 `order_id`，⑥ 直接用它，**不需要订单表** | 这行只是告诉你走的是最省事那条路；没有它才会退回昵称反查（需要订单表） |
| 某一单日志停住不动（界面转圈但 CPU 不高），几十秒后报「等响应超过 60s」 | 服务端把那条请求**挂住了**（连续请求触发风控时会出现）；老版本还会在这里**永久死等** —— `awaitPromise` 死等 + CDP 的 `recv` 被源源不断的事件消息无限续命 | 2026-10-15 已修：`call()` 的 `timeout` 是**总时限**（到点抛 `ZiniaoTimeout`），单次请求 60s 上限；API 层拿到超时**不再原层重试**（否则只是把 60s 变 3×60s）。这一单会被记成失败、继续下一条，整批不会卡死。对策：等一会儿再重跑，或降低抓取速度（分批跑） |
| ⑥ 报「既没订单号列、也没认出买家昵称列」 | 评分文件列名变了 | 把真实列名加进 `tiktok_export.RATING_ORDER_COLUMNS`（首选）或 `RATING_BUYER_COLUMNS`；日志里会打印前 10 个列名 |
| 日志报「认不出后台店铺 X」 | 「后台店铺」框里填的名字不在内置市场表里 | 填显示名（`United States`）、区域码（`US`）或直接贴后台 URL（`…?shop_region=US`）。认不出时**不会**乱导航，继续用当前页面 |
| 日志报「导航到 X 后落在 <URL>（想要 shop_region=XX）」 | 导航过去了但页面把请求重定向到了**别的区域站**（真机见过「开 eu 落在 us」），或店铺不在这家账号下 | 看日志里回填的**实际落地 URL**：如果那家店才是你要的，就把「后台店铺」填成它；这条只影响本次，后面身份核对与区域站扫描还有兜底 |
| 日志报「导航到 X 失败」 | 网络/代理不通，或店铺浏览器没登录 | 先手动打开后台确认能进，再重跑。导航失败不会中断流水线，「强制重启客户端」可用来清理卡住的实例 |
| 单独点「抓取差评客户」弹「找不到评分表」 | 还没导出过评分，或「评分 Excel」框里的路径不对 | 先点「导出商品评分」，或把框里改成已有文件（也认同目录下的 `_评分原始.xlsx`） |
| 单独点「抓取差评客户」弹「找不到订单表」 | 评分文件里**没有订单号列**（老版「仅昵称」文件），只能靠昵称回订单表换 —— 而订单表不在了 | 先跑一次「导出订单」，或把「订单 Excel」指向已有文件。**若你的评分是⑤新导的（带 `order_id`），根本不该看到这个弹窗**（旧版守卫会误拦，已修） |
| ⑥ 报「没认出买家昵称列」 / 「没有一条差评能对应到订单」 | 评分文件里的昵称列改名，或昵称被平台脱敏成 `K*****`，与订单表的明文 `Kristi` 对不上 | 前者：把列名加进 `te.RATING_BUYER_COLUMNS`；后者**无解**（脱敏不可逆），日志会照实报「N/M 条没匹配到订单」 |
| ⑥ 说「订单表里没有可用的买家昵称列」 | 走的是昵称反查那条老路，而③ 没跑过、或「订单 Excel」里的文件没有 `Buyer Nickname` 列 | 先跑一次③导出订单，再跑⑤+⑥；或直接用带 `order_id` 的评分文件（那样根本不查订单表） |

## 目录结构

```
customer_gui.py         GUI 主程序（六步流水线：③④导出订单/抓客户 + ⑤商品评分 + ⑥差评客户）
tiktok_export.py        ★「导出订单」「商品评分」按钮自动化：
                        OrderExporter  = 面板操作 / 等生成 / 拦 URL / 取字节 / 解析
                        RatingExporter = 评分页那颗导出按钮（★真机是异步任务：读 export/job 的 file_url，
                        跨域 → 借同域 robots.txt 当壳子取字节；CSV 解析）
ziniao_bridge.py        紫鸟桥接：BrowserSession（原生 CDP + 断线自愈 + Fetch 拦截 + 页面内取字节）
                        ★ `call()` 的 timeout 是**总时限**（事件消息不能无限续命，到点抛 ZiniaoTimeout）
                        ★ BrowserTransport 单次页面请求 60s 上限（wait_timeout）
tiktok_client.py        ④ 的 order/get 客户端：参数 / 解析 / Excel 读写
detail.py / list.py     原始抓包脚本（仅留档，接口直连思路已废弃）
devtools/export_check.py 真机端到端验证脚本
devtools/diagnose_store.py 98001008 诊断：页面 / cookie / 订单 三者身份对照
devtools/capture_order_request.py 抓后台真实订单请求 + 定位发送位置（并产出模板）
devtools/captured_order_request.json 抓包模板（④ 会自动照抄它）
devtools/live_rating_probe.py 真机探测：评分页点了「导出」之后浏览器到底发了什么
                        （默认只读；`--api` 留接口响应体；`--app` 端到端跑一遍导出）
devtools/live_bad_probe.py 真机验证：评分 CSV 的 order_id → order/get → 客户
devtools/check_bad_guard.py ⑥ 前置守卫的一次性验证（用**真实文件**跑：只有评分表、
                        没有订单表时不弹窗；老版「仅昵称」文件才拦。不联网）
devtools/smoke_gui.py   离屏冒烟测试（14 段：构建/布局/「浏览…」换目录/流水线/换来源/换写法/商品评分/⑥单独跑/店铺导航/评分真机形状…）
                        ★ 必须桩掉紫鸟 IP 检测：假会话里那个「检测成功」按钮永远不出现，
                          真跑 check_ip 会**每段白等 30s 超时**（14 段 ≈ 十几分钟，看着像卡死）
devtools/test_rating_export.py 商品评分导出回归测试（假 CDP 会话，52 条断言）
devtools/test_transport.py 「页面来源 × 写法」择路 + 超时处置回归测试
                        （51 条断言；第 10 组「超时不在原层重试」，第 11 组「事件不续命」）
devtools/test_customer_sku.py 客户表「SKU ID」列回归测试（30 条断言：响应用 sku_id / 不拿 Seller SKU 顶 / 订单表兜底）
devtools/test_bad_review.py 差评链路回归测试（75 条断言：1~3 星筛选 / 别名列名 / order_id 直连 / 昵称反查兜底 / ⑥单独跑）
devtools/test_shop_switch.py 后台店铺回归测试（43 条断言：显示名/区域码/URL→后台 URL 解析、已是目标幂等、落错站回填实际 URL、导航失败不中断）
devtools/live_shop_probe.py  真机探测脚本（连已开的店铺浏览器：现在在哪家店 + 目标会导航到哪个 URL；默认只读，--go 才真跳）
devtools/test_windows_paths.py Windows 客户端定位回归测试（假文件系统，任何平台可跑）
devtools/make_release.py 组装发布目录（exe + 旁挂紫鸟项目），可 --zip
build.spec              PyInstaller 配置（自动扫描紫鸟客户端的依赖）
.github/workflows/build.yml  CI：windows-latest 上构建 exe + 冻结产物自检
vendor/ziniao_webdriver_demo/ 内联的紫鸟协议客户端（只保留真正用到的 2 个 .py）
    util/ziniao_client.py   ★ 唯一真正干活的（972 行，含 Windows 客户端定位）
    util/dp.py              被加载但零方法调用（DrissionPage 兼容层，留作后备）
    config/config.example.yaml  账号模板（真配置 config.yaml 已被 .gitignore 挡住）
docs/screenshots/       界面截图
```

## 实测结论（2026-10-13）

- **接管方式**：DrissionPage 不行 —— 它接管 ziniao 店铺浏览器后，页面一导航其 CDP
  WebSocket 就被 reset，且**不会自愈**，之后所有页面操作恒定 `PageDisconnectedError`。
  换**原生 `websocket-client`** 直连 `webSocketDebuggerUrl` 则完全正常。
- **订单只能靠页面导出按钮**（见上），导出的 Excel 含全部客户字段，616 行实测 48s 拿到。
- **导出记录接口** `GET /api/fulfillment/na/order/export_record/get` 返回
  `{file_name, download, file_key, export_time, export_task_id}`，是判断「导出好了没」的依据；
  但它**只能在页面里被调用**（我们通过注入 hook 读它的响应，重放会 404）。
- **seller_id 必须来自该店铺自己的 cookie**（`oec_seller_id_unified_seller_env`）。
  实测「汉维2-圣诞树」= `7494876782645708274`，「嘉恒业」= `7496000161034505075`。
- `order/get` 本地直连 → `code=0`（`aid=4068` + seller_id 即可，不需要签名）。
- `api16-normal-no1a.tiktokshopglobalselling.com` 是本账号唯一可解析的接口宿主。

## 实测结论（2026-10-14，`98001008` 收口）

一次完整真机日志把「页面内请求也被拒」的**最后一层**定住了 —— 答案是**区域站**：

| 页面所在站 + 同源 `/api/fulfillment/order/get` | 结果 |
|---|---|
| `seller.tiktokshopglobalselling.com`（通用站） | `98001008` 需要权限 |
| `seller.eu.tiktokshopglobalselling.com` | HTTP 200 `"No matching route"` |
| `seller.us.tiktokshopglobalselling.com` | **`code=0`** ✅（返回完整 `main_order`，含 buyer/收件人/地址） |

同一份 cookie、同一个 `seller_id`、同样的参数，**只因为页面停的站不同，结果完全不同**。
两个把这件事搞砸的实现细节：① `ensure_on` 按「站点尾巴」比较，把各区域站视为同站 → **从不换站**；
② 成功后记的是「我们请求的地址」而非「页面实际落地的址」（区域站之间会互相重定向，打开 eu 会落在 us），
于是后续每条都用错站发。现在：`navigate` + host 严格复核 + `location` 实测 + 「写法 × 站」一起锁定。

顺带确认（对照实验）：`seller_id` / `oec_seller_id` 换成别的店 → `98001004`；**整段删掉** → 照样
`98001008`；改 `aid`、加 `X-Tts-Oec-Bsid`、换 host → 都不改变结论。**参数不是权限开关**，
服务端按会话 + 页面站判定。

## 实测结论（2026-10-14 追加，`98001008` 的第二层：**页面必须停在订单页**）

只有**站对**还不够 —— 用户真机复现：同一个 exe，Mac 上能抓，Windows 上整批 `98001008`，
两边唯一稳定的差别是**页面当时停在哪个页面**。结论是：

> 同源 `fetch` 要求页面**本身就停在订单页**（`<站>/order?selected_sort=6&tab=all`）上。
> 站选对了、但页面停在首页 / 工作台 → 一样 `98001008 需要权限`。

所以现在是**两层**约束，缺一不可：

| | 站对 | 页对 | 结果 |
|---|---|---|---|
| ① 换区域站 | ✅ `seller.us.*` | — | 上一节解决的问题 |
| ② 换页面 | ✅ | ✅ `<站>/order?…` | **本节解决的问题** |

落地方式（三处，互相兜底）：

1. `BrowserTransport(page_path="/order?…")` —— 换站时**直接打开「该站的订单页」**，
   而不是只到站点根；已经在目标站但不在订单页时补一次导航。
2. `PipelineThread._ensure_order_page()` —— ④ 抓客户**之前**显式把页面停回订单页
   （用 ③ 导出成功时那个站优先）。已经在订单页就不导航、不等，避免多余的掉线风险。
3. `TransportChain` 的 fallthrough 码加上 `98001008` 这类权限码 —— 浏览器内那条路
   真拿不到数据时，让**本地直连**（`order/get` 实测可用）有机会兜住客户信息。

回归防线：`devtools/test_transport.py` 第 9 组断言（CI 每次构建都跑）。

## 实测结论（2026-10-14，Windows 客户端定位）

Windows 上紫鸟的安装目录是**安装时用户自选**的，没有固定位置。官方 demo 示例里写的
`D:\ziniao\ziniao.exe` 换台机器多半不存在 —— 照抄就报「找不到紫鸟客户端可执行文件」。
真机见到的位置就有 `D:\紫鸟\ziniao\ziniao.exe`、`C:\Program Files\ZiniaoBrowser\...` 等。

| 你填的 `client_path` | 行为 |
|---|---|
| **留空**（推荐） | 自动探测：环境变量 `ZINIAO_EXE_PATH` → 运行中的进程 → 注册表（卸载信息 / App Paths）→ 常见安装目录 → 各盘根浅扫 `*ziniao*` / `*紫鸟*` |
| `D:\紫鸟\ziniao`（目录） | 在目录里找 `ziniao.exe`；填再上一层（`D:\紫鸟`）也会往下钻一层 |
| `D:\紫鸟\ziniao.exe` | 直接用 |
| `"D:\...\ziniao.exe"` / `d:/紫鸟/...` / 末尾带 `\` | 引号、正斜杠、大小写、尾部分隔符都容忍 |
| `D:\紫鸟\ziniao`（**隐藏扩展名**，真身是 `ziniao.exe`） | 自动补一次 `.exe` 认到它（见下） |
| 路径写错 / 换台机器失效 | **不报错**，静默回落到自动探测 |

三个反直觉的点，都是被真机教出来的：

- **刻意不查 PATH**。紫鸟客户端从不往 PATH 里注册，而 `which` 会把任何同名程序当成客户端 ——
  开发机上就撞到过 `~/.local/bin/ziniao` 这种无关 shim，直接选错文件。
- **探测结果会缓存，所以时机很关键**。启动流程会先杀掉客户端进程，那时按「运行中的进程」
  就找不到了。因此「启动环境」是在 kill **之前**强制重探一次（进程还在，路径最准），
  之后复用该结果。
- **「省略 .exe」必须单独兜一层**。Windows 资源管理器默认隐藏已知扩展名，用户从快捷方式
  属性里抄到的是 `D:\紫鸟\ziniao`（看着像目录，实际是 `ziniao.exe`）。只做
  `isfile` + `isdir` 两步的话，这种写法两个都是假 → 被当成「路径写错」丢给自动探测；
  而这个目录又不在常见位置表里（比如装在 `E:\soft\ziniao.exe`）→ 自动探测也找不到，
  最后报「找不到客户端」。所以解析时会额外试一次 `raw + ".exe"`。

**回归测试**：`python devtools/test_windows_paths.py` —— 在假文件系统上把上面每种写法
（含 `C:\Program Files\...`、中文 exe 名、目录里混着卸载器、E 盘不在位置表里…）跑一遍，
任何平台都能跑，CI 每次构建都会执行。

**CI 哨兵**：冻结产物自检报告里固定带一行 `客户端程序 : ...`。这一行会真的跑完 Windows
的自动探测链路（环境变量 → 进程 → `winreg` 注册表 → 目录 → 盘根扫描）—— runner 上没装
紫鸟，结果必然是「未命中」，但**只要这行出现了，就说明整条链路在冻结产物里没崩**；
它依赖的 `winreg` / `ntpath` / `subprocess` 恰好是最容易漏打进包的东西（静态分析看不见
动态加载文件的 import）。缺了这一行，CI 直接判失败。

## 已知限制

- 导出文件里的手机号 / 姓名 / 地址**带平台脱敏**（`*****`），与 `order/get` 返回的一致。
- 店铺未登录该平台时订单页会跳登录，导出按钮不可用。
- 关店时只断开 CDP，不关浏览器进程；浏览器由紫鸟 `stopBrowser` 负责。
- 订单表格最多渲染前 800 行（`MAX_TABLE_ROWS`），完整数据看导出的 Excel。
- ⑥ 依赖买家昵称做关联：昵称在评分文件里被脱敏成 `K*****` 时匹配不上，
  那条差评就不会进客户差评表（会照实在日志里报条数）。
