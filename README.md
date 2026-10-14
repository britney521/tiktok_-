# TikTok 订单 / 客户信息导出 GUI（紫鸟店铺浏览器版）

四步流水线：**① 启动紫鸟环境 → ② 获取店铺列表 → ③ 选店铺导出订单 → ④ 读订单 Excel 抓客户信息**。

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

## ④ 客户信息

从订单 Excel 的 `Order ID` 列逐条调 `order/get`（该接口本地直连就通），写 `<店铺名>_客户.xlsx`（17 列）。

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
QT_QPA_PLATFORM=offscreen python devtools/smoke_gui.py   # 冒烟测试（桩化紫鸟，零网络）
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

1. **启动环境**：读紫鸟 `config/config.yaml`（企业账号 / 端口 / `--user-data-dir`），
   终止旧进程 → 启动客户端（web_driver + http IPC）→ 更新内核
2. **获取店铺列表**：表格列按真实返回字段自动推导，支持关键词过滤，双击看原始 JSON
3. **选中店铺 → 导出订单**：打开订单页 → 点「导出」按钮 → 等生成 → 拦截 URL 取回文件，
   写 `<店铺名>_订单.xlsx`（列名 = 导出文件表头，全量保留），原始文件另存 `_导出原始.xlsx`
4. **抓客户信息**：从订单 Excel 读 `Order ID` 列，逐条 `order/get`，写 `<店铺名>_客户.xlsx`（17 列）

> 输出文件名默认带店铺名（选中店铺后自动填入），一个店一份，不同店铺互不覆盖。
> 想用别的名字就点「另存为…」——手填过的路径不会被自动改名覆盖。

「一键完成」会把 ③④ 连起来跑；中途可「停止」，已抓部分照常落盘（每 10 条写一次）。

## 关键参数

③ 区只保留「订单 Excel」一个文件框，其余参数都已固定为默认值、不再出现在界面上。

| 参数 | 说明 |
| --- | --- |
| 订单 Excel | ③ 唯一的框：导出文件名，默认 `<店铺名>_订单.xlsx`（选中店铺后自动填入），可「另存为…」 |
| 导出范围 | **固定**「筛选出的订单」（= 页面当前筛选结果） |
| 导出格式 | **固定** Excel |
| 传输通道 | **固定**自动（浏览器内优先，失败转本地直连） |
| 接口主机 | **固定** `api16-normal-no1a.tiktokshopglobalselling.com`（本账号实测唯一可用） |
| 页面来源 | **固定**留空自动（后台首页 → 内置候选 → 接口 origin） |
| seller_id | **固定**从店铺 cookie `oec_seller_id_unified_seller_env` 自动取 |
| 请求间隔 | ④ 抓客户每条之间的等待，**最小 1.5s**（默认 1.5s，防限流） |
| 跳过 IP 检测 | 默认会等紫鸟的 ip 检测页通过，确认代理生效（失败不中断） |

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
| `[code=98001008] 需要权限` | 两种可能：① **cookie 不是导出这张订单表的那个店铺的**（不是过期）—— 实测「汉维2-圣诞树 cookie + 嘉恒业订单」必报此码；② **页面没停在正确的区域站**（同一个店铺，`seller.*` 通用站看不到 `seller.us.*` 的订单，见下两行） | ①换成本店铺的 cookie：`获取客户/sync_cookie.py --store N` 从紫鸟当前店铺同步，或在上级 GUI 里跑（自动取当前店铺 cookie）；②正常跑主 GUI 即可 —— ④ 会自动扫区域站并锁定能通的那个 |
| `98001008` 时想「改 `params` 里的 `seller_id` 修好」 | **改不动**（2026-10-14 实测）：同一份 cookie 下把 `seller_id`/`oec_seller_id` 换成别的店 → `98001004`；把这两个参数**整段删掉** → 照样 `98001008`；改 `aid` / 加 `X-Tts-Oec-Bsid` / 换 `seller.us` 域名 host → 全是 `98001008`。服务端是**按会话（cookie）判定**你能看哪家店，参数不是权限开关 | 别在参数上试了，先跑 `python devtools/diagnose_store.py` 把三者身份打出来（见下） |
| 抓客户全批 `98001008`，但订单确实是这个店导出的 | **店铺身份取错**：`Storage.getCookies` 把多个区域店铺的同名 cookie 混在一起，合并时覆盖成了另一个店（实测：cookie 读到「13 单」那家的 id，而页面实际开着「619 单」那家） | 已修：③ 改按店铺后台地址精确取 cookie（`Network.getCookies`）+ 等身份稳定，导出后用**页面自己请求里的 seller_id** 复核，不一致以页面为准。升级后重跑即可；日志会打印 `⚠ 店铺身份不一致` 与页面实际用的 id |
| `[code=98001004] 部分信息填写错误` | 参数里的 `seller_id` 与 cookie 里的店铺对不上 | 清空界面上的 seller_id；注意硬改 cookie 里的 seller_id 字段也无效（服务端会拦），必须换真的那份 cookie |
| 换了正确店铺仍全批 `98001008`，页面内请求也一起被拒 | **页面没停在正确的「区域站」上**：同一个接口路径在 `seller.*`（通用站）回 `98001008`、`seller.eu.*` 回 `No matching route`、**只有 `seller.us.*` 通**（2026-10-14 真机日志）。其次才是 host / 路径前缀（`/na/` 之类） | 已自动兜底：④ 按 **「区域站 × 写法（模板/同源/绝对）」** 逐个试，命中即**锁定「写法 + 该站」组合**，之后每条只发 1 次。要做实可先跑 `python devtools/capture_order_request.py --store N --wait 30` 抓一次真实请求，再正常跑 `customer_gui.py` |
| 日志里明明有一次 `code=0`，紧接着又全是 `98001008` / `No matching route` | 曾经的实现缺陷（已修）：只记住「写法」没记住「站」，于是下一条跑到 eu 站还用同源路径；换站用的是 `ensure_on`（按站点尾巴比较，各区域站被视为同站 → **从不换站**）；且记录的是**请求地址**而非**页面实际落地的址**（区域站会互相重定向） | 已修：`_ensure_origin` 改用 `navigate` + 按 host 严格复核；每次导航后回读 `location` 取**实测站**；`winner_label` + `winner_origin` 一起锁定 |

## 目录结构

```
customer_gui.py         GUI 主程序（四步流水线）
tiktok_export.py        ★「导出订单」按钮自动化：面板操作 / 等生成 / 拦 URL / 取字节 / 解析
ziniao_bridge.py        紫鸟桥接：BrowserSession（原生 CDP + 断线自愈 + Fetch 拦截 + 页面内取字节）
tiktok_client.py        ④ 的 order/get 客户端：参数 / 解析 / Excel 读写
detail.py / list.py     原始抓包脚本（仅留档，接口直连思路已废弃）
devtools/export_check.py 真机端到端验证脚本
devtools/diagnose_store.py 98001008 诊断：页面 / cookie / 订单 三者身份对照
devtools/capture_order_request.py 抓后台真实订单请求 + 定位发送位置（并产出模板）
devtools/captured_order_request.json 抓包模板（④ 会自动照抄它）
devtools/smoke_gui.py   离屏冒烟测试
devtools/make_release.py 组装发布目录（exe + 旁挂紫鸟项目），可 --zip
build.spec              PyInstaller 配置（自动扫描紫鸟客户端的依赖）
.github/workflows/build.yml  CI：windows-latest 上构建 exe + 冻结产物自检
vendor/ziniao_webdriver_demo/ 内联的紫鸟协议客户端（只保留真正用到的 2 个 .py）
    util/ziniao_client.py   ★ 唯一真正干活的（604 行）
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

## 已知限制

- 导出文件里的手机号 / 姓名 / 地址**带平台脱敏**（`*****`），与 `order/get` 返回的一致。
- 店铺未登录该平台时订单页会跳登录，导出按钮不可用。
- 关店时只断开 CDP，不关浏览器进程；浏览器由紫鸟 `stopBrowser` 负责。
- 订单表格最多渲染前 800 行（`MAX_TABLE_ROWS`），完整数据看导出的 Excel。
