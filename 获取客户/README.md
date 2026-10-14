# 获取客户（独立小工具）

只做一件事：**读订单 Excel 的 `Order ID` 列 → 逐个调 `order/get` → 写 `客户.xlsx`**。

从上级目录 `customer_gui.py`（四步流水线）里把「获取客户」这一段单独拆出来，做成
零依赖紫鸟、可独立运行的窗口。

## 为什么能独立运行

实测 `order/get`（订单详情）**本地 requests 直连即可**：
`aid=4068` + `seller_id` → `code=0`，不需要 X-Bogus / X-Tts-Oec-Bsid 签名头。

> 边界：`order/list`（订单列表）不适用本结论 —— 本地直连恒定 `code=10000`，
> 必须在店铺浏览器里带签名发起。订单导出请用上级目录的 `customer_gui.py`。

## 运行

```bash
python get_customers_gui.py            # 打开界面
python get_customers_gui.py --check    # 离屏自检（不联网）
QT_QPA_PLATFORM=offscreen python devtools/smoke_gui.py   # 冒烟测试（零网络）
```

依赖：`PyQt6`、`requests`、`openpyxl`（见上级 requirements.txt）。

## 界面

1. **数据来源**：订单 Excel（平台导出表 / 上级工具导出的订单.xlsx 均可）；
   自动列工作表与列名，选中 `Order ID` 列后实时预览读到的 ID 数
   （自动去重、过滤非 ID 杂质行）
2. **输出与节奏**：客户 Excel 输出路径、请求间隔（防限流）、失败重试
3. **凭据**：接口主机、seller_id（留空从 cookie 推导）、Cookie 编辑

## ★ cookie 必须是「订单所属那个店铺」的（最容易踩的坑）

一个店铺的登录态**只能看它自己的订单** —— cookie 里的
`oec_seller_id_unified_seller_env` 决定了「当前在哪个店」。

2026-10-13 实测（同一份 cookie，只换订单号）：

| 请求 | 结果 |
| --- | --- |
| 汉维2-圣诞树 cookie + 汉维2-圣诞树订单 | `code=0` success，拿到买家昵称 |
| 汉维2-圣诞树 cookie + **嘉恒业**订单 | `code=98001008` **需要权限** |
| 把 cookie 里的 seller_id 字段硬改成嘉恒业 | `code=98001004` 服务端校验拦下，改不动 |
| 汉维 cookie + 参数 seller_id 填嘉恒业 | `code=98001004` 参数与 cookie 不一致 |

→ 所以 **`98001008 需要权限` 不是 cookie 过期，是 cookie 拿错店了**；
换店铺就得换那份 cookie。`X-Tts-Oec-Bsid` 加不加都不影响这个结论。

同一份 cookie 单独取 `seller_id` 也可能取错：统一商家环境下同一个 profile 会同时存在
多个区域店铺（`seller.us.*` / `seller.eu.*`）的**同名 cookie**，合并成 dict 时后写覆盖前写。
上级 GUI 已改成按店铺后台地址精确取（`Network.getCookies`）；本目录工具请确保 `cookies.json`
里那份就是目标店铺的。

两个防呆：

- **🔍 凭据预检**（界面上）：只发 1 次请求验一下 cookie 能不能看这批订单，
  别等跑完几百条才发现全错。
- **提前中止**：一旦命中 98001008 / 98001004，工具立刻停手（其余行标
  `SKIPPED`），不再对同一个错误重复请求几百次。

## Cookie 管理

- `工具 → 编辑 Cookie…`：粘贴 `name=value; name2=value2` 或 JSON 对象
- `工具 → 保存 Cookie 到 cookies.json`：持久化，下次启动自动加载
- 同目录 `cookies.json` 存在时优先于内置默认值
- **`python sync_cookie.py --store 2`**：从紫鸟浏览器**当前店铺**直接同步一份
  cookie 到 `cookies.json`，再也不用手工从 DevTools 抄：

  ```bash
  python sync_cookie.py                        # 列出店铺，交互式选
  python sync_cookie.py --store 2              # 序号选（与上级 GUI 同序）
  python sync_cookie.py --name 嘉恒业           # 按店名模糊匹配
  python sync_cookie.py --port 9222            # 浏览器已在跑，直连调试端口
  python sync_cookie.py --store 2 --check-orders ../订单.xlsx
                                               # 同步完顺手预检第一条订单
  ```

  对照：**汉维2-圣诞树 = `7494876782645708274`**，**嘉恒业 = `7496000161034505075`**。
  会话过期（报 `code=1000003`）时也用它重新同步即可。

## 文件

```
get_customers_gui.py   GUI 主程序（含「🔍 凭据预检」）
customer_api.py        order/get 客户端 + probe 预检 + 解析 + Excel 读写（可独立复用）
sync_cookie.py         从紫鸟当前店铺同步 cookie 到 cookies.json
cookies.json           本地 cookie（优先于内置默认值）
客户.xlsx              默认输出
devtools/smoke_gui.py  离屏冒烟测试
```

## 已知限制

- 接口返回的手机号 / 邮箱 / 收件人姓名本身带平台脱敏（`*****`），非工具问题。
- 失败的订单会写出占位行（`Fetch Status = API_ERROR / EMPTY / EXCEPTION / SKIPPED`），
  行数与读取到的 ID 数一致，方便对照补抓。`SKIPPED` 表示因凭据问题提前中止、没发请求。
- 每抓 10 条自动落盘一次，中途停止/崩溃不丢已抓数据。
