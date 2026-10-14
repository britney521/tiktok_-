"""DrissionPage 兼容层。

目标：把原本基于 selenium 的脚本平滑迁到 DrissionPage，**改动面压到最小**。

- `By`：取值与 selenium 完全一致，脚本里 `By.XPATH` / `By.TAG_NAME` 等写法原样保留
- `to_locator()`：把 selenium 的 `(by, value)` 翻译成 DrissionPage 定位串
- `DpElement` / `DpShadowRoot`：把 DrissionPage 元素包成 selenium 风格的最小接口
  （`text` / `get_attribute()` / `find_element()` / `find_elements()` / `shadow_root` / `click()` …）

因此 `amazon/*.py` 里除了把
    from selenium.webdriver.common.by import By
换成
    from util.dp import By
其余一行都不用动。

定位串语法（DrissionPage 4.x）：
    xpath://div      css:.cls      tag:tbody      @id=xxx       text:内容
在元素上查找时，以 `//` 开头的 xpath 会自动补 `.` 变成相对查找，
避免元素内查找变成整文档查找。
"""

from __future__ import annotations

import time
from typing import Any

DEFAULT_TIMEOUT = 10.0


class By:
    """与 selenium.webdriver.common.by.By 取值一致。"""

    ID = "id"
    XPATH = "xpath"
    LINK_TEXT = "link text"
    PARTIAL_LINK_TEXT = "partial link text"
    NAME = "name"
    TAG_NAME = "tag name"
    CLASS_NAME = "class name"
    CSS_SELECTOR = "css selector"


class ElementNotFound(Exception):
    """找不到单个元素（对应 selenium 的 NoSuchElementException）。"""


class ElementsNotFound(Exception):
    """一组元素一个都没找到（对应 selenium 的 TimeoutException）。"""


def _xpath_escape(value: str) -> str:
    if '"' not in value:
        return f'"{value}"'
    if "'" not in value:
        return f"'{value}'"
    parts = value.split('"')
    return "concat(" + ', \'"\', '.join(f'"{p}"' for p in parts) + ")"


def to_locator(by: str, value: str, under_element: bool = False) -> str:
    """selenium 的 (by, value) → DrissionPage 定位串。

    under_element=True 时用于元素内查找，此时以 `//` 开头的 xpath 补上 `.`。
    """
    if by == By.XPATH:
        v = value.strip()
        if under_element and v.startswith("//"):
            v = "." + v
        return f"xpath:{v}"
    if by == By.ID:
        return f'xpath://*[@id={_xpath_escape(value)}]'
    if by == By.NAME:
        return f'xpath://*[@name={_xpath_escape(value)}]'
    if by == By.TAG_NAME:
        return f"tag:{value}"
    if by == By.CLASS_NAME:
        # selenium 的 CLASS_NAME 是「按 class 精确匹配一个类名」
        return f'xpath://*[contains(concat(" ", normalize-space(@class), " "), {_xpath_escape(" " + value.strip() + " ")})]'
    if by == By.CSS_SELECTOR:
        return f"css:{value}"
    if by == By.LINK_TEXT:
        return f'xpath://a[normalize-space(.)={_xpath_escape(value)}]'
    if by == By.PARTIAL_LINK_TEXT:
        return f'xpath://a[contains(normalize-space(.), {_xpath_escape(value)})]'
    # 未知 by：当成 xpath 直接透传
    return f"xpath:{value}"


# ----------------------------------------------------------------------------
# 元素包装
# ----------------------------------------------------------------------------
class DpElement:
    """selenium 风格的元素最小接口，底层是 DrissionPage 的 ChromiumElement。"""

    def __init__(self, ele: Any, timeout: float = DEFAULT_TIMEOUT):
        self._ele = ele
        self._timeout = timeout

    # -- 原生对象 ---------------------------------------------------------
    @property
    def raw(self) -> Any:
        return self._ele

    def __bool__(self) -> bool:
        return bool(self._ele)

    def __repr__(self) -> str:
        return f"<DpElement {self._ele!r}>"

    # -- selenium 风格属性 -------------------------------------------------
    @property
    def text(self) -> str:
        return self._ele.text

    @property
    def tag_name(self) -> str:
        return self._ele.tag

    @property
    def shadow_root(self) -> "DpShadowRoot | None":
        sr = self._ele.shadow_root
        return DpShadowRoot(sr, self._timeout) if sr else None

    # -- selenium 风格方法 -------------------------------------------------
    def get_attribute(self, name: str) -> str:
        """selenium 找不到属性返回 None；这里统一返回空串，避免 `in None` 崩掉。"""
        value = self._ele.attr(name)
        return "" if value is None else str(value)

    def is_displayed(self) -> bool:
        try:
            return bool(self._ele.states.is_displayed)
        except Exception:
            return False

    def is_enabled(self) -> bool:
        try:
            return bool(self._ele.states.is_enabled)
        except Exception:
            return False

    def click(self) -> None:
        # by_js=None：先模拟点击，被遮挡时自动改用 JS 点击
        self._ele.click(by_js=None)

    def send_keys(self, text: str) -> None:
        self._ele.input(text)

    def clear(self) -> None:
        self._ele.clear()

    def find_element(self, by: str = By.ID, value: str | None = None) -> "DpElement":
        loc = to_locator(by, value or "", under_element=True)
        ele = self._ele.ele(loc, timeout=self._timeout)
        if not ele:
            raise ElementNotFound(f"元素内未找到：{by}={value}")
        return DpElement(ele, self._timeout)

    def find_elements(self, by: str = By.ID, value: str | None = None) -> list["DpElement"]:
        loc = to_locator(by, value or "", under_element=True)
        return [DpElement(e, self._timeout) for e in self._ele.eles(loc, timeout=self._timeout)]

    # -- 透传 DrissionPage 的高阶能力（脚本里可直接用）----------------------
    def scroll_into_view(self) -> None:
        try:
            self._ele.scroll.to_see()
        except Exception:
            pass

    def hover(self) -> None:
        self._ele.hover()

    def click_by_js(self) -> None:
        self._ele.click(by_js=True)


class DpShadowRoot:
    """Shadow DOM 包装，提供 find_element / find_elements。"""

    def __init__(self, shadow_root: Any, timeout: float = DEFAULT_TIMEOUT):
        self._sr = shadow_root
        self._timeout = timeout

    @property
    def raw(self) -> Any:
        return self._sr

    def __bool__(self) -> bool:
        return self._sr is not None

    def find_element(self, by: str = By.ID, value: str | None = None) -> DpElement:
        loc = to_locator(by, value or "", under_element=True)
        ele = self._sr.ele(loc, timeout=self._timeout)
        if not ele:
            raise ElementNotFound(f"shadow root 内未找到：{by}={value}")
        return DpElement(ele, self._timeout)

    def find_elements(self, by: str = By.ID, value: str | None = None) -> list[DpElement]:
        loc = to_locator(by, value or "", under_element=True)
        return [DpElement(e, self._timeout) for e in self._sr.eles(loc, timeout=self._timeout)]


def wait_until(fn, timeout: float, interval: float = 0.5, message: str = "等待超时"):
    """轮询等待条件成立，超时抛异常。"""
    end = time.time() + timeout
    last = None
    while time.time() < end:
        try:
            result = fn()
            if result:
                return result
        except Exception as exc:
            last = exc
        time.sleep(interval)
    raise ElementsNotFound(message if last is None else f"{message}（{last}）")


# ----------------------------------------------------------------------------
# 浏览器包装
# ----------------------------------------------------------------------------
class DpBrowser:
    """把 DrissionPage 的 `Chromium` 包装成「浏览器 + 当前页」二合一对象。

    为什么需要它：DrissionPage 的 `Chromium` 只是**浏览器控制器**，本身没有
    `get()` / `ele()` / `url` 这些页面方法（那些在 `latest_tab` 上）。而原脚本里
    `driver` 是 selenium 的 WebDriver，页面方法与浏览器方法混在一起用。
    这里做一层转发，让 `driver.get(...)`、`driver.latest_tab` 同时可用。

    另外 `quit()` 被改成**只断开 CDP 连接**，不关闭浏览器——紫鸟的店铺浏览器
    由客户端统一管理，关闭交给 `stopBrowser`（即 `close_store`）。
    """

    def __init__(self, chromium: Any, timeout: float = DEFAULT_TIMEOUT):
        self._chromium = chromium
        self._timeout = timeout

    # -- 原生对象 ---------------------------------------------------------
    @property
    def raw(self) -> Any:
        return self._chromium

    def __repr__(self) -> str:
        return f"<DpBrowser {self._chromium!r}>"

    # -- 浏览器级 ---------------------------------------------------------
    @property
    def latest_tab(self):
        """当前活动标签页。"""
        return self._chromium.latest_tab

    # 语义别名，脚本里两种写法都能用
    tab = latest_tab

    def get_tabs(self) -> list:
        return self._chromium.get_tabs()

    def new_tab(self, url: str | None = None, **kwargs):
        return self._chromium.new_tab(url, **kwargs)

    def close_tabs(self, *args, **kwargs):
        return self._chromium.close_tabs(*args, **kwargs)

    @property
    def tab_ids(self) -> list:
        return self._chromium.tab_ids

    # -- 页面级：转发到当前标签页 -----------------------------------------
    def get(self, url: str, **kwargs):
        return self._chromium.latest_tab.get(url, **kwargs)

    def ele(self, *args, **kwargs):
        return self._chromium.latest_tab.ele(*args, **kwargs)

    def eles(self, *args, **kwargs):
        return self._chromium.latest_tab.eles(*args, **kwargs)

    @property
    def url(self) -> str:
        return self._chromium.latest_tab.url

    @property
    def title(self) -> str:
        return self._chromium.latest_tab.title

    @property
    def html(self) -> str:
        return self._chromium.latest_tab.html

    def refresh(self):
        return self._chromium.latest_tab.refresh()

    # -- 兼容 selenium 语义 -----------------------------------------------
    def implicitly_wait(self, timeout: float) -> None:
        """对齐 selenium 的隐式等待：设置全局元素查找默认超时（秒）。"""
        self._timeout = float(timeout)
        try:
            self._chromium.latest_tab.set.timeouts(base=float(timeout))
        except Exception:
            pass

    def quit(self) -> None:
        """断开 CDP 连接。

        故意**不**调用 `Chromium.quit()`——那会真的把浏览器关掉（发 Browser.close），
        而这里应该由紫鸟客户端的 stopBrowser 来关。所以只断开本地驱动连接。
        """
        try:
            self._chromium._driver.stop()
        except Exception:
            try:
                self._chromium.reconnect()
            except Exception:
                pass

    def __getattr__(self, item: str) -> Any:
        # 没显式实现的一律转发给当前标签页，保证老脚本不因缺方法而崩
        if item.startswith("_"):
            raise AttributeError(item)
        return getattr(self._chromium.latest_tab, item)
