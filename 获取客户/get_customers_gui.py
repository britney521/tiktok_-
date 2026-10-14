"""TikTok Shop 客户信息抓取工具（独立 GUI）。

功能只有一件事，做成一个能单独运行的窗口：
    读订单 Excel 的 ``Order ID`` 列 → 逐个调 order/get → 写 ``客户.xlsx``

为什么能独立（不需要紫鸟浏览器）:
    实测 ``order/get`` 本地 requests 直连即可（``aid=4068`` + ``seller_id`` → code=0）。
    订单导出（order/list）才必须走店铺浏览器，那部分在上级目录的 customer_gui.py。

★ cookie 必须是「订单所属那个店铺」的:
    一个店铺的登录态只能看它自己的订单。拿错店 → ``code=98001008 需要权限``
    （这**不是** cookie 过期，是「不是这家店」）。换店铺就得换 cookie：
        · 界面「工具 → 编辑 Cookie」粘贴新 cookie；或
        · ``python sync_cookie.py --store 2`` 从紫鸟当前店铺直接同步一份。
    跑整批前建议先点「🔍 凭据预检」，只花 1 次请求就能验证。

数据来源与输出:
    · 订单 Excel：任意含 ``Order ID`` 列的表（平台导出的订单表 / 上级工具导出的订单.xlsx）
    · 客户 Excel：完整 17 列（含 Fetch Status）

运行:
    python get_customers_gui.py
离屏自检（不联网，仅构建界面）:
    python get_customers_gui.py --check
"""

from __future__ import annotations

import json
import re
import sys
import threading
import traceback
from pathlib import Path
from typing import Any

from PyQt6.QtCore import QObject, Qt, QThread, pyqtSignal
from PyQt6.QtGui import QAction, QFont
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QDockWidget,
    QDoubleSpinBox,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QSpinBox,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

import customer_api as api

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_CUSTOMER_OUT = BASE_DIR / "客户.xlsx"

# 抓客户的最小请求间隔（秒）。太密容易被平台风控，界面上不允许调到比这更小。
MIN_CUSTOMER_INTERVAL = 1.5


def order_candidates() -> list[Path]:
    """上级 GUI 可能导出的订单表，新导出的排在前面。

    上级 GUI 现在按店铺命名（``<店铺名>_订单.xlsx``，一个店铺一份），所以先按
    修改时间倒序收这批，再回退到旧名 / 平台下载的原始表。
    """
    parent = BASE_DIR.parent
    named = sorted(parent.glob("*_订单.xlsx"),
                   key=lambda p: p.stat().st_mtime, reverse=True)
    return [*named,
            parent / "订单.xlsx",
            parent / "全部 笔订单-2026-10-13-06_43.xlsx"]


def store_name_from_order(path: Path | str) -> str:
    """从「<店铺名>_订单.xlsx」里取店铺名；取不到返回空串。"""
    text = str(path or "").strip()
    if not text:
        return ""
    m = re.match(r"^(.+)_订单$", Path(text).stem)
    return m.group(1) if m else ""


DEFAULT_ORDER_IN = next((p for p in order_candidates() if p.exists()),
                        BASE_DIR.parent / "订单.xlsx")


# ----------------------------------------------------------------------------
# 日志总线（工作线程 → GUI 线程）
# ----------------------------------------------------------------------------
class LogBus(QObject):
    """把任意线程的日志请求转成 Qt 信号，排队回 GUI 线程。

    QPlainTextEdit 不是线程安全的，工作线程直接写会 SIGSEGV。
    """

    message = pyqtSignal(str)


# ----------------------------------------------------------------------------
# 后台线程
# ----------------------------------------------------------------------------
class FetchThread(QThread):
    """读 ID → 逐条抓取 → 落盘。"""

    log = pyqtSignal(str)
    progress = pyqtSignal(int, int)          # 已完成, 总数
    row = pyqtSignal(int, dict)              # 行号, 客户信息
    done = pyqtSignal(dict)
    failed = pyqtSignal(str)

    def __init__(self, opts: dict[str, Any], parent=None):
        super().__init__(parent)
        self.opts = dict(opts)
        self.stop_event = threading.Event()

    def stop(self) -> None:
        self.stop_event.set()

    def run(self) -> None:  # noqa: D102
        o = self.opts
        try:
            self.log.emit("正在读取 Order ID…")
            src = Path(o["order_in"])
            ids = api.read_order_ids(src, o.get("sheet") or None,
                                     o.get("column") or "Order ID")
            if not ids:
                raise api.ApiError(
                    "no_ids",
                    f"在「{o.get('column') or 'Order ID'}」列没有读到有效 Order ID（{src}）",
                )
            self.log.emit(f"共读取 {len(ids)} 个 Order ID（已去重）")

            cookies = o.get("cookies") or api.load_cookies(o.get("cookie_file"))
            client = api.CustomerClient.from_cookies(
                cookies,
                api_base=(o.get("api_base") or api.DEFAULT_API_BASE).rstrip("/"),
                retries=int(o.get("retries") or 2),
            )
            if o.get("seller_id"):
                client.seller_id = str(o["seller_id"]).strip()
            self.log.emit(f"接口主机：{client.api_base}；seller_id={client.seller_id}")

            result = api.fetch_customers(
                client, ids,
                on_row=lambda i, r: self.row.emit(i, r),
                on_progress=lambda d, t: self.progress.emit(d, t),
                should_stop=self.stop_event.is_set,
                interval=float(o.get("interval") or 0.0),
                flush_every=int(o.get("flush_every") or 10),
                out_path=o.get("customer_out") or None,
                on_log=self.log.emit,
            )
            self.done.emit(result)
        except Exception as exc:
            self.log.emit(traceback.format_exc(limit=8))
            self.failed.emit(f"{type(exc).__name__}: {exc}")


class ProbeThread(QThread):
    """凭据预检：只发 1 次请求，验证 cookie 能否看这批订单。"""

    log = pyqtSignal(str)
    done = pyqtSignal(dict)

    def __init__(self, client: api.CustomerClient, order_id: str, parent=None):
        super().__init__(parent)
        self.client = client
        self.order_id = order_id

    def run(self) -> None:  # noqa: D102
        self.log.emit(f"预检：用订单 {self.order_id} 试一次 "
                      f"order/get（seller_id={self.client.seller_id or '(空)'}）")
        try:
            self.done.emit(self.client.probe(self.order_id))
        except Exception as exc:  # 兜底，绝不让线程静默死掉
            self.done.emit({"ok": False, "code": "exception",
                            "message": f"{type(exc).__name__}: {exc}",
                            "hint": "", "order_id": self.order_id})


# ----------------------------------------------------------------------------
# 对话框
# ----------------------------------------------------------------------------
class CookieDialog(QDialog):
    """粘贴 / 查看当前 Cookie。"""

    def __init__(self, cookies: dict[str, str], parent=None):
        super().__init__(parent)
        self.setWindowTitle("Cookie")
        self.resize(760, 480)
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(
            "支持两种格式：\n"
            "  · 浏览器复制的 name=value; name2=value2\n"
            "  · JSON 对象 {\"name\": \"value\"}"))
        self.edit = QPlainTextEdit()
        self.edit.setFont(QFont("Menlo", 11))
        self.edit.setPlainText(json.dumps(cookies, ensure_ascii=False, indent=2))
        layout.addWidget(self.edit, 1)
        box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        box.accepted.connect(self.accept)
        box.rejected.connect(self.reject)
        layout.addWidget(box)

    def cookies(self) -> dict[str, str]:
        return api.parse_cookie_text(self.edit.toPlainText())


# ----------------------------------------------------------------------------
# 主窗口
# ----------------------------------------------------------------------------
class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("TikTok 客户信息抓取（order/get）")
        self.resize(1180, 860)

        self.cookies: dict[str, str] = api.load_cookies()
        self.worker: Any = None

        self.log_bus = LogBus()
        self.log_bus.message.connect(self._append_log)

        self._build_ui()
        self._build_log_dock()
        self._build_menu()
        self._set_busy(False)
        self._refresh_meta()
        self.log_line("就绪：选订单 Excel → 核对 Order ID 列 → 建议先「凭据预检」→「开始抓取」。")
        _sid = api.seller_id_from_cookies(self.cookies)
        _name = api.KNOWN_STORES.get(_sid, "")
        self.log_line(
            f"Cookie：{len(self.cookies)} 条，店铺 seller_id="
            f"{_sid or '(未取到)'}{f'（{_name}）' if _name else ''}"
            f"（可用「工具 → 编辑 Cookie」更换）")
        self.log_line(
            "⚠ 一个店铺的 cookie 只能看它自己的订单：cookie 里的 seller_id 必须与"
            "订单表所属店铺一致，否则 order/get 返回 98001008「需要权限」。")

    # -- 界面构建 -----------------------------------------------------------
    def _build_ui(self) -> None:
        top = QWidget()
        tv = QVBoxLayout(top)
        tv.setContentsMargins(8, 8, 8, 8)
        tv.setSpacing(8)
        tv.addWidget(self._build_input_box())
        tv.addWidget(self._build_output_box())
        tv.addWidget(self._build_cred_box())
        tv.addLayout(self._build_run_bar())
        tv.addStretch(0)

        self.tbl = QTableWidget(0, len(api.PREVIEW_COLUMNS))
        self.tbl.setHorizontalHeaderLabels(api.PREVIEW_COLUMNS)
        self.tbl.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.tbl.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.tbl.setAlternatingRowColors(True)
        self.tbl.verticalHeader().setDefaultSectionSize(24)
        h = self.tbl.horizontalHeader()
        h.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        h.setStretchLastSection(True)

        main = QSplitter(Qt.Orientation.Vertical)
        main.addWidget(top)
        main.addWidget(self.tbl)
        main.setStretchFactor(0, 0)
        main.setStretchFactor(1, 1)
        self.setCentralWidget(main)

    def _build_input_box(self) -> QGroupBox:
        box = QGroupBox("① 数据来源（订单 Excel）")
        grid = QGridLayout(box)
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(6)

        grid.addWidget(QLabel("订单 Excel："), 0, 0)
        self.ed_in = QLineEdit(str(DEFAULT_ORDER_IN) if DEFAULT_ORDER_IN.exists() else "")
        self.ed_in.setPlaceholderText("含 Order ID 列的 Excel")
        self.ed_in.setMinimumWidth(160)
        self.ed_in.textChanged.connect(self._refresh_meta)
        grid.addWidget(self.ed_in, 0, 1, 1, 3)
        btn = QPushButton("浏览…")
        btn.clicked.connect(self._pick_input)
        grid.addWidget(btn, 0, 4)

        grid.addWidget(QLabel("工作表："), 1, 0)
        self.cb_sheet = QComboBox()
        self.cb_sheet.setMinimumWidth(120)
        self.cb_sheet.currentTextChanged.connect(self._refresh_columns)
        grid.addWidget(self.cb_sheet, 1, 1)
        grid.addWidget(QLabel("Order ID 列："), 1, 2)
        self.cb_column = QComboBox()
        self.cb_column.setEditable(True)
        self.cb_column.setMinimumWidth(120)
        self.cb_column.addItem("Order ID")
        self.cb_column.currentTextChanged.connect(self._update_id_preview)
        grid.addWidget(self.cb_column, 1, 3, 1, 2)

        self.lbl_ids = QLabel("Order ID：—")
        self.lbl_ids.setWordWrap(True)
        grid.addWidget(self.lbl_ids, 2, 0, 1, 5)

        grid.setColumnStretch(1, 1)
        return box

    def _build_output_box(self) -> QGroupBox:
        box = QGroupBox("② 输出与节奏")
        grid = QGridLayout(box)
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(6)

        grid.addWidget(QLabel("客户 Excel："), 0, 0)
        self.ed_out = QLineEdit(str(DEFAULT_CUSTOMER_OUT))
        self.ed_out.setMinimumWidth(160)
        grid.addWidget(self.ed_out, 0, 1, 1, 3)
        btn = QPushButton("另存为…")
        btn.clicked.connect(self._pick_output)
        grid.addWidget(btn, 0, 4)

        grid.addWidget(QLabel("请求间隔(秒)："), 1, 0)
        self.sp_interval = QDoubleSpinBox()
        self.sp_interval.setRange(MIN_CUSTOMER_INTERVAL, 60.0)
        self.sp_interval.setSingleStep(0.5)
        self.sp_interval.setValue(MIN_CUSTOMER_INTERVAL)
        self.sp_interval.setToolTip(f"每条之间的等待，防限流；最小 {MIN_CUSTOMER_INTERVAL}s")
        self.sp_interval.setMinimumWidth(100)
        grid.addWidget(self.sp_interval, 1, 1)
        grid.addWidget(QLabel("失败重试："), 1, 2)
        self.sp_retry = QSpinBox()
        self.sp_retry.setRange(0, 10)
        self.sp_retry.setValue(2)
        self.sp_retry.setMinimumWidth(100)
        grid.addWidget(self.sp_retry, 1, 3)

        grid.setColumnStretch(1, 1)
        return box

    def _build_cred_box(self) -> QGroupBox:
        box = QGroupBox("③ 凭据")
        grid = QGridLayout(box)
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(6)

        grid.addWidget(QLabel("接口主机："), 0, 0)
        self.ed_api = QLineEdit(api.DEFAULT_API_BASE)
        self.ed_api.setMinimumWidth(160)
        grid.addWidget(self.ed_api, 0, 1, 1, 4)

        grid.addWidget(QLabel("seller_id："), 1, 0)
        self.ed_seller = QLineEdit()
        self.ed_seller.setPlaceholderText("留空则从 Cookie 自动取")
        self.ed_seller.setMinimumWidth(160)
        grid.addWidget(self.ed_seller, 1, 1, 1, 4)

        self.lbl_cookie = QLabel("Cookie：—")
        self.lbl_cookie.setWordWrap(True)
        grid.addWidget(self.lbl_cookie, 2, 0, 1, 3)
        bar = QHBoxLayout()
        btn_edit = QPushButton("编辑 Cookie…")
        btn_edit.clicked.connect(self._edit_cookies)
        bar.addWidget(btn_edit)
        btn_reload = QPushButton("重载 cookies.json")
        btn_reload.clicked.connect(self._reload_cookies)
        bar.addWidget(btn_reload)
        bar.addStretch(1)
        grid.addLayout(bar, 2, 3, 1, 2)

        grid.setColumnStretch(1, 1)
        return box

    def _build_run_bar(self) -> QHBoxLayout:
        bar = QHBoxLayout()
        self.btn_run = QPushButton("▶ 开始抓取")
        self.btn_run.setMinimumHeight(36)
        self.btn_run.clicked.connect(self._start)
        bar.addWidget(self.btn_run, 2)

        self.btn_probe = QPushButton("🔍 凭据预检")
        self.btn_probe.setMinimumHeight(36)
        self.btn_probe.setToolTip(
            "只发 1 次请求，验证当前 cookie 能不能看这批订单。\n"
            "报 98001008 就是 cookie 不是这个店铺的。")
        self.btn_probe.clicked.connect(self._preflight)
        bar.addWidget(self.btn_probe)

        self.btn_stop = QPushButton("■ 停止")
        self.btn_stop.setMinimumHeight(36)
        self.btn_stop.clicked.connect(self._stop)
        bar.addWidget(self.btn_stop)

        self.lbl_stat = QLabel("等待开始")
        bar.addWidget(self.lbl_stat, 1)

        self.bar_progress = QProgressBar()
        self.bar_progress.setMinimumWidth(220)
        self.bar_progress.setRange(0, 100)
        self.bar_progress.setValue(0)
        bar.addWidget(self.bar_progress, 2)
        return bar

    def _build_log_dock(self) -> None:
        dock = QDockWidget("日志", self)
        dock.setAllowedAreas(Qt.DockWidgetArea.BottomDockWidgetArea)
        self.log_box = QPlainTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setFont(QFont("Menlo", 11))
        self.log_box.setMaximumBlockCount(4000)
        dock.setWidget(self.log_box)
        self.addDockWidget(Qt.DockWidgetArea.BottomDockWidgetArea, dock)
        self.resizeDocks([dock], [200], Qt.Orientation.Vertical)

    def _build_menu(self) -> None:
        m = self.menuBar().addMenu("工具")
        act_cookie = QAction("编辑 Cookie…", self)
        act_cookie.triggered.connect(self._edit_cookies)
        m.addAction(act_cookie)
        act_save = QAction("保存 Cookie 到 cookies.json", self)
        act_save.triggered.connect(self._save_cookies)
        m.addAction(act_save)
        act_open = QAction("打开输出目录", self)
        act_open.triggered.connect(self._open_output_dir)
        m.addAction(act_open)
        act_clear = QAction("清空日志", self)
        act_clear.triggered.connect(self.log_box.clear)
        m.addAction(act_clear)
        act_check = QAction("环境自检", self)
        act_check.triggered.connect(self._self_check)
        m.addAction(act_check)

    # -- 日志 ---------------------------------------------------------------
    def log_line(self, msg: str) -> None:
        app = QApplication.instance()
        if app is None or QThread.currentThread() is app.thread():
            self._append_log(msg)
        else:
            self.log_bus.message.emit(str(msg))

    def _append_log(self, msg: str) -> None:
        if not hasattr(self, "log_box"):
            return
        self.log_box.appendPlainText(str(msg))

    # -- 状态 ---------------------------------------------------------------
    def _set_busy(self, busy: bool) -> None:
        for w in (self.btn_run, self.btn_probe, self.ed_in, self.cb_sheet,
                  self.cb_column, self.ed_out, self.sp_interval, self.sp_retry,
                  self.ed_api, self.ed_seller):
            w.setEnabled(not busy)
        self.btn_stop.setEnabled(busy)
        if not busy:
            self.lbl_stat.setText("就绪")

    # -- 元信息 -------------------------------------------------------------
    def _sync_out_name(self) -> None:
        """订单表带店铺名时（``<店铺名>_订单.xlsx``），客户表也跟着带店铺名。

        只在用户没手动改过「客户 Excel」时才覆盖 —— 框里仍是默认名、或仍是
        上一次自动填进去的名字。用户「另存为」过就原样保留。
        """
        line = getattr(self, "ed_out", None)
        if line is None:
            return
        name = store_name_from_order(self.ed_in.text().strip())
        target = (DEFAULT_CUSTOMER_OUT.with_name(f"{name}_客户.xlsx")
                  if name else DEFAULT_CUSTOMER_OUT)
        cur = line.text().strip()
        last_auto = getattr(self, "_auto_out", None)
        if cur in ("", str(DEFAULT_CUSTOMER_OUT)) or (last_auto and cur == last_auto):
            line.setText(str(target))
            self._auto_out = str(target)

    def _refresh_meta(self) -> None:
        self._sync_out_name()
        self._refresh_sheets()
        n = len(self.cookies)
        seller = api.seller_id_from_cookies(self.cookies) or ""
        known = api.KNOWN_STORES.get(seller, "")
        label = f"{seller}（{known}）" if known else (seller or "(未取到，cookie 里没有登录态)")
        self.lbl_cookie.setText(f"Cookie：{n} 条；店铺 seller_id={label}")

    def _refresh_sheets(self) -> None:
        if not hasattr(self, "cb_sheet"):
            return
        # 注意：Path("") 恰好是当前目录，exists() 为 True，必须先判空
        text = self.ed_in.text().strip()
        if not text:
            self.lbl_ids.setText("Order ID：—（请选择订单 Excel）")
            return
        path = Path(text)
        if not path.exists():
            self.lbl_ids.setText("Order ID：—（请选择存在的订单 Excel）")
            return
        try:
            sheets = api.list_sheets(path)
        except Exception as exc:
            self.log_line(f"读取工作表失败：{exc}")
            return
        current = self.cb_sheet.currentText()
        self.cb_sheet.blockSignals(True)
        self.cb_sheet.clear()
        self.cb_sheet.addItems(sheets)
        if current in sheets:
            self.cb_sheet.setCurrentText(current)
        self.cb_sheet.blockSignals(False)
        self._refresh_columns()

    def _refresh_columns(self) -> None:
        text = self.ed_in.text().strip()
        if not text:
            return
        path = Path(text)
        if not path.exists():
            return
        sheet = self.cb_sheet.currentText() or None
        try:
            cols = api.read_header(path, sheet, row=1)
        except Exception as exc:
            self.log_line(f"读取列名失败：{exc}")
            return
        if not cols:
            return
        current = self.cb_column.currentText()
        self.cb_column.blockSignals(True)
        self.cb_column.clear()
        self.cb_column.addItems(cols)
        self.cb_column.setCurrentText(
            current if current in cols else ("Order ID" if "Order ID" in cols else cols[0]))
        self.cb_column.blockSignals(False)
        self._update_id_preview()

    def _update_id_preview(self) -> None:
        text = self.ed_in.text().strip()
        if not text:
            self.lbl_ids.setText("Order ID：—（请选择订单 Excel）")
            return
        path = Path(text)
        column = self.cb_column.currentText().strip() or "Order ID"
        if not path.exists():
            self.lbl_ids.setText("Order ID：—（请选择存在的订单 Excel）")
            return
        try:
            ids = api.read_order_ids(path, self.cb_sheet.currentText() or None, column)
        except Exception as exc:
            self.lbl_ids.setText(f"Order ID：读取失败（{exc}）")
            return
        if not ids:
            self.lbl_ids.setText(f"Order ID：在「{column}」列没读到有效 ID")
            return
        head = ", ".join(ids[:3])
        more = " …" if len(ids) > 3 else ""
        self.lbl_ids.setText(f"Order ID：共 {len(ids)} 个（{head}{more}）")

    # -- Cookie -------------------------------------------------------------
    def _edit_cookies(self) -> None:
        dlg = CookieDialog(self.cookies, self)
        if dlg.exec():
            parsed = dlg.cookies()
            if not parsed:
                QMessageBox.warning(self, "Cookie 为空", "没解析出任何 cookie。")
                return
            self.cookies = parsed
            self._refresh_meta()
            self.log_line(f"Cookie 已更新：{len(parsed)} 条")

    def _reload_cookies(self) -> None:
        self.cookies = api.load_cookies()
        self._refresh_meta()
        self.log_line(f"已重载 cookies.json：{len(self.cookies)} 条")

    def _save_cookies(self) -> None:
        try:
            path = api.save_cookies(self.cookies)
            self.log_line(f"Cookie 已保存：{path}")
        except Exception as exc:
            QMessageBox.critical(self, "保存失败", str(exc))

    # -- 文件 ---------------------------------------------------------------
    def _pick_input(self) -> None:
        start = self.ed_in.text().strip() or str(BASE_DIR.parent)
        path, _ = QFileDialog.getOpenFileName(
            self, "选择订单 Excel", start, "Excel 文件 (*.xlsx *.xlsm *.xls)")
        if path:
            self.ed_in.setText(path)

    def _pick_output(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self, "保存为", self.ed_out.text().strip() or str(DEFAULT_CUSTOMER_OUT),
            "Excel 文件 (*.xlsx)")
        if path:
            self.ed_out.setText(path)

    def _open_output_dir(self) -> None:
        import subprocess

        p = Path(self.ed_out.text().strip() or str(DEFAULT_CUSTOMER_OUT))
        target = p.parent if p.parent.exists() else BASE_DIR
        subprocess.Popen(["open", str(target)])

    # -- 凭据 ---------------------------------------------------------------
    def _make_client(self) -> api.CustomerClient:
        """按界面当前设置造客户端（预检与正式抓取共用）。"""
        client = api.CustomerClient.from_cookies(
            dict(self.cookies),
            api_base=(self.ed_api.text().strip() or api.DEFAULT_API_BASE).rstrip("/"),
            retries=1,
        )
        sid = self.ed_seller.text().strip()
        if sid:
            client.seller_id = sid
        return client

    def _first_order_id(self) -> str:
        text = self.ed_in.text().strip()
        if not text or not Path(text).exists():
            return ""
        try:
            ids = api.read_order_ids(Path(text), self.cb_sheet.currentText() or None,
                                     self.cb_column.currentText().strip() or "Order ID")
        except Exception:
            return ""
        return ids[0] if ids else ""

    def _preflight(self) -> None:
        oid = self._first_order_id()
        if not oid:
            QMessageBox.warning(self, "未就绪", "请先选择订单 Excel（需要能读到 Order ID）。")
            return
        client = self._make_client()
        self._set_busy(True)
        self.lbl_stat.setText("预检中…")
        self.worker = ProbeThread(client, oid, self)
        self.worker.log.connect(self.log_line)
        self.worker.done.connect(self._on_probe)
        self.worker.start()

    def _on_probe(self, res: dict) -> None:
        self._set_busy(False)
        code = str(res.get("code"))
        if res.get("ok"):
            self.lbl_stat.setText("预检通过")
            self.log_line(f"✅ 预检通过：cookie 可用（订单 {res.get('order_id')}"
                          f" 买家 {res.get('buyer') or '—'}）。可以开始抓取。")
            return

        self.lbl_stat.setText("预检未通过")
        self.log_line(f"❌ 预检未通过：code={code} {res.get('message') or ''}")
        if res.get("hint"):
            self.log_line(f"   → {res['hint']}")

        body = (f"code={code}\n{res.get('message') or ''}")
        if res.get("hint"):
            body += f"\n\n→ {res['hint']}"
        if code in ("98001008", "98001004"):
            body += ("\n\n常见原因：这份 cookie 属于【别的店铺】。\n"
                     "一个店铺的 cookie 只能看它自己的订单 ——\n"
                     "请点「编辑 Cookie…」粘贴导出该订单表的那个店铺的 cookie，\n"
                     "或改用上级目录 customer_gui.py（它会自动取当前店铺的 cookie）。")
        elif code == "no_seller":
            body += ("\n\ncookie 里没有该平台的登录态，请换一份目标店铺的 cookie。")
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Icon.Warning)
        box.setWindowTitle("凭据预检未通过")
        box.setText(body)
        edit_btn = box.addButton("编辑 Cookie…", QMessageBox.ButtonRole.AcceptRole)
        box.addButton("知道了", QMessageBox.ButtonRole.RejectRole)
        box.exec()
        if box.clickedButton() is edit_btn:
            self._edit_cookies()

    # -- 运行 ---------------------------------------------------------------
    def _start(self) -> None:
        src = Path(self.ed_in.text().strip())
        if not src.exists():
            QMessageBox.warning(self, "未就绪", "请选择存在的订单 Excel。")
            return
        out = Path(self.ed_out.text().strip() or str(DEFAULT_CUSTOMER_OUT))
        if out.suffix.lower() not in (".xlsx", ".xlsm"):
            out = out.with_suffix(".xlsx")
            self.ed_out.setText(str(out))

        opts = {
            "order_in": str(src),
            "sheet": self.cb_sheet.currentText().strip(),
            "column": self.cb_column.currentText().strip() or "Order ID",
            "customer_out": str(out),
            "interval": float(self.sp_interval.value()),
            "retries": int(self.sp_retry.value()),
            "api_base": self.ed_api.text().strip() or api.DEFAULT_API_BASE,
            "seller_id": self.ed_seller.text().strip(),
            "cookies": dict(self.cookies),
        }

        self._set_busy(True)
        self.bar_progress.setRange(0, 0)
        self.bar_progress.setValue(0)
        self.tbl.setRowCount(0)
        self.log_line("=" * 60)
        self.log_line(f"开始抓取：来源={src.name}　列={opts['column']}　输出={out.name}")

        self.worker = FetchThread(opts, self)
        self.worker.log.connect(self.log_line)
        self.worker.progress.connect(self._on_progress)
        self.worker.row.connect(self._on_row)
        self.worker.done.connect(self._on_done)
        self.worker.failed.connect(self._on_failed)
        self.worker.start()

    def _on_progress(self, done: int, total: int) -> None:
        self.bar_progress.setRange(0, max(1, total))
        self.bar_progress.setValue(done)
        self.lbl_stat.setText(f"抓取中：{done}/{total}")

    def _on_row(self, idx: int, row: dict) -> None:
        if idx < 0:
            return
        if self.tbl.rowCount() <= idx:
            self.tbl.setRowCount(idx + 1)
        for c, col in enumerate(api.PREVIEW_COLUMNS):
            self.tbl.setItem(idx, c, QTableWidgetItem(str(row.get(col, ""))))

    def _on_done(self, result: dict) -> None:
        self._set_busy(False)
        self.bar_progress.setRange(0, 100)
        self.bar_progress.setValue(100)
        text = (f"完成：成功 {result.get('ok', 0)}，失败 {result.get('fail', 0)}"
                + (f" → {result['out']}" if result.get("out") else ""))
        if result.get("stopped"):
            text += "（用户中途停止）"
        if result.get("fatal"):
            text += "（因凭据问题提前中止）"
        self.log_line(text)
        self.lbl_stat.setText("完成")
        self.worker = None
        if result.get("fatal"):
            QMessageBox.warning(
                self, "已提前中止",
                "这批订单的 cookie 不对，后面的都没发请求。\n\n"
                f"{result['fatal']}")

    def _on_failed(self, message: str) -> None:
        self._set_busy(False)
        self.bar_progress.setRange(0, 100)
        self.bar_progress.setValue(0)
        self.lbl_stat.setText("失败")
        self.log_line(f"抓取失败：{message}")
        QMessageBox.critical(self, "抓取失败", message)
        self.worker = None

    def _stop(self) -> None:
        if self.worker is not None and self.worker.isRunning():
            self.worker.stop()
            self.btn_stop.setEnabled(False)
            self.log_line("已发送停止信号，等待当前请求返回…")

    # -- 其它 ---------------------------------------------------------------
    def _self_check(self) -> None:
        lines = [
            f"Python    : {sys.version.split()[0]}",
            f"工作目录  : {BASE_DIR}",
            f"订单 Excel: {self.ed_in.text().strip() or '(未选)'}",
            f"客户 Excel: {self.ed_out.text().strip()}",
            f"接口主机  : {self.ed_api.text().strip()}",
            f"seller_id : {self.ed_seller.text().strip() or '(从 Cookie 取)'}",
            f"Cookie    : {len(self.cookies)} 条",
            f"客户列数  : {len(api.CUSTOMER_COLUMNS)}",
        ]
        QMessageBox.information(self, "环境自检", "\n".join(lines))
        for line in lines:
            self.log_line(line)

    def closeEvent(self, event) -> None:  # noqa: N802
        if self.worker is not None and self.worker.isRunning():
            self.worker.stop()
            self.worker.wait(5000)
        event.accept()


# ----------------------------------------------------------------------------
# 入口
# ----------------------------------------------------------------------------
def main(argv: list[str] | None = None) -> int:
    argv = list(sys.argv if argv is None else argv)
    if "--check" in argv:
        app = QApplication.instance() or QApplication(argv)
        win = MainWindow()
        win.show()
        if len(argv) > argv.index("--check") + 1:
            win.ed_in.setText(argv[argv.index("--check") + 1])
        app.processEvents()
        print("GUI 构建成功")
        print("客户列   :", len(api.CUSTOMER_COLUMNS))
        print("Cookie   :", len(win.cookies), "条")
        print("接口主机 :", win.ed_api.text().strip())
        print("订单 Excel:", win.ed_in.text().strip() or "(未选)")
        win.close()
        return 0

    app = QApplication(argv)
    win = MainWindow()
    win.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
