from __future__ import annotations

import os
import logging
import sys
import traceback
from pathlib import Path

from config import BASE_DIR
from app.logging_config import (
    flush_logs,
    install_thread_exception_logging,
    log_startup_stage,
    setup_logging,
    shutdown_native_crash_capture,
)

LOG_PATH = setup_logging(BASE_DIR)
logger = logging.getLogger(__name__)
install_thread_exception_logging(logger)
log_startup_stage(logger, "before PySide6 imports")

from PySide6.QtCore import (
    QEasingCurve,
    QPropertyAnimation,
    Qt,
    QSize,
    QRect,
    QEvent,
    QtMsgType,
    qInstallMessageHandler,
)
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QApplication,
    QGraphicsDropShadowEffect,
    QLabel,
    QMainWindow,
    QSizeGrip,
    QVBoxLayout,
    QWidget,
    QComboBox,
    QListView,
    QStyledItemDelegate,
)

log_startup_stage(logger, "PySide6 imports complete")


def _qt_message_handler(message_type, context, message):
    """Mirror Qt diagnostics into app.log instead of losing them on console."""
    level = logging.INFO
    if message_type == QtMsgType.QtWarningMsg:
        level = logging.WARNING
    elif message_type in (QtMsgType.QtCriticalMsg, QtMsgType.QtFatalMsg):
        level = logging.CRITICAL

    location = ""
    if context is not None:
        file_name = getattr(context, "file", None)
        line = getattr(context, "line", 0)
        function = getattr(context, "function", None)
        pieces = []
        if file_name:
            pieces.append(str(file_name))
        if line:
            pieces.append(str(line))
        if function:
            pieces.append(str(function))
        if pieces:
            location = " | " + ":".join(pieces)

    logger.log(level, "QT | %s%s", message, location)
    flush_logs()


qInstallMessageHandler(_qt_message_handler)
log_startup_stage(logger, "Qt message handler installed")

from app.ui import resource_rc  # noqa: F401
from app.ui.main_window_ui import Ui_MainWindow
from app.ui.table_style import apply_global_table_display_rules
from app.ui.table_scale import (
    DEFAULT_TABLE_SCALE,
    MAX_TABLE_SCALE,
    MIN_TABLE_SCALE,
    get_table_scale_manager,
    initialise_table_scale_manager,
    shutdown_table_scale_manager,
)

# ВАЖНО: страницы не импортируем при старте программы.
# Импорт тяжелых страниц (pandas/openpyxl/win32com и т.д.) делаем только при первом открытии раздела.
import importlib


PAGE_STYLESHEET = ""


def global_exception_handler(exc_type, exc_value, exc_traceback):
    traceback.print_exception(exc_type, exc_value, exc_traceback)
    logger.critical(
        "Необработанное исключение, приложение будет закрыто. Лог: %s",
        LOG_PATH,
        exc_info=(exc_type, exc_value, exc_traceback),
    )
    flush_logs()

    print("\n" + "=" * 80)
    print("APPLICATION CRASHED")
    print(f"Подробности сохранены в: {LOG_PATH}")
    print("=" * 80)

    input("\nPress Enter to close...")
    sys.exit(1)


sys.excepthook = global_exception_handler

_QT_SHUTDOWN_PREPARED = False


def _prepare_qt_shutdown() -> None:
    """Remove Python callbacks from Qt before native object destruction starts."""
    global _QT_SHUTDOWN_PREPARED
    if _QT_SHUTDOWN_PREPARED:
        return
    _QT_SHUTDOWN_PREPARED = True

    logger.info("SHUTDOWN | preparing Qt teardown")
    flush_logs()

    # The table scale manager is a long-lived Python QObject installed as an
    # event filter on many child widgets.  Detach it while those widgets and
    # QApplication are still alive.
    try:
        shutdown_table_scale_manager()
    except Exception:
        logger.exception("SHUTDOWN | failed to stop table scale manager")
        flush_logs()

    # qInstallMessageHandler stores a process-global callback in native Qt.
    # Never leave a Python callback installed while Python/Qt are finalising.
    try:
        qInstallMessageHandler(None)
    except Exception:
        logger.exception("SHUTDOWN | failed to remove Qt message handler")

    logger.info("SHUTDOWN | Python Qt callbacks detached")
    flush_logs()


def lazy_page(module_name: str, class_name: str):
    def factory():
        module = importlib.import_module(module_name)
        page_class = getattr(module, class_name)
        page = page_class()
        from app.utils.product_selection_filter import attach_product_selection_filter
        attach_product_selection_filter(page)
        return page

    return factory


# alembic revision --autogenerate -m "changed SupplierPriceCalculation"
# alembic upgrade head
# pyside6-rcc -o resource_rc.py resource.qrc


WINDOW_SIZE = 0


class PlaceholderPage(QWidget):
    def __init__(self, title: str):
        super().__init__()
        layout = QVBoxLayout(self)
        # label = QLabel(f'Раздел "{title}" пока не реализован')
        # label.setAlignment(Qt.AlignCenter)
        # layout.addWidget(label)


class CompactComboDelegate(QStyledItemDelegate):
    def sizeHint(self, option, index):
        size = super().sizeHint(option, index)
        return QSize(size.width(), 18)


class MyWindow(QMainWindow):
    def __init__(self):
        super(MyWindow, self).__init__()
        # Do NOT use WA_DeleteOnClose here.  The main window owns many PySide
        # children with Python callbacks/event filters.  Qt deleting that tree
        # during the native close event can race our shutdown cleanup and cause
        # Windows 0xC0000005 before aboutToQuit is emitted.

        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        self.setWindowFlag(Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)

        self.shadow = self._create_shadow_effect()
        self.ui.centralwidget.setGraphicsEffect(self.shadow)

        self.ui.minimizeAppBtn.clicked.connect(lambda: self.showMinimized())
        self.ui.maximizeRestoreAppBtn.clicked.connect(lambda: self.maximize_restore())
        self.ui.closeAppBtn.clicked.connect(lambda: self.close())
        self.ui.toggleButton.clicked.connect(lambda: self.toggleMenu())

        self.table_scale_manager = get_table_scale_manager()
        if self.table_scale_manager is None:
            self.table_scale_manager = initialise_table_scale_manager(QApplication.instance())

        self.ui.tableScaleMinusBtn.clicked.connect(self.table_scale_manager.decrease)
        self.ui.tableScalePlusBtn.clicked.connect(self.table_scale_manager.increase)
        self.table_scale_manager.scale_changed.connect(self._update_table_scale_controls)
        self.ui.tableScaleLabel.installEventFilter(self)
        self._update_table_scale_controls(self.table_scale_manager.scale_percent)

        self.ui.search_widget.mouseMoveEvent = self.MoveWindow

        self.resize_border = 8
        self.resizing = False
        self.resize_edge = None
        self.drag_pos = None
        self.start_geometry = None
        self.oldPos = None

        self.setMouseTracking(True)
        self.ui.centralwidget.setMouseTracking(True)
        self.ui.search_widget.setMouseTracking(True)
        self.ui.menu_widget.setMouseTracking(True)
        self.ui.tabWidget.setMouseTracking(True)

        self.installEventFilter(self)
        self.ui.centralwidget.installEventFilter(self)
        self.ui.search_widget.installEventFilter(self)
        self.ui.menu_widget.installEventFilter(self)
        self.ui.tabWidget.installEventFilter(self)

        self.sizegrip = QSizeGrip(self.ui.frame_size_grip)
        self.sizegrip.setStyleSheet("width: 20px; height: 20px; margin: 0px; padding: 0px;")
        self.sizegrip_shadow = self._create_shadow_effect()
        self.ui.frame_size_grip.setGraphicsEffect(self.sizegrip_shadow)

        self.home_btn = self.ui.btn_Home
        self.btn_product = self.ui.btn_Products
        self.btn_articles = self.ui.btn_Articles
        self.btn_product_mapping = self.ui.btn_ProductMapping
        self.btn_supplier = self.ui.btn_Supplier
        self.btn_exchange_rates = self.ui.btn_ExchangeRates
        self.btn_fixed_costs = self.ui.btn_FixedCosts
        self.btn_marking_rates = self.ui.btn_MarkingRates
        self.btn_pack_types = self.ui.btn_PackTypes
        self.btn_price_history = self.ui.btn_PriceHistory
        self.btn_product_uc3 = self.ui.btn_ProductUc3
        self.btn_product_search = self.ui.btn_ProdSearchDB
        self.btn_supplier_price = self.ui.btn_SupplierPrice
        self.btn_customer_cost = self.ui.btn_CustomerCost
        self.btn_target_price = self.ui.btn_TargetPrice
        self.btn_target_price_report = getattr(self.ui, "btn_TargetPriceReport", None)
        self.btn_customer_cost_report = self.ui.btn_CustCostReport
        self.btn_product_stock = self.ui.btn_Stock
        self.btn_quick_cost_calc = self.ui.btn_QuickCostCalc
        self.btn_price_reports = self.ui.btn_PriceReports
        self.btn_order_planning = self.ui.btn_OrderPlanning

        self.menu_btns_list = {
            self.home_btn: lambda: PlaceholderPage("Главная"),
            self.btn_product: lazy_page("app.page_functions.products_page", "ProductsPage"),
            self.btn_articles: lazy_page("app.page_functions.product_articles_page", "ProductArticlesPage"),
            self.btn_product_mapping: lazy_page("app.page_functions.product_mapping_page", "ProductMappingPage"),
            self.btn_supplier: lazy_page("app.page_functions.suppliers_page", "SuppliersPage"),
            self.btn_exchange_rates: lazy_page("app.page_functions.exchange_rates_page", "ExchangeRatesPage"),
            self.btn_fixed_costs: lazy_page("app.page_functions.fixed_costs_page", "FixedCostsPage"),
            self.btn_marking_rates: lazy_page("app.page_functions.marking_rates_page", "MarkingRatesPage"),
            self.btn_pack_types: lazy_page("app.page_functions.pack_types_page", "PackTypesPage"),
            self.btn_price_history: lazy_page("app.page_functions.price_history_page", "PriceHistoryPage"),
            self.btn_product_uc3: lazy_page("app.page_functions.product_uc3_page", "ProductUc3Page"),
            self.btn_product_search: lazy_page("app.page_functions.product_search_page", "ProductSearchPage"),
            self.btn_supplier_price: lazy_page("app.page_functions.supplier_prices_page", "SupplierPricesPage"),
            self.btn_customer_cost: lazy_page("app.page_functions.customer_costs_page", "CustomerCostsPage"),
            self.btn_target_price: lazy_page("app.page_functions.target_prices_page", "TargetPricesPage"),
            self.btn_customer_cost_report: lazy_page("app.page_functions.customer_costs_reports_page", "CustomerCostsReportsPage"),
            self.btn_product_stock: lazy_page("app.page_functions.product_stock_page", "ProductStockPage"),
            self.btn_quick_cost_calc: lazy_page("app.page_functions.quick_cost_calc_page", "QuickCostCalcPage"),
            self.btn_price_reports: lazy_page("app.page_functions.price_reports_page", "PriceReportsPage"),
            self.btn_order_planning: lazy_page("app.page_functions.order_planning_page", "OrderPlanningPage"),
        }


        if self.btn_target_price_report is not None:
            self.menu_btns_list[self.btn_target_price_report] = lazy_page(
                "app.page_functions.target_price_history_page",
                "TargetPriceHistoryPage",
            )

        self.show_home_window()

        self.ui.tabWidget.tabCloseRequested.connect(self.close_tab)
        for button in self.menu_btns_list.keys():
            button.clicked.connect(self.show_selected_window)

    def _update_table_scale_controls(self, value: int):
        self.ui.tableScaleLabel.setText(f"{value}%")
        self.ui.tableScaleMinusBtn.setEnabled(value > MIN_TABLE_SCALE)
        self.ui.tableScalePlusBtn.setEnabled(value < MAX_TABLE_SCALE)
        self.ui.tableScaleLabel.setToolTip(
            f"Масштаб таблиц: {value}%. "
            f"Двойной щелчок возвращает {DEFAULT_TABLE_SCALE}%"
        )

    def _create_shadow_effect(self):
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(17)
        shadow.setXOffset(0)
        shadow.setYOffset(0)
        shadow.setColor(QColor(0, 0, 0, 150))
        return shadow

    def show_home_window(self):
        result = self.open_tab_flag(self.home_btn.text())
        self.set_btn_checked(self.home_btn)

        if result[0]:
            self.ui.tabWidget.setCurrentIndex(result[1])
        else:
            title = self.home_btn.text()
            page_factory = self.menu_btns_list[self.home_btn]
            page = page_factory()
            if hasattr(page, "ui") and isinstance(page.ui, QWidget):
                page.ui.setStyleSheet("")
                if PAGE_STYLESHEET:
                    page.ui.setStyleSheet(PAGE_STYLESHEET)
            else:
                if PAGE_STYLESHEET:
                    page.setStyleSheet(PAGE_STYLESHEET)
            self.setup_all_compact_comboboxes(page)
            apply_global_table_display_rules(page)
            self.table_scale_manager.register_tables(page)
            cur_index = self.ui.tabWidget.addTab(page, title)
            self.ui.tabWidget.setCurrentIndex(cur_index)
            self.ui.tabWidget.setVisible(True)

    def show_selected_window(self):
        button = self.sender()
        result = self.open_tab_flag(button.text())
        self.set_btn_checked(button)

        if result[0]:
            self.ui.tabWidget.setCurrentIndex(result[1])
        else:
            title = button.text()
            page_factory = self.menu_btns_list[button]
            page = page_factory()
            if hasattr(page, "ui") and isinstance(page.ui, QWidget):
                page.ui.setStyleSheet("")
                if PAGE_STYLESHEET:
                    page.ui.setStyleSheet(PAGE_STYLESHEET)
            else:
                if PAGE_STYLESHEET:
                    page.setStyleSheet(PAGE_STYLESHEET)
            self.setup_all_compact_comboboxes(page)
            apply_global_table_display_rules(page)
            self.table_scale_manager.register_tables(page)
            cur_index = self.ui.tabWidget.addTab(page, title)
            self.ui.tabWidget.setCurrentIndex(cur_index)
            self.ui.tabWidget.setVisible(True)

    def close_tab(self, index):
        self.ui.tabWidget.removeTab(index)
        if self.ui.tabWidget.count() == 0:
            self.ui.toolBox.setCurrentIndex(0)
            self.show_home_window()

    def open_tab_flag(self, tab):
        open_tab_count = self.ui.tabWidget.count()
        for i in range(open_tab_count):
            tab_name = self.ui.tabWidget.tabText(i)
            if tab_name == tab:
                return True, i
        return False, -1

    def set_btn_checked(self, btn):
        for button in self.menu_btns_list.keys():
            button.setChecked(button == btn)

    def maximize_restore(self):
        global WINDOW_SIZE
        status = WINDOW_SIZE

        if status == 0:
            WINDOW_SIZE = 1
            self.showMaximized()
            self.ui.appMargins.setContentsMargins(0, 0, 0, 0)
            self.ui.maximizeRestoreAppBtn.setToolTip("Restore")
        else:
            WINDOW_SIZE = 0
            self.showNormal()
            self.resize(self.width() + 1, self.height() + 1)
            self.ui.appMargins.setContentsMargins(4, 4, 4, 4)
            self.ui.maximizeRestoreAppBtn.setToolTip("Maximize")

    def get_resize_edge(self, pos):
        x = pos.x()
        y = pos.y()
        w = self.width()
        h = self.height()
        m = self.resize_border

        left = x <= m
        right = x >= w - m
        top = y <= m
        bottom = y >= h - m

        if top and left:
            return "top_left"
        if top and right:
            return "top_right"
        if bottom and left:
            return "bottom_left"
        if bottom and right:
            return "bottom_right"
        if left:
            return "left"
        if right:
            return "right"
        if top:
            return "top"
        if bottom:
            return "bottom"
        return None

    def update_cursor(self, edge):
        if edge in ("left", "right"):
            self.setCursor(Qt.SizeHorCursor)
        elif edge in ("top", "bottom"):
            self.setCursor(Qt.SizeVerCursor)
        elif edge in ("top_left", "bottom_right"):
            self.setCursor(Qt.SizeFDiagCursor)
        elif edge in ("top_right", "bottom_left"):
            self.setCursor(Qt.SizeBDiagCursor)
        else:
            self.setCursor(Qt.ArrowCursor)

    def resize_window(self, global_pos):
        if not self.resizing or not self.resize_edge or not self.start_geometry or not self.drag_pos:
            return

        dx = int(global_pos.x() - self.drag_pos.x())
        dy = int(global_pos.y() - self.drag_pos.y())

        geo = QRect(self.start_geometry)

        min_w = max(self.minimumWidth(), 200)
        min_h = max(self.minimumHeight(), 150)

        if "left" in self.resize_edge:
            new_left = geo.left() + dx
            if geo.right() - new_left + 1 >= min_w:
                geo.setLeft(new_left)

        if "right" in self.resize_edge:
            new_width = geo.width() + dx
            if new_width >= min_w:
                geo.setWidth(new_width)

        if "top" in self.resize_edge:
            new_top = geo.top() + dy
            if geo.bottom() - new_top + 1 >= min_h:
                geo.setTop(new_top)

        if "bottom" in self.resize_edge:
            new_height = geo.height() + dy
            if new_height >= min_h:
                geo.setHeight(new_height)

        self.setGeometry(geo)

    def eventFilter(self, obj, event):
        if obj is self.ui.tableScaleLabel and event.type() == QEvent.MouseButtonDblClick:
            self.table_scale_manager.reset()
            return True

        if event.type() == QEvent.MouseMove and not self.isMaximized():
            try:
                global_pos = event.globalPosition().toPoint()
            except AttributeError:
                return super().eventFilter(obj, event)

            local_pos = self.mapFromGlobal(global_pos)

            if self.resizing:
                self.resize_window(global_pos)
                return True

            edge = self.get_resize_edge(local_pos)
            self.update_cursor(edge)

        elif event.type() == QEvent.Leave and not self.resizing:
            self.setCursor(Qt.ArrowCursor)

        return super().eventFilter(obj, event)

    def closeEvent(self, event):
        # Detach Python event filters BEFORE Qt starts processing the native
        # close path.  These filters point back to this Python QMainWindow and
        # must not receive events while its child widgets are being torn down.
        logger.info("SHUTDOWN | main window closeEvent begin")
        flush_logs()
        for target in (
            self,
            self.ui.centralwidget,
            self.ui.search_widget,
            self.ui.menu_widget,
            self.ui.tabWidget,
            self.ui.tableScaleLabel,
        ):
            try:
                target.removeEventFilter(self)
            except RuntimeError:
                pass

        # Detach the global table event filters/timers and Qt's process-global
        # Python message handler while all widgets are still valid.
        _prepare_qt_shutdown()
        logger.info("SHUTDOWN | main window closeEvent cleanup complete")
        flush_logs()
        super().closeEvent(event)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton and not self.isMaximized():
            edge = self.get_resize_edge(event.position().toPoint())
            if edge:
                self.resizing = True
                self.resize_edge = edge
                self.drag_pos = event.globalPosition().toPoint()
                self.start_geometry = self.geometry()
                event.accept()
                return

        if event.button() == Qt.LeftButton:
            self.oldPos = event.globalPosition().toPoint()

        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self.isMaximized():
            self.setCursor(Qt.ArrowCursor)
            super().mouseMoveEvent(event)
            return

        if self.resizing:
            self.resize_window(event.globalPosition().toPoint())
            event.accept()
            return

        edge = self.get_resize_edge(event.position().toPoint())
        self.update_cursor(edge)

        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        self.resizing = False
        self.resize_edge = None
        self.drag_pos = None
        self.start_geometry = None
        self.setCursor(Qt.ArrowCursor)
        super().mouseReleaseEvent(event)

    def MoveWindow(self, event):
        if self.isMaximized() or self.resizing:
            return

        if event.buttons() & Qt.LeftButton and self.oldPos is not None:
            delta = event.globalPosition().toPoint() - self.oldPos
            self.move(self.pos() + delta)
            self.oldPos = event.globalPosition().toPoint()
            event.accept()

    def toggleMenu(self):
        width = self.ui.menu_widget.width()
        max_extend = 220
        standard = 60

        width_extended = max_extend if width == 60 else standard

        self.animation = QPropertyAnimation(self.ui.menu_widget, b"minimumWidth")
        self.animation.setDuration(500)
        self.animation.setStartValue(width)
        self.animation.setEndValue(width_extended)
        self.animation.setEasingCurve(QEasingCurve.InOutQuart)
        self.animation.start()

    def setup_all_compact_comboboxes(self, root_widget):
        for combo in root_widget.findChildren(QComboBox):
            view = QListView()
            view.setSpacing(0)
            view.setItemDelegate(CompactComboDelegate(view))
            combo.setView(view)


if __name__ == "__main__":
    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
    os.environ["QT_QPA_PLATFORM"] = "windows"

    log_startup_stage(logger, "creating QApplication")
    app = QApplication(sys.argv)
    app.aboutToQuit.connect(_prepare_qt_shutdown)
    log_startup_stage(logger, "QApplication created")

    log_startup_stage(logger, "initialising table scale manager")
    initialise_table_scale_manager(app)
    log_startup_stage(logger, "table scale manager initialised")

    style_path = Path(__file__).resolve().parent / "app" / "ui" / "styles" / "app_styles.qss"
    PAGE_STYLESHEET = style_path.read_text(encoding="utf-8") if style_path.exists() else ""
    log_startup_stage(logger, "stylesheet loaded")

    log_startup_stage(logger, "creating main window")
    window = MyWindow()
    log_startup_stage(logger, "main window created")

    window.show()
    log_startup_stage(logger, "main window shown; entering event loop")

    exit_code = app.exec()
    logger.info("Application event loop finished, exit code: %s", exit_code)
    flush_logs()

    # aboutToQuit normally performed this while the event loop was active.
    # Keep a fallback for unusual exits where that signal was not delivered.
    _prepare_qt_shutdown()

    # Drop Python references explicitly instead of leaving PySide/Shiboken
    # wrappers to be finalised in arbitrary order at interpreter shutdown.
    try:
        del window
    except NameError:
        pass

    # Destroy QApplication in a controlled point while native crash capture is
    # still active. If Qt itself fails here, faulthandler will still record it.
    del app

    logger.info("SHUTDOWN | QApplication released cleanly")
    flush_logs()

    # At this point Qt has already shut down cleanly. On Windows/PySide the
    # remaining interpreter finalisation can still destroy global Shiboken/Qt
    # wrappers in an unsafe order and cause 0xC0000005 after all application
    # shutdown checkpoints have completed. Close logging explicitly and leave
    # the process without running Python's remaining module/global destructors.
    shutdown_native_crash_capture()
    try:
        logging.shutdown()
    except Exception:
        pass
    try:
        sys.stdout.flush()
        sys.stderr.flush()
    except Exception:
        pass
    os._exit(int(exit_code))
