from __future__ import annotations

import os
import sys
import traceback
from pathlib import Path

from PySide6.QtCore import QEasingCurve, QPropertyAnimation, Qt, QSize, QRect, QEvent
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

from app.ui import resource_rc  # noqa: F401
from app.ui.main_window_ui import Ui_MainWindow

# ВАЖНО: страницы не импортируем при старте программы.
# Импорт тяжелых страниц (pandas/openpyxl/win32com и т.д.) делаем только при первом открытии раздела.
import importlib


PAGE_STYLESHEET = ""


def global_exception_handler(exc_type, exc_value, exc_traceback):
    traceback.print_exception(exc_type, exc_value, exc_traceback)

    print("\n" + "=" * 80)
    print("APPLICATION CRASHED")
    print("=" * 80)

    input("\nPress Enter to close...")
    sys.exit(1)


sys.excepthook = global_exception_handler

def lazy_page(module_name: str, class_name: str):
    def factory():
        module = importlib.import_module(module_name)
        page_class = getattr(module, class_name)
        return page_class()

    return factory


# alembic revision --autogenerate -m "changed SupplierPriceCalculation"
# alembic upgrade head
# pyside6-rcc -o resource_rc.py resource.qrc


WINDOW_SIZE = 0


class PlaceholderPage(QWidget):
    def __init__(self, title: str):
        super().__init__()
        layout = QVBoxLayout(self)
        label = QLabel(f'Раздел "{title}" пока не реализован')
        label.setAlignment(Qt.AlignCenter)
        layout.addWidget(label)


class CompactComboDelegate(QStyledItemDelegate):
    def sizeHint(self, option, index):
        size = super().sizeHint(option, index)
        return QSize(size.width(), 18)


class MyWindow(QMainWindow):
    def __init__(self):
        super(MyWindow, self).__init__()

        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        self.setWindowFlag(Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)

        self.shadow = QGraphicsDropShadowEffect(self)
        self.shadow.setBlurRadius(17)
        self.shadow.setXOffset(0)
        self.shadow.setYOffset(0)
        self.shadow.setColor(QColor(0, 0, 0, 150))
        self.ui.centralwidget.setGraphicsEffect(self.shadow)

        self.ui.minimizeAppBtn.clicked.connect(lambda: self.showMinimized())
        self.ui.maximizeRestoreAppBtn.clicked.connect(lambda: self.maximize_restore())
        self.ui.closeAppBtn.clicked.connect(lambda: self.close())
        self.ui.toggleButton.clicked.connect(lambda: self.toggleMenu())

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
        self.ui.frame_size_grip.setGraphicsEffect(self.shadow)

        self.home_btn = self.ui.btn_Home
        self.btn_product = self.ui.btn_Products
        self.btn_articles = self.ui.btn_Articles
        self.btn_supplier = self.ui.btn_Supplier
        self.btn_exchange_rates = self.ui.btn_ExchangeRates
        self.btn_fixed_costs = self.ui.btn_FixedCosts
        self.btn_marking_rates = self.ui.btn_MarkingRates
        self.btn_pack_types = self.ui.btn_PackTypes
        self.btn_price_history = self.ui.btn_PriceHistory
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
            self.btn_supplier: lazy_page("app.page_functions.suppliers_page", "SuppliersPage"),
            self.btn_exchange_rates: lazy_page("app.page_functions.exchange_rates_page", "ExchangeRatesPage"),
            self.btn_fixed_costs: lazy_page("app.page_functions.fixed_costs_page", "FixedCostsPage"),
            self.btn_marking_rates: lazy_page("app.page_functions.marking_rates_page", "MarkingRatesPage"),
            self.btn_pack_types: lazy_page("app.page_functions.pack_types_page", "PackTypesPage"),
            self.btn_price_history: lazy_page("app.page_functions.price_history_page", "PriceHistoryPage"),
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

    app = QApplication(sys.argv)

    style_path = Path(__file__).resolve().parent / "app" / "ui" / "styles" / "app_styles.qss"
    PAGE_STYLESHEET = style_path.read_text(encoding="utf-8") if style_path.exists() else ""

    window = MyWindow()
    window.show()

    sys.exit(app.exec())