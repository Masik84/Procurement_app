from __future__ import annotations

import os
import sys
from pathlib import Path

from PySide6.QtCore import QEasingCurve, QPropertyAnimation, Qt, QSize
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

from app.page_functions.customer_costs_page import CustomerCostsPage
from app.page_functions.exchange_rates_page import ExchangeRatesPage
from app.page_functions.fixed_costs_page import FixedCostsPage
from app.page_functions.marking_rates_page import MarkingRatesPage
from app.page_functions.pack_types_page import PackTypesPage
from app.page_functions.price_history_page import PriceHistoryPage
from app.page_functions.price_reports_page import PriceReportsPage
from app.page_functions.product_articles_page import ProductArticlesPage
from app.page_functions.product_search_page import ProductSearchPage
from app.page_functions.product_stock_page import ProductStockPage
from app.page_functions.products_page import ProductsPage
from app.page_functions.quick_cost_calc_page import QuickCostCalcPage
from app.page_functions.supplier_prices_page import SupplierPricesPage
from app.page_functions.suppliers_page import SuppliersPage

try:
    from app.db.db import Base, engine
except Exception:
    Base = None
    engine = None


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
        # self.ui.appMargins.setContentsMargins(4, 4, 4, 4)
        # self.ui.appMargins.setStyleSheet("background: transparent; border: none;")

        self.ui.minimizeAppBtn.clicked.connect(lambda: self.showMinimized())
        self.ui.maximizeRestoreAppBtn.clicked.connect(lambda: self.maximize_restore())
        self.ui.closeAppBtn.clicked.connect(lambda: self.close())
        self.ui.toggleButton.clicked.connect(lambda: self.toggleMenu())

        self.ui.search_widget.mouseMoveEvent = self.MoveWindow

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
        self.btn_product_stock = self.ui.btn_Stock
        self.btn_quick_cost_calc = self.ui.btn_QuickCostCalc
        self.btn_price_reports = self.ui.btn_PriceReports

        self.menu_btns_list = {
            self.home_btn: lambda: PlaceholderPage("Главная"),
            self.btn_product: ProductsPage,
            self.btn_articles: ProductArticlesPage,
            self.btn_supplier: SuppliersPage,
            self.btn_exchange_rates: ExchangeRatesPage,
            self.btn_fixed_costs: FixedCostsPage,
            self.btn_marking_rates: MarkingRatesPage,
            self.btn_pack_types: PackTypesPage,
            self.btn_price_history: PriceHistoryPage,
            self.btn_product_search: ProductSearchPage,
            self.btn_supplier_price: SupplierPricesPage,
            self.btn_customer_cost: CustomerCostsPage,
            self.btn_product_stock: ProductStockPage,
            self.btn_quick_cost_calc: QuickCostCalcPage,
            self.btn_price_reports: PriceReportsPage,
        }

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

    def mousePressEvent(self, event):
        self.oldPos = self.window().mapFromGlobal(event.globalPosition())

    def MoveWindow(self, event):
        if self.isMaximized() is False:
            delta = self.window().mapFromGlobal(event.globalPosition()) - self.oldPos
            self.move(self.x() + delta.x(), self.y() + delta.y())
            self.oldPos = self.window().mapFromGlobal(event.globalPosition())

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

    # style_path = Path(__file__).resolve().parent / "app" / "ui" / "styles" / "app_styles.qss"
    # if style_path.exists():
    #     app.setStyleSheet(style_path.read_text(encoding="utf-8"))

    style_path = Path(__file__).resolve().parent / "app" / "ui" / "styles" / "app_styles.qss"
    PAGE_STYLESHEET = style_path.read_text(encoding="utf-8") if style_path.exists() else ""
    
    if Base is not None and engine is not None:
        try:
            Base.metadata.create_all(bind=engine)
        except Exception:
            pass

    window = MyWindow()
    window.show()

    sys.exit(app.exec())
