from __future__ import annotations

from pathlib import Path
from openpyxl import Workbook
from sqlalchemy.orm import Session

from app.db.models import TempIsImport, TempStockImport, TempSupplierOrdersImport
from app.utils.excel_export_format import write_openpyxl_dict_sheet


class ProductStockExporter:
    def __init__(self, session: Session):
        self.session = session

    def _write(self, rows: list[dict], output_path: str | Path) -> Path:
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Sheet1"
        write_openpyxl_dict_sheet(worksheet, rows)
        workbook.save(output_path)
        return output_path

    def export_stock_product_issues(self, batch_id: str, imported_by: str, output_path: str | Path) -> Path | None:
        rows = self.session.query(TempStockImport).filter(
            TempStockImport.batch_id == batch_id,
            TempStockImport.imported_by == imported_by,
            TempStockImport.selected_product_id.is_(None),
        ).order_by(TempStockImport.import_row_no.asc()).all()

        out = []
        for row in rows:
            out.append({
                "ImportRowNo": row.import_row_no,
                "SourceArticle": row.source_article,
                "SourceSKU": row.source_sku,
                "SourceProductName": row.source_product_name,
                "Comment": "Не найден SelectedProductID. Требуется сопоставление или создание нового продукта.",
            })

        if not out:
            return None
        return self._write(out, output_path)

    def export_stock_lpc_warnings(self, batch_id: str, imported_by: str, output_path: str | Path) -> Path | None:
        rows = self.session.query(TempStockImport).filter(
            TempStockImport.batch_id == batch_id,
            TempStockImport.imported_by == imported_by,
            TempStockImport.has_lpc_warning.is_(True),
        ).order_by(TempStockImport.import_row_no.asc()).all()

        out = []
        for row in rows:
            out.append({
                "ImportRowNo": row.import_row_no,
                "SourceArticle": row.source_article,
                "SourceSKU": row.source_sku,
                "SourceProductName": row.source_product_name,
                "StockQty": row.stock_qty,
                "MarkdownQty": row.markdown_qty,
                "ReserveQty": row.reserve_qty,
                "ReserveECommQty": getattr(row, "reserve_ecomm_qty", 0),
                "LPC": row.lpc,
                "Comment": "Есть остаток, но LPC в файле пустой или 0. Итоговый LPC рассчитывается по БД продаж.",
            })

        if not out:
            return None
        return self._write(out, output_path)

    def export_supplier_orders_product_issues(self, batch_id: str, imported_by: str, output_path: str | Path) -> Path | None:
        rows = self.session.query(TempSupplierOrdersImport).filter(
            TempSupplierOrdersImport.batch_id == batch_id,
            TempSupplierOrdersImport.imported_by == imported_by,
            TempSupplierOrdersImport.selected_product_id.is_(None),
        ).order_by(TempSupplierOrdersImport.import_row_no.asc()).all()

        out = []
        for row in rows:
            out.append({
                "ImportRowNo": row.import_row_no,
                "SourceArticle": row.source_article,
                "SourceProductName": row.source_product_name,
                "OrderQty": row.order_qty,
                "CoralOrderQty": row.is_order_qty,
                "CoralConfirmedQty": row.is_confirmed_order_qty,
                "Comment": "Не найден SelectedProductID. Требуется сопоставление или создание нового продукта.",
            })

        if not out:
            return None
        return self._write(out, output_path)

    def export_is_product_issues(self, batch_id: str, imported_by: str, output_path: str | Path) -> Path | None:
        rows = self.session.query(TempIsImport).filter(
            TempIsImport.batch_id == batch_id,
            TempIsImport.imported_by == imported_by,
            TempIsImport.selected_product_id.is_(None),
        ).order_by(TempIsImport.import_row_no.asc()).all()

        out = []
        for row in rows:
            out.append({
                "ImportRowNo": row.import_row_no,
                "SourceArticle": row.source_article,
                "SourceProductName": row.source_product_name,
                "ConfirmedQty": row.confirmed_qty,
                "RemainsQty": row.remains_qty,
                "StockQty": row.stock_qty,
                "Comment": "Не найден SelectedProductID. Требуется сопоставление или создание нового продукта.",
            })

        if not out:
            return None
        return self._write(out, output_path)
# Backward-compatible alias.
ProductStockImportExport = ProductStockExporter
