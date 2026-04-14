from __future__ import annotations

import uuid
from pathlib import Path
from sqlalchemy.orm import Session

from app.exports.product_stock_import_export import ProductStockImportExport
from app.imports.is_importer import ISImporter
from app.imports.stock_importer import StockImporter
from app.imports.supplier_orders_importer import SupplierOrdersImporter
from app.services.product_stock_import import ProductStockImportService


class ProductStockImportRun:
    def __init__(self, session: Session):
        self.session = session
        self.stock_importer = StockImporter()
        self.supplier_orders_importer = SupplierOrdersImporter()
        self.is_importer = ISImporter()
        self.service = ProductStockImportService(session)
        self.exporter = ProductStockImportExport(session)

    def start_batch(self) -> str:
        return str(uuid.uuid4())

    def import_stock(self, file_path: str | Path, imported_by: str, batch_id: str | None = None) -> dict:
        batch_id = batch_id or self.start_batch()
        self.service.cleanup_old_temp_stock()
        rows = self.stock_importer.read_excel(file_path)
        imported_count = self.service.import_stock_rows(rows=rows, batch_id=batch_id, imported_by=imported_by, replace_existing=True)
        matched_count = self.service.automatch_stock_rows(batch_id=batch_id, imported_by=imported_by)
        return {"batch_id": batch_id, "imported_count": imported_count, "matched_count": matched_count}

    def save_stock(self, batch_id: str, imported_by: str) -> dict:
        self.service.validate_new_stock_products_before_save(batch_id, imported_by)
        created_products_count = self.service.create_stock_products_from_temp(batch_id, imported_by)
        product_articles_count = self.service.create_or_update_stock_product_articles(batch_id, imported_by)
        saved_count = self.service.save_stock_to_product_stock(batch_id, imported_by)
        self.service.delete_stock_rows(batch_id, imported_by)
        return {
            "batch_id": batch_id,
            "created_products_count": created_products_count,
            "product_articles_count": product_articles_count,
            "saved_count": saved_count,
        }

    def import_supplier_orders(self, file_path: str | Path, imported_by: str, batch_id: str | None = None) -> dict:
        batch_id = batch_id or self.start_batch()
        self.service.cleanup_old_temp_supplier_orders()
        rows = self.supplier_orders_importer.read_excel(file_path)
        imported_count = self.service.import_supplier_orders_rows(rows=rows, batch_id=batch_id, imported_by=imported_by, replace_existing=True)
        matched_count = self.service.automatch_supplier_orders_rows(batch_id=batch_id, imported_by=imported_by)
        return {"batch_id": batch_id, "imported_count": imported_count, "matched_count": matched_count}

    def save_supplier_orders(self, batch_id: str, imported_by: str) -> dict:
        self.service.validate_new_supplier_orders_products_before_save(batch_id, imported_by)
        created_products_count = self.service.create_supplier_orders_products_from_temp(batch_id, imported_by)
        product_articles_count = self.service.create_or_update_supplier_orders_product_articles(batch_id, imported_by)
        saved_count = self.service.save_supplier_orders_to_product_stock(batch_id, imported_by)
        self.service.delete_supplier_orders_rows(batch_id, imported_by)
        return {
            "batch_id": batch_id,
            "created_products_count": created_products_count,
            "product_articles_count": product_articles_count,
            "saved_count": saved_count,
        }

    def import_is(self, file_path: str | Path, imported_by: str, batch_id: str | None = None) -> dict:
        batch_id = batch_id or self.start_batch()
        self.service.cleanup_old_temp_is()
        rows = self.is_importer.read_excel(file_path)
        imported_count = self.service.import_is_rows(rows=rows, batch_id=batch_id, imported_by=imported_by, replace_existing=True)
        matched_count = self.service.automatch_is_rows(batch_id=batch_id, imported_by=imported_by)
        return {"batch_id": batch_id, "imported_count": imported_count, "matched_count": matched_count}

    def save_is(self, batch_id: str, imported_by: str) -> dict:
        self.service.validate_new_is_products_before_save(batch_id, imported_by)
        created_products_count = self.service.create_is_products_from_temp(batch_id, imported_by)
        product_articles_count = self.service.create_or_update_is_product_articles(batch_id, imported_by)
        saved_count = self.service.save_is_to_product_stock(batch_id, imported_by)
        self.service.delete_is_rows(batch_id, imported_by)
        return {
            "batch_id": batch_id,
            "created_products_count": created_products_count,
            "product_articles_count": product_articles_count,
            "saved_count": saved_count,
        }

    def export_stock_product_issues(self, batch_id: str, imported_by: str, output_path: str | Path):
        return self.exporter.export_stock_product_issues(batch_id, imported_by, output_path)

    def export_stock_lpc_warnings(self, batch_id: str, imported_by: str, output_path: str | Path):
        return self.exporter.export_stock_lpc_warnings(batch_id, imported_by, output_path)

    def export_supplier_orders_product_issues(self, batch_id: str, imported_by: str, output_path: str | Path):
        return self.exporter.export_supplier_orders_product_issues(batch_id, imported_by, output_path)

    def export_is_product_issues(self, batch_id: str, imported_by: str, output_path: str | Path):
        return self.exporter.export_is_product_issues(batch_id, imported_by, output_path)