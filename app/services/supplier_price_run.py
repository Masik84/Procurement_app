from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional

from sqlalchemy.orm import Session

from app.imports.supplier_price_importer import SupplierPriceImporter
from app.services.supplier import SupplierService, SupplierUpsertData
from app.services.supplier_price_import import SupplierPriceImportService


@dataclass(slots=True)
class SupplierPriceImportResult:
    supplier_id: int
    supplier_name: str
    batch_id: str
    imported_by: str
    import_file: str
    imported_count: int
    matched_count: int
    created_products_count: int
    product_articles_count: int
    filled_prices_count: int
    saved_prices_count: int
    saved_calculations_count: int


class SupplierPriceImportRun:
    def __init__(self, session: Session) -> None:
        self.session = session
        self.supplier_service = SupplierService(session)
        self.import_service = SupplierPriceImportService(session)
        self.importer = SupplierPriceImporter()

    def run_from_excel(
        self,
        *,
        file_path: str | Path,
        imported_by: str,
        supplier_data: SupplierUpsertData,
        supplier_id: Optional[int] = None,
        import_date: Optional[datetime] = None,
        save_exchange_rate: bool = False,
        explicit_fx_rate: Optional[float] = None,
    ) -> SupplierPriceImportResult:
        supplier = self.supplier_service.ensure_supplier(supplier_id=supplier_id, data=supplier_data)
        currency_code = supplier.base_currency
        fx_rate: Optional[float] = None

        if explicit_fx_rate is not None:
            fx_rate = float(explicit_fx_rate)
            if save_exchange_rate:
                self.supplier_service.save_exchange_rate(currency_code, fx_rate)
        else:
            fx_rate = self.supplier_service.get_rate_to_rub(currency_code)

        if fx_rate is None or float(fx_rate) == 0:
            raise ValueError(f"Для валюты '{currency_code}' не найден корректный курс rate_to_rub.")

        rows = self.importer.read_excel(file_path)
        batch_id = self.import_service.start_batch()

        stats = self.import_service.run_full_import_pipeline(
            supplier_id=supplier.id,
            batch_id=batch_id,
            imported_by=imported_by,
            rows=rows,
            currency_code=currency_code,
            fx_rate=fx_rate,
            import_date=import_date,
            replace_existing_batch_rows=True,
        )

        return SupplierPriceImportResult(
            supplier_id=supplier.id,
            supplier_name=supplier.name,
            batch_id=batch_id,
            imported_by=imported_by,
            import_file=str(Path(file_path)),
            imported_count=stats["imported_count"],
            matched_count=stats["matched_count"],
            created_products_count=stats["created_products_count"],
            product_articles_count=stats["product_articles_count"],
            filled_prices_count=stats["filled_prices_count"],
            saved_prices_count=stats["saved_prices_count"],
            saved_calculations_count=stats["saved_calculations_count"],
        )