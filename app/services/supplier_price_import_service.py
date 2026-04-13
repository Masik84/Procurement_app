from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Iterable, Optional

from sqlalchemy.orm import Session

from app.db.models import (
    CurrentSupplierPrice,
    PriceHistory,
    SupplierPriceCalculation,
    TempPriceImport,
)
from app.services.cost_calculation_service import CostCalculationService
from app.services.product_matching_service import ProductMatchingService
from app.services.supplier_service import SupplierService
from app.utils.batch import generate_import_batch_id


@dataclass(slots=True)
class SupplierPriceImportRowData:
    supplier_article: Optional[str] = None
    product_name: Optional[str] = None
    price: Optional[float] = None
    price_pack: Optional[float] = None
    import_row_no: Optional[int] = None


class SupplierPriceImportService:
    def __init__(self, session: Session) -> None:
        self.session = session
        self.product_matching_service = ProductMatchingService(session)
        self.supplier_service = SupplierService(session)
        self.cost_calculation_service = CostCalculationService(session)

    # =========================================================
    # Batch / temp helpers
    # =========================================================

    def start_batch(self) -> str:
        return generate_import_batch_id()

    def delete_temp_rows(self, batch_id: str, imported_by: str) -> None:
        (
            self.session.query(TempPriceImport)
            .filter(
                TempPriceImport.batch_id == batch_id,
                TempPriceImport.imported_by == imported_by,
            )
            .delete(synchronize_session=False)
        )
        self.session.flush()

    def get_temp_rows(self, batch_id: str, imported_by: str) -> list[TempPriceImport]:
        return (
            self.session.query(TempPriceImport)
            .filter(
                TempPriceImport.batch_id == batch_id,
                TempPriceImport.imported_by == imported_by,
            )
            .order_by(TempPriceImport.import_row_no.asc(), TempPriceImport.id.asc())
            .all()
        )

    # =========================================================
    # Import into temp
    # =========================================================

    def add_temp_row(
        self,
        *,
        supplier_id: int,
        batch_id: str,
        imported_by: str,
        import_date: datetime,
        row_data: SupplierPriceImportRowData,
    ) -> TempPriceImport:
        row = TempPriceImport(
            supplier_article=row_data.supplier_article,
            product_name=row_data.product_name,
            price=row_data.price,
            price_pack=row_data.price_pack,
            supplier_id=supplier_id,
            import_date=import_date,
            batch_id=batch_id,
            imported_by=imported_by,
            import_row_no=row_data.import_row_no,
            selected_product_id=None,
            new_product_name=None,
            new_brand=None,
            new_pack=None,
            new_is_excise=None,
        )
        self.session.add(row)
        self.session.flush()
        return row

    def import_rows_to_temp(
        self,
        *,
        supplier_id: int,
        batch_id: str,
        imported_by: str,
        rows: Iterable[SupplierPriceImportRowData],
        import_date: Optional[datetime] = None,
        replace_existing_batch_rows: bool = True,
    ) -> list[TempPriceImport]:
        dt = import_date or datetime.now()

        if replace_existing_batch_rows:
            self.delete_temp_rows(batch_id, imported_by)

        created_rows: list[TempPriceImport] = []

        for row_data in rows:
            created_row = self.add_temp_row(
                supplier_id=supplier_id,
                batch_id=batch_id,
                imported_by=imported_by,
                import_date=dt,
                row_data=row_data,
            )
            created_rows.append(created_row)

        return created_rows

    # =========================================================
    # Matching
    # =========================================================

    def automatch_temp_rows(self, batch_id: str, imported_by: str) -> int:
        rows = (
            self.session.query(TempPriceImport)
            .filter(
                TempPriceImport.batch_id == batch_id,
                TempPriceImport.imported_by == imported_by,
                TempPriceImport.selected_product_id.is_(None),
            )
            .order_by(TempPriceImport.import_row_no.asc(), TempPriceImport.id.asc())
            .all()
        )

        matched_count = 0

        for row in rows:
            product = self.product_matching_service.find_price_import_product(
                supplier_article=row.supplier_article,
                supplier_product_name=row.product_name,
            )
            if product is not None:
                row.selected_product_id = product.id
                matched_count += 1

        self.session.flush()
        return matched_count

    # =========================================================
    # New products from temp
    # =========================================================

    def validate_new_products_before_save(self, batch_id: str, imported_by: str) -> None:
        rows = (
            self.session.query(TempPriceImport)
            .filter(
                TempPriceImport.batch_id == batch_id,
                TempPriceImport.imported_by == imported_by,
                TempPriceImport.selected_product_id.is_(None),
                TempPriceImport.new_product_name.isnot(None),
            )
            .all()
        )

        for row in rows:
            if row.new_product_name is None or not str(row.new_product_name).strip():
                continue

            if row.new_is_excise is None:
                raise ValueError(
                    f"Для нового продукта '{row.new_product_name}' не заполнено поле new_is_excise."
                )

            if row.new_brand is None or not str(row.new_brand).strip():
                raise ValueError(
                    f"Для нового продукта '{row.new_product_name}' не заполнен new_brand."
                )

            if row.new_pack is None:
                raise ValueError(
                    f"Для нового продукта '{row.new_product_name}' не заполнен new_pack."
                )

    def create_products_from_temp(self, batch_id: str, imported_by: str) -> int:
        rows = (
            self.session.query(TempPriceImport)
            .filter(
                TempPriceImport.batch_id == batch_id,
                TempPriceImport.imported_by == imported_by,
                TempPriceImport.selected_product_id.is_(None),
                TempPriceImport.new_product_name.isnot(None),
                TempPriceImport.new_brand.isnot(None),
                TempPriceImport.new_pack.isnot(None),
                TempPriceImport.new_is_excise.isnot(None),
            )
            .order_by(TempPriceImport.import_row_no.asc(), TempPriceImport.id.asc())
            .all()
        )

        created_count = 0

        for row in rows:
            product = self.product_matching_service.get_or_create_product(
                name=row.new_product_name,
                brand=row.new_brand,
                pack=row.new_pack,
                is_excise=bool(row.new_is_excise),
            )
            row.selected_product_id = product.id
            created_count += 1

        self.session.flush()
        return created_count

    # =========================================================
    # ProductArticle sync
    # =========================================================

    def create_or_update_product_articles(self, batch_id: str, imported_by: str) -> int:
        rows = (
            self.session.query(TempPriceImport)
            .filter(
                TempPriceImport.batch_id == batch_id,
                TempPriceImport.imported_by == imported_by,
                TempPriceImport.selected_product_id.isnot(None),
            )
            .order_by(TempPriceImport.import_row_no.asc(), TempPriceImport.id.asc())
            .all()
        )

        processed_count = 0

        for row in rows:
            self.product_matching_service.save_product_articles_by_split_articles(
                product_id=row.selected_product_id,
                supplier_article=row.supplier_article,
                supplier_product_name=row.product_name,
            )
            processed_count += 1

        self.session.flush()
        return processed_count

    # =========================================================
    # Price helpers
    # =========================================================

    def fill_price_from_price_pack(self, batch_id: str, imported_by: str) -> int:
        rows = (
            self.session.query(TempPriceImport)
            .filter(
                TempPriceImport.batch_id == batch_id,
                TempPriceImport.imported_by == imported_by,
                TempPriceImport.selected_product_id.isnot(None),
                TempPriceImport.price.is_(None),
                TempPriceImport.price_pack.isnot(None),
            )
            .all()
        )

        filled_count = 0

        for row in rows:
            product = (
                self.session.query(TempPriceImport)
                .filter(TempPriceImport.id == row.id)
                .first()
            )
            if product is None or product.selected_product is None:
                continue

            pack = float(product.selected_product.pack)
            if pack == 0:
                continue

            row.price = round(float(row.price_pack) / pack, 4)
            filled_count += 1

        self.session.flush()
        return filled_count

    def get_current_supplier_price(
        self,
        *,
        supplier_id: int,
        product_id: int,
    ) -> Optional[CurrentSupplierPrice]:
        return (
            self.session.query(CurrentSupplierPrice)
            .filter(
                CurrentSupplierPrice.supplier_id == supplier_id,
                CurrentSupplierPrice.product_id == product_id,
            )
            .first()
        )

    # =========================================================
    # Save prices
    # =========================================================

    def save_prices_to_history_and_current(
        self,
        *,
        batch_id: str,
        imported_by: str,
        currency_code: str,
    ) -> int:
        rows = (
            self.session.query(TempPriceImport)
            .filter(
                TempPriceImport.batch_id == batch_id,
                TempPriceImport.imported_by == imported_by,
                TempPriceImport.selected_product_id.isnot(None),
                TempPriceImport.price.isnot(None),
            )
            .order_by(TempPriceImport.import_row_no.asc(), TempPriceImport.id.asc())
            .all()
        )

        saved_count = 0

        for row in rows:
            history_row = PriceHistory(
                supplier_id=row.supplier_id,
                product_id=row.selected_product_id,
                price_date=row.import_date,
                price=float(row.price),
                currency=currency_code,
            )
            self.session.add(history_row)

            current_row = self.get_current_supplier_price(
                supplier_id=row.supplier_id,
                product_id=row.selected_product_id,
            )

            if current_row is None:
                current_row = CurrentSupplierPrice(
                    supplier_id=row.supplier_id,
                    product_id=row.selected_product_id,
                    price=float(row.price),
                    currency=currency_code,
                    last_update=row.import_date,
                )
                self.session.add(current_row)
            else:
                if current_row.last_update is None or current_row.last_update <= row.import_date:
                    current_row.price = float(row.price)
                    current_row.currency = currency_code
                    current_row.last_update = row.import_date

            saved_count += 1

        self.session.flush()
        return saved_count

    # =========================================================
    # Save calculations
    # =========================================================

    def save_supplier_price_calculations(
        self,
        *,
        batch_id: str,
        imported_by: str,
        fx_rate: float,
        currency_code: str,
    ) -> int:
        rows = (
            self.session.query(TempPriceImport)
            .filter(
                TempPriceImport.batch_id == batch_id,
                TempPriceImport.imported_by == imported_by,
                TempPriceImport.selected_product_id.isnot(None),
                TempPriceImport.price.isnot(None),
            )
            .order_by(TempPriceImport.import_row_no.asc(), TempPriceImport.id.asc())
            .all()
        )

        saved_count = 0

        for row in rows:
            calc_result = self.cost_calculation_service.calculate_supplier_costs(
                supplier_id=row.supplier_id,
                product_id=row.selected_product_id,
                supplier_price=float(row.price),
                fx_rate=float(fx_rate),
                currency_code=currency_code,
            )

            calc_row = SupplierPriceCalculation(
                calc_date=datetime.now(),
                batch_id=batch_id,
                imported_by=imported_by,
                supplier_id=row.supplier_id,
                product_id=row.selected_product_id,
                supplier_article=row.supplier_article,
                supplier_product_name=row.product_name,
                supplier_price=calc_result.supplier_price,
                cost_novo_wvat=calc_result.cost_novo_wvat,
                full_cost_msk=calc_result.full_cost_msk,
                currency_code=calc_result.currency_code,
                fx_rate_used=calc_result.fx_rate_used,
                fx_markup_used=calc_result.fx_markup_used,
                transport_used=calc_result.transport_used,
                reexport_used=calc_result.reexport_used,
                has_customs_used=calc_result.has_customs_used,
                via_novo_used=calc_result.via_novo_used,
                bank_fee_used=calc_result.bank_fee_used,
                customs_fee_used=calc_result.customs_fee_used,
                move_novo_used=calc_result.move_novo_used,
                move_msk_used=calc_result.move_msk_used,
                is_excise_used=calc_result.is_excise_used,
                additional_customs_used=calc_result.additional_customs_used,
                storage_used=calc_result.storage_used,
                marking_used=calc_result.marking_used,
            )
            self.session.add(calc_row)
            saved_count += 1

        self.session.flush()
        return saved_count

    # =========================================================
    # Full pipeline
    # =========================================================

    def run_full_import_pipeline(
        self,
        *,
        supplier_id: int,
        batch_id: str,
        imported_by: str,
        rows: Iterable[SupplierPriceImportRowData],
        currency_code: str,
        fx_rate: float,
        import_date: Optional[datetime] = None,
        replace_existing_batch_rows: bool = True,
    ) -> dict:
        self.import_rows_to_temp(
            supplier_id=supplier_id,
            batch_id=batch_id,
            imported_by=imported_by,
            rows=rows,
            import_date=import_date,
            replace_existing_batch_rows=replace_existing_batch_rows,
        )

        matched_count = self.automatch_temp_rows(batch_id, imported_by)
        self.validate_new_products_before_save(batch_id, imported_by)
        created_products_count = self.create_products_from_temp(batch_id, imported_by)
        product_articles_count = self.create_or_update_product_articles(batch_id, imported_by)
        filled_prices_count = self.fill_price_from_price_pack(batch_id, imported_by)
        saved_prices_count = self.save_prices_to_history_and_current(
            batch_id=batch_id,
            imported_by=imported_by,
            currency_code=currency_code,
        )
        saved_calculations_count = self.save_supplier_price_calculations(
            batch_id=batch_id,
            imported_by=imported_by,
            fx_rate=fx_rate,
            currency_code=currency_code,
        )

        return {
            "matched_count": matched_count,
            "created_products_count": created_products_count,
            "product_articles_count": product_articles_count,
            "filled_prices_count": filled_prices_count,
            "saved_prices_count": saved_prices_count,
            "saved_calculations_count": saved_calculations_count,
        }