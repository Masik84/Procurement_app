from __future__ import annotations

from datetime import datetime, date
from decimal import Decimal, ROUND_HALF_UP
from typing import Optional

from sqlalchemy.orm import Session

from app.db.models import SupplierPriceCalculation, TempPriceImport
from app.services.cost_calculation import CostCalculationService
from app.services.price_repository import PriceRepository
from app.services.product_matching import ProductMatchingService
from app.utils.batch import generate_import_batch_id


class SupplierPriceImportService:
    def __init__(self, session: Session) -> None:
        self.session = session
        self.product_matching_service = ProductMatchingService(session)
        self.cost_calculation_service = CostCalculationService(session)
        self.price_repository = PriceRepository(session)

    @staticmethod
    def _to_decimal(value: object) -> Decimal:
        if value is None:
            return Decimal("0")
        if isinstance(value, Decimal):
            return value
        return Decimal(str(value))

    @staticmethod
    def _round4(value: Decimal) -> Decimal:
        return value.quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP)

    def start_batch(self) -> str:
        return generate_import_batch_id()

    def delete_temp_rows(self, batch_id: str, imported_by: str) -> int:
        deleted_count = (
            self.session.query(TempPriceImport)
            .filter(
                TempPriceImport.batch_id == batch_id,
                TempPriceImport.imported_by == imported_by,
            )
            .delete(synchronize_session=False)
        )
        self.session.flush()
        return int(deleted_count or 0)

    def delete_supplier_price_calculations(self, batch_id: str, imported_by: str) -> int:
        deleted_count = (
            self.session.query(SupplierPriceCalculation)
            .filter(
                SupplierPriceCalculation.batch_id == batch_id,
                SupplierPriceCalculation.imported_by == imported_by,
            )
            .delete(synchronize_session=False)
        )
        self.session.flush()
        return int(deleted_count or 0)

    def reset_batch(self, batch_id: str, imported_by: str) -> None:
        self.delete_supplier_price_calculations(batch_id, imported_by)
        self.delete_temp_rows(batch_id, imported_by)
        self.session.flush()

    def cleanup_old_temp_rows(self, imported_by: str, before_date: Optional[date] = None) -> int:
        cutoff = before_date or date.today()
        deleted_count = (
            self.session.query(TempPriceImport)
            .filter(
                TempPriceImport.imported_by == imported_by,
                TempPriceImport.import_date < datetime.combine(cutoff, datetime.min.time()),
            )
            .delete(synchronize_session=False)
        )
        self.session.flush()
        return int(deleted_count or 0)

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

    def create_empty_temp_row(
        self,
        *,
        supplier_id: int,
        batch_id: str,
        imported_by: str,
        import_date: Optional[datetime] = None,
    ) -> TempPriceImport:
        if import_date is None:
            import_date = datetime.now()

        last_row_no = (
            self.session.query(TempPriceImport.import_row_no)
            .filter(
                TempPriceImport.batch_id == batch_id,
                TempPriceImport.imported_by == imported_by,
            )
            .order_by(TempPriceImport.import_row_no.desc(), TempPriceImport.id.desc())
            .first()
        )
        next_row_no = int(last_row_no[0]) + 1 if last_row_no and last_row_no[0] is not None else 1

        row = TempPriceImport(
            supplier_article=None,
            product_name=None,
            price=None,
            price_pack=None,
            supplier_id=supplier_id,
            import_date=import_date,
            batch_id=batch_id,
            imported_by=imported_by,
            import_row_no=next_row_no,
            selected_product_id=None,
            new_product_name=None,
            new_brand=None,
            new_pack=None,
            new_is_excise=False,
        )
        self.session.add(row)
        self.session.flush()
        return row

    def import_rows_to_temp(
        self,
        supplier_id: int,
        batch_id: str,
        imported_by: str,
        rows: list[dict],
        import_date: Optional[datetime] = None,
        replace_existing_batch_rows: bool = True,
    ) -> int:
        if import_date is None:
            import_date = datetime.now()

        if replace_existing_batch_rows:
            self.delete_temp_rows(batch_id, imported_by)
            self.delete_supplier_price_calculations(batch_id, imported_by)

        created_rows = []

        for row in rows:
            temp_row = TempPriceImport(
                supplier_article=row.get("supplier_article") or None,
                product_name=row.get("product_name") or None,
                price=row.get("price"),
                price_pack=row.get("price_pack"),
                supplier_id=supplier_id,
                import_date=import_date,
                batch_id=batch_id,
                imported_by=imported_by,
                import_row_no=row.get("import_row_no"),
                selected_product_id=None,
                new_product_name=None,
                new_brand=None,
                new_pack=None,
                new_is_excise=None,
            )
            created_rows.append(temp_row)

        if created_rows:
            self.session.add_all(created_rows)

        self.session.flush()
        return len(created_rows)

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

            self.product_matching_service.validate_new_product_fields(
                product_name=row.new_product_name,
                brand=row.new_brand,
                pack=row.new_pack,
                is_excise=row.new_is_excise,
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

    def _normalize_supplier_price_for_calc(
        self,
        *,
        supplier_id: int,
        raw_price: object,
        rf_prices_include_vat: bool,
    ) -> Decimal | None:
        if raw_price is None:
            return None

        price_value = self._to_decimal(raw_price)
        if price_value == Decimal("0"):
            return Decimal("0")

        supplier = self.cost_calculation_service.get_supplier(supplier_id)
        if not supplier.is_rf or not rf_prices_include_vat:
            return price_value

        fixed = self.cost_calculation_service.get_fixed_costs()
        vat = self._to_decimal(fixed.vat)

        if vat <= Decimal("-1"):
            raise ValueError("Некорректное значение НДС в fixed_costs.")

        return self._round4(price_value / (Decimal("1") + vat))

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
            if row.selected_product is None:
                continue

            pack = self._to_decimal(row.selected_product.pack)
            if pack == Decimal("0"):
                continue

            row.price = self._round4(self._to_decimal(row.price_pack) / pack)
            filled_count += 1

        self.session.flush()
        return filled_count

    def save_prices_to_history_and_current(
        self,
        batch_id: str,
        imported_by: str,
        currency_code: str,
        rf_prices_include_vat: bool = False,
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
            normalized_price = self._normalize_supplier_price_for_calc(
                supplier_id=row.supplier_id,
                raw_price=row.price,
                rf_prices_include_vat=rf_prices_include_vat,
            )

            self.price_repository.save_supplier_price(
                supplier_id=row.supplier_id,
                product_id=row.selected_product_id,
                price=normalized_price,
                currency_code=currency_code,
                price_date=row.import_date,
            )
            saved_count += 1

        self.session.flush()
        return saved_count

    def save_supplier_price_calculations(
        self,
        batch_id: str,
        imported_by: str,
        fx_rate: Decimal,
        currency_code: str,
        rf_prices_include_vat: bool = False,
    ) -> int:
        self.delete_supplier_price_calculations(batch_id, imported_by)

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
            normalized_price = self._normalize_supplier_price_for_calc(
                supplier_id=row.supplier_id,
                raw_price=row.price,
                rf_prices_include_vat=rf_prices_include_vat,
            )

            calc_result = self.cost_calculation_service.calculate_supplier_costs(
                supplier_id=row.supplier_id,
                product_id=row.selected_product_id,
                supplier_price=self._to_decimal(normalized_price),
                fx_rate=self._to_decimal(fx_rate),
                currency_code=currency_code,
            )

            calc_row = SupplierPriceCalculation(
                calc_date=datetime.now(),
                batch_id=batch_id,
                imported_by=imported_by,
                import_row_no=row.import_row_no,
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
                agent_fee_used=calc_result.agent_fee_used,
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

    def run_full_import_pipeline(
        self,
        supplier_id: int,
        batch_id: str,
        imported_by: str,
        rows: list[dict],
        currency_code: str,
        fx_rate: Decimal,
        import_date: Optional[datetime] = None,
        replace_existing_batch_rows: bool = True,
        rf_prices_include_vat: bool = False,
    ) -> dict:
        imported_count = self.import_rows_to_temp(
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
            rf_prices_include_vat=rf_prices_include_vat,
        )
        saved_calculations_count = self.save_supplier_price_calculations(
            batch_id=batch_id,
            imported_by=imported_by,
            fx_rate=fx_rate,
            currency_code=currency_code,
            rf_prices_include_vat=rf_prices_include_vat,
        )

        return {
            "imported_count": imported_count,
            "matched_count": matched_count,
            "created_products_count": created_products_count,
            "product_articles_count": product_articles_count,
            "filled_prices_count": filled_prices_count,
            "saved_prices_count": saved_prices_count,
            "saved_calculations_count": saved_calculations_count,
        }