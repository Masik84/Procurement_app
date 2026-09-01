from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, date
from pathlib import Path
from decimal import Decimal, ROUND_HALF_UP
from typing import Optional

from sqlalchemy.orm import Session, joinedload

from app.db.models import Product, SupplierPriceCalculation, TempPriceImport
from app.exports.supplier_price_exporter import SupplierPriceExporter
from app.imports.supplier_price_importer import SupplierPriceImporter
from app.services.supplier_service import SupplierService, SupplierUpsertData
from app.services.supplier_currency_cost_service import SupplierCurrencyCostService
from app.services.cost_calculation_service import CostCalculationService
from app.services.price_repository import PriceRepository
from app.services.product_matching_service import ProductMatchingService
from app.services.temp_cleanup_service import TempCleanupService
from app.utils.batch import generate_import_batch_id
from app.services.qty_in_box_service import (
    calculate_qty_in_box_candidates,
    default_qty_in_box_for_pack,
    normalize_qty_in_box,
    whole_qty_in_box_candidate,
)
from app.utils.text import clean_multi_spaces


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


class SupplierPriceService:
    def __init__(self, session: Session) -> None:
        self.session = session
        self.product_matching_service = ProductMatchingService(session)
        self.cost_calculation_service = CostCalculationService(session)
        self.price_repository = PriceRepository(session)
        self.supplier_service = SupplierService(session)
        self.currency_cost_service = SupplierCurrencyCostService(
            session,
            cost_calculation=self.cost_calculation_service,
        )
        self.importer = SupplierPriceImporter()
        self.exporter = SupplierPriceExporter(session)
        self.last_create_products_debug: list[dict] = []
        self.last_validate_products_debug: list[dict] = []

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

    @classmethod
    def _is_positive_price(cls, value: object) -> bool:
        """Returns True only for a real positive price.

        Empty, zero and negative values are treated as rows without a price:
        these rows may still be used for product matching/article links,
        but they must not be written to price history or cost calculations.
        """
        if value is None:
            return False

        try:
            return cls._to_decimal(value) > Decimal("0")
        except Exception:
            return False

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

    def delete_temp_rows_for_user(self, imported_by: str) -> int:
        return TempCleanupService(self.session).delete_current_user(
            imported_by=imported_by,
            tables=(TempPriceImport,),
        )

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

    def cleanup_old_temp_rows(self, imported_by: str | None = None, before_date: Optional[date] = None) -> int:
        # Daily cleanup is global by date: all temp rows older than today are stale.
        # Current user temp rows are cleaned after successful save.
        return TempCleanupService(self.session).cleanup_old_for_all(before_date=before_date)

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
            price_box=None,
            qty_pcs=None,
            qty_box=None,
            volume_l=None,
            supplier_id=supplier_id,
            import_date=import_date,
            batch_id=batch_id,
            imported_by=imported_by,
            import_row_no=next_row_no,
            selected_product_id=None,
            new_product_name=None,
            new_brand=None,
            new_pack=None,
            new_qty_in_box=None,
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
                price_box=row.get("price_box"),
                qty_pcs=row.get("qty_pcs"),
                qty_box=row.get("qty_box"),
                volume_l=row.get("volume_l"),
                supplier_id=supplier_id,
                import_date=import_date,
                batch_id=batch_id,
                imported_by=imported_by,
                import_row_no=row.get("import_row_no"),
                selected_product_id=None,
                new_product_name=None,
                new_brand=None,
                new_pack=None,
                new_qty_in_box=None,
                new_is_excise=False,
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
                row.new_product_name = None
                try:
                    row.new_qty_in_box = normalize_qty_in_box(product.qty_in_box)
                except ValueError:
                    row.new_qty_in_box = None
                row.new_is_excise = bool(product.is_excise)
                matched_count += 1

        self.session.flush()
        return matched_count

    @staticmethod
    def _has_new_product_data(row: TempPriceImport) -> bool:
        return any([
            bool((row.new_product_name or "").strip()),
            bool((row.new_brand or "").strip()),
            row.new_pack is not None,
        ])

    def validate_new_products_before_save(self, batch_id: str, imported_by: str) -> None:
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

        self.last_validate_products_debug = []

        for row in rows:
            has_any_new_data = self._has_new_product_data(row)
            if not has_any_new_data:
                continue

            dbg = {
                "row_id": row.id,
                "import_row_no": row.import_row_no,
                "article": row.supplier_article,
                "supplier_product_name": row.product_name,
                "new_product_name": row.new_product_name,
                "new_brand": row.new_brand,
                "new_pack": row.new_pack,
                "new_qty_in_box": row.new_qty_in_box,
                "new_is_excise": row.new_is_excise,
                "selected_product_id": row.selected_product_id,
            }
            self.last_validate_products_debug.append(dbg)

            try:
                self.product_matching_service.validate_new_product_fields(
                    product_name=row.new_product_name,
                    brand=row.new_brand,
                    pack=row.new_pack,
                    is_excise=row.new_is_excise,
                    qty_in_box=row.new_qty_in_box,
                )
            except Exception as e:
                raise ValueError(
                    f"[DEBUG validate_new_products_before_save] "
                    f"row_id={row.id}, import_row_no={row.import_row_no}, "
                    f"article={row.supplier_article!r}, source_name={row.product_name!r}, "
                    f"new_product_name={row.new_product_name!r}, new_brand={row.new_brand!r}, "
                    f"new_pack={row.new_pack!r}, new_is_excise={row.new_is_excise!r}. "
                    f"Ошибка: {e}"
                ) from e

    def create_products_from_temp(self, batch_id: str, imported_by: str) -> int:
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

        created_count = 0
        self.last_create_products_debug = []

        for row in rows:
            debug_row = {
                "row_id": row.id,
                "import_row_no": row.import_row_no,
                "article": row.supplier_article,
                "supplier_product_name": row.product_name,
                "new_product_name": row.new_product_name,
                "new_brand": row.new_brand,
                "new_pack": row.new_pack,
                "new_qty_in_box": row.new_qty_in_box,
                "new_is_excise": row.new_is_excise,
                "selected_product_id_before": row.selected_product_id,
                "status": "",
                "product_id_after": None,
            }

            has_any_new_data = self._has_new_product_data(row)
            if not has_any_new_data:
                debug_row["status"] = "skip:no_new_product_data"
                self.last_create_products_debug.append(debug_row)
                continue

            if not (row.new_product_name or "").strip():
                debug_row["status"] = "skip:empty_new_product_name"
                self.last_create_products_debug.append(debug_row)
                continue

            if not (row.new_brand or "").strip():
                debug_row["status"] = "skip:empty_new_brand"
                self.last_create_products_debug.append(debug_row)
                continue

            if row.new_pack is None:
                debug_row["status"] = "skip:empty_new_pack"
                self.last_create_products_debug.append(debug_row)
                continue

            try:
                product = self.product_matching_service.get_or_create_product(
                    name=row.new_product_name,
                    brand=row.new_brand,
                    pack=row.new_pack,
                    qty_in_box=row.new_qty_in_box,
                    is_excise=bool(row.new_is_excise) if row.new_is_excise is not None else False,
                )
                row.selected_product_id = product.id
                debug_row["status"] = "created_or_found"
                debug_row["product_id_after"] = product.id
                created_count += 1
                self.last_create_products_debug.append(debug_row)
            except Exception as e:
                debug_row["status"] = f"error:{e}"
                self.last_create_products_debug.append(debug_row)
                raise ValueError(
                    f"[DEBUG create_products_from_temp] "
                    f"row_id={row.id}, import_row_no={row.import_row_no}, "
                    f"article={row.supplier_article!r}, source_name={row.product_name!r}, "
                    f"new_product_name={row.new_product_name!r}, new_brand={row.new_brand!r}, "
                    f"new_pack={row.new_pack!r}, new_is_excise={row.new_is_excise!r}. "
                    f"Ошибка: {e}"
                ) from e

        self.session.flush()
        return created_count

    def prepare_box_data_and_update_products(self, batch_id: str, imported_by: str) -> list[dict]:
        """Apply permitted product edits, resolve Qty in Box and fill box quantities."""
        rows = (
            self.session.query(TempPriceImport)
            .options(joinedload(TempPriceImport.selected_product))
            .filter(
                TempPriceImport.batch_id == batch_id,
                TempPriceImport.imported_by == imported_by,
                TempPriceImport.selected_product_id.isnot(None),
            )
            .order_by(TempPriceImport.import_row_no.asc(), TempPriceImport.id.asc())
            .all()
        )
        warnings: list[dict] = []
        missing_rows: list[str] = []

        def add_warning(row, product, *, db_value, calculated, source, comment):
            warnings.append({
                "Import row": row.import_row_no,
                "Product ID": int(product.id),
                "Product Name": product.name or "",
                "Pack": product.pack,
                "Qty in Box DB": db_value,
                "Qty in Box calculated": calculated,
                "Source": source,
                "Comment": comment,
            })

        for row in rows:
            product = row.selected_product
            if product is None:
                continue

            new_name = clean_multi_spaces(row.new_product_name).upper()
            if new_name and new_name != clean_multi_spaces(product.name).upper():
                duplicate = (
                    self.session.query(Product)
                    .filter(Product.id != product.id, Product.name == new_name)
                    .first()
                )
                if duplicate is not None:
                    raise ValueError(
                        f"Строка {row.import_row_no}: название '{new_name}' уже используется продуктом id={duplicate.id}."
                    )
                self.product_matching_service.validate_product_name_pack_format(
                    product_name=new_name,
                    pack_value=product.pack,
                )
                product.name = new_name
                product.family = self.product_matching_service.build_product_family_from_name(
                    new_name, product.pack
                )

            if row.new_is_excise is not None:
                product.is_excise = bool(row.new_is_excise)

            db_qty_raw = product.qty_in_box
            try:
                db_qty = normalize_qty_in_box(db_qty_raw)
            except ValueError:
                db_qty = None
                add_warning(
                    row,
                    product,
                    db_value=db_qty_raw,
                    calculated=None,
                    source="DB",
                    comment="В БД указано нулевое, отрицательное или дробное Qty in Box.",
                )

            user_qty = normalize_qty_in_box(
                row.new_qty_in_box,
                field_name="Qty in Box (for new)",
            )
            explicit_user_change = user_qty is not None and user_qty != db_qty
            if explicit_user_change:
                product.qty_in_box = user_qty
                db_qty = user_qty

            default_qty = default_qty_in_box_for_pack(self.session, product.pack)
            formula_candidates = calculate_qty_in_box_candidates(
                qty_pcs=row.qty_pcs,
                qty_box=row.qty_box,
                volume_l=row.volume_l,
                pack=product.pack,
            )
            whole_candidates = {
                source: whole_qty_in_box_candidate(value)
                for source, value in formula_candidates.items()
            }

            formula_values = {value for value in formula_candidates.values()}
            if len(formula_values) > 1:
                add_warning(
                    row,
                    product,
                    db_value=db_qty_raw,
                    calculated="; ".join(f"{source} = {value}" for source, value in formula_candidates.items()),
                    source="Несколько формул",
                    comment="Расчёты Qty in Box дают разные значения; БД не изменена.",
                )
            elif formula_candidates:
                source, calculated_raw = next(iter(formula_candidates.items()))
                calculated = whole_candidates[source]
                if calculated is None:
                    add_warning(
                        row,
                        product,
                        db_value=db_qty_raw,
                        calculated=calculated_raw,
                        source=source,
                        comment="Расчёт дал нецелое или неположительное значение; округление не выполнялось.",
                    )
                elif default_qty is not None and calculated != default_qty:
                    add_warning(
                        row,
                        product,
                        db_value=db_qty_raw,
                        calculated=calculated,
                        source=source,
                        comment="Для упаковки бочка/ведро/куб/кега стандарт Qty in Box = 1; рассчитано другое значение.",
                    )
                elif not explicit_user_change and db_qty is not None and calculated != db_qty:
                    add_warning(
                        row,
                        product,
                        db_value=db_qty_raw,
                        calculated=calculated,
                        source=source,
                        comment="Рассчитанное значение отличается от БД; БД не перезаписана.",
                    )
                elif db_qty is None and default_qty is None:
                    product.qty_in_box = calculated
                    db_qty = calculated

            if not explicit_user_change and default_qty is not None:
                if db_qty is None:
                    product.qty_in_box = default_qty
                    db_qty = default_qty
                elif db_qty != default_qty:
                    add_warning(
                        row,
                        product,
                        db_value=db_qty_raw,
                        calculated=default_qty,
                        source="Виды Упаковок",
                        comment="Для упаковки бочка/ведро/куб/кега ожидается Qty in Box = 1; БД не перезаписана.",
                    )

            effective_qty = normalize_qty_in_box(product.qty_in_box) if product.qty_in_box is not None else None
            row.new_product_name = None
            row.new_qty_in_box = effective_qty
            row.new_is_excise = bool(product.is_excise)

            qty_box = self._to_decimal(row.qty_box)
            if qty_box > 0 and effective_qty is not None:
                if self._to_decimal(row.qty_pcs) <= 0:
                    row.qty_pcs = qty_box * Decimal(effective_qty)
                if self._to_decimal(row.volume_l) <= 0 and self._to_decimal(row.qty_pcs) > 0:
                    row.volume_l = self._to_decimal(row.qty_pcs) * self._to_decimal(product.pack)

            if self._is_positive_price(row.price_box) and effective_qty is None:
                missing_rows.append(
                    f"строка {row.import_row_no}: id={product.id}, {product.name}"
                )

        if missing_rows:
            raise ValueError(
                "Для расчёта Price, box не заполнено Qty in Box. Заполните 'Qty in Box (for new)':\n"
                + "\n".join(missing_rows)
            )

        self.session.flush()
        return warnings

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
            .options(joinedload(TempPriceImport.selected_product))
            .filter(
                TempPriceImport.batch_id == batch_id,
                TempPriceImport.imported_by == imported_by,
                TempPriceImport.selected_product_id.isnot(None),
            )
            .all()
        )

        filled_count = 0

        for row in rows:
            if row.selected_product is None:
                continue

            # If Price, L is already a real positive value, keep it.
            # If it is empty/0 but Price, pack is positive, calculate Price, L from pack.
            if self._is_positive_price(row.price):
                continue

            pack = self._to_decimal(row.selected_product.pack)
            if pack <= Decimal("0"):
                continue

            if self._is_positive_price(row.price_pack):
                row.price = self._round4(self._to_decimal(row.price_pack) / pack)
            elif self._is_positive_price(row.price_box):
                qty_in_box = normalize_qty_in_box(row.selected_product.qty_in_box)
                if qty_in_box is None:
                    continue
                row.price = self._round4(
                    self._to_decimal(row.price_box) / pack / Decimal(qty_in_box)
                )
            else:
                continue
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

        prices_to_save: list[dict[str, object]] = []

        for row in rows:
            if not self._is_positive_price(row.price):
                continue

            normalized_price = self._normalize_supplier_price_for_calc(
                supplier_id=row.supplier_id,
                raw_price=row.price,
                rf_prices_include_vat=rf_prices_include_vat,
            )
            if not self._is_positive_price(normalized_price):
                continue

            prices_to_save.append({
                "supplier_id": row.supplier_id,
                "product_id": row.selected_product_id,
                "price": normalized_price,
                "currency_code": currency_code,
                "price_date": row.import_date,
            })

        return self.price_repository.save_supplier_prices_batch(prices_to_save)

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
            if not self._is_positive_price(row.price):
                continue

            normalized_price = self._normalize_supplier_price_for_calc(
                supplier_id=row.supplier_id,
                raw_price=row.price,
                rf_prices_include_vat=rf_prices_include_vat,
            )
            if not self._is_positive_price(normalized_price):
                continue

            calc_result = self.currency_cost_service.calculate_costs_for_price_record(
                supplier_id=row.supplier_id,
                product_id=row.selected_product_id,
                supplier_price=self._to_decimal(normalized_price),
                price_currency_code=currency_code,
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
                fx_markup_abs_used=calc_result.fx_markup_abs_used,
                transport_used=calc_result.transport_used,
                reexport_used=calc_result.reexport_used,
                insurance_used=calc_result.insurance_used,
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
        self.prepare_box_data_and_update_products(batch_id, imported_by)
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

    def run_from_excel(
        self,
        *,
        file_path: str | Path,
        imported_by: str,
        supplier_data: SupplierUpsertData,
        supplier_id: int | None = None,
        import_date: datetime | None = None,
        save_exchange_rate: bool = False,
        explicit_fx_rate: float | None = None,
        rf_prices_include_vat: bool = False,
    ) -> SupplierPriceImportResult:
        supplier = self.supplier_service.ensure_supplier(
            supplier_id=supplier_id,
            data=supplier_data,
        )
        currency_code = supplier.base_currency
        fx_rate: float | None = None

        if explicit_fx_rate is not None:
            fx_rate = float(explicit_fx_rate)
            if save_exchange_rate:
                self.supplier_service.save_exchange_rate(currency_code, fx_rate)
        else:
            fx_rate = self.supplier_service.get_rate_to_rub(currency_code)

        if fx_rate is None or float(fx_rate) == 0:
            raise ValueError(f"Для валюты '{currency_code}' не найден корректный курс rate_to_rub.")

        rows = self.importer.read_excel(file_path)
        batch_id = self.start_batch()

        stats = self.run_full_import_pipeline(
            supplier_id=supplier.id,
            batch_id=batch_id,
            imported_by=imported_by,
            rows=rows,
            currency_code=currency_code,
            fx_rate=fx_rate,
            import_date=import_date,
            replace_existing_batch_rows=True,
            rf_prices_include_vat=rf_prices_include_vat,
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

    def export_calculated(
        self,
        batch_id: str,
        imported_by: str,
        supplier_id: int,
        output_path: str | Path | None = None,
        quick_order_months: int | None = None,
        safe_stock_months: int | None = None,
    ) -> Path:
        return self.exporter.export_calculated(
            batch_id=batch_id,
            imported_by=imported_by,
            supplier_id=supplier_id,
            output_path=output_path,
            quick_order_months=quick_order_months,
            safe_stock_months=safe_stock_months,
        )


# Backward-compatible aliases for old imports.
SupplierPriceImportService = SupplierPriceService
SupplierPriceImportRun = SupplierPriceService
