from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
import uuid

from sqlalchemy import text
from sqlalchemy.orm import Session

from app.db.models import (
    FixedCosts,
    Product,
    ProductStock,
    Supplier,
    TargetPriceCalculation,
    TempTargetPriceImport,
    TempTargetPriceOption,
)
from app.imports.target_price_importer import TargetPriceImporter
from app.services.cost_calculation_service import CostCalculationService
from app.services.price_repository import PriceRepository
from app.services.product_matching_service import ProductMatchingService
from app.services.supplier_service import SupplierService
from app.services.supplier_currency_cost_service import SupplierCurrencyCostService
from app.services.temp_cleanup_service import TempCleanupService
from app.utils.text import clean_multi_spaces


class TargetPriceService:
    def __init__(self, session: Session):
        self.session = session
        self.product_matching = ProductMatchingService(session)
        self.price_repository = PriceRepository(session)
        self.cost_calculation = CostCalculationService(session)
        self.supplier_service = SupplierService(session)
        self.currency_cost_service = SupplierCurrencyCostService(
            session,
            cost_calculation=self.cost_calculation,
        )
        self.importer = TargetPriceImporter()

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
        return f"TP_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}_{uuid.uuid4().hex[:6]}"

    def delete_temp_options(self, batch_id: str, imported_by: str) -> int:
        deleted = self.session.query(TempTargetPriceOption).filter(
            TempTargetPriceOption.batch_id == batch_id,
            TempTargetPriceOption.imported_by == imported_by,
        ).delete(synchronize_session=False)
        self.session.flush()
        return int(deleted or 0)

    def delete_temp_rows(self, batch_id: str, imported_by: str) -> int:
        self.delete_temp_options(batch_id, imported_by)
        deleted = self.session.query(TempTargetPriceImport).filter(
            TempTargetPriceImport.batch_id == batch_id,
            TempTargetPriceImport.imported_by == imported_by,
        ).delete(synchronize_session=False)
        self.session.flush()
        return int(deleted or 0)

    def delete_temp_rows_for_user(self, imported_by: str) -> int:
        return TempCleanupService(self.session).delete_current_user(
            imported_by=imported_by,
            tables=(TempTargetPriceOption, TempTargetPriceImport),
        )

    def cleanup_old_temp_rows(self, imported_by: str | None = None, before_date: date | None = None) -> int:
        # Daily cleanup is global by date: all temp rows older than today are stale.
        # Current user temp rows are cleaned after successful save.
        return TempCleanupService(self.session).cleanup_old_for_all(before_date=before_date)

    def import_rows(
        self,
        *,
        rows: list[dict],
        batch_id: str,
        imported_by: str,
        supplier_id: int | None,
        replace_existing: bool = True,
    ) -> int:
        if replace_existing:
            self.delete_temp_rows(batch_id, imported_by)
        count = 0
        for r in rows:
            row = TempTargetPriceImport(
                batch_id=batch_id,
                imported_by=imported_by,
                import_date=datetime.utcnow(),
                import_row_no=r.get("import_row_no"),
                target_supplier_id=supplier_id,
                supplier_article=r.get("supplier_article"),
                product_name=r.get("product_name"),
            )
            self.session.add(row)
            count += 1
        self.session.flush()
        return count

    def automatch_temp_rows(self, batch_id: str, imported_by: str) -> int:
        rows = self.session.query(TempTargetPriceImport).filter(
            TempTargetPriceImport.batch_id == batch_id,
            TempTargetPriceImport.imported_by == imported_by,
            TempTargetPriceImport.selected_product_id.is_(None),
        ).all()
        matched = 0
        for row in rows:
            product = self.product_matching.find_price_import_product(
                supplier_article=row.supplier_article,
                supplier_product_name=row.product_name,
            )
            if product is not None:
                row.selected_product_id = product.id
                matched += 1
        self.session.flush()
        return matched

    def validate_new_products_before_save(self, batch_id: str, imported_by: str) -> None:
        rows = self.session.query(TempTargetPriceImport).filter(
            TempTargetPriceImport.batch_id == batch_id,
            TempTargetPriceImport.imported_by == imported_by,
            TempTargetPriceImport.selected_product_id.is_(None),
            TempTargetPriceImport.new_product_name.isnot(None),
        ).all()
        for row in rows:
            if not clean_multi_spaces(row.new_product_name):
                continue
            self.product_matching.validate_new_product_fields(
                product_name=row.new_product_name,
                brand=row.new_brand,
                pack=row.new_pack,
                is_excise=row.new_is_excise,
                qty_in_box=row.new_qty_in_box,
            )

    def create_products_from_temp(self, batch_id: str, imported_by: str) -> int:
        rows = self.session.query(TempTargetPriceImport).filter(
            TempTargetPriceImport.batch_id == batch_id,
            TempTargetPriceImport.imported_by == imported_by,
            TempTargetPriceImport.selected_product_id.is_(None),
            TempTargetPriceImport.new_product_name.isnot(None),
            TempTargetPriceImport.new_brand.isnot(None),
            TempTargetPriceImport.new_pack.isnot(None),
            TempTargetPriceImport.new_is_excise.isnot(None),
        ).order_by(TempTargetPriceImport.import_row_no.asc(), TempTargetPriceImport.id.asc()).all()
        created = 0
        for row in rows:
            product = self.product_matching.get_or_create_product(
                name=row.new_product_name,
                brand=row.new_brand,
                pack=row.new_pack,
                is_excise=bool(row.new_is_excise),
                qty_in_box=row.new_qty_in_box,
            )
            row.selected_product_id = product.id
            created += 1
        self.session.flush()
        return created

    def create_or_update_product_articles(self, batch_id: str, imported_by: str) -> int:
        rows = self.session.query(TempTargetPriceImport).filter(
            TempTargetPriceImport.batch_id == batch_id,
            TempTargetPriceImport.imported_by == imported_by,
            TempTargetPriceImport.selected_product_id.isnot(None),
        ).order_by(TempTargetPriceImport.import_row_no.asc(), TempTargetPriceImport.id.asc()).all()
        processed = 0
        for row in rows:
            self.product_matching.save_product_articles_by_split_articles(
                product_id=row.selected_product_id,
                supplier_article=row.supplier_article,
                supplier_product_name=row.product_name,
            )
            processed += 1
        self.session.flush()
        return processed

    def _get_fx_rate_for_currency(self, currency_code: str) -> Decimal:
        rate = self.supplier_service.get_rate_to_rub(currency_code)
        if rate is None or float(rate) == 0:
            raise ValueError(f"Для валюты '{currency_code}' не найден корректный курс rate_to_rub.")
        return self._to_decimal(rate)

    def build_supplier_options(
        self,
        batch_id: str,
        imported_by: str,
        supplier_price_age_months: int | None = None,
    ) -> int:
        self.delete_temp_options(batch_id, imported_by)
        min_price_date = self.price_repository.supplier_price_cutoff_from_months(supplier_price_age_months)
        rows = self.session.query(TempTargetPriceImport).filter(
            TempTargetPriceImport.batch_id == batch_id,
            TempTargetPriceImport.imported_by == imported_by,
            TempTargetPriceImport.selected_product_id.isnot(None),
        ).order_by(TempTargetPriceImport.import_row_no.asc(), TempTargetPriceImport.id.asc()).all()

        prices_by_product = self.price_repository.get_supplier_prices_for_products(
            (row.selected_product_id for row in rows),
            only_rating_calc=True,
            min_price_date=min_price_date,
        )
        supplier_ids = {
            price.supplier_id
            for prices in prices_by_product.values()
            for price in prices
        }
        self.currency_cost_service.preload_reference_data(
            product_ids=(row.selected_product_id for row in rows),
            supplier_ids=supplier_ids,
        )

        created = 0
        for row in rows:
            for supplier_price in prices_by_product.get(int(row.selected_product_id), []):
                self._create_option_from_snapshot(row, batch_id, imported_by, supplier_price)
                created += 1

        self.session.flush()
        self.session.execute(text("""
            WITH ranked AS (
                SELECT
                    id,
                    ROW_NUMBER() OVER (
                        PARTITION BY temp_import_id
                        ORDER BY cost_novo_wvat ASC, full_cost_msk ASC, supplier_name ASC, id ASC
                    ) AS new_rank
                FROM temp_target_price_options
                WHERE batch_id = :batch_id AND imported_by = :imported_by
            )
            UPDATE temp_target_price_options AS target
            SET opt_rank = ranked.new_rank
            FROM ranked
            WHERE target.id = ranked.id
        """), {"batch_id": batch_id, "imported_by": imported_by})
        ordered_options = self.session.query(TempTargetPriceOption).filter(
            TempTargetPriceOption.batch_id == batch_id,
            TempTargetPriceOption.imported_by == imported_by,
        ).order_by(
            TempTargetPriceOption.temp_import_id.asc(),
            TempTargetPriceOption.opt_rank.asc(),
            TempTargetPriceOption.id.asc(),
        ).populate_existing().all()
        options_by_row = {}
        for option in ordered_options:
            row_options = options_by_row.setdefault(int(option.temp_import_id), [])
            row_options.append(option)

        for row in rows:
            options = options_by_row.get(int(row.id), [])
            row.selected_option_id = options[0].id if len(options) == 1 else None
        self.session.flush()
        return created

    def _create_option_from_snapshot(self, row: TempTargetPriceImport, batch_id: str, imported_by: str, supplier_price) -> TempTargetPriceOption:
        calc = self.currency_cost_service.calculate_costs_for_price_record(
            supplier_id=supplier_price.supplier_id,
            product_id=row.selected_product_id,
            supplier_price=supplier_price.price,
            price_currency_code=supplier_price.currency_code,
        )

        def safe(v):
            return v if v is not None else Decimal("0")

        option = TempTargetPriceOption(
            temp_import_id=row.id,
            batch_id=batch_id,
            imported_by=imported_by,
            calc_date=datetime.utcnow(),
            supplier_id=supplier_price.supplier_id,
            product_id=row.selected_product_id,
            supplier_name=supplier_price.supplier_name,
            supplier_article=getattr(supplier_price, "supplier_article", None) or row.supplier_article,
            supplier_product_name=getattr(supplier_price, "supplier_product_name", None) or row.product_name,
            supplier_price=calc.supplier_price,
            price_date_used=supplier_price.price_date,
            cost_novo_wvat=calc.cost_novo_wvat,
            full_cost_msk=calc.full_cost_msk,
            currency_code=calc.currency_code,
            fx_rate_used=safe(calc.fx_rate_used),
            fx_markup_used=safe(calc.fx_markup_used),
            fx_markup_abs_used=safe(calc.fx_markup_abs_used),
            transport_used=safe(calc.transport_used),
            reexport_used=safe(calc.reexport_used),
            insurance_used=safe(calc.insurance_used),
            agent_fee_used=safe(calc.agent_fee_used),
            has_customs_used=calc.has_customs_used,
            via_novo_used=calc.via_novo_used,
            bank_fee_used=safe(calc.bank_fee_used),
            customs_fee_used=safe(calc.customs_fee_used),
            move_novo_used=safe(calc.move_novo_used),
            move_msk_used=safe(calc.move_msk_used),
            is_excise_used=calc.is_excise_used,
            additional_customs_used=safe(calc.additional_customs_used),
            storage_used=safe(calc.storage_used),
            marking_used=safe(calc.marking_used),
            opt_rank=None,
        )
        self.session.add(option)
        return option

    def rank_supplier_options(self, batch_id: str, imported_by: str) -> None:
        options = self.session.query(TempTargetPriceOption).filter(
            TempTargetPriceOption.batch_id == batch_id,
            TempTargetPriceOption.imported_by == imported_by,
        ).order_by(
            TempTargetPriceOption.temp_import_id.asc(),
            TempTargetPriceOption.cost_novo_wvat.asc(),
            TempTargetPriceOption.full_cost_msk.asc(),
            TempTargetPriceOption.supplier_name.asc(),
            TempTargetPriceOption.id.asc(),
        ).all()
        current_row_id: int | None = None
        rank = 0
        for option in options:
            row_id = int(option.temp_import_id)
            if row_id != current_row_id:
                current_row_id = row_id
                rank = 0
            rank += 1
            option.opt_rank = rank
        self.session.flush()

    def select_best_options(self, batch_id: str, imported_by: str) -> None:
        rows = self.session.query(TempTargetPriceImport).filter(
            TempTargetPriceImport.batch_id == batch_id,
            TempTargetPriceImport.imported_by == imported_by,
        ).all()
        options = self.session.query(TempTargetPriceOption).filter(
            TempTargetPriceOption.batch_id == batch_id,
            TempTargetPriceOption.imported_by == imported_by,
        ).order_by(TempTargetPriceOption.opt_rank.asc(), TempTargetPriceOption.id.asc()).all()
        options_by_row: dict[int, list[TempTargetPriceOption]] = {}
        for option in options:
            options_by_row.setdefault(int(option.temp_import_id), []).append(option)
        for row in rows:
            row_options = options_by_row.get(int(row.id), [])
            row.selected_option_id = row_options[0].id if len(row_options) == 1 else None
        self.session.flush()

    def _get_manual_supplier(self) -> Supplier:
        supplier = self.session.query(Supplier).filter(Supplier.name == "Manual").first()
        if supplier is None:
            raise ValueError("В справочнике поставщиков должен быть технический поставщик Manual.")
        return supplier

    def apply_manual_full_costs(self, batch_id: str, imported_by: str, manual_full_costs: dict[int, Decimal] | None) -> int:
        if not manual_full_costs:
            return 0
        manual_supplier = self._get_manual_supplier()
        fixed = self.cost_calculation.get_fixed_costs()
        values_by_row_id = {
            int(row_id): self._to_decimal(value)
            for row_id, value in manual_full_costs.items()
            if self._to_decimal(value) != 0
        }
        if not values_by_row_id:
            return 0
        rows = self.session.query(TempTargetPriceImport).filter(
            TempTargetPriceImport.id.in_(values_by_row_id),
            TempTargetPriceImport.batch_id == batch_id,
            TempTargetPriceImport.imported_by == imported_by,
            TempTargetPriceImport.selected_product_id.isnot(None),
        ).all()
        row_ids = [int(row.id) for row in rows]
        if row_ids:
            # Manual overrides all ordinary supplier options for this row.
            self.session.query(TempTargetPriceOption).filter(
                TempTargetPriceOption.batch_id == batch_id,
                TempTargetPriceOption.imported_by == imported_by,
                TempTargetPriceOption.temp_import_id.in_(row_ids),
            ).delete(synchronize_session=False)

        new_options: list[tuple[TempTargetPriceImport, TempTargetPriceOption]] = []
        for row in rows:
            option = TempTargetPriceOption(
                temp_import_id=row.id,
                batch_id=batch_id,
                imported_by=imported_by,
                calc_date=datetime.utcnow(),
                supplier_id=manual_supplier.id,
                product_id=row.selected_product_id,
                supplier_name="Manual",
                supplier_article=row.supplier_article,
                supplier_product_name=row.product_name,
                supplier_price=Decimal("0"),
                price_date_used=None,
                cost_novo_wvat=Decimal("0"),
                full_cost_msk=values_by_row_id[int(row.id)],
                currency_code=manual_supplier.base_currency or "USD",
                fx_rate_used=Decimal("0"),
                fx_markup_used=Decimal("0"),
                fx_markup_abs_used=Decimal("0"),
                transport_used=Decimal("0"),
                reexport_used=Decimal("0"),
                insurance_used=Decimal("0"),
                agent_fee_used=Decimal("0"),
                has_customs_used=False,
                via_novo_used=False,
                bank_fee_used=self._to_decimal(getattr(fixed, "bank_fee", 0) if fixed else 0),
                customs_fee_used=self._to_decimal(getattr(fixed, "customs_fee", 0) if fixed else 0),
                move_novo_used=self._to_decimal(getattr(fixed, "move_novo_tamozh", 0) if fixed else 0),
                move_msk_used=self._to_decimal(getattr(fixed, "move_tamozh_chekhov", 0) if fixed else 0),
                is_excise_used=False,
                additional_customs_used=self._to_decimal(getattr(fixed, "additional_customs", 0) if fixed else 0),
                storage_used=self._to_decimal(getattr(fixed, "storage", 0) if fixed else 0),
                marking_used=Decimal("0"),
                opt_rank=1,
            )
            self.session.add(option)
            new_options.append((row, option))

        self.session.flush()
        for row, option in new_options:
            row.selected_option_id = option.id
        self.session.flush()
        return len(new_options)

    def ensure_single_options_selected(self, batch_id: str, imported_by: str) -> None:
        rows = self.session.query(TempTargetPriceImport).filter(
            TempTargetPriceImport.batch_id == batch_id,
            TempTargetPriceImport.imported_by == imported_by,
        ).all()
        options = self.session.query(TempTargetPriceOption).filter(
            TempTargetPriceOption.batch_id == batch_id,
            TempTargetPriceOption.imported_by == imported_by,
        ).order_by(TempTargetPriceOption.opt_rank.asc(), TempTargetPriceOption.id.asc()).all()
        options_by_row: dict[int, list[TempTargetPriceOption]] = {}
        for option in options:
            options_by_row.setdefault(int(option.temp_import_id), []).append(option)
        for row in rows:
            if row.selected_option_id is None:
                row_options = options_by_row.get(int(row.id), [])
                if len(row_options) == 1:
                    row.selected_option_id = row_options[0].id
        self.session.flush()

    def run_calculation(
        self,
        batch_id: str,
        imported_by: str,
        manual_full_costs: dict[int, Decimal] | None = None,
        supplier_price_age_months: int | None = None,
    ) -> dict:
        self.validate_new_products_before_save(batch_id, imported_by)
        created_products = self.create_products_from_temp(batch_id, imported_by)
        product_articles = self.create_or_update_product_articles(batch_id, imported_by)
        options = self.build_supplier_options(
            batch_id,
            imported_by,
            supplier_price_age_months=supplier_price_age_months,
        )
        manual_options = self.apply_manual_full_costs(batch_id, imported_by, manual_full_costs or {})
        self.select_best_options(batch_id, imported_by)
        return {
            "created_products_count": created_products,
            "product_articles_count": product_articles,
            "options_count": options + manual_options,
        }

    def reverse_calculate_target_price(
        self,
        *,
        target_supplier_id: int,
        product_id: int,
        full_cost_msk: object,
        currency_code: str,
        fx_rate: object,
        transport: object,
        reexport: object,
        insurance: object,
        fx_markup: object,
        fx_markup_abs: object,
        has_customs: bool,
        via_novo: bool,
        agent_fee: object | None = None,
    ) -> tuple[Decimal, Decimal]:
        supplier = self.cost_calculation.get_supplier(target_supplier_id)
        product = self.cost_calculation.get_product(product_id)
        fixed = self.cost_calculation.get_fixed_costs()

        d_full_cost = self._to_decimal(full_cost_msk)
        d_fx_rate = self._to_decimal(fx_rate)
        if d_fx_rate == 0:
            raise ValueError("Курс валюты не может быть 0.")
        d_transport = self._to_decimal(transport)
        d_reexport = self._to_decimal(reexport)
        d_insurance = self._to_decimal(insurance)
        d_fx_markup = self._to_decimal(fx_markup)
        d_fx_markup_abs = self._to_decimal(fx_markup_abs)
        d_effective_fx_rate = d_fx_rate * (Decimal("1") + d_fx_markup) + d_fx_markup_abs
        d_agent_fee = self._to_decimal(
            getattr(supplier, "agent_fee", None) if agent_fee is None else agent_fee
        )
        d_vat = self._to_decimal(fixed.vat)
        d_money = self._to_decimal(fixed.money)
        d_storage = self._to_decimal(fixed.storage)
        d_move_novo = self._to_decimal(fixed.move_novo_tamozh)
        d_move_msk = self._to_decimal(fixed.move_tamozh_chekhov)
        d_customs_clearance = self._to_decimal(fixed.customs_clearance)
        d_bank_fee = self._to_decimal(fixed.bank_fee)
        d_customs_fee = self._to_decimal(fixed.customs_fee)
        d_additional_customs = self._to_decimal(fixed.additional_customs)
        d_excise = self._to_decimal(fixed.excise)
        d_eco_fee = self._to_decimal(fixed.eco_fee)

        supplier_is_rf = bool(supplier.is_rf)
        logistics = d_storage
        if not supplier_is_rf:
            logistics += d_move_msk
            if via_novo:
                logistics += d_move_novo

        cost_novo_wvat = (d_full_cost - logistics * (Decimal("1") + d_vat)) / (Decimal("1") + d_money)

        marking = Decimal("0") if supplier.marks_for_us else self.cost_calculation.get_marking_cost(product_id)
        customs_multiplier = Decimal("1") + d_customs_clearance if has_customs else Decimal("1")
        customs_and_insurance_multiplier = customs_multiplier + d_insurance

        base = cost_novo_wvat / (Decimal("1") + d_vat)
        if supplier_is_rf:
            numerator = base - marking - (d_agent_fee * d_fx_rate)
            denominator = (Decimal("1") + d_reexport) * customs_and_insurance_multiplier * d_effective_fx_rate
        else:
            numerator = (
                base
                - d_customs_fee
                - (d_excise if bool(product.is_excise) else Decimal("0"))
                - d_eco_fee
                - d_additional_customs
                - marking
                - (d_agent_fee * d_fx_rate)
            )
            denominator = (
                (Decimal("1") + d_reexport)
                * customs_and_insurance_multiplier
                * (Decimal("1") + d_bank_fee)
                * d_effective_fx_rate
            )

        if denominator == 0:
            raise ValueError("Деление на 0 в обратном расчете target price.")
        target_price_l = numerator / denominator - d_transport
        return self._round4(cost_novo_wvat), self._round4(target_price_l)

    def save_target_calculations(
        self,
        *,
        batch_id: str,
        imported_by: str,
        target_supplier_id: int,
        currency_code: str,
        fx_rate: Decimal,
        transport: Decimal,
        reexport: Decimal,
        insurance: Decimal,
        fx_markup: Decimal,
        fx_markup_abs: Decimal,
        has_customs: bool,
        via_novo: bool,
        manual_full_costs: dict[int, Decimal] | None = None,
    ) -> int:
        self.session.query(TargetPriceCalculation).filter(
            TargetPriceCalculation.batch_id == batch_id,
            TargetPriceCalculation.imported_by == imported_by,
        ).delete(synchronize_session=False)

        self.apply_manual_full_costs(batch_id, imported_by, manual_full_costs or {})
        self.ensure_single_options_selected(batch_id, imported_by)

        rows = self.session.query(TempTargetPriceImport).filter(
            TempTargetPriceImport.batch_id == batch_id,
            TempTargetPriceImport.imported_by == imported_by,
        ).order_by(TempTargetPriceImport.import_row_no.asc(), TempTargetPriceImport.id.asc()).all()

        option_ids = {
            int(row.selected_option_id)
            for row in rows
            if row.selected_option_id is not None
        }
        options_by_id = {
            int(option.id): option
            for option in (
                self.session.query(TempTargetPriceOption)
                .filter(TempTargetPriceOption.id.in_(option_ids))
                .all()
                if option_ids else []
            )
        }
        product_ids = {
            int(row.selected_product_id)
            for row in rows
            if row.selected_product_id is not None
        }
        self.cost_calculation.preload_reference_data(
            product_ids=product_ids,
            supplier_ids=(target_supplier_id,),
        )
        target_supplier = self.cost_calculation.get_supplier(target_supplier_id)
        fixed = self.cost_calculation.get_fixed_costs()

        saved = 0
        for row in rows:
            if row.selected_product_id is None or row.selected_option_id is None:
                raise ValueError("Для всех строк нужно выбрать Our Product Name и финального поставщика.")
            option = options_by_id.get(int(row.selected_option_id))
            if option is None:
                raise ValueError("Выбранный финальный поставщик не найден.")
            product = self.cost_calculation.get_product(row.selected_product_id)
            pack = self._to_decimal(product.pack)
            cost_novo_wvat, target_price_l = self.reverse_calculate_target_price(
                target_supplier_id=target_supplier_id,
                product_id=row.selected_product_id,
                full_cost_msk=option.full_cost_msk,
                currency_code=currency_code,
                fx_rate=fx_rate,
                transport=transport,
                reexport=reexport,
                insurance=insurance,
                fx_markup=fx_markup,
                fx_markup_abs=fx_markup_abs,
                has_customs=has_customs,
                via_novo=via_novo,
            )
            target_price_pack = self._round4(target_price_l * pack) if pack else Decimal("0")
            calc_row = TargetPriceCalculation(
                calc_date=datetime.utcnow(),
                batch_id=batch_id,
                imported_by=imported_by,
                import_row_no=row.import_row_no,
                target_supplier_id=target_supplier_id,
                donor_supplier_id=option.supplier_id,
                product_id=row.selected_product_id,
                supplier_article=row.supplier_article,
                supplier_product_name=row.product_name,
                target_price_l=target_price_l,
                target_price_pack=target_price_pack,
                currency_code=currency_code,
                fx_rate_used=fx_rate,
                full_cost_msk_source=option.full_cost_msk,
                cost_novo_wvat=cost_novo_wvat,
                fx_markup_used=fx_markup,
                fx_markup_abs_used=fx_markup_abs,
                transport_used=transport,
                reexport_used=reexport,
                insurance_used=insurance,
                agent_fee_used=self._to_decimal(getattr(target_supplier, "agent_fee", None)),
                has_customs_used=has_customs,
                via_novo_used=via_novo,
                bank_fee_used=option.bank_fee_used,
                customs_fee_used=option.customs_fee_used,
                additional_customs_used=option.additional_customs_used,
                storage_used=option.storage_used,
                move_novo_used=option.move_novo_used,
                move_msk_used=option.move_msk_used,
                marking_used=option.marking_used,
                is_excise_used=option.is_excise_used,
                vat_used=self._to_decimal(fixed.vat),
                money_used=self._to_decimal(fixed.money),
                price_date_used=option.price_date_used,
            )
            self.session.add(calc_row)
            saved += 1
        self.session.flush()
        return saved
