from __future__ import annotations

from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
import uuid

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
from app.utils.text import clean_multi_spaces


class TargetPriceService:
    def __init__(self, session: Session):
        self.session = session
        self.product_matching = ProductMatchingService(session)
        self.price_repository = PriceRepository(session)
        self.cost_calculation = CostCalculationService(session)
        self.supplier_service = SupplierService(session)
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

    def cleanup_old_temp_rows(self, imported_by: str) -> int:
        # Keep behaviour conservative: remove only current user's old batches not referenced by active page.
        deleted_options = self.session.query(TempTargetPriceOption).filter(
            TempTargetPriceOption.imported_by == imported_by
        ).delete(synchronize_session=False)
        deleted_rows = self.session.query(TempTargetPriceImport).filter(
            TempTargetPriceImport.imported_by == imported_by
        ).delete(synchronize_session=False)
        self.session.flush()
        return int(deleted_options or 0) + int(deleted_rows or 0)

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

        created = 0
        for row in rows:
            seen: set[int] = set()
            current_prices = self.price_repository.get_suppliers_with_current_prices_for_product(
                product_id=row.selected_product_id,
                only_rating_calc=True,
                min_price_date=min_price_date,
            )
            for supplier_price in current_prices:
                self._create_option_from_snapshot(row, batch_id, imported_by, supplier_price)
                seen.add(supplier_price.supplier_id)
                created += 1
            latest_history = self.price_repository.get_latest_history_prices_for_product(
                product_id=row.selected_product_id,
                only_rating_calc=True,
                min_price_date=min_price_date,
            )
            for supplier_price in latest_history:
                if supplier_price.supplier_id in seen:
                    continue
                self._create_option_from_snapshot(row, batch_id, imported_by, supplier_price)
                seen.add(supplier_price.supplier_id)
                created += 1

        self.session.flush()
        self.rank_supplier_options(batch_id, imported_by)
        self.select_best_options(batch_id, imported_by)
        return created

    def _create_option_from_snapshot(self, row: TempTargetPriceImport, batch_id: str, imported_by: str, supplier_price) -> None:
        calc = self.cost_calculation.calculate_supplier_costs(
            supplier_id=supplier_price.supplier_id,
            product_id=row.selected_product_id,
            supplier_price=supplier_price.price,
            fx_rate=self._get_fx_rate_for_currency(supplier_price.currency_code),
            currency_code=supplier_price.currency_code,
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
            transport_used=safe(calc.transport_used),
            reexport_used=safe(calc.reexport_used),
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

    def rank_supplier_options(self, batch_id: str, imported_by: str) -> None:
        rows = self.session.query(TempTargetPriceImport).filter(
            TempTargetPriceImport.batch_id == batch_id,
            TempTargetPriceImport.imported_by == imported_by,
        ).all()
        for row in rows:
            options = self.session.query(TempTargetPriceOption).filter(
                TempTargetPriceOption.batch_id == batch_id,
                TempTargetPriceOption.imported_by == imported_by,
                TempTargetPriceOption.temp_import_id == row.id,
            ).order_by(
                TempTargetPriceOption.cost_novo_wvat.asc(),
                TempTargetPriceOption.full_cost_msk.asc(),
                TempTargetPriceOption.supplier_name.asc(),
                TempTargetPriceOption.id.asc(),
            ).all()
            for rank, option in enumerate(options, start=1):
                option.opt_rank = rank
        self.session.flush()

    def select_best_options(self, batch_id: str, imported_by: str) -> None:
        rows = self.session.query(TempTargetPriceImport).filter(
            TempTargetPriceImport.batch_id == batch_id,
            TempTargetPriceImport.imported_by == imported_by,
        ).all()
        for row in rows:
            options = self.session.query(TempTargetPriceOption).filter(
                TempTargetPriceOption.batch_id == batch_id,
                TempTargetPriceOption.imported_by == imported_by,
                TempTargetPriceOption.temp_import_id == row.id,
            ).order_by(TempTargetPriceOption.opt_rank.asc(), TempTargetPriceOption.id.asc()).all()
            row.selected_option_id = options[0].id if len(options) == 1 else None
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
        fixed = self.session.query(FixedCosts).order_by(FixedCosts.id.asc()).first()
        created_or_updated = 0
        for row_id, value in manual_full_costs.items():
            full_cost = self._to_decimal(value)
            if full_cost == 0:
                continue
            row = self.session.query(TempTargetPriceImport).filter(
                TempTargetPriceImport.id == int(row_id),
                TempTargetPriceImport.batch_id == batch_id,
                TempTargetPriceImport.imported_by == imported_by,
            ).first()
            if row is None or row.selected_product_id is None:
                continue
            # Manual overrides all ordinary supplier options for this row.
            self.session.query(TempTargetPriceOption).filter(
                TempTargetPriceOption.batch_id == batch_id,
                TempTargetPriceOption.imported_by == imported_by,
                TempTargetPriceOption.temp_import_id == row.id,
            ).delete(synchronize_session=False)
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
                full_cost_msk=full_cost,
                currency_code=manual_supplier.base_currency or "USD",
                fx_rate_used=Decimal("0"),
                fx_markup_used=Decimal("0"),
                transport_used=Decimal("0"),
                reexport_used=Decimal("0"),
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
            self.session.flush()
            row.selected_option_id = option.id
            created_or_updated += 1
        self.session.flush()
        return created_or_updated

    def ensure_single_options_selected(self, batch_id: str, imported_by: str) -> None:
        rows = self.session.query(TempTargetPriceImport).filter(
            TempTargetPriceImport.batch_id == batch_id,
            TempTargetPriceImport.imported_by == imported_by,
        ).all()
        for row in rows:
            if row.selected_option_id is not None:
                continue
            options = self.session.query(TempTargetPriceOption).filter(
                TempTargetPriceOption.batch_id == batch_id,
                TempTargetPriceOption.imported_by == imported_by,
                TempTargetPriceOption.temp_import_id == row.id,
            ).order_by(TempTargetPriceOption.opt_rank.asc(), TempTargetPriceOption.id.asc()).all()
            if len(options) == 1:
                row.selected_option_id = options[0].id
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
        fx_markup: object,
        has_customs: bool,
        via_novo: bool,
    ) -> tuple[Decimal, Decimal]:
        supplier = self.session.query(Supplier).filter(Supplier.id == target_supplier_id).first()
        if supplier is None:
            raise ValueError("Поставщик для target price не найден.")
        product = self.session.query(Product).filter(Product.id == product_id).first()
        if product is None:
            raise ValueError("Продукт для target price не найден.")
        fixed = self.session.query(FixedCosts).order_by(FixedCosts.id.asc()).first()
        if fixed is None:
            raise ValueError("В таблице fixed_costs нет данных.")

        d_full_cost = self._to_decimal(full_cost_msk)
        d_fx_rate = self._to_decimal(fx_rate)
        if d_fx_rate == 0:
            raise ValueError("Курс валюты не может быть 0.")
        d_transport = self._to_decimal(transport)
        d_reexport = self._to_decimal(reexport)
        d_fx_markup = self._to_decimal(fx_markup)
        d_agent_fee = self._to_decimal(getattr(supplier, "agent_fee", None))
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

        base = cost_novo_wvat / (Decimal("1") + d_vat)
        if supplier_is_rf:
            numerator = base - marking - (d_agent_fee * d_fx_rate)
            denominator = (Decimal("1") + d_reexport) * customs_multiplier * d_fx_rate * (Decimal("1") + d_fx_markup)
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
                * customs_multiplier
                * (Decimal("1") + d_bank_fee)
                * d_fx_rate
                * (Decimal("1") + d_fx_markup)
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
        fx_markup: Decimal,
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

        saved = 0
        for row in rows:
            if row.selected_product_id is None or row.selected_option_id is None:
                raise ValueError("Для всех строк нужно выбрать Our Product Name и финального поставщика.")
            option = self.session.query(TempTargetPriceOption).filter(TempTargetPriceOption.id == row.selected_option_id).first()
            if option is None:
                raise ValueError("Выбранный финальный поставщик не найден.")
            product = self.session.query(Product).filter(Product.id == row.selected_product_id).first()
            pack = self._to_decimal(product.pack if product else 0)
            cost_novo_wvat, target_price_l = self.reverse_calculate_target_price(
                target_supplier_id=target_supplier_id,
                product_id=row.selected_product_id,
                full_cost_msk=option.full_cost_msk,
                currency_code=currency_code,
                fx_rate=fx_rate,
                transport=transport,
                reexport=reexport,
                fx_markup=fx_markup,
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
                cost_novo_wvat_recalculated=cost_novo_wvat,
                fx_markup_used=fx_markup,
                transport_used=transport,
                reexport_used=reexport,
                agent_fee_used=self._to_decimal(getattr(self.session.query(Supplier).filter(Supplier.id == target_supplier_id).first(), "agent_fee", None)),
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
                vat_used=self._to_decimal(self.cost_calculation.get_fixed_costs().vat),
                money_used=self._to_decimal(self.cost_calculation.get_fixed_costs().money),
                price_date_used=option.price_date_used,
            )
            self.session.add(calc_row)
            saved += 1
        self.session.flush()
        return saved
