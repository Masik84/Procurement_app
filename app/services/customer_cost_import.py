from __future__ import annotations

from datetime import datetime
from decimal import Decimal
from sqlalchemy.orm import Session

from app.db.models import CustomerPriceCalculation, TempCustomerCostImport, TempCustomerCostOption
from app.services.cost_calculation import CostCalculationService
from app.services.price_repository import PriceRepository
from app.services.product_matching import ProductMatchingService
from app.services.supplier import SupplierService


class CustomerCostImportService:
    def __init__(self, session: Session):
        self.session = session
        self.product_matching = ProductMatchingService(session)
        self.price_repository = PriceRepository(session)
        self.cost_calculation = CostCalculationService(session)
        self.supplier_service = SupplierService(session)

    @staticmethod
    def _to_decimal(value: object) -> Decimal:
        if value is None:
            return Decimal("0")
        if isinstance(value, Decimal):
            return value
        return Decimal(str(value))

    def delete_temp_options(self, batch_id: str, imported_by: str) -> int:
        deleted_count = self.session.query(TempCustomerCostOption).filter(
            TempCustomerCostOption.batch_id == batch_id,
            TempCustomerCostOption.imported_by == imported_by,
        ).delete(synchronize_session=False)
        self.session.flush()
        return int(deleted_count or 0)

    def delete_temp_rows(self, batch_id: str, imported_by: str) -> int:
        self.delete_temp_options(batch_id, imported_by)
        deleted_count = self.session.query(TempCustomerCostImport).filter(
            TempCustomerCostImport.batch_id == batch_id,
            TempCustomerCostImport.imported_by == imported_by,
        ).delete(synchronize_session=False)
        self.session.flush()
        return int(deleted_count or 0)

    def import_rows(self, rows: list[dict], batch_id: str, imported_by: str, replace_existing: bool = True) -> int:
        if replace_existing:
            self.delete_temp_rows(batch_id, imported_by)

        count = 0
        for r in rows:
            row = TempCustomerCostImport(
                batch_id=batch_id,
                imported_by=imported_by,
                import_date=datetime.utcnow(),
                import_row_no=r.get("import_row_no"),
                request_date=r.get("RequestDate"),
                manager_name=r.get("ManagerName"),
                customer_name=r.get("CustomerName"),
                supplier_article=r.get("SupplierArticle"),
                product_name=r.get("ProductName"),
                pack=r.get("Pack"),
                qty_pcs=r.get("QtyPcs"),
                volume_l=r.get("VolumeL"),
                purchase_type=r.get("PurchaseType"),
                payment_terms=r.get("PaymentTerms"),
                comments=r.get("Comments"),
                selected_product_id=None,
                selected_option_id=None,
                new_product_name=None,
                new_brand=None,
                new_pack=None,
                new_is_excise=None,
            )
            self.session.add(row)
            count += 1

        self.session.flush()
        return count

    def get_temp_rows(self, batch_id: str, imported_by: str) -> list[TempCustomerCostImport]:
        return self.session.query(TempCustomerCostImport).filter(
            TempCustomerCostImport.batch_id == batch_id,
            TempCustomerCostImport.imported_by == imported_by,
        ).order_by(TempCustomerCostImport.import_row_no.asc(), TempCustomerCostImport.id.asc()).all()

    def automatch_temp_rows(self, batch_id: str, imported_by: str) -> int:
        rows = self.session.query(TempCustomerCostImport).filter(
            TempCustomerCostImport.batch_id == batch_id,
            TempCustomerCostImport.imported_by == imported_by,
            TempCustomerCostImport.selected_product_id.is_(None),
        ).order_by(TempCustomerCostImport.import_row_no.asc(), TempCustomerCostImport.id.asc()).all()

        matched_count = 0
        for row in rows:
            product = self.product_matching.find_customer_product(
                supplier_article=row.supplier_article,
                product_name=row.product_name,
                pack=row.pack,
            )
            if product is not None:
                row.selected_product_id = product.id
                matched_count += 1

        self.session.flush()
        return matched_count

    def validate_new_products_before_save(self, batch_id: str, imported_by: str) -> None:
        rows = self.session.query(TempCustomerCostImport).filter(
            TempCustomerCostImport.batch_id == batch_id,
            TempCustomerCostImport.imported_by == imported_by,
            TempCustomerCostImport.selected_product_id.is_(None),
            TempCustomerCostImport.new_product_name.isnot(None),
        ).all()

        for row in rows:
            if row.new_product_name is None or not str(row.new_product_name).strip():
                continue
            if row.new_is_excise is None:
                raise ValueError(f"Для нового продукта '{row.new_product_name}' не заполнено поле new_is_excise.")
            if row.new_brand is None or not str(row.new_brand).strip():
                raise ValueError(f"Для нового продукта '{row.new_product_name}' не заполнен new_brand.")
            if row.new_pack is None:
                raise ValueError(f"Для нового продукта '{row.new_product_name}' не заполнен new_pack.")

    def create_products_from_temp(self, batch_id: str, imported_by: str) -> int:
        rows = self.session.query(TempCustomerCostImport).filter(
            TempCustomerCostImport.batch_id == batch_id,
            TempCustomerCostImport.imported_by == imported_by,
            TempCustomerCostImport.selected_product_id.is_(None),
            TempCustomerCostImport.new_product_name.isnot(None),
            TempCustomerCostImport.new_brand.isnot(None),
            TempCustomerCostImport.new_pack.isnot(None),
            TempCustomerCostImport.new_is_excise.isnot(None),
        ).order_by(TempCustomerCostImport.import_row_no.asc(), TempCustomerCostImport.id.asc()).all()

        created_count = 0
        for row in rows:
            product = self.product_matching.get_or_create_product(
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
        rows = self.session.query(TempCustomerCostImport).filter(
            TempCustomerCostImport.batch_id == batch_id,
            TempCustomerCostImport.imported_by == imported_by,
            TempCustomerCostImport.selected_product_id.isnot(None),
        ).order_by(TempCustomerCostImport.import_row_no.asc(), TempCustomerCostImport.id.asc()).all()

        processed_count = 0
        for row in rows:
            self.product_matching.save_product_articles_by_split_articles(
                product_id=row.selected_product_id,
                supplier_article=row.supplier_article,
                supplier_product_name=row.product_name,
            )
            processed_count += 1

        self.session.flush()
        return processed_count

    def _get_fx_rate_for_currency(self, currency_code: str) -> Decimal:
        rate = self.supplier_service.get_rate_to_rub(currency_code)
        if rate is None or float(rate) == 0:
            raise ValueError(f"Для валюты '{currency_code}' не найден корректный курс rate_to_rub.")
        return self._to_decimal(rate)

    def build_supplier_options(self, batch_id: str, imported_by: str) -> int:
        self.delete_temp_options(batch_id, imported_by)

        rows = self.session.query(TempCustomerCostImport).filter(
            TempCustomerCostImport.batch_id == batch_id,
            TempCustomerCostImport.imported_by == imported_by,
            TempCustomerCostImport.selected_product_id.isnot(None),
        ).order_by(TempCustomerCostImport.import_row_no.asc(), TempCustomerCostImport.id.asc()).all()

        created_count = 0
        for row in rows:
            seen_supplier_ids: set[int] = set()

            current_prices = self.price_repository.get_suppliers_with_current_prices_for_product(
                product_id=row.selected_product_id,
                only_rating_calc=True,
            )
            for supplier_price in current_prices:
                self._create_option_from_snapshot(row, batch_id, imported_by, supplier_price)
                seen_supplier_ids.add(supplier_price.supplier_id)
                created_count += 1

            latest_history = self.price_repository.get_latest_history_prices_for_product(
                product_id=row.selected_product_id,
                only_rating_calc=True,
            )
            for supplier_price in latest_history:
                if supplier_price.supplier_id in seen_supplier_ids:
                    continue
                self._create_option_from_snapshot(row, batch_id, imported_by, supplier_price)
                seen_supplier_ids.add(supplier_price.supplier_id)
                created_count += 1

        self.session.flush()
        self.rank_supplier_options(batch_id, imported_by)
        self.select_best_options(batch_id, imported_by)
        return created_count

    def _create_option_from_snapshot(self, row: TempCustomerCostImport, batch_id: str, imported_by: str, supplier_price) -> None:
        calc = self.cost_calculation.calculate_supplier_costs(
            supplier_id=supplier_price.supplier_id,
            product_id=row.selected_product_id,
            supplier_price=supplier_price.price,
            fx_rate=self._get_fx_rate_for_currency(supplier_price.currency_code),
            currency_code=supplier_price.currency_code,
        )
        option = TempCustomerCostOption(
            temp_import_id=row.id,
            batch_id=batch_id,
            imported_by=imported_by,
            calc_date=datetime.utcnow(),
            supplier_id=supplier_price.supplier_id,
            product_id=row.selected_product_id,
            supplier_name=supplier_price.supplier_name,
            supplier_article=row.supplier_article,
            supplier_product_name=row.product_name,
            supplier_price=calc.supplier_price,
            price_date_used=supplier_price.price_date,
            cost_novo_wvat=calc.cost_novo_wvat,
            full_cost_msk=calc.full_cost_msk,
            currency_code=calc.currency_code,
            fx_rate_used=calc.fx_rate_used,
            fx_markup_used=calc.fx_markup_used,
            transport_used=calc.transport_used,
            reexport_used=calc.reexport_used,
            has_customs_used=calc.has_customs_used,
            via_novo_used=calc.via_novo_used,
            bank_fee_used=calc.bank_fee_used,
            customs_fee_used=calc.customs_fee_used,
            move_novo_used=calc.move_novo_used,
            move_msk_used=calc.move_msk_used,
            is_excise_used=calc.is_excise_used,
            additional_customs_used=calc.additional_customs_used,
            storage_used=calc.storage_used,
            marking_used=calc.marking_used,
            opt_rank=None,
        )
        self.session.add(option)

    def rank_supplier_options(self, batch_id: str, imported_by: str) -> None:
        rows = self.session.query(TempCustomerCostImport).filter(
            TempCustomerCostImport.batch_id == batch_id,
            TempCustomerCostImport.imported_by == imported_by,
        ).all()
        for row in rows:
            options = self.session.query(TempCustomerCostOption).filter(
                TempCustomerCostOption.batch_id == batch_id,
                TempCustomerCostOption.imported_by == imported_by,
                TempCustomerCostOption.temp_import_id == row.id,
            ).order_by(
                TempCustomerCostOption.full_cost_msk.asc(),
                TempCustomerCostOption.supplier_name.asc(),
                TempCustomerCostOption.id.asc(),
            ).all()
            rank = 1
            for option in options:
                option.opt_rank = rank
                rank += 1
        self.session.flush()

    def select_best_options(self, batch_id: str, imported_by: str) -> None:
        rows = self.session.query(TempCustomerCostImport).filter(
            TempCustomerCostImport.batch_id == batch_id,
            TempCustomerCostImport.imported_by == imported_by,
        ).all()
        for row in rows:
            best_option = self.session.query(TempCustomerCostOption).filter(
                TempCustomerCostOption.batch_id == batch_id,
                TempCustomerCostOption.imported_by == imported_by,
                TempCustomerCostOption.temp_import_id == row.id,
            ).order_by(
                TempCustomerCostOption.opt_rank.asc(),
                TempCustomerCostOption.id.asc(),
            ).first()
            row.selected_option_id = best_option.id if best_option is not None else None
        self.session.flush()

    def delete_saved_calculations(self, batch_id: str, imported_by: str) -> int:
        deleted_count = self.session.query(CustomerPriceCalculation).filter(
            CustomerPriceCalculation.batch_id == batch_id,
            CustomerPriceCalculation.imported_by == imported_by,
        ).delete(synchronize_session=False)
        self.session.flush()
        return int(deleted_count or 0)

    def save_calculations(self, batch_id: str, imported_by: str) -> int:
        self.delete_saved_calculations(batch_id, imported_by)
        rows = self.session.query(TempCustomerCostImport).filter(
            TempCustomerCostImport.batch_id == batch_id,
            TempCustomerCostImport.imported_by == imported_by,
        ).order_by(TempCustomerCostImport.import_row_no.asc(), TempCustomerCostImport.id.asc()).all()

        saved_count = 0
        for row in rows:
            if row.selected_product_id is None or row.selected_option_id is None:
                continue
            option = self.session.query(TempCustomerCostOption).filter(
                TempCustomerCostOption.id == row.selected_option_id
            ).first()
            if option is None:
                continue

            calc_row = CustomerPriceCalculation(
                calc_date=datetime.utcnow(),
                batch_id=batch_id,
                imported_by=imported_by,
                manager_name=row.manager_name,
                customer_name=row.customer_name,
                request_date=row.request_date,
                supplier_id=option.supplier_id,
                product_id=row.selected_product_id,
                supplier_article=row.supplier_article,
                supplier_product_name=row.product_name,
                pack=row.pack,
                qty_pcs=row.qty_pcs,
                volume_l=row.volume_l,
                comments=row.comments,
                supplier_price=option.supplier_price,
                cost_novo_wvat=option.cost_novo_wvat,
                full_cost_msk=option.full_cost_msk,
                currency_code=option.currency_code,
                fx_rate_used=option.fx_rate_used,
                fx_markup_used=option.fx_markup_used,
                transport_used=option.transport_used,
                reexport_used=option.reexport_used,
                has_customs_used=option.has_customs_used,
                via_novo_used=option.via_novo_used,
                bank_fee_used=option.bank_fee_used,
                customs_fee_used=option.customs_fee_used,
                additional_customs_used=option.additional_customs_used,
                storage_used=option.storage_used,
                move_novo_used=option.move_novo_used,
                move_msk_used=option.move_msk_used,
                marking_used=option.marking_used,
                is_excise_used=option.is_excise_used,
                price_date_used=option.price_date_used,
                import_row_no=row.import_row_no,
            )
            self.session.add(calc_row)
            saved_count += 1

        self.session.flush()
        return saved_count

    def run_calculation(self, batch_id: str, imported_by: str) -> dict:
        self.validate_new_products_before_save(batch_id, imported_by)
        created_products_count = self.create_products_from_temp(batch_id, imported_by)
        product_articles_count = self.create_or_update_product_articles(batch_id, imported_by)
        options_count = self.build_supplier_options(batch_id, imported_by)
        return {
            "created_products_count": created_products_count,
            "product_articles_count": product_articles_count,
            "options_count": options_count,
        }
