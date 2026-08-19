from __future__ import annotations

from collections import defaultdict
from datetime import date, datetime
import uuid
from pathlib import Path
from decimal import Decimal
from sqlalchemy.orm import Session

from app.exports.product_stock_exporter import ProductStockExporter
from app.imports.is_importer import ISImporter
from app.imports.stock_importer import StockImporter
from app.imports.supplier_orders_importer import SupplierOrdersImporter
from app.db.models import Product, ProductArticle, ProductStock, TempIsImport, TempStockImport, TempSupplierOrdersImport
from app.services.product_matching_service import ProductCreateData, ProductMatchingService
from app.services.sales_stock_metrics_service import ProductStockMetrics, SalesStockMetricsService
from app.services.temp_cleanup_service import TempCleanupService
from app.utils.text import clean_multi_spaces


class ProductStockService:
    AUTOMATCH_FLUSH_BATCH_SIZE = 250

    def __init__(self, session: Session):
        self.session = session
        self.product_matching = ProductMatchingService(session)
        self.stock_importer = StockImporter()
        self.supplier_orders_importer = SupplierOrdersImporter()
        self.is_importer = ISImporter()
        self.exporter = ProductStockExporter(session)
        self.sales_metrics_service = SalesStockMetricsService(session)

    @staticmethod
    def _to_decimal(value: object) -> Decimal:
        if value is None:
            return Decimal("0")
        if isinstance(value, Decimal):
            return value
        return Decimal(str(value))

    def cleanup_old_temp_rows(self, imported_by: str | None = None, before_date: date | None = None) -> int:
        # Daily cleanup is global by date: all temp rows older than today are stale.
        # Current user temp rows are cleaned after successful save.
        return TempCleanupService(self.session).cleanup_old_for_all(before_date=before_date)

    def cleanup_old_temp_stock(self) -> int:
        return self.cleanup_old_temp_rows()

    def cleanup_old_temp_supplier_orders(self) -> int:
        return self.cleanup_old_temp_rows()

    def cleanup_old_temp_is(self) -> int:
        return self.cleanup_old_temp_rows()

    def delete_stock_rows(self, batch_id: str, imported_by: str) -> int:
        deleted = self.session.query(TempStockImport).filter(
            TempStockImport.batch_id == batch_id,
            TempStockImport.imported_by == imported_by,
        ).delete(synchronize_session=False)
        self.session.flush()
        return int(deleted or 0)

    def delete_stock_rows_for_user(self, imported_by: str) -> int:
        return TempCleanupService(self.session).delete_current_user(
            imported_by=imported_by,
            tables=(TempStockImport,),
        )

    def delete_supplier_orders_rows(self, batch_id: str, imported_by: str) -> int:
        deleted = self.session.query(TempSupplierOrdersImport).filter(
            TempSupplierOrdersImport.batch_id == batch_id,
            TempSupplierOrdersImport.imported_by == imported_by,
        ).delete(synchronize_session=False)
        self.session.flush()
        return int(deleted or 0)

    def delete_supplier_orders_rows_for_user(self, imported_by: str) -> int:
        return TempCleanupService(self.session).delete_current_user(
            imported_by=imported_by,
            tables=(TempSupplierOrdersImport,),
        )

    def delete_is_rows(self, batch_id: str, imported_by: str) -> int:
        deleted = self.session.query(TempIsImport).filter(
            TempIsImport.batch_id == batch_id,
            TempIsImport.imported_by == imported_by,
        ).delete(synchronize_session=False)
        self.session.flush()
        return int(deleted or 0)

    def delete_is_rows_for_user(self, imported_by: str) -> int:
        return TempCleanupService(self.session).delete_current_user(
            imported_by=imported_by,
            tables=(TempIsImport,),
        )

    def import_stock_rows(self, rows: list[dict], batch_id: str, imported_by: str, replace_existing: bool = True) -> int:
        if replace_existing:
            self.delete_stock_rows(batch_id, imported_by)

        created = []
        now = datetime.now()

        for r in rows:
            created.append(TempStockImport(
                batch_id=batch_id,
                imported_by=imported_by,
                import_date=now,
                import_row_no=r.get("import_row_no"),
                source_article=r.get("source_article"),
                source_sku=r.get("source_sku"),
                source_product_name=r.get("source_product_name"),
                source_origin=r.get("source_origin"),
                source_brand_group=r.get("source_brand_group"),
                abc_category=clean_multi_spaces(r.get("abc_category")) or "-",
                lpc=r.get("lpc"),
                landed_cost=r.get("landed_cost"),
                distr_price=r.get("distr_price"),
                promo_price=r.get("promo_price"),
                stock_qty=r.get("stock_qty") or 0,
                transit_qty=r.get("transit_qty") or 0,
                markdown_qty=r.get("markdown_qty") or 0,
                reserve_qty=r.get("reserve_qty") or 0,
                reserve_ecomm_qty=r.get("reserve_ecomm_qty") or 0,
                selected_product_id=None,
                has_lpc_warning=bool(r.get("has_lpc_warning")),
                new_product_name=r.get("new_product_name"),
                new_brand=r.get("new_brand"),
                new_pack=r.get("new_pack"),
                new_is_excise=r.get("new_is_excise"),
            ))

        if created:
            self.session.add_all(created)
        self.session.flush()
        return len(created)

    def import_supplier_orders_rows(self, rows: list[dict], batch_id: str, imported_by: str, replace_existing: bool = True) -> int:
        if replace_existing:
            self.delete_supplier_orders_rows(batch_id, imported_by)

        created = []
        now = datetime.now()

        for r in rows:
            created.append(TempSupplierOrdersImport(
                batch_id=batch_id,
                imported_by=imported_by,
                import_date=now,
                import_row_no=r.get("import_row_no"),
                source_article=r.get("source_article"),
                source_our_product_name=r.get("source_our_product_name"),
                source_product_name=r.get("source_product_name"),
                abc_category=clean_multi_spaces(r.get("abc_category")) or "-",
                order_qty=r.get("order_qty") or 0,
                is_order_qty=r.get("is_order_qty") or 0,
                is_confirmed_order_qty=r.get("is_confirmed_order_qty") or 0,
                selected_product_id=None,
                new_product_name=r.get("new_product_name"),
                new_brand=r.get("new_brand"),
                new_pack=r.get("new_pack"),
                new_is_excise=r.get("new_is_excise"),
            ))

        if created:
            self.session.add_all(created)
        self.session.flush()
        return len(created)

    def import_is_rows(self, rows: list[dict], batch_id: str, imported_by: str, replace_existing: bool = True) -> int:
        if replace_existing:
            self.delete_is_rows(batch_id, imported_by)

        created = []
        now = datetime.now()

        for r in rows:
            created.append(TempIsImport(
                batch_id=batch_id,
                imported_by=imported_by,
                import_date=now,
                import_row_no=r.get("import_row_no"),
                source_article=r.get("source_article"),
                source_product_name=r.get("source_product_name"),
                confirmed_qty=r.get("confirmed_qty") or 0,
                remains_qty=r.get("remains_qty") or 0,
                stock_qty=r.get("stock_qty") or 0,
                selected_product_id=None,
                new_product_name=None,
                new_brand=None,
                new_pack=None,
                new_is_excise=None,
            ))

        if created:
            self.session.add_all(created)
        self.session.flush()
        return len(created)

    def _validate_new_rows(self, rows, name_attr: str = "new_product_name"):
        for row in rows:
            if getattr(row, name_attr) is None or not str(getattr(row, name_attr)).strip():
                continue

            self.product_matching.validate_new_product_fields(
                product_name=row.new_product_name,
                brand=row.new_brand,
                pack=row.new_pack,
                is_excise=row.new_is_excise,
            )

    def validate_new_stock_products_before_save(self, batch_id: str, imported_by: str) -> None:
        rows = self.session.query(TempStockImport).filter(
            TempStockImport.batch_id == batch_id,
            TempStockImport.imported_by == imported_by,
            TempStockImport.selected_product_id.is_(None),
            TempStockImport.new_product_name.isnot(None),
        ).all()
        self._validate_new_rows(rows)

    def validate_new_supplier_orders_products_before_save(self, batch_id: str, imported_by: str) -> None:
        rows = self.session.query(TempSupplierOrdersImport).filter(
            TempSupplierOrdersImport.batch_id == batch_id,
            TempSupplierOrdersImport.imported_by == imported_by,
            TempSupplierOrdersImport.selected_product_id.is_(None),
            TempSupplierOrdersImport.new_product_name.isnot(None),
        ).all()
        self._validate_new_rows(rows)

    def validate_new_is_products_before_save(self, batch_id: str, imported_by: str) -> None:
        rows = self.session.query(TempIsImport).filter(
            TempIsImport.batch_id == batch_id,
            TempIsImport.imported_by == imported_by,
            TempIsImport.selected_product_id.is_(None),
            TempIsImport.new_product_name.isnot(None),
        ).all()
        self._validate_new_rows(rows)

    def _create_products_from_rows(self, rows, id_attr: str):
        products = self.product_matching.get_or_create_products_batch([
            ProductCreateData(
                name=row.new_product_name,
                brand=row.new_brand,
                pack=row.new_pack,
                is_excise=bool(row.new_is_excise),
            )
            for row in rows
        ])

        for row, product in zip(rows, products):
            setattr(row, "selected_product_id", product.id)

        self.product_matching.create_product_articles_if_missing_batch([
            (
                int(product.id),
                getattr(row, "source_article", None),
                getattr(row, "source_product_name", None),
            )
            for row, product in zip(rows, products)
        ])
        self.session.flush()
        return len(rows)

    @staticmethod
    def can_create_new_product(row: object) -> bool:
        return bool(
            clean_multi_spaces(getattr(row, "new_product_name", None))
            and clean_multi_spaces(getattr(row, "new_brand", None))
            and getattr(row, "new_pack", None) is not None
            and getattr(row, "new_is_excise", None) is not None
        )

    def create_stock_products_from_temp(self, batch_id: str, imported_by: str) -> int:
        rows = self.session.query(TempStockImport).filter(
            TempStockImport.batch_id == batch_id,
            TempStockImport.imported_by == imported_by,
            TempStockImport.selected_product_id.is_(None),
            TempStockImport.new_product_name.isnot(None),
            TempStockImport.new_brand.isnot(None),
            TempStockImport.new_pack.isnot(None),
            TempStockImport.new_is_excise.isnot(None),
        ).order_by(TempStockImport.import_row_no.asc(), TempStockImport.id.asc()).all()
        return self._create_products_from_rows(rows, "id")

    def create_supplier_orders_products_from_temp(self, batch_id: str, imported_by: str) -> int:
        rows = self.session.query(TempSupplierOrdersImport).filter(
            TempSupplierOrdersImport.batch_id == batch_id,
            TempSupplierOrdersImport.imported_by == imported_by,
            TempSupplierOrdersImport.selected_product_id.is_(None),
            TempSupplierOrdersImport.new_product_name.isnot(None),
            TempSupplierOrdersImport.new_brand.isnot(None),
            TempSupplierOrdersImport.new_pack.isnot(None),
            TempSupplierOrdersImport.new_is_excise.isnot(None),
        ).order_by(TempSupplierOrdersImport.import_row_no.asc(), TempSupplierOrdersImport.id.asc()).all()
        return self._create_products_from_rows(rows, "id")

    def create_is_products_from_temp(self, batch_id: str, imported_by: str) -> int:
        rows = self.session.query(TempIsImport).filter(
            TempIsImport.batch_id == batch_id,
            TempIsImport.imported_by == imported_by,
            TempIsImport.selected_product_id.is_(None),
            TempIsImport.new_product_name.isnot(None),
            TempIsImport.new_brand.isnot(None),
            TempIsImport.new_pack.isnot(None),
            TempIsImport.new_is_excise.isnot(None),
        ).order_by(TempIsImport.import_row_no.asc(), TempIsImport.id.asc()).all()
        return self._create_products_from_rows(rows, "id")

    def _flush_automatch_batch(self, matched_count: int) -> None:
        if matched_count and matched_count % self.AUTOMATCH_FLUSH_BATCH_SIZE == 0:
            self.session.flush()

    def _flush_automatch_remainder(self, matched_count: int) -> None:
        if matched_count % self.AUTOMATCH_FLUSH_BATCH_SIZE:
            self.session.flush()

    @staticmethod
    def _assign_matched_product(row, product_id: int) -> None:
        row.selected_product_id = product_id
        row.new_product_name = None
        row.new_brand = None
        row.new_pack = None
        row.new_is_excise = None

    def automatch_stock_rows(self, batch_id: str, imported_by: str) -> int:
        rows = self.session.query(TempStockImport).filter(
            TempStockImport.batch_id == batch_id,
            TempStockImport.imported_by == imported_by,
            TempStockImport.selected_product_id.is_(None),
        ).order_by(TempStockImport.import_row_no.asc(), TempStockImport.id.asc()).all()

        matched = 0
        for row in rows:
            product = self.product_matching.find_stock_product(
                source_article=row.source_article,
                source_product_name=row.source_product_name,
            )
            if product is not None:
                self._assign_matched_product(row, product.id)
                matched += 1
                self._flush_automatch_batch(matched)
        self._flush_automatch_remainder(matched)
        return matched

    def automatch_supplier_orders_rows(self, batch_id: str, imported_by: str) -> int:
        rows = self.session.query(TempSupplierOrdersImport).filter(
            TempSupplierOrdersImport.batch_id == batch_id,
            TempSupplierOrdersImport.imported_by == imported_by,
            TempSupplierOrdersImport.selected_product_id.is_(None),
        ).order_by(TempSupplierOrdersImport.import_row_no.asc(), TempSupplierOrdersImport.id.asc()).all()

        matched = 0
        for row in rows:
            product = self.product_matching.find_by_normalized_product_name(row.source_our_product_name)
            if product is None:
                product = self.product_matching.find_stock_product(
                    source_article=row.source_article,
                    source_product_name=row.source_product_name,
                )
            if product is not None:
                self._assign_matched_product(row, product.id)
                matched += 1
                self._flush_automatch_batch(matched)
        self._flush_automatch_remainder(matched)
        return matched

    def automatch_is_rows(self, batch_id: str, imported_by: str) -> int:
        rows = self.session.query(TempIsImport).filter(
            TempIsImport.batch_id == batch_id,
            TempIsImport.imported_by == imported_by,
            TempIsImport.selected_product_id.is_(None),
        ).order_by(TempIsImport.import_row_no.asc(), TempIsImport.id.asc()).all()

        matched = 0
        for row in rows:
            product = self.product_matching.find_is_product(
                source_article=row.source_article,
                source_product_name=row.source_product_name,
            )
            if product is not None:
                self._assign_matched_product(row, product.id)
                matched += 1
                self._flush_automatch_batch(matched)
        self._flush_automatch_remainder(matched)
        return matched

    def _create_or_update_articles(self, rows) -> int:
        matched_rows = [row for row in rows if row.selected_product_id]
        self.product_matching.create_product_articles_if_missing_batch([
            (
                int(row.selected_product_id),
                getattr(row, "source_article", None),
                getattr(row, "source_product_name", None),
            )
            for row in matched_rows
        ])
        self.session.flush()
        return len(matched_rows)

    def create_or_update_stock_product_articles(self, batch_id: str, imported_by: str) -> int:
        rows = self.session.query(TempStockImport).filter(
            TempStockImport.batch_id == batch_id,
            TempStockImport.imported_by == imported_by,
            TempStockImport.selected_product_id.isnot(None),
        ).all()
        return self._create_or_update_articles(rows)

    def create_or_update_supplier_orders_product_articles(self, batch_id: str, imported_by: str) -> int:
        rows = self.session.query(TempSupplierOrdersImport).filter(
            TempSupplierOrdersImport.batch_id == batch_id,
            TempSupplierOrdersImport.imported_by == imported_by,
            TempSupplierOrdersImport.selected_product_id.isnot(None),
        ).order_by(TempSupplierOrdersImport.import_row_no.asc(), TempSupplierOrdersImport.id.asc()).all()
        return self._create_or_update_articles(rows)

    def create_or_update_is_product_articles(self, batch_id: str, imported_by: str) -> int:
        rows = self.session.query(TempIsImport).filter(
            TempIsImport.batch_id == batch_id,
            TempIsImport.imported_by == imported_by,
            TempIsImport.selected_product_id.isnot(None),
        ).all()
        return self._create_or_update_articles(rows)

    def _load_products_map(self, product_ids: list[int]) -> dict[int, Product]:
        if not product_ids:
            return {}
        rows = self.session.query(Product).filter(Product.id.in_(product_ids)).all()
        return {int(row.id): row for row in rows}

    def _load_product_stock_map(self, product_ids: list[int]) -> dict[int, ProductStock]:
        if not product_ids:
            return {}
        rows = self.session.query(ProductStock).filter(ProductStock.product_id.in_(product_ids)).all()
        return {int(row.product_id): row for row in rows}

    def _is_origin_ru(self, value: object) -> bool:
        s = str(value or "").strip().upper()
        return s in {"RU", "РОССИЯ", "RUSSIA"} or s.startswith("RU")


    def _abc_category_by_product(self, batch_id: str, imported_by: str, product_id: int) -> str:
        rows = self.session.query(TempStockImport.abc_category).filter(
            TempStockImport.batch_id == batch_id,
            TempStockImport.imported_by == imported_by,
            TempStockImport.selected_product_id == product_id,
        ).order_by(TempStockImport.import_row_no.asc(), TempStockImport.id.asc()).all()

        return self._abc_category_from_values(value for (value,) in rows)

    @staticmethod
    def _abc_category_from_values(values) -> str:
        for value in values:
            category = clean_multi_spaces(value)
            if category and category != "-":
                return category
        return "-"

    def _supplier_orders_abc_category_by_product(
        self,
        batch_id: str,
        imported_by: str,
        product_id: int,
    ) -> str:
        rows = self.session.query(TempSupplierOrdersImport.abc_category).filter(
            TempSupplierOrdersImport.batch_id == batch_id,
            TempSupplierOrdersImport.imported_by == imported_by,
            TempSupplierOrdersImport.selected_product_id == product_id,
        ).order_by(TempSupplierOrdersImport.import_row_no.asc(), TempSupplierOrdersImport.id.asc()).all()

        for (value,) in rows:
            category = clean_multi_spaces(value)
            if category and category != "-":
                return category
        return "-"

    def _min_price_by_product(self, batch_id: str, imported_by: str, product_id: int, field_name: str) -> Decimal:
        rows = self.session.query(TempStockImport).filter(
            TempStockImport.batch_id == batch_id,
            TempStockImport.imported_by == imported_by,
            TempStockImport.selected_product_id == product_id,
        ).all()

        return self._min_price_from_rows(rows, field_name)

    def _min_price_from_rows(self, rows, field_name: str) -> Decimal:
        vals_any = []
        vals_import_non_ru = []
        positive_any = []
        positive_import_non_ru = []

        for row in rows:
            price_val = getattr(row, field_name)
            if price_val is None:
                continue
            d = self._to_decimal(price_val)
            vals_any.append(d)
            if d > 0:
                positive_any.append(d)
            brand_group = str(row.source_brand_group or "").strip().lower()
            if brand_group == "import" and not self._is_origin_ru(row.source_origin):
                vals_import_non_ru.append(d)
                if d > 0:
                    positive_import_non_ru.append(d)

        # Empty prices are imported as 0. They must not make the final distributor/promo
        # price equal to 0 when the same product also has a valid positive price.
        if positive_import_non_ru:
            return min(positive_import_non_ru)
        if positive_any:
            return min(positive_any)
        if vals_import_non_ru:
            return min(vals_import_non_ru)
        if vals_any:
            return min(vals_any)
        return Decimal("0")


    def _apply_sales_metrics_to_product_stock(
        self,
        metrics_by_product: dict[int, ProductStockMetrics],
        now: datetime,
    ) -> int:
        if not metrics_by_product:
            return 0

        product_ids = list(metrics_by_product.keys())
        products = self._load_products_map(product_ids)
        stocks = self._load_product_stock_map(product_ids)

        saved = 0
        for product_id, metric in metrics_by_product.items():
            product = products.get(product_id)
            if product is None:
                continue

            stock = stocks.get(product_id)
            if stock is None:
                stock = ProductStock(
                    product_id=product_id,
                    product_name=product.name or "",
                    stock_update_date=now,
                    supplier_orders_update_date=now,
                    is_update_date=now,
                    stock_qty=0,
                    markdown_qty=0,
                    reserve_qty=0,
                    reserve_ecomm_qty=0,
                    lpc=0,
                    landed_cost=0,
                    distr_price=0,
                    promo_price=0,
                    volume_py=0,
                    volume_3m=0,
                    uc3_py=0,
                    uc3_3m=0,
                    transit_qty=0,
                    order_qty=0,
                    is_order_qty=0,
                    is_confirmed_order_qty=0,
                    is_stock_qty=0,
                )
                self.session.add(stock)
                stocks[product_id] = stock

            stock.product_name = product.name or ""
            stock.stock_update_date = now
            stock.lpc = metric.lpc
            stock.volume_py = metric.volume_py
            stock.volume_3m = metric.volume_3m
            stock.uc3_py = metric.uc3_py
            stock.uc3_3m = metric.uc3_3m
            saved += 1

        return saved

    def _weighted_field_by_product(self, batch_id: str, imported_by: str, product_id: int, field_name: str) -> Decimal:
        rows = self.session.query(TempStockImport).filter(
            TempStockImport.batch_id == batch_id,
            TempStockImport.imported_by == imported_by,
            TempStockImport.selected_product_id == product_id,
        ).all()

        return self._weighted_field_from_rows(rows, field_name)

    def _weighted_field_from_rows(self, rows, field_name: str) -> Decimal:
        weighted_sum = Decimal("0")
        total_weight = Decimal("0")
        avg_sum = Decimal("0")
        avg_cnt = 0

        for row in rows:
            row_weight = (
                self._to_decimal(row.stock_qty)
                + self._to_decimal(row.transit_qty)
                + self._to_decimal(row.markdown_qty)
                + self._to_decimal(row.reserve_qty)
            )
            val = getattr(row, field_name)
            d_val = self._to_decimal(val)

            if row_weight > 0:
                weighted_sum += d_val * row_weight
                total_weight += row_weight
            elif d_val != 0:
                avg_sum += d_val
                avg_cnt += 1

        if total_weight > 0:
            return weighted_sum / total_weight
        if avg_cnt > 0:
            return avg_sum / Decimal(str(avg_cnt))
        return Decimal("0")

    def save_stock_to_product_stock(self, batch_id: str, imported_by: str) -> int:
        now = datetime.combine(datetime.now().date(), datetime.min.time())
        try:
            sales_metrics = self.sales_metrics_service.calculate(update_date=now.date())
        except Exception as exc:
            raise ValueError(f"Не удалось рассчитать LPC/uC3 по БД продаж: {exc}") from exc

        # ABC category is refreshed from the same stock file. Products absent from
        # the current file must explicitly receive "-".
        self.session.query(Product).update({Product.abc_category: "-"}, synchronize_session=False)

        self.session.query(ProductStock).update({
            ProductStock.stock_qty: 0,
            ProductStock.transit_qty: 0,
            ProductStock.markdown_qty: 0,
            ProductStock.reserve_qty: 0,
            ProductStock.reserve_ecomm_qty: 0,
            ProductStock.lpc: 0,
            ProductStock.landed_cost: 0,
            ProductStock.distr_price: 0,
            ProductStock.promo_price: 0,
            ProductStock.volume_py: 0,
            ProductStock.volume_3m: 0,
            ProductStock.uc3_py: 0,
            ProductStock.uc3_3m: 0,
            ProductStock.stock_update_date: now,
        }, synchronize_session=False)

        stock_rows = self.session.query(TempStockImport).filter(
            TempStockImport.batch_id == batch_id,
            TempStockImport.imported_by == imported_by,
            TempStockImport.selected_product_id.isnot(None),
        ).order_by(TempStockImport.import_row_no.asc(), TempStockImport.id.asc()).all()

        if not stock_rows:
            raise ValueError("SaveStockToProductStock: не найдено ни одной строки с SelectedProductID для текущего batch.")

        rows_by_product = defaultdict(list)
        for row in stock_rows:
            rows_by_product[int(row.selected_product_id)].append(row)

        product_ids = list(rows_by_product.keys())
        products = self._load_products_map(product_ids)
        stocks = self._load_product_stock_map(product_ids)

        saved = 0

        for product_id in product_ids:
            rows = rows_by_product[product_id]
            stock_qty = sum(self._to_decimal(x.stock_qty) for x in rows)
            transit_qty = sum(self._to_decimal(x.transit_qty) for x in rows)
            markdown_qty = sum(self._to_decimal(x.markdown_qty) for x in rows)
            reserve_qty = sum(self._to_decimal(x.reserve_qty) for x in rows)
            reserve_ecomm_qty = sum(self._to_decimal(getattr(x, "reserve_ecomm_qty", 0)) for x in rows)

            metric = sales_metrics.get(int(product_id))
            lpc_val = metric.lpc if metric else Decimal("0")
            landed_val = self._weighted_field_from_rows(rows, "landed_cost")
            distr_val = self._min_price_from_rows(rows, "distr_price")
            promo_val = self._min_price_from_rows(rows, "promo_price")
            volume_py = metric.volume_py if metric else Decimal("0")
            volume_3m = metric.volume_3m if metric else Decimal("0")
            uc3_py = metric.uc3_py if metric else Decimal("0")
            uc3_3m = metric.uc3_3m if metric else Decimal("0")

            product = products.get(product_id)
            abc_category = self._abc_category_from_values(row.abc_category for row in rows)
            if product is not None:
                product.abc_category = abc_category
            p_name = product.name if product else ""

            stock = stocks.get(product_id)
            if stock is None:
                stock = ProductStock(
                    product_id=product_id,
                    product_name=p_name,
                    stock_update_date=now,
                    supplier_orders_update_date=now,
                    is_update_date=now,
                    stock_qty=stock_qty,
                    markdown_qty=markdown_qty,
                    reserve_qty=reserve_qty,
                    reserve_ecomm_qty=reserve_ecomm_qty,
                    lpc=lpc_val,
                    landed_cost=landed_val,
                    distr_price=distr_val,
                    promo_price=promo_val,
                    volume_py=volume_py,
                    volume_3m=volume_3m,
                    uc3_py=uc3_py,
                    uc3_3m=uc3_3m,
                    transit_qty=transit_qty,
                    order_qty=0,
                    is_order_qty=0,
                    is_confirmed_order_qty=0,
                    is_stock_qty=0,
                )
                self.session.add(stock)
                stocks[product_id] = stock
            else:
                stock.product_name = p_name
                stock.stock_update_date = now
                stock.stock_qty = stock_qty
                stock.markdown_qty = markdown_qty
                stock.reserve_qty = reserve_qty
                stock.reserve_ecomm_qty = reserve_ecomm_qty
                stock.transit_qty = transit_qty
                stock.lpc = lpc_val
                stock.landed_cost = landed_val
                stock.distr_price = distr_val
                stock.promo_price = promo_val
                stock.volume_py = volume_py
                stock.volume_3m = volume_3m
                stock.uc3_py = uc3_py
                stock.uc3_3m = uc3_3m
            saved += 1

        # SessionLocal uses autoflush=False. Flush newly created ProductStock rows
        # before the metrics pass, otherwise the following query cannot see them
        # and may create a second row with the same product_id.
        self.session.flush()
        self._apply_sales_metrics_to_product_stock(sales_metrics, now)
        self.session.flush()
        if saved == 0:
            raise ValueError("SaveStockToProductStock: не была сохранена ни одна строка.")
        return saved

    def save_supplier_orders_to_product_stock(self, batch_id: str, imported_by: str) -> int:
        now = datetime.combine(datetime.now().date(), datetime.min.time())

        self.session.query(ProductStock).update({
            ProductStock.order_qty: 0,
            ProductStock.is_order_qty: 0,
            ProductStock.is_confirmed_order_qty: 0,
            ProductStock.supplier_orders_update_date: now,
        }, synchronize_session=False)

        rows = self.session.query(TempSupplierOrdersImport).filter(
            TempSupplierOrdersImport.batch_id == batch_id,
            TempSupplierOrdersImport.imported_by == imported_by,
            TempSupplierOrdersImport.selected_product_id.isnot(None),
        ).order_by(TempSupplierOrdersImport.import_row_no.asc(), TempSupplierOrdersImport.id.asc()).all()

        if not rows:
            self.session.flush()
            return 0

        grouped = defaultdict(lambda: {
            "order_qty": Decimal("0"),
            "is_order_qty": Decimal("0"),
            "is_confirmed_order_qty": Decimal("0"),
            "abc_category": "-",
        })
        for row in rows:
            totals = grouped[int(row.selected_product_id)]
            totals["order_qty"] += self._to_decimal(row.order_qty)
            totals["is_order_qty"] += self._to_decimal(row.is_order_qty)
            totals["is_confirmed_order_qty"] += self._to_decimal(row.is_confirmed_order_qty)
            row_abc = clean_multi_spaces(row.abc_category)
            if totals["abc_category"] == "-" and row_abc and row_abc != "-":
                totals["abc_category"] = row_abc

        product_ids = list(grouped.keys())
        products = {
            product.id: product
            for product in self.session.query(Product).filter(Product.id.in_(product_ids)).all()
        }
        stocks = {
            stock.product_id: stock
            for stock in self.session.query(ProductStock).filter(ProductStock.product_id.in_(product_ids)).all()
        }

        saved = 0
        for product_id, totals in grouped.items():
            product = products.get(product_id)
            if product is not None:
                abc_category = totals["abc_category"]
                # Заказы поставщиков используются только для перепроверки ABC.
                # Пустая ячейка импорта нормализуется в "-" и не должна
                # стирать категорию, ранее загруженную из файла остатков.
                if abc_category and abc_category != "-":
                    product.abc_category = abc_category
            p_name = product.name if product else ''
            stock = stocks.get(product_id)
            if stock is None:
                stock = ProductStock(
                    product_id=product_id,
                    product_name=p_name,
                    stock_update_date=now,
                    supplier_orders_update_date=now,
                    is_update_date=now,
                    stock_qty=0,
                    markdown_qty=0,
                    reserve_qty=0,
                    reserve_ecomm_qty=0,
                    lpc=0,
                    landed_cost=0,
                    distr_price=0,
                    promo_price=0,
                    volume_py=0,
                    volume_3m=0,
                    uc3_py=0,
                    uc3_3m=0,
                    transit_qty=0,
                    order_qty=totals["order_qty"],
                    is_order_qty=totals["is_order_qty"],
                    is_confirmed_order_qty=totals["is_confirmed_order_qty"],
                    is_stock_qty=0,
                )
                self.session.add(stock)
                stocks[product_id] = stock
            else:
                stock.product_name = p_name
                stock.supplier_orders_update_date = now
                stock.order_qty = totals["order_qty"]
                stock.is_order_qty = totals["is_order_qty"]
                stock.is_confirmed_order_qty = totals["is_confirmed_order_qty"]
            saved += 1

        self.session.flush()
        return saved

    def save_is_to_product_stock(self, batch_id: str, imported_by: str) -> int:
        now = datetime.combine(datetime.now().date(), datetime.min.time())

        self.session.query(ProductStock).update({
            ProductStock.is_update_date: now,
            ProductStock.is_stock_qty: 0,
        }, synchronize_session=False)

        rows = self.session.query(TempIsImport).filter(
            TempIsImport.batch_id == batch_id,
            TempIsImport.imported_by == imported_by,
            TempIsImport.selected_product_id.isnot(None),
        ).all()
        totals_by_product: dict[int, Decimal] = defaultdict(lambda: Decimal("0"))
        for row in rows:
            totals_by_product[int(row.selected_product_id)] += self._to_decimal(row.stock_qty)

        product_ids = list(totals_by_product)
        products = self._load_products_map(product_ids)
        stocks = self._load_product_stock_map(product_ids)

        saved = 0

        for product_id in product_ids:
            stock_val = totals_by_product[product_id]
            product = products.get(product_id)
            p_name = product.name if product else ""

            stock = stocks.get(product_id)
            if stock is None:
                stock = ProductStock(
                    product_id=product_id,
                    product_name=p_name,
                    stock_update_date=now,
                    supplier_orders_update_date=now,
                    is_update_date=now,
                    stock_qty=0,
                    markdown_qty=0,
                    reserve_qty=0,
                    reserve_ecomm_qty=0,
                    lpc=0,
                    landed_cost=0,
                    distr_price=0,
                    promo_price=0,
                    volume_py=0,
                    volume_3m=0,
                    uc3_py=0,
                    uc3_3m=0,
                    transit_qty=0,
                    order_qty=0,
                    is_order_qty=0,
                    is_confirmed_order_qty=0,
                    is_stock_qty=stock_val,
                )
                self.session.add(stock)
                stocks[product_id] = stock
            else:
                stock.product_name = p_name
                stock.is_update_date = now
                stock.is_stock_qty = stock_val
            saved += 1

        self.session.flush()
        return saved
    def start_batch(self) -> str:
        return str(uuid.uuid4())

    def import_stock(self, file_path: str | Path, imported_by: str, batch_id: str | None = None) -> dict:
        batch_id = batch_id or self.start_batch()
        self.cleanup_old_temp_stock()
        rows = self.stock_importer.read_excel(file_path)
        imported_count = self.import_stock_rows(rows=rows, batch_id=batch_id, imported_by=imported_by, replace_existing=True)
        matched_count = self.automatch_stock_rows(batch_id=batch_id, imported_by=imported_by)
        return {"batch_id": batch_id, "imported_count": imported_count, "matched_count": matched_count}

    def save_stock(self, batch_id: str, imported_by: str) -> dict:
        self.validate_new_stock_products_before_save(batch_id, imported_by)
        created_products_count = self.create_stock_products_from_temp(batch_id, imported_by)
        product_articles_count = self.create_or_update_stock_product_articles(batch_id, imported_by)
        saved_count = self.save_stock_to_product_stock(batch_id, imported_by)
        self.delete_stock_rows_for_user(imported_by)
        return {
            "batch_id": batch_id,
            "created_products_count": created_products_count,
            "product_articles_count": product_articles_count,
            "saved_count": saved_count,
        }

    def import_supplier_orders(self, file_path: str | Path, imported_by: str, batch_id: str | None = None) -> dict:
        batch_id = batch_id or self.start_batch()
        self.cleanup_old_temp_supplier_orders()
        rows = self.supplier_orders_importer.read_excel(file_path)
        imported_count = self.import_supplier_orders_rows(rows=rows, batch_id=batch_id, imported_by=imported_by, replace_existing=True)
        matched_count = self.automatch_supplier_orders_rows(batch_id=batch_id, imported_by=imported_by)
        return {"batch_id": batch_id, "imported_count": imported_count, "matched_count": matched_count}

    def save_supplier_orders(self, batch_id: str, imported_by: str) -> dict:
        self.validate_new_supplier_orders_products_before_save(batch_id, imported_by)
        created_products_count = self.create_supplier_orders_products_from_temp(batch_id, imported_by)
        product_articles_count = self.create_or_update_supplier_orders_product_articles(batch_id, imported_by)
        saved_count = self.save_supplier_orders_to_product_stock(batch_id, imported_by)
        self.delete_supplier_orders_rows_for_user(imported_by)
        return {
            "batch_id": batch_id,
            "created_products_count": created_products_count,
            "product_articles_count": product_articles_count,
            "saved_count": saved_count,
        }

    def import_is(self, file_path: str | Path, imported_by: str, batch_id: str | None = None) -> dict:
        batch_id = batch_id or self.start_batch()
        self.cleanup_old_temp_is()
        rows = self.is_importer.read_excel(file_path)
        imported_count = self.import_is_rows(rows=rows, batch_id=batch_id, imported_by=imported_by, replace_existing=True)
        matched_count = self.automatch_is_rows(batch_id=batch_id, imported_by=imported_by)
        return {"batch_id": batch_id, "imported_count": imported_count, "matched_count": matched_count}

    def save_is(self, batch_id: str, imported_by: str) -> dict:
        self.validate_new_is_products_before_save(batch_id, imported_by)
        created_products_count = self.create_is_products_from_temp(batch_id, imported_by)
        product_articles_count = self.create_or_update_is_product_articles(batch_id, imported_by)
        saved_count = self.save_is_to_product_stock(batch_id, imported_by)
        self.delete_is_rows_for_user(imported_by)
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


# Backward-compatible aliases for old imports.
ProductStockImportService = ProductStockService
ProductStockImportRun = ProductStockService
