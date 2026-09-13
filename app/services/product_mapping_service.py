from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
from typing import Optional

import pandas as pd
from sqlalchemy import create_engine, text
from sqlalchemy.orm import Session, joinedload

from app.db.models import Product, SalesProductLink
from app.services.product_matching_service import ProductMatchingService
from app.services.qty_in_box_service import normalize_qty_in_box
from app.utils.money import to_decimal
from app.utils.text import clean_multi_spaces


SALES_DB_URI = "postgresql+psycopg2://postgres:qwerty@localhost:5432/report_db?client_encoding=utf8"
EXCLUDED_SALES_TYPES = ("SPARES", "FILTER")


@dataclass(slots=True)
class ProductMappingCheckResult:
    rows: list[dict]
    auto_matched_count: int
    new_count: int
    changed_count: int


class ProductMappingService:
    """Synchronises the sales DB product catalogue with local Products.

    ``check_products`` reads the source catalogue and runs matching. ``search_links``
    only displays already saved links; it does not call the matching algorithm.
    """

    def __init__(self, session: Session, sales_db_uri: str = SALES_DB_URI) -> None:
        self.session = session
        self.sales_db_uri = sales_db_uri
        self.matcher = ProductMatchingService(session)

    @staticmethod
    def _to_decimal(value) -> Decimal:
        return to_decimal(value)

    @staticmethod
    def _normalize_sales_name(value: object) -> str:
        text_value = clean_multi_spaces(value).replace("л", "L").replace("Л", "L").replace("кг", "KG").replace("КГ", "KG")
        return text_value.upper()

    @staticmethod
    def _bool_from_sales_excise(value: object) -> Optional[bool]:
        text_value = clean_multi_spaces(value).lower()
        if not text_value:
            return None
        if text_value in {"да", "yes", "true", "1", "+"}:
            return True
        if text_value in {"нет", "no", "false", "0", "-"}:
            return False
        return None

    @staticmethod
    def _qty_in_box_from_sales(value: object) -> int | None:
        try:
            return normalize_qty_in_box(value, field_name="Кол_во_в_упак")
        except (TypeError, ValueError, InvalidOperation):
            return None

    def _sales_engine(self):
        return create_engine(self.sales_db_uri, pool_pre_ping=True)

    def read_sales_products(self) -> pd.DataFrame:
        """Read the whole source catalogue except Type SPARES / Filter."""
        query = text(
            '''
            SELECT
                "Код", "Артикул", "Продукт_упаковка", "Упаковка",
                "Кол_во_в_упак", "Brand", "Акциз_да_нет", "Type"
            FROM products
            WHERE UPPER(TRIM(COALESCE("Type", ''))) <> ALL(:excluded_types)
            '''
        )
        with self._sales_engine().connect() as conn:
            df = pd.read_sql(query, conn, params={"excluded_types": list(EXCLUDED_SALES_TYPES)})
        if not df.empty:
            df["Продукт_упаковка"] = df["Продукт_упаковка"].map(self._normalize_sales_name)
        return df

    def _clean_sales_frame(self, sales_df: pd.DataFrame) -> pd.DataFrame:
        if sales_df.empty:
            return pd.DataFrame(columns=[
                "sales_code", "sales_article", "sales_name", "sales_pack",
                "sales_qty_in_box", "sales_brand", "sales_excise", "sales_type",
            ])
        return pd.DataFrame({
            "sales_code": sales_df["Код"].map(clean_multi_spaces),
            "sales_article": sales_df["Артикул"].map(clean_multi_spaces),
            "sales_name": sales_df["Продукт_упаковка"].map(clean_multi_spaces),
            "sales_pack": sales_df["Упаковка"],
            "sales_qty_in_box": sales_df["Кол_во_в_упак"].map(self._qty_in_box_from_sales),
            "sales_brand": sales_df["Brand"].map(clean_multi_spaces),
            "sales_excise": sales_df["Акциз_да_нет"].map(self._bool_from_sales_excise),
            "sales_type": sales_df["Type"].map(clean_multi_spaces),
        })

    def product_map(self) -> dict[int, Product]:
        products = self.session.query(Product).all()
        return {int(product.id): product for product in products}

    def link_map(self, *, load_products: bool = False) -> dict[str, SalesProductLink]:
        query = self.session.query(SalesProductLink)
        if load_products:
            query = query.options(joinedload(SalesProductLink.product))
        return {row.sales_code: row for row in query.all()}

    def get_brand_values(self) -> list[str]:
        rows = (
            self.session.query(Product.brand)
            .filter(Product.brand.isnot(None), Product.brand != "")
            .distinct()
            .order_by(Product.brand.asc())
            .all()
        )
        return [row[0] for row in rows if row[0]]

    def get_products_for_combo(self, brand_filter: str = "", text_filter: str = "") -> list[Product]:
        query = self.session.query(Product).filter(Product.name.isnot(None), Product.name != "")
        if brand_filter and brand_filter != "-":
            query = query.filter(Product.brand == brand_filter)
        if text_filter:
            query = query.filter(Product.name.ilike(f"%{text_filter}%"))
        return query.order_by(Product.name.asc()).all()

    def sales_link_matches_source(
        self,
        link: SalesProductLink | None,
        *,
        article: object,
        product_name: object,
        pack: object,
        brand: object,
        is_excise: Optional[bool],
    ) -> bool:
        if link is None:
            return False
        return (
            clean_multi_spaces(link.sales_article) == clean_multi_spaces(article)
            and clean_multi_spaces(link.sales_product_name) == clean_multi_spaces(product_name)
            and self._to_decimal(link.sales_pack) == self._to_decimal(pack)
            and clean_multi_spaces(link.sales_brand) == clean_multi_spaces(brand)
            and link.sales_is_excise == is_excise
        )

    @staticmethod
    def _status_text(*, is_new: bool, auto_found: bool, source_changed: bool, qty_missing: bool, qty_mismatch: bool) -> str:
        parts: list[str] = []
        if is_new:
            parts.append("Автоподбор" if auto_found else "Новый")
        elif source_changed:
            parts.append("Изменён источник")
        elif auto_found:
            parts.append("Автоподбор")
        if qty_missing:
            parts.append("Qty in Box пуст")
        elif qty_mismatch:
            parts.append("Qty in Box отличается")
        return "; ".join(parts) or "Изменение"

    def _row_payload(
        self,
        *,
        sales_code: str,
        sales_article: str,
        sales_name: str,
        sales_pack,
        sales_qty_in_box: int | None,
        sales_brand: str,
        sales_excise: Optional[bool],
        sales_type: str,
        product: Product | None,
        status: str,
        is_new: bool,
        auto_found: bool = False,
    ) -> dict:
        product_qty_in_box = None
        if product is not None:
            try:
                product_qty_in_box = normalize_qty_in_box(product.qty_in_box)
            except ValueError:
                product_qty_in_box = None
        unmatched = product is None
        return {
            "sales_code": sales_code,
            "sales_article": sales_article,
            "sales_product_name": sales_name,
            "sales_pack": self._to_decimal(sales_pack),
            "sales_qty_in_box": sales_qty_in_box,
            "sales_brand": sales_brand,
            "sales_is_excise": sales_excise,
            "sales_type": sales_type,
            "status": status,
            "product_id": int(product.id) if product else None,
            "product_name": product.name if product else "",
            "product_qty_in_box": product_qty_in_box,
            "new_product_name": sales_name if unmatched else "",
            "new_brand": sales_brand if unmatched else "",
            "new_pack": self._to_decimal(sales_pack) if unmatched else None,
            "new_qty_in_box": sales_qty_in_box if unmatched else None,
            "new_is_excise": bool(sales_excise) if unmatched and sales_excise is not None else (False if unmatched else None),
            "is_auto_matched": bool(auto_found),
            "is_new": bool(is_new),
        }

    def check_products(self) -> ProductMappingCheckResult:
        """Run the source-vs-local check and matching algorithm."""
        sales_df = self.read_sales_products()
        if sales_df.empty:
            return ProductMappingCheckResult([], 0, 0, 0)

        product_map = self.product_map()
        links = self.link_map(load_products=False)
        cleaned = self._clean_sales_frame(sales_df)

        rows: list[dict] = []
        auto_matched = 0
        new_count = 0
        changed_count = 0

        for source in cleaned.itertuples(index=False):
            sales_code = source.sales_code
            if not sales_code:
                continue
            link = links.get(sales_code)
            link_matches = self.sales_link_matches_source(
                link,
                article=source.sales_article,
                product_name=source.sales_name,
                pack=source.sales_pack,
                brand=source.sales_brand,
                is_excise=source.sales_excise,
            )
            linked_product = (
                product_map.get(int(link.product_id))
                if link_matches and link and link.product_id is not None
                else None
            )
            product = linked_product
            auto_found = False
            if product is None:
                product = self.matcher.find_customer_product(
                    source.sales_article,
                    source.sales_name,
                    source.sales_pack,
                    brand=source.sales_brand,
                )
                if product is not None:
                    auto_found = True
                    auto_matched += 1

            product_qty = None
            if product is not None:
                try:
                    product_qty = normalize_qty_in_box(product.qty_in_box)
                except ValueError:
                    product_qty = None
            qty_missing = product is not None and source.sales_qty_in_box is not None and product_qty is None
            qty_mismatch = (
                product is not None
                and source.sales_qty_in_box is not None
                and product_qty is not None
                and product_qty != source.sales_qty_in_box
            )

            is_new = link is None
            source_changed = link is not None and not link_matches
            is_changed = source_changed or (linked_product is None and product is not None) or qty_missing or qty_mismatch
            if is_new:
                new_count += 1
            elif is_changed:
                changed_count += 1

            if not is_new and not is_changed and not auto_found:
                continue

            rows.append(self._row_payload(
                sales_code=sales_code,
                sales_article=source.sales_article,
                sales_name=source.sales_name,
                sales_pack=source.sales_pack,
                sales_qty_in_box=source.sales_qty_in_box,
                sales_brand=source.sales_brand,
                sales_excise=source.sales_excise,
                sales_type=source.sales_type,
                product=product,
                status=self._status_text(
                    is_new=is_new,
                    auto_found=auto_found,
                    source_changed=source_changed,
                    qty_missing=qty_missing,
                    qty_mismatch=qty_mismatch,
                ),
                is_new=is_new,
                auto_found=auto_found,
            ))

        return ProductMappingCheckResult(rows, auto_matched, new_count, changed_count)

    def search_links(self) -> list[dict]:
        """Show only mappings already stored in the local Procurement DB.

        ``Search`` must never connect to the sales DB and must never run
        ProductMatchingService.  It is a read-only view of ``SalesProductLink``
        plus the linked local ``Product`` record.
        """
        links = self.link_map(load_products=True)
        if not links:
            return []

        rows: list[dict] = []
        for code, link in sorted(links.items(), key=lambda item: item[0]):
            product = link.product
            rows.append(self._row_payload(
                sales_code=code,
                sales_article=clean_multi_spaces(link.sales_article),
                sales_name=clean_multi_spaces(link.sales_product_name),
                sales_pack=link.sales_pack,
                # Qty in Box from the sales DB is intentionally not refreshed
                # here: it is external source data and Search is local-only.
                sales_qty_in_box=None,
                sales_brand=clean_multi_spaces(link.sales_brand),
                sales_excise=link.sales_is_excise,
                sales_type="",
                product=product,
                status="Сопоставлен" if product is not None else "Без продукта",
                is_new=False,
                auto_found=False,
            ))
        return rows

    def _ensure_product_for_row(self, row: dict) -> Product | None:
        product_id = row.get("product_id")
        sales_qty_in_box = self._qty_in_box_from_sales(row.get("sales_qty_in_box"))
        if product_id:
            product = self.session.query(Product).filter(Product.id == int(product_id)).first()
            if product:
                if product.qty_in_box is None and sales_qty_in_box is not None:
                    product.qty_in_box = sales_qty_in_box
                row["product_id"] = product.id
                row["product_name"] = product.name
                row["product_qty_in_box"] = product.qty_in_box
                return product

        product_name = clean_multi_spaces(row.get("new_product_name") or row.get("product_name")).upper()
        if not product_name:
            return None
        brand = clean_multi_spaces(row.get("new_brand") or row.get("sales_brand"))
        pack = row.get("new_pack")
        if pack in (None, ""):
            pack = row.get("sales_pack")
        qty_in_box = row.get("new_qty_in_box")
        if qty_in_box in (None, ""):
            qty_in_box = sales_qty_in_box
        is_excise = row.get("new_is_excise")
        if is_excise is None:
            is_excise = row.get("sales_is_excise")
        if is_excise is None:
            is_excise = False

        product = self.matcher.get_or_create_product(
            name=product_name,
            brand=brand,
            pack=pack,
            qty_in_box=qty_in_box,
            is_excise=bool(is_excise),
        )
        row["product_id"] = product.id
        row["product_name"] = product.name
        row["product_qty_in_box"] = product.qty_in_box
        return product

    def _upsert_sales_link_for_row(self, row: dict, product: Product | None) -> bool:
        code = clean_multi_spaces(row.get("sales_code"))
        if not code:
            return False
        link = self.session.query(SalesProductLink).filter(SalesProductLink.sales_code == code).first()
        if link is None:
            link = SalesProductLink(sales_code=code)
            self.session.add(link)
        link.product_id = product.id if product else None
        link.sales_article = clean_multi_spaces(row.get("sales_article")) or None
        link.sales_product_name = clean_multi_spaces(row.get("sales_product_name")) or None
        link.sales_pack = self._to_decimal(row.get("sales_pack"))
        link.sales_brand = clean_multi_spaces(row.get("sales_brand")) or None
        link.sales_is_excise = row.get("sales_is_excise")
        link.updated_at = datetime.now()
        if product is not None:
            self.matcher.create_article_link_from_source(
                product_id=int(product.id),
                source_article=row.get("sales_article"),
                source_name=row.get("sales_product_name"),
            )
        return True

    def save_rows(self, rows: list[dict]) -> int:
        saved = 0
        for row in rows:
            product = self._ensure_product_for_row(row)
            if product is None:
                raise ValueError(
                    f"Не выбран продукт для {clean_multi_spaces(row.get('sales_product_name')) or clean_multi_spaces(row.get('sales_code'))}"
                )
            if self._upsert_sales_link_for_row(row, product):
                saved += 1
        self.session.flush()
        return saved
