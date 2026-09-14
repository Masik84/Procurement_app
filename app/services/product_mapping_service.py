from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
from typing import Optional

import pandas as pd
from sqlalchemy import text
from sqlalchemy.orm import Session, joinedload

from app.db.models import Product, SalesProductLink
from app.services.product_matching_service import ProductMatchingService
from app.services.qty_in_box_service import normalize_qty_in_box
from app.utils.money import to_decimal
from app.utils.text import clean_multi_spaces


@dataclass(slots=True)
class ProductMappingCheckResult:
    rows: list[dict]
    auto_matched_count: int
    new_count: int
    changed_count: int


class ProductMappingService:
    """Synchronise the imported 1C Portfolio with local Products.

    ``check_products`` works only with the Portfolio DataFrame passed by the GUI.
    ``search_links`` displays already saved local mappings and never reads the
    external source workbook.
    """

    def __init__(self, session: Session) -> None:
        self.session = session
        self.matcher = ProductMatchingService(session)

    @staticmethod
    def _to_decimal(value) -> Decimal:
        return to_decimal(value)

    @staticmethod
    def _source_text(value: object) -> str:
        """Convert source cells to text without stripping/collapsing spaces."""
        if value is None:
            return ""
        try:
            if pd.isna(value):
                return ""
        except Exception:
            pass
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value)

    @classmethod
    def _source_upper_text(cls, value: object) -> str:
        return cls._source_text(value).upper()

    @classmethod
    def _code_key(cls, value: object) -> str:
        return cls._source_text(value).casefold()

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

    def _prepare_source_frame(self, source_df: pd.DataFrame) -> pd.DataFrame:
        if source_df.empty:
            return pd.DataFrame(columns=[
                "sales_code", "sales_article", "sales_name", "sales_pack",
                "sales_qty_in_box", "sales_brand", "sales_excise", "sales_type",
                "source_pack_type", "source_unit", "source_density", "source_net_weight",
            ])

        return pd.DataFrame({
            "sales_code": source_df["Код"].map(self._source_text),
            "sales_article": source_df["Артикул"].map(self._source_text),
            "sales_name": source_df["Продукт_упаковка"].map(self._source_upper_text),
            "sales_pack": source_df["Упаковка"],
            "sales_qty_in_box": source_df["Кол_во_в_упак"].map(self._qty_in_box_from_sales),
            "sales_brand": source_df["Brand"].map(self._source_upper_text),
            "sales_excise": source_df["Акциз_да_нет"].map(self._bool_from_sales_excise),
            "sales_type": source_df["Type"].map(self._source_text),
            "source_pack_type": source_df["Вид упаковки"].map(self._source_text),
            "source_unit": source_df["УЕ"].map(self._source_text),
            "source_density": source_df["Плотность"],
            "source_net_weight": source_df["Вес Нетто кг"],
        })

    def cleanup_stale_new_links(self, source_df: pd.DataFrame) -> int:
        """Delete obsolete local links whose Portfolio code starts with ``new``.

        A ``new...`` mapping is invalid when that code disappeared from the new
        Portfolio, or when the same code now points to a different English
        product name. Product rows themselves are never deleted here.
        """
        prepared = self._prepare_source_frame(source_df)
        current_new_names: dict[str, str] = {}
        for source in prepared.itertuples(index=False):
            code = self._source_text(source.sales_code)
            if code.casefold().startswith("new"):
                current_new_names[self._code_key(code)] = self._source_text(source.sales_name)

        links = (
            self.session.query(SalesProductLink)
            .filter(SalesProductLink.sales_code.ilike("new%"))
            .all()
        )
        deleted = 0
        for link in links:
            current_name = current_new_names.get(self._code_key(link.sales_code))
            saved_name = self._source_text(link.sales_product_name)
            if current_name is None or saved_name != current_name:
                self.session.delete(link)
                deleted += 1

        if deleted:
            self.session.flush()
        return deleted

    def product_map(self) -> dict[int, Product]:
        products = self.session.query(Product).all()
        return {int(product.id): product for product in products}

    def link_map(self, *, load_products: bool = False) -> dict[str, SalesProductLink]:
        query = self.session.query(SalesProductLink)
        if load_products:
            query = query.options(joinedload(SalesProductLink.product))
        return {row.sales_code: row for row in query.all()}

    def saved_sales_qty_in_box_map(self) -> dict[str, int | None]:
        """Return the saved Portfolio Qty in Box snapshot from our DB only.

        The field is intentionally read with SQL instead of the ORM model so
        this patch does not require changing model declarations just for a
        display-only source snapshot.
        """
        rows = self.session.execute(
            text(
                """
                SELECT sales_code, sales_qty_in_box
                FROM sales_product_links
                """
            )
        ).all()
        result: dict[str, int | None] = {}
        for code, value in rows:
            source_code = self._source_text(code)
            if source_code:
                result[source_code] = self._qty_in_box_from_sales(value)
        return result

    def ignored_sales_codes(self) -> set[str]:
        """Return source product codes explicitly excluded by the user.

        ``is_ignored`` is kept as a persisted flag in ``sales_product_links``.
        The ORM model is intentionally not required for this small source-state
        flag, so older project branches can still apply this patch cleanly.
        """
        rows = self.session.execute(
            text(
                """
                SELECT sales_code
                FROM sales_product_links
                WHERE COALESCE(is_ignored, FALSE) = TRUE
                """
            )
        ).all()
        return {
            self._source_text(code)
            for (code,) in rows
            if self._source_text(code)
        }

    def set_ignored(self, rows: list[dict], *, ignored: bool) -> int:
        """Persist the user's decision to ignore/unignore source products.

        Ignoring a new row never creates a Product in our catalogue.  It only
        saves the source snapshot and the ignore flag, so future checks can
        skip it.  Existing mapped Products are not deleted.
        """
        changed = 0
        for row in rows:
            code = self._source_text(row.get("sales_code"))
            if not code:
                continue

            link = (
                self.session.query(SalesProductLink)
                .filter(SalesProductLink.sales_code == code)
                .first()
            )
            if link is None:
                link = SalesProductLink(sales_code=code)
                self.session.add(link)

            # Keep an existing mapping when the source row is ignored. For a
            # brand-new ignored row product_id remains NULL: no Product is made.
            link.sales_article = self._source_text(row.get("sales_article")) or None
            link.sales_product_name = self._source_text(row.get("sales_product_name")) or None
            link.sales_pack = self._to_decimal(row.get("sales_pack"))
            link.sales_brand = self._source_text(row.get("sales_brand")) or None
            link.sales_is_excise = row.get("sales_is_excise")
            link.updated_at = datetime.now()
            self.session.flush()

            self.session.execute(
                text(
                    """
                    UPDATE sales_product_links
                    SET sales_qty_in_box = :qty_in_box,
                        is_ignored = :is_ignored
                    WHERE sales_code = :sales_code
                    """
                ),
                {
                    "sales_code": code,
                    "qty_in_box": self._qty_in_box_from_sales(row.get("sales_qty_in_box")),
                    "is_ignored": bool(ignored),
                },
            )
            changed += 1

        self.session.flush()
        return changed

    def delete_links(self, sales_codes: set[str] | list[str]) -> int:
        codes = [self._source_text(code) for code in sales_codes if self._source_text(code)]
        if not codes:
            return 0
        return (
            self.session.query(SalesProductLink)
            .filter(SalesProductLink.sales_code.in_(codes))
            .delete(synchronize_session=False)
        )

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
            self._source_text(link.sales_article) == self._source_text(article)
            and self._source_text(link.sales_product_name) == self._source_text(product_name)
            and self._to_decimal(link.sales_pack) == self._to_decimal(pack)
            and self._source_text(link.sales_brand) == self._source_text(brand)
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
        source_pack_type: str = "",
        source_unit: str = "",
        source_density=None,
        source_net_weight=None,
        product: Product | None = None,
        status: str,
        is_new: bool,
        auto_found: bool = False,
        is_ignored: bool = False,
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
            "source_pack_type": source_pack_type,
            "source_unit": source_unit,
            "source_density": source_density,
            "source_net_weight": source_net_weight,
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
            "is_ignored": bool(is_ignored),
        }

    def check_products(self, source_df: pd.DataFrame) -> ProductMappingCheckResult:
        """Run Portfolio-vs-local matching for the already imported source."""
        if source_df is None or source_df.empty:
            return ProductMappingCheckResult([], 0, 0, 0)

        product_map = self.product_map()
        links = self.link_map(load_products=False)
        saved_sales_qty = self.saved_sales_qty_in_box_map()
        ignored_codes = self.ignored_sales_codes()
        cleaned = self._prepare_source_frame(source_df)

        rows: list[dict] = []
        auto_matched = 0
        new_count = 0
        changed_count = 0

        for source in cleaned.itertuples(index=False):
            sales_code = source.sales_code
            if not sales_code or sales_code in ignored_codes:
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
            saved_qty = saved_sales_qty.get(sales_code)
            qty_snapshot_changed = link is not None and saved_qty != source.sales_qty_in_box
            source_changed = link is not None and (not link_matches or qty_snapshot_changed)
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
                source_pack_type=source.source_pack_type,
                source_unit=source.source_unit,
                source_density=source.source_density,
                source_net_weight=source.source_net_weight,
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

    def search_links(self, *, include_ignored: bool = False) -> list[dict]:
        """Show mappings already stored in the local Procurement DB only.

        ``Search`` never reads the Portfolio file and never runs automatic
        matching. Ignored rows are hidden by default, but can be shown from the
        table context menu so the user can restore them.
        """
        links = self.link_map(load_products=True)
        if not links:
            return []

        saved_sales_qty = self.saved_sales_qty_in_box_map()
        ignored_codes = self.ignored_sales_codes()
        rows: list[dict] = []
        for code, link in sorted(links.items(), key=lambda item: item[0]):
            is_ignored = code in ignored_codes
            if is_ignored and not include_ignored:
                continue
            product = link.product
            rows.append(self._row_payload(
                sales_code=code,
                sales_article=self._source_text(link.sales_article),
                sales_name=self._source_text(link.sales_product_name),
                sales_pack=link.sales_pack,
                sales_qty_in_box=saved_sales_qty.get(code),
                sales_brand=self._source_text(link.sales_brand),
                sales_excise=link.sales_is_excise,
                sales_type="",
                product=product,
                status=(
                    "Игнорируется"
                    if is_ignored
                    else ("Сопоставлен" if product is not None else "Без продукта")
                ),
                is_new=False,
                auto_found=False,
                is_ignored=is_ignored,
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

        product_name = self._source_upper_text(row.get("new_product_name") or row.get("product_name"))
        if not product_name:
            return None
        brand = self._source_upper_text(row.get("new_brand") or row.get("sales_brand"))
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
            pack_type=row.get("source_pack_type"),
            density=row.get("source_density"),
            net_weight=row.get("source_net_weight"),
            unit=row.get("source_unit"),
            preserve_text_spacing=True,
        )
        row["product_id"] = product.id
        row["product_name"] = product.name
        row["product_qty_in_box"] = product.qty_in_box
        return product

    def _upsert_sales_link_for_row(self, row: dict, product: Product | None) -> bool:
        code = self._source_text(row.get("sales_code"))
        if not code:
            return False
        link = self.session.query(SalesProductLink).filter(SalesProductLink.sales_code == code).first()
        if link is None:
            link = SalesProductLink(sales_code=code)
            self.session.add(link)
        link.product_id = product.id if product else None
        link.sales_article = self._source_text(row.get("sales_article")) or None
        link.sales_product_name = self._source_text(row.get("sales_product_name")) or None
        link.sales_pack = self._to_decimal(row.get("sales_pack"))
        link.sales_brand = self._source_text(row.get("sales_brand")) or None
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
        qty_snapshots: list[tuple[str, int | None]] = []
        for row in rows:
            product = self._ensure_product_for_row(row)
            if product is None:
                raise ValueError(
                    f"Не выбран продукт для {clean_multi_spaces(row.get('sales_product_name')) or clean_multi_spaces(row.get('sales_code'))}"
                )
            if self._upsert_sales_link_for_row(row, product):
                saved += 1
                code = self._source_text(row.get("sales_code"))
                if code:
                    qty_snapshots.append(
                        (code, self._qty_in_box_from_sales(row.get("sales_qty_in_box")))
                    )

        # New SalesProductLink ORM rows must exist before the local snapshot
        # column can be updated by sales_code.
        self.session.flush()
        for code, qty_in_box in qty_snapshots:
            self.session.execute(
                text(
                    """
                    UPDATE sales_product_links
                    SET sales_qty_in_box = :qty_in_box,
                        is_ignored = FALSE
                    WHERE sales_code = :sales_code
                    """
                ),
                {"sales_code": code, "qty_in_box": qty_in_box},
            )
        self.session.flush()
        return saved
