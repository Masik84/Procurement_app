from __future__ import annotations

import os
import math
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Optional

import pandas as pd
from sqlalchemy import create_engine, text
from sqlalchemy.orm import Session, joinedload

from app.db.models import OrderPlanningCalculation, Product, ProductStock, SalesProductLink
from app.services.product_matching_service import ProductMatchingService
from app.utils.text import clean_multi_spaces


SALES_DB_URI = "postgresql+psycopg2://postgres:qwerty@localhost:5432/report_db?client_encoding=utf8"


@dataclass(slots=True)
class ProductCheckResult:
    rows: list[dict]
    auto_matched_count: int
    new_count: int
    changed_count: int


@dataclass(slots=True)
class CalculationResult:
    rows: list[dict]
    period_from: date
    period_to: date
    auto_matched_count: int
    unmatched_count: int


class OrderPlanningService:
    def __init__(self, session: Session, sales_db_uri: str = SALES_DB_URI) -> None:
        self.session = session
        self.sales_db_uri = sales_db_uri
        self.matcher = ProductMatchingService(session)

    # ------------------------------------------------------------------
    # Basic helpers
    # ------------------------------------------------------------------
    @staticmethod
    def _to_decimal(value: object, default: Decimal = Decimal("0")) -> Decimal:
        if value is None or value == "":
            return default
        if isinstance(value, Decimal):
            return value
        try:
            return Decimal(str(value))
        except Exception:
            return default

    @staticmethod
    def _round4(value: Decimal) -> Decimal:
        return value.quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP)

    @staticmethod
    def _ceil_decimal(value: Decimal) -> Decimal:
        if value <= 0:
            return Decimal("0")
        return Decimal(int(math.ceil(float(value))))

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

    def _sales_engine(self):
        return create_engine(self.sales_db_uri, pool_pre_ping=True)

    def _read_sales_products(self, sales_codes: Optional[list[str]] = None) -> pd.DataFrame:
        cols = ['"Код"', '"Артикул"', '"Продукт_упаковка"', '"Упаковка"', '"Brand"', '"Акциз_да_нет"', '"Группа_бренда"']
        query = (
            f"SELECT {', '.join(cols)} FROM products "
            "WHERE \"Группа_бренда\" = 'Import' "
            "AND COALESCE(\"Brand\", '') NOT IN ('-', 'Phoenix Oil', 'MEVAG')"
        )
        params: dict[str, object] = {}
        if sales_codes:
            query += ' AND "Код" = ANY(:codes)'
            params["codes"] = sales_codes
        with self._sales_engine().connect() as conn:
            df = pd.read_sql(text(query), conn, params=params)
        if df.empty:
            return df
        df["Продукт_упаковка"] = df["Продукт_упаковка"].map(self._normalize_sales_name)
        return df

    def _read_sales_data(self, period_from: date, period_to: date) -> pd.DataFrame:
        query = """
            SELECT "Код", "Дата", "Статус", "Кол_во_л"
            FROM full_data
            WHERE "Дата" >= :date_from
              AND "Дата" <= :date_to
              AND "Статус" = 'Факт'
        """
        with self._sales_engine().connect() as conn:
            return pd.read_sql(text(query), conn, params={"date_from": period_from, "date_to": period_to})

    def _product_map(self) -> dict[int, Product]:
        products = self.session.query(Product).options(joinedload(Product.stock)).all()
        return {int(product.id): product for product in products}

    def _link_map(self) -> dict[str, SalesProductLink]:
        rows = self.session.query(SalesProductLink).options(joinedload(SalesProductLink.product)).all()
        return {row.sales_code: row for row in rows}

    def get_brand_values(self) -> list[str]:
        rows = (
            self.session.query(Product.brand)
            .filter(Product.brand.isnot(None), Product.brand != "")
            .distinct()
            .order_by(Product.brand.asc())
            .all()
        )
        return [row[0] for row in rows if row[0]]

    def get_family_values(self, brands: Optional[list[str]] = None) -> list[str]:
        query = self.session.query(Product.family).filter(Product.family.isnot(None), Product.family != "")
        if brands:
            query = query.filter(Product.brand.in_(brands))
        rows = query.distinct().order_by(Product.family.asc()).all()
        return [row[0] for row in rows if row[0]]

    def get_products_for_combo(self, brand_filter: str = "", text_filter: str = "") -> list[Product]:
        query = self.session.query(Product).filter(Product.name.isnot(None), Product.name != "")
        if brand_filter and brand_filter != "-":
            query = query.filter(Product.brand == brand_filter)
        if text_filter:
            query = query.filter(Product.name.ilike(f"%{text_filter}%"))
        return query.order_by(Product.name.asc()).all()

    # ------------------------------------------------------------------
    # Product check / linking
    # ------------------------------------------------------------------
    def check_products(self) -> ProductCheckResult:
        sales_df = self._read_sales_products()
        if sales_df.empty:
            return ProductCheckResult([], 0, 0, 0)

        links = self._link_map()
        rows: list[dict] = []
        auto_matched = 0
        new_count = 0
        changed_count = 0

        for _, source in sales_df.iterrows():
            sales_code = clean_multi_spaces(source.get("Код"))
            if not sales_code:
                continue

            sales_article = clean_multi_spaces(source.get("Артикул"))
            sales_name = clean_multi_spaces(source.get("Продукт_упаковка"))
            sales_pack = source.get("Упаковка")
            sales_brand = clean_multi_spaces(source.get("Brand"))
            sales_excise = self._bool_from_sales_excise(source.get("Акциз_да_нет"))

            link = links.get(sales_code)
            linked_product = link.product if link and link.product else None
            product = linked_product
            auto_found = False

            if product is None:
                product = self.matcher.find_customer_product(sales_article, sales_name, sales_pack)
                if product is not None:
                    auto_found = True
                    auto_matched += 1

            is_new = link is None
            is_changed = False
            if is_new:
                new_count += 1
            else:
                if (link.sales_article or "") != (sales_article or ""):
                    is_changed = True
                if (link.sales_product_name or "") != (sales_name or ""):
                    is_changed = True
                if self._to_decimal(link.sales_pack) != self._to_decimal(sales_pack):
                    is_changed = True
                if (link.sales_brand or "") != (sales_brand or ""):
                    is_changed = True
                if link.sales_is_excise != sales_excise:
                    is_changed = True
                if linked_product is None and product is not None:
                    is_changed = True
                if is_changed:
                    changed_count += 1

            if not is_new and not is_changed and not auto_found:
                continue

            rows.append({
                "sales_code": sales_code,
                "sales_article": sales_article,
                "sales_product_name": sales_name,
                "sales_pack": self._to_decimal(sales_pack),
                "sales_brand": sales_brand,
                "sales_is_excise": sales_excise,
                "product_id": product.id if product else None,
                "product_name": product.name if product else "",
                "is_auto_matched": auto_found,
                "is_new": is_new,
            })

        return ProductCheckResult(rows, auto_matched, new_count, changed_count)

    def _ensure_product_for_row(self, row: dict) -> Product | None:
        """Return selected product or create a new Product from manually typed Product Name."""
        product_id = row.get("product_id")
        if product_id:
            product = self.session.query(Product).filter(Product.id == int(product_id)).first()
            if product:
                row["product_id"] = product.id
                row["product_name"] = product.name
                row["brand"] = product.brand
                row["family"] = product.family
                row["pack"] = product.pack
                return product

        product_name = clean_multi_spaces(row.get("product_name")).upper()
        if not product_name:
            return None

        brand = clean_multi_spaces(row.get("sales_brand") or row.get("brand"))
        pack = row.get("sales_pack") if row.get("sales_pack") not in (None, "") else row.get("pack")
        is_excise = row.get("sales_is_excise")
        if is_excise is None:
            is_excise = False

        product = self.matcher.get_or_create_product(
            name=product_name,
            brand=brand,
            pack=pack,
            is_excise=bool(is_excise),
        )

        row["product_id"] = product.id
        row["product_name"] = product.name
        row["brand"] = product.brand
        row["family"] = product.family
        row["pack"] = product.pack
        return product

    def _sales_codes_from_row(self, row: dict) -> list[str]:
        raw = clean_multi_spaces(row.get("sales_code"))
        if not raw:
            return []
        codes: list[str] = []
        for part in raw.replace(",", ";").split(";"):
            code = clean_multi_spaces(part)
            if code and code not in codes:
                codes.append(code)
        return codes

    def _upsert_sales_link_for_row(self, row: dict, product: Product | None, sales_code: str | None = None) -> bool:
        code = clean_multi_spaces(sales_code or row.get("sales_code"))
        if not code:
            return False

        link = (
            self.session.query(SalesProductLink)
            .filter(SalesProductLink.sales_code == code)
            .first()
        )
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

    def save_product_links(self, rows: list[dict]) -> int:
        saved = 0
        for row in rows:
            product = self._ensure_product_for_row(row)
            for sales_code in self._sales_codes_from_row(row):
                if self._upsert_sales_link_for_row(row, product, sales_code):
                    saved += 1
        self.session.flush()
        return saved

    # ------------------------------------------------------------------
    # Calculation
    # ------------------------------------------------------------------
    def calculate(self, period_from: date, period_to: date) -> CalculationResult:
        sales_df = self._read_sales_data(period_from, period_to)
        product_df = self._read_sales_products()
        if sales_df.empty or product_df.empty:
            return CalculationResult([], period_from, period_to, 0, 0)

        sales_df["Кол_во_л"] = pd.to_numeric(sales_df["Кол_во_л"], errors="coerce").fillna(0)
        merged = sales_df.merge(product_df, on="Код", how="left")
        merged = merged[merged["Группа_бренда"].fillna("").eq("Import")].copy()
        if merged.empty:
            return CalculationResult([], period_from, period_to, 0, 0)

        grouped = (
            merged.groupby(["Код", "Артикул", "Продукт_упаковка", "Упаковка", "Brand", "Акциз_да_нет"], dropna=False)["Кол_во_л"]
            .sum()
            .reset_index()
        )

        days_count = max((period_to - period_from).days + 1, 1)
        grouped["avg_sales_month"] = grouped["Кол_во_л"] / days_count * 30

        links = self._link_map()
        auto_matched = 0
        unmatched = 0
        rows: list[dict] = []

        for _, source in grouped.iterrows():
            sales_code = clean_multi_spaces(source.get("Код"))
            sales_article = clean_multi_spaces(source.get("Артикул"))
            sales_name = clean_multi_spaces(source.get("Продукт_упаковка"))
            sales_pack = source.get("Упаковка")
            sales_brand = clean_multi_spaces(source.get("Brand"))
            sales_excise = self._bool_from_sales_excise(source.get("Акциз_да_нет"))
            avg_sales = self._round4(self._to_decimal(source.get("avg_sales_month")))

            product = None
            link = links.get(sales_code)
            if link and link.product:
                product = link.product
            else:
                product = self.matcher.find_customer_product(sales_article, sales_name, sales_pack)
                if product:
                    auto_matched += 1
                else:
                    unmatched += 1

            rows.append({
                "sales_code": sales_code,
                "sales_article": sales_article,
                "sales_product_name": sales_name,
                "sales_pack": self._to_decimal(sales_pack),
                "sales_brand": sales_brand,
                "sales_is_excise": sales_excise,
                "product_id": product.id if product else None,
                "product_name": product.name if product else "",
                "is_auto_matched": bool(product and (not link or not link.product)),
                "avg_sales_month": avg_sales,
            })

        return CalculationResult(self.build_display_rows(rows), period_from, period_to, auto_matched, unmatched)

    def build_display_rows(
        self,
        base_rows: list[dict],
        quick_months: Decimal | int | float = Decimal("0"),
        safe_months: Decimal | int | float = Decimal("0"),
        brand_filter: Optional[list[str]] = None,
        family_filter: Optional[list[str]] = None,
        vol_not_null: bool = False,
    ) -> list[dict]:
        # group rows by selected Product.id; unmatched rows remain separate by sales_code
        grouped: dict[object, dict] = {}
        for row in base_rows:
            product_id = row.get("product_id")
            key = ("product", int(product_id)) if product_id else ("sales", row.get("sales_code"))
            if key not in grouped:
                grouped[key] = dict(row)
            else:
                grouped[key]["avg_sales_month"] = self._to_decimal(grouped[key].get("avg_sales_month")) + self._to_decimal(row.get("avg_sales_month"))
                grouped[key]["sales_code"] = f"{grouped[key].get('sales_code')}; {row.get('sales_code')}"

        product_map = self._product_map()
        quick_m = self._to_decimal(quick_months)
        safe_m = self._to_decimal(safe_months)
        result: list[dict] = []

        for item in grouped.values():
            product_id = item.get("product_id")
            product = product_map.get(int(product_id)) if product_id else None
            stock = product.stock if product and product.stock else None

            brand = product.brand if product else item.get("sales_brand", "")
            family = product.family if product else ""
            if brand_filter and brand not in brand_filter:
                continue
            if family_filter and family not in family_filter:
                continue

            avg_sales = self._to_decimal(item.get("avg_sales_month"))
            pack = self._to_decimal(product.pack if product else item.get("sales_pack"), Decimal("0"))

            stock_qty = self._to_decimal(getattr(stock, "stock_qty", None))
            transit_qty = self._to_decimal(getattr(stock, "transit_qty", None))
            order_qty = self._to_decimal(getattr(stock, "order_qty", None))
            is_order_qty = self._to_decimal(getattr(stock, "is_order_qty", None))
            reserve_qty = self._to_decimal(getattr(stock, "reserve_qty", None))
            reserve_ecomm_qty = self._to_decimal(getattr(stock, "reserve_ecomm_qty", None))
            markdown_qty = self._to_decimal(getattr(stock, "markdown_qty", None))
            free_base = stock_qty
            total_stock = stock_qty

            free_st = free_base
            # Safe Stock (st+tr) = Stock + Transit
            # Safe Stock (+ord) = Stock + Transit + Purchase Order + Order IS
            free_st_tr = free_base + transit_qty
            free_ord = free_base + transit_qty + order_qty + is_order_qty

            safe_st_month = self._round4(free_st / avg_sales) if avg_sales > 0 else Decimal("0")
            safe_st_tr_month = self._round4(free_st_tr / avg_sales) if avg_sales > 0 else Decimal("0")
            safe_ord_month = self._round4(free_ord / avg_sales) if avg_sales > 0 else Decimal("0")

            std_l_raw = (safe_m - safe_ord_month) * avg_sales
            quick_l_raw = (quick_m - safe_st_tr_month) * avg_sales
            std_l_raw = Decimal("0") if std_l_raw < 0 else std_l_raw
            quick_l_raw = Decimal("0") if quick_l_raw < 0 else quick_l_raw

            if pack > 0:
                std_pcs = self._ceil_decimal(std_l_raw / pack)
                quick_pcs = self._ceil_decimal(quick_l_raw / pack)
            else:
                std_pcs = Decimal("0")
                quick_pcs = Decimal("0")

            std_l = std_pcs * pack
            quick_l = quick_pcs * pack
            is_new_or_unchecked = not bool(product_id) or bool(item.get("is_auto_matched"))
            if vol_not_null and std_l <= 0 and quick_l <= 0 and not is_new_or_unchecked:
                continue

            result.append({
                "sales_code": item.get("sales_code", ""),
                "sales_article": item.get("sales_article", ""),
                "sales_product_name": "" if product else item.get("sales_product_name", ""),
                "sales_pack": item.get("sales_pack"),
                "sales_brand": item.get("sales_brand", ""),
                "sales_is_excise": item.get("sales_is_excise"),
                "product_id": product.id if product else item.get("product_id"),
                "brand": brand or "",
                "family": family or "",
                "product_name": product.name if product else item.get("product_name", ""),
                "pack": pack,
                "avg_sales_month": self._round4(avg_sales),
                "safe_stock_st_month": safe_st_month,
                "safe_stock_st_tr_month": safe_st_tr_month,
                "safe_stock_ord_month": safe_ord_month,
                "quick_order_pcs": quick_pcs,
                "quick_order_l": self._round4(quick_l),
                "std_order_pcs": std_pcs,
                "std_order_l": self._round4(std_l),
                "distr_price": self._to_decimal(getattr(stock, "distr_price", None)),
                "promo_price": self._to_decimal(getattr(stock, "promo_price", None)),
                "free_stock_st": free_st,
                "free_stock_st_tr": free_st_tr,
                "free_stock_ord": free_ord,
                "stock": total_stock,
                "transit": transit_qty,
                "purchase_order": order_qty,
                "order_is": is_order_qty,
                "stock_is": self._to_decimal(getattr(stock, "is_stock_qty", None)),
                "reserve": self._to_decimal(getattr(stock, "reserve_qty", None)),
                "reserve_ecomm": self._to_decimal(getattr(stock, "reserve_ecomm_qty", None)),
                "markdown": markdown_qty,
                "is_auto_matched": bool(item.get("is_auto_matched")),
            })

        result.sort(key=lambda r: ((r.get("brand") or ""), (r.get("family") or ""), (r.get("product_name") or ""), str(r.get("sales_code") or "")))
        return result

    def load_saved_base_rows(self) -> tuple[list[dict], Optional[date], Optional[date]]:
        rows = (
            self.session.query(OrderPlanningCalculation, Product)
            .join(Product, Product.id == OrderPlanningCalculation.product_id)
            .order_by(Product.brand.asc(), Product.family.asc(), Product.name.asc())
            .all()
        )
        if not rows:
            return [], None, None
        period_from = rows[0][0].period_from
        period_to = rows[0][0].period_to
        base_rows = []
        for calc, product in rows:
            base_rows.append({
                "sales_code": "",
                "sales_article": "",
                "sales_product_name": "",
                "product_id": product.id,
                "product_name": product.name,
                "is_auto_matched": False,
                "avg_sales_month": self._to_decimal(calc.avg_sales_month),
            })
        return base_rows, period_from, period_to

    def save_calculation(self, display_rows: list[dict], period_from: date, period_to: date) -> int:
        # rows must already contain the user-selected products. Re-group before saving.
        prepared_rows: list[dict] = []
        for row in display_rows:
            row_copy = dict(row)
            product = self._ensure_product_for_row(row_copy)
            if product is not None:
                for sales_code in self._sales_codes_from_row(row_copy):
                    self._upsert_sales_link_for_row(row_copy, product, sales_code)
                prepared_rows.append(row_copy)

        if not prepared_rows:
            raise ValueError("Нет строк с выбранным или введенным Product name для сохранения")

        grouped: dict[int, Decimal] = {}
        for row in prepared_rows:
            product_id = int(row["product_id"])
            grouped[product_id] = grouped.get(product_id, Decimal("0")) + self._to_decimal(row.get("avg_sales_month"))

        self.session.query(OrderPlanningCalculation).delete(synchronize_session=False)
        for product_id, avg_sales in grouped.items():
            self.session.add(OrderPlanningCalculation(
                product_id=product_id,
                period_from=period_from,
                period_to=period_to,
                avg_sales_month=self._round4(avg_sales),
            ))
        self.session.flush()
        return len(grouped)
