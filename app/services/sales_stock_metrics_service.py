from __future__ import annotations

import calendar
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from typing import Iterable

from sqlalchemy import create_engine, text
from sqlalchemy.orm import Session

from app.db.models import FixedCosts, Product, SalesProductLink
from app.services.cost_calculation_service import CostCalculationService
from app.services.product_matching_service import ProductMatchingService
from app.utils.text import clean_multi_spaces


# Same sales DB that is used by order planning.
SALES_DB_URI = "postgresql+psycopg2://postgres:qwerty@localhost:5432/report_db?client_encoding=utf8"


@dataclass(slots=True)
class ProductStockMetrics:
    product_id: int
    lpc: Decimal = Decimal("0")
    landed_cost: Decimal = Decimal("0")
    volume_py: Decimal = Decimal("0")
    volume_3m: Decimal = Decimal("0")
    uc3_py: Decimal = Decimal("0")
    uc3_3m: Decimal = Decimal("0")


@dataclass(slots=True)
class SalesMetricPeriods:
    prev_year_from: date
    prev_year_to: date
    last_3m_from: date
    last_3m_to: date


class SalesStockMetricsService:
    """Calculate stock-related metrics from the sales/report DB.

    The service reads the external sales DB tables `full_data` and `products`,
    maps sales products to local `Product` rows, and returns values that are
    stored in `product_stock` during stock update.
    """

    PURCHASE_STATUSES = ("Закуп", "Склад", "Транзит")
    TRANSIT_STATUS = "Транзит"
    SALES_STATUS = "Факт"
    PERIOD_PREV_YEAR = "prev.year"
    PERIOD_3_MONTHS = "3 mnth"

    def __init__(self, session: Session, sales_db_uri: str = SALES_DB_URI) -> None:
        self.session = session
        self.sales_db_uri = sales_db_uri
        self.matcher = ProductMatchingService(session)
        self.cost_calculation = CostCalculationService(session)

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
            if isinstance(value, str):
                value = value.strip().replace("\u00a0", "").replace(" ", "").replace(",", ".")
            return Decimal(str(value))
        except (InvalidOperation, ValueError, TypeError):
            return default

    @staticmethod
    def _normalize_sales_name(value: object) -> str:
        text_value = clean_multi_spaces(value)
        text_value = text_value.replace("л", "L").replace("Л", "L").replace("кг", "KG").replace("КГ", "KG")
        return text_value.upper()

    @staticmethod
    def _first_day_month_shift(base_date: date, months_shift: int) -> date:
        month_index = (base_date.year * 12 + base_date.month - 1) + months_shift
        year = month_index // 12
        month = month_index % 12 + 1
        return date(year, month, 1)

    @staticmethod
    def _last_day_of_month(base_date: date) -> date:
        return date(base_date.year, base_date.month, calendar.monthrange(base_date.year, base_date.month)[1])

    @classmethod
    def build_periods(cls, update_date: date | None = None) -> SalesMetricPeriods:
        current = update_date or datetime.now().date()
        prev_year = current.year - 1
        prev_year_from = date(prev_year, 1, 1)
        prev_year_to = date(prev_year, 12, 31)

        last_day = calendar.monthrange(current.year, current.month)[1]
        is_last_3_days = current.day >= last_day - 2
        if is_last_3_days:
            # In the last 3 days of the month, the current month is considered complete enough:
            # e.g. 2026-06-30 -> 2026-04-01..2026-06-30.
            last_3m_from = cls._first_day_month_shift(current, -2)
            last_3m_to = cls._last_day_of_month(current)
        else:
            # Otherwise use the last three fully closed months:
            # e.g. 2026-06-15 -> 2026-03-01..2026-05-31.
            last_3m_from = cls._first_day_month_shift(current, -3)
            prev_month_first = cls._first_day_month_shift(current, -1)
            last_3m_to = cls._last_day_of_month(prev_month_first)

        return SalesMetricPeriods(
            prev_year_from=prev_year_from,
            prev_year_to=prev_year_to,
            last_3m_from=last_3m_from,
            last_3m_to=last_3m_to,
        )

    def _sales_engine(self):
        return create_engine(self.sales_db_uri, pool_pre_ping=True)

    def _link_map(self) -> dict[str, SalesProductLink]:
        rows = self.session.query(SalesProductLink).all()
        return {clean_multi_spaces(row.sales_code): row for row in rows if clean_multi_spaces(row.sales_code)}

    def _split_sales_codes(self, value: object) -> list[str]:
        raw = clean_multi_spaces(value)
        if not raw:
            return []
        result: list[str] = []
        seen: set[str] = set()
        for token in raw.replace(",", ";").split(";"):
            code = clean_multi_spaces(token)
            if code and code not in seen:
                seen.add(code)
                result.append(code)
        return result

    def _resolve_product_id(self, row: dict, links: dict[str, SalesProductLink]) -> int | None:
        for code in self._split_sales_codes(row.get("sales_codes")):
            link = links.get(code)
            if link and link.product_id:
                return int(link.product_id)

        product = self.matcher.find_customer_product(
            supplier_article=row.get("sales_article"),
            product_name=row.get("product_pack"),
            pack=row.get("sales_pack"),
        )
        if product is not None:
            return int(product.id)
        return None

    @staticmethod
    def _is_better_base_row(candidate: dict, current: dict | None) -> bool:
        if current is None:
            return True
        # Prefer rows that contain an article/pack because product matching is more reliable.
        cand_score = int(bool(candidate.get("sales_article"))) + int(bool(candidate.get("sales_pack")))
        curr_score = int(bool(current.get("sales_article"))) + int(bool(current.get("sales_pack")))
        return cand_score > curr_score

    def _merge_lpc_rows_by_normalized_name(self, rows: Iterable[dict]) -> list[dict]:
        grouped: dict[str, dict] = {}
        for source in rows:
            product_name = self._normalize_sales_name(source.get("product_pack"))
            if not product_name:
                continue

            target = grouped.setdefault(
                product_name,
                {
                    "product_pack": product_name,
                    "sales_article": source.get("sales_article"),
                    "sales_pack": source.get("sales_pack"),
                    "sales_codes": "",
                    "purchase_volume": Decimal("0"),
                    "purchase_supply_cost": Decimal("0"),
                    "fact_volume": Decimal("0"),
                    "fact_landed_cost": Decimal("0"),
                    "fallback_volume": Decimal("0"),
                    "fallback_supply_cost": Decimal("0"),
                    "_base_row": None,
                },
            )
            if self._is_better_base_row(source, target.get("_base_row")):
                target["sales_article"] = source.get("sales_article")
                target["sales_pack"] = source.get("sales_pack")
                target["_base_row"] = source

            codes = self._split_sales_codes(target.get("sales_codes"))
            seen_codes = set(codes)
            for code in self._split_sales_codes(source.get("sales_codes")):
                if code not in seen_codes:
                    codes.append(code)
                    seen_codes.add(code)
            target["sales_codes"] = ";".join(codes)
            target["purchase_volume"] += self._to_decimal(source.get("purchase_volume"))
            target["purchase_supply_cost"] += self._to_decimal(source.get("purchase_supply_cost"))
            target["fact_volume"] += self._to_decimal(source.get("fact_volume"))
            target["fact_landed_cost"] += self._to_decimal(source.get("fact_landed_cost"))
            target["fallback_volume"] += self._to_decimal(source.get("fallback_volume"))
            target["fallback_supply_cost"] += self._to_decimal(source.get("fallback_supply_cost"))

        for row in grouped.values():
            row.pop("_base_row", None)
        return list(grouped.values())

    def _merge_uc3_rows_by_normalized_name(self, rows: Iterable[dict]) -> list[dict]:
        grouped: dict[tuple[str, str], dict] = {}
        for source in rows:
            product_name = self._normalize_sales_name(source.get("product_pack"))
            period_label = clean_multi_spaces(source.get("period_label"))
            if not product_name or not period_label:
                continue
            key = (product_name, period_label)
            target = grouped.setdefault(
                key,
                {
                    "period_label": period_label,
                    "product_pack": product_name,
                    "sales_article": source.get("sales_article"),
                    "sales_pack": source.get("sales_pack"),
                    "sales_codes": "",
                    "volume": Decimal("0"),
                    "margin_c3": Decimal("0"),
                    "_base_row": None,
                },
            )
            if self._is_better_base_row(source, target.get("_base_row")):
                target["sales_article"] = source.get("sales_article")
                target["sales_pack"] = source.get("sales_pack")
                target["_base_row"] = source

            codes = self._split_sales_codes(target.get("sales_codes"))
            seen_codes = set(codes)
            for code in self._split_sales_codes(source.get("sales_codes")):
                if code not in seen_codes:
                    codes.append(code)
                    seen_codes.add(code)
            target["sales_codes"] = ";".join(codes)
            target["volume"] += self._to_decimal(source.get("volume"))
            target["margin_c3"] += self._to_decimal(source.get("margin_c3"))

        for row in grouped.values():
            row.pop("_base_row", None)
        return list(grouped.values())

    # ------------------------------------------------------------------
    # Sales DB reads
    # ------------------------------------------------------------------
    def _read_lpc_rows(self, date_to: date | None = None) -> list[dict]:
        query = text(
            """
            WITH base AS (
                SELECT
                    fd.id,
                    COALESCE(p."Продукт_упаковка", '') AS product_pack,
                    COALESCE(p."Артикул", '') AS sales_article,
                    p."Упаковка" AS sales_pack,
                    fd."Код"::text AS sales_code,
                    fd."Дата" AS document_date,
                    fd."Статус" AS status,
                    fd."Документ" AS document,
                    COALESCE(fd."Кол_во_л", 0) AS volume,
                    COALESCE(fd."Себ_ть_поставки_партии", 0) AS supply_cost,
                    COALESCE(fd."Себ_ть_до_склада_партии", 0) AS cost
                FROM full_data fd
                LEFT JOIN products p ON p."Код" = fd."Код"
                WHERE fd."Статус" = ANY(:all_statuses)
                  AND (fd."Статус" = :transit_status OR :date_to IS NULL OR fd."Дата" <= :date_to)
            ),
            last_purchase AS (
                SELECT DISTINCT ON (product_pack)
                    product_pack,
                    document_date AS last_purchase_date,
                    document AS last_purchase_document,
                    id AS last_purchase_id
                FROM base
                WHERE status = :purchase_status
                ORDER BY product_pack, document_date DESC NULLS LAST, id DESC
            )
            SELECT
                b.product_pack,
                MIN(b.sales_article) AS sales_article,
                MIN(b.sales_pack) AS sales_pack,
                STRING_AGG(DISTINCT b.sales_code, ';') AS sales_codes,
                SUM(CASE WHEN b.status = ANY(:purchase_statuses)
                         THEN b.volume ELSE 0 END) AS purchase_volume,
                SUM(CASE WHEN b.status = ANY(:purchase_statuses)
                         THEN b.supply_cost ELSE 0 END) AS purchase_supply_cost,
                SUM(CASE WHEN b.status = :sales_status
                         THEN b.volume ELSE 0 END) AS fact_volume,
                SUM(CASE WHEN b.status = :sales_status
                         THEN b.cost ELSE 0 END) AS fact_landed_cost,
                SUM(CASE WHEN
                    (
                        b.status = :purchase_status
                        AND b.document_date IS NOT DISTINCT FROM lp.last_purchase_date
                        AND b.document IS NOT DISTINCT FROM lp.last_purchase_document
                    )
                    OR (
                        b.status = :warehouse_status
                        AND (
                            b.document_date > lp.last_purchase_date
                            OR (
                                b.document_date = lp.last_purchase_date
                                AND b.id > lp.last_purchase_id
                            )
                        )
                    )
                    THEN b.volume ELSE 0 END) AS fallback_volume,
                SUM(CASE WHEN
                    (
                        b.status = :purchase_status
                        AND b.document_date IS NOT DISTINCT FROM lp.last_purchase_date
                        AND b.document IS NOT DISTINCT FROM lp.last_purchase_document
                    )
                    OR (
                        b.status = :warehouse_status
                        AND (
                            b.document_date > lp.last_purchase_date
                            OR (
                                b.document_date = lp.last_purchase_date
                                AND b.id > lp.last_purchase_id
                            )
                        )
                    )
                    THEN b.supply_cost ELSE 0 END) AS fallback_supply_cost
            FROM base b
            LEFT JOIN last_purchase lp ON lp.product_pack = b.product_pack
            GROUP BY b.product_pack
            """
        )
        params = {
            "purchase_statuses": list(self.PURCHASE_STATUSES),
            "purchase_status": "Закуп",
            "warehouse_status": "Склад",
            "sales_status": self.SALES_STATUS,
            "transit_status": self.TRANSIT_STATUS,
            "all_statuses": list(self.PURCHASE_STATUSES) + [self.SALES_STATUS],
            "date_to": date_to,
        }
        with self._sales_engine().connect() as conn:
            result = conn.execute(query, params).mappings().all()
        return [dict(row) for row in result]

    def _read_uc3_rows_for_period(self, label: str, date_from: date, date_to: date) -> list[dict]:
        query = text(
            """
            SELECT
                :period_label AS period_label,
                COALESCE(p."Продукт_упаковка", '') AS product_pack,
                MIN(COALESCE(p."Артикул", '')) AS sales_article,
                MIN(p."Упаковка") AS sales_pack,
                STRING_AGG(DISTINCT fd."Код"::text, ';') AS sales_codes,
                SUM(COALESCE(fd."Кол_во_л", 0)) AS volume,
                SUM(COALESCE(fd."Маржа_C3_партии", 0)) AS margin_c3
            FROM full_data fd
            LEFT JOIN products p ON p."Код" = fd."Код"
            WHERE fd."Дата" >= :date_from
              AND fd."Дата" <= :date_to
              AND fd."Статус" = :sales_status
            GROUP BY COALESCE(p."Продукт_упаковка", '')
            """
        )
        params = {
            "period_label": label,
            "date_from": date_from,
            "date_to": date_to,
            "sales_status": self.SALES_STATUS,
        }
        with self._sales_engine().connect() as conn:
            result = conn.execute(query, params).mappings().all()
        return [dict(row) for row in result]

    # ------------------------------------------------------------------
    # Calculation
    # ------------------------------------------------------------------
    def _vat_multiplier(self) -> Decimal:
        fixed = self.session.query(FixedCosts).order_by(FixedCosts.id.asc()).first()
        if fixed is None:
            return Decimal("1")

        vat = self._to_decimal(getattr(fixed, "vat", None))
        if vat <= Decimal("-1"):
            return Decimal("1")
        return Decimal("1") + vat

    def _pack_for_lpc_row(self, row: dict, product_id: int) -> Decimal:
        pack = self._to_decimal(row.get("sales_pack"))
        if pack > 0:
            return pack

        product = self.session.query(Product).filter(Product.id == product_id).first()
        if product is None:
            return Decimal("0")
        return self._to_decimal(product.pack)

    @classmethod
    def calc_lpc(
        cls,
        purchase_volume: Decimal,
        purchase_supply_cost: Decimal,
        fact_volume: Decimal,
        fact_landed_cost: Decimal,
        fallback_volume: Decimal,
        fallback_supply_cost: Decimal,
        vat_multiplier: Decimal = Decimal("1"),
    ) -> Decimal:
        """Calculate LPC from purchase and fact cost-per-litre components.

        When purchased/transit/warehouse volume is fully consumed by Fact,
        use the latest purchase document and subsequent warehouse movements.
        """
        if purchase_volume - fact_volume == 0:
            if fallback_volume == 0:
                return Decimal("0")
            return fallback_supply_cost / fallback_volume * vat_multiplier

        purchase_part = (
            purchase_supply_cost / purchase_volume
            if purchase_volume != 0 else Decimal("0")
        )
        fact_part = (
            fact_landed_cost / fact_volume
            if fact_volume != 0 else Decimal("0")
        )
        return (purchase_part + fact_part) * vat_multiplier

    @classmethod
    def calc_uc3(cls, volume: Decimal, margin_c3: Decimal) -> Decimal:
        if volume == 0:
            return Decimal("0")
        return margin_c3 / volume

    def calculate(self, update_date: date | None = None) -> dict[int, ProductStockMetrics]:
        links = self._link_map()
        metrics: dict[int, ProductStockMetrics] = {}

        vat_multiplier = self._vat_multiplier()
        lpc_rows = self._merge_lpc_rows_by_normalized_name(self._read_lpc_rows(date_to=update_date))
        for row in lpc_rows:
            product_id = self._resolve_product_id(row, links)
            if not product_id:
                continue
            metric = metrics.setdefault(product_id, ProductStockMetrics(product_id=product_id))
            metric.lpc = self.calc_lpc(
                self._to_decimal(row.get("purchase_volume")),
                self._to_decimal(row.get("purchase_supply_cost")),
                self._to_decimal(row.get("fact_volume")),
                self._to_decimal(row.get("fact_landed_cost")),
                self._to_decimal(row.get("fallback_volume")),
                self._to_decimal(row.get("fallback_supply_cost")),
                vat_multiplier,
            )
            metric.landed_cost = self.cost_calculation.calc_landed_cost_from_lpc(metric.lpc)

        periods = self.build_periods(update_date)
        uc3_source_rows = []
        uc3_source_rows.extend(
            self._read_uc3_rows_for_period(self.PERIOD_PREV_YEAR, periods.prev_year_from, periods.prev_year_to)
        )
        uc3_source_rows.extend(
            self._read_uc3_rows_for_period(self.PERIOD_3_MONTHS, periods.last_3m_from, periods.last_3m_to)
        )

        uc3_rows = self._merge_uc3_rows_by_normalized_name(uc3_source_rows)
        for row in uc3_rows:
            product_id = self._resolve_product_id(row, links)
            if not product_id:
                continue
            metric = metrics.setdefault(product_id, ProductStockMetrics(product_id=product_id))
            volume = self._to_decimal(row.get("volume"))
            margin_c3 = self._to_decimal(row.get("margin_c3"))
            uc3 = self.calc_uc3(volume=volume, margin_c3=margin_c3)
            if row.get("period_label") == self.PERIOD_PREV_YEAR:
                metric.volume_py = volume
                metric.uc3_py = uc3
            elif row.get("period_label") == self.PERIOD_3_MONTHS:
                metric.volume_3m = volume
                metric.uc3_3m = uc3

        return metrics
