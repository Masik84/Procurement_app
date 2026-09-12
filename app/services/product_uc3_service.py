from __future__ import annotations

from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Iterable

from sqlalchemy.orm import Session

from app.db.models import Product, ProductStock, ProductUc3History
from app.utils.text import normalize_product_name


SOURCE_SUPPLIER = "SUPPLIER"
SOURCE_MANUAL = "MANUAL"
SOURCE_TARGET_UC3 = "TARGET_UC3"
SOURCE_WALK_AWAY_UC3 = "WALK_AWAY_UC3"

SOURCE_LABELS = {
    SOURCE_SUPPLIER: "Supplier",
    SOURCE_MANUAL: "Manual",
    SOURCE_TARGET_UC3: "Target uC3",
    SOURCE_WALK_AWAY_UC3: "Walk-Away uC3",
}


class ProductUc3Service:
    def __init__(self, session: Session):
        self.session = session

    @staticmethod
    def round_integer(value) -> int | None:
        if value is None or value == "":
            return None
        try:
            number = Decimal(str(value).replace(" ", "").replace(",", "."))
            if not number.is_finite():
                return None
            return int(number.to_integral_value(rounding=ROUND_HALF_UP))
        except (InvalidOperation, ValueError, TypeError):
            raise ValueError(f"Некорректное значение uC3: {value}")

    def get_current(self, product_id: int) -> ProductUc3History | None:
        return (
            self.session.query(ProductUc3History)
            .filter(ProductUc3History.product_id == int(product_id))
            .order_by(ProductUc3History.change_date.desc(), ProductUc3History.id.desc())
            .first()
        )

    def get_current_map(self, product_ids: Iterable[int]) -> dict[int, ProductUc3History]:
        ids = sorted({int(value) for value in product_ids if value is not None})
        if not ids:
            return {}
        rows = (
            self.session.query(ProductUc3History)
            .filter(ProductUc3History.product_id.in_(ids))
            .order_by(
                ProductUc3History.product_id.asc(),
                ProductUc3History.change_date.desc(),
                ProductUc3History.id.desc(),
            )
            .all()
        )
        result: dict[int, ProductUc3History] = {}
        for row in rows:
            result.setdefault(int(row.product_id), row)
        return result

    def save_values(
        self,
        *,
        product_id: int,
        target_uc3,
        walk_away_uc3,
        change_date: datetime | None = None,
    ) -> tuple[ProductUc3History | None, bool]:
        target_value = self.round_integer(target_uc3)
        walk_value = self.round_integer(walk_away_uc3)
        current = self.get_current(product_id)
        if current is not None:
            current_target = self.round_integer(current.target_uc3)
            current_walk = self.round_integer(current.walk_away_uc3)
            if current_target == target_value and current_walk == walk_value:
                return current, False
        elif target_value is None and walk_value is None:
            return None, False

        row = ProductUc3History(
            product_id=int(product_id),
            target_uc3=target_value,
            walk_away_uc3=walk_value,
            change_date=change_date or datetime.now(),
        )
        self.session.add(row)
        self.session.flush()
        return row, True

    def import_rows(self, rows: list[dict]) -> dict:
        if not rows:
            return {"created": 0, "unchanged": 0, "missing_products": []}

        products = self.session.query(Product).filter(Product.name.isnot(None)).all()
        exact = {str(product.name).strip(): product for product in products if product.name}
        normalized = {}
        for product in products:
            key = normalize_product_name(product.name)
            if key and key not in normalized:
                normalized[key] = product

        # If a product occurs more than once in the workbook, the last row wins.
        prepared: dict[int, tuple[Product, dict]] = {}
        missing: list[str] = []
        for src in rows:
            name = str(src.get("product_name") or "").strip()
            product = exact.get(name)
            if product is None:
                key = normalize_product_name(name)
                product = normalized.get(key) if key else None
            if product is None:
                if name and name not in missing:
                    missing.append(name)
                continue
            prepared[int(product.id)] = (product, src)

        created = 0
        unchanged = 0
        now = datetime.now()
        for product_id, (_product, src) in prepared.items():
            _row, was_created = self.save_values(
                product_id=product_id,
                target_uc3=src.get("target_uc3"),
                walk_away_uc3=src.get("walk_away_uc3"),
                change_date=now,
            )
            if was_created:
                created += 1
            else:
                unchanged += 1

        return {"created": created, "unchanged": unchanged, "missing_products": missing}

    @staticmethod
    def _positive_decimal(value) -> Decimal | None:
        if value is None or value == "":
            return None
        try:
            result = Decimal(str(value))
        except Exception:
            return None
        return result if result > 0 else None

    @classmethod
    def sales_reference_price(cls, stock: ProductStock | None) -> Decimal | None:
        if stock is None:
            return None
        values = [
            value
            for value in (
                cls._positive_decimal(getattr(stock, "distr_price", None)),
                cls._positive_decimal(getattr(stock, "promo_price", None)),
            )
            if value is not None
        ]
        return min(values) if values else None

    @classmethod
    def full_cost_from_uc3(cls, *, stock: ProductStock | None, uc3_value, vat) -> Decimal | None:
        sale_price = cls.sales_reference_price(stock)
        if sale_price is None or uc3_value is None:
            return None
        try:
            uc3 = Decimal(str(uc3_value))
            vat_value = Decimal(str(vat))
        except Exception:
            return None
        return sale_price - uc3 * (Decimal("1") + vat_value)

    def full_cost_for_product_source(self, *, product_id: int, source_type: str, vat) -> Decimal | None:
        current = self.get_current(product_id)
        if current is None:
            return None
        if source_type == SOURCE_TARGET_UC3:
            uc3_value = current.target_uc3
        elif source_type == SOURCE_WALK_AWAY_UC3:
            uc3_value = current.walk_away_uc3
        else:
            return None
        stock = self.session.query(ProductStock).filter(ProductStock.product_id == int(product_id)).first()
        return self.full_cost_from_uc3(stock=stock, uc3_value=uc3_value, vat=vat)
