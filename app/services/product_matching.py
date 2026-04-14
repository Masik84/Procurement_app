from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Optional
from decimal import Decimal

from sqlalchemy.orm import Session

from app.db.models import Product, ProductArticle
from app.utils.parsers import parse_loose_number
from app.utils.text import (
    clean_multi_spaces,
    normalize_customer_product_name,
    normalize_product_name,
)


@dataclass(slots=True)
class ProductCreateData:
    name: str
    brand: str
    pack: float
    is_excise: bool


class ProductMatchingService:
    def __init__(self, session: Session) -> None:
        self.session = session

    # =========================================================
    # Article helpers
    # =========================================================

    @staticmethod
    def article_token_normalize(value: object) -> str:
        return clean_multi_spaces(value)

    @classmethod
    def split_article_tokens(cls, article_text: object) -> list[str]:
        raw_text = cls.article_token_normalize(article_text)
        if not raw_text:
            return []

        raw_text = raw_text.replace(",", "/")
        parts = raw_text.split("/")

        result: list[str] = []
        seen: set[str] = set()

        for part in parts:
            token = cls.article_token_normalize(part)
            if not token:
                continue

            token_key = token.casefold()
            if token_key not in seen:
                seen.add(token_key)
                result.append(token)

        return result

    @staticmethod
    def is_pseudo_new_article(article: object) -> bool:
        s = clean_multi_spaces(article).lower()
        return s.startswith("new")

    @classmethod
    def should_create_article_link(cls, article: object) -> bool:
        s = clean_multi_spaces(article)
        return bool(s) and not cls.is_pseudo_new_article(s)

    # =========================================================
    # Pack / family helpers
    # =========================================================

    @staticmethod
    def _normalize_pack_text(value: object) -> str:
        s = clean_multi_spaces(value)
        s = s.replace(",", ".")
        return s

    @staticmethod
    def _get_trailing_number(value: str) -> str:
        match = re.search(r"([0-9]+(?:[.,][0-9]+)?)\s*$", value.strip(), flags=re.IGNORECASE)
        if not match:
            return ""
        return match.group(1).replace(",", ".")

    @classmethod
    def build_name_with_pack_unit(cls, product_name: object, pack_value: object, unit_text: str) -> str:
        s = normalize_customer_product_name(product_name)
        p = cls._normalize_pack_text(pack_value)
        unit_text = clean_multi_spaces(unit_text).upper()

        if not s:
            return ""

        if not p:
            return s

        if re.search(r"([0-9]+(?:[.,][0-9]+)?)\s*(L|KG|л|Л|кг|КГ)\s*$", s, flags=re.IGNORECASE):
            return re.sub(
                r"([0-9]+(?:[.,][0-9]+)?)\s*(L|KG|л|Л|кг|КГ)\s*$",
                rf"\1 {unit_text}",
                s,
                flags=re.IGNORECASE,
            )

        tail_num = cls._get_trailing_number(s)
        if tail_num and tail_num == p:
            return f"{s} {unit_text}"

        return f"{s} {p} {unit_text}"

    @staticmethod
    def build_product_family_from_name(product_name: str, pack_value: object) -> str:
        s_name = clean_multi_spaces(product_name).upper()
        s_pack = clean_multi_spaces(pack_value)

        if not s_name:
            raise ValueError("Не заполнен ProductName.")

        if not s_pack:
            raise ValueError("Не заполнен Pack.")

        marker_l = f" {s_pack}L"
        marker_kg = f" {s_pack}KG"

        pos = s_name.find(marker_l)
        if pos == -1:
            pos = s_name.find(marker_kg)

        if pos == -1:
            raise ValueError(f"Проверьте корректность упаковки {s_name}")

        return s_name[:pos]

    # =========================================================
    # Search helpers
    # =========================================================

    def _get_product_by_exact_name(self, product_name: str) -> Optional[Product]:
        return (
            self.session.query(Product)
            .filter(Product.name == product_name)
            .first()
        )

    def _get_article_link_by_exact_article(self, article: str) -> Optional[ProductArticle]:
        return (
            self.session.query(ProductArticle)
            .filter(ProductArticle.article == article)
            .order_by(ProductArticle.id.asc())
            .first()
        )

    def _get_article_link_by_exact_name(self, supplier_name: str) -> Optional[ProductArticle]:
        return (
            self.session.query(ProductArticle)
            .filter(ProductArticle.name == supplier_name)
            .order_by(ProductArticle.id.asc())
            .first()
        )

    def find_by_normalized_product_name(self, source_name: object) -> Optional[Product]:
        target = normalize_product_name(source_name)
        if not target:
            return None

        products = (
            self.session.query(Product)
            .filter(Product.name.isnot(None))
            .all()
        )

        for product in products:
            if normalize_product_name(product.name) == target:
                return product

        return None

    def find_product_by_split_articles(self, supplier_article: object) -> Optional[Product]:
        first_product: Optional[Product] = None

        for token in self.split_article_tokens(supplier_article):
            link = self._get_article_link_by_exact_article(token)
            if link and link.product:
                if first_product is None:
                    first_product = link.product

        return first_product

    # =========================================================
    # Public matching methods
    # =========================================================

    def find_customer_product(
        self,
        supplier_article: object,
        product_name: object,
        pack: object,
    ) -> Optional[Product]:
        article = clean_multi_spaces(supplier_article)
        name = clean_multi_spaces(product_name)
        normalized_name = normalize_customer_product_name(name)

        if article:
            link = self._get_article_link_by_exact_article(article)
            if link and link.product:
                return link.product

        if name:
            link = self._get_article_link_by_exact_name(name)
            if link and link.product:
                return link.product

            product = self.find_by_normalized_product_name(name)
            if product:
                return product

        if normalized_name and normalized_name != name:
            link = self._get_article_link_by_exact_name(normalized_name)
            if link and link.product:
                return link.product

            product = self.find_by_normalized_product_name(normalized_name)
            if product:
                return product

        name_l = self.build_name_with_pack_unit(name, pack, "L")
        if name_l:
            link = self._get_article_link_by_exact_name(name_l)
            if link and link.product:
                return link.product

            product = self.find_by_normalized_product_name(name_l)
            if product:
                return product

        name_kg = self.build_name_with_pack_unit(name, pack, "KG")
        if name_kg:
            link = self._get_article_link_by_exact_name(name_kg)
            if link and link.product:
                return link.product

            product = self.find_by_normalized_product_name(name_kg)
            if product:
                return product

        return None

    def find_price_import_product(
        self,
        supplier_article: object,
        supplier_product_name: object,
    ) -> Optional[Product]:
        article = clean_multi_spaces(supplier_article)
        name = clean_multi_spaces(supplier_product_name)

        if article:
            product = self.find_product_by_split_articles(article)
            if product:
                return product

        if name:
            link = self._get_article_link_by_exact_name(name)
            if link and link.product:
                return link.product

        if name:
            product = self.find_by_normalized_product_name(name)
            if product:
                return product

        return None

    def find_stock_product(
        self,
        source_article: object,
        source_product_name: object,
    ) -> Optional[Product]:
        name = clean_multi_spaces(source_product_name)
        article = clean_multi_spaces(source_article)

        if name:
            product = self.find_by_normalized_product_name(name)
            if product:
                return product

        if article:
            link = self._get_article_link_by_exact_article(article)
            if link and link.product:
                return link.product

        return None

    # =========================================================
    # Create / link methods
    # =========================================================

    def get_or_create_product(
        self,
        *,
        name: str,
        brand: str,
        pack: object,
        is_excise: bool,
    ) -> Product:
        clean_name = clean_multi_spaces(name).upper()
        clean_brand = clean_multi_spaces(brand).upper()
        pack_num = parse_loose_number(pack)

        if not clean_name:
            raise ValueError("Не заполнен NewProductName.")

        if not clean_brand:
            raise ValueError("Не заполнен NewBrand.")

        if pack_num is None:
            raise ValueError("NewPack должен быть числом")

        family = self.build_product_family_from_name(clean_name, pack_num)

        existing = self._get_product_by_exact_name(clean_name)
        if existing:
            return existing

        product = Product(
            name=clean_name,
            brand=clean_brand,
            family=family,
            pack=pack_num,
            is_excise=bool(is_excise),
        )
        self.session.add(product)
        self.session.flush()

        return product

    def create_product_article_if_missing(
        self,
        *,
        product_id: int,
        article: object,
        supplier_name: object,
    ) -> Optional[ProductArticle]:
        clean_article = clean_multi_spaces(article)
        clean_name = clean_multi_spaces(supplier_name)

        if not clean_article and not clean_name:
            return None

        existing = (
            self.session.query(ProductArticle)
            .filter(
                ProductArticle.product_id == product_id,
                ProductArticle.article == (clean_article or None),
                ProductArticle.name == (clean_name or None),
            )
            .first()
        )
        if existing:
            return existing

        row = ProductArticle(
            product_id=product_id,
            article=clean_article or None,
            name=clean_name or None,
        )
        self.session.add(row)
        self.session.flush()
        return row

    def save_product_articles_by_split_articles(
        self,
        *,
        product_id: int,
        supplier_article: object,
        supplier_product_name: object,
    ) -> None:
        safe_name = clean_multi_spaces(supplier_product_name)
        tokens = self.split_article_tokens(supplier_article)

        if not tokens:
            if safe_name:
                self.create_product_article_if_missing(
                    product_id=product_id,
                    article=None,
                    supplier_name=safe_name,
                )
            return

        for token in tokens:
            self.create_product_article_if_missing(
                product_id=product_id,
                article=token,
                supplier_name=safe_name or None,
            )

    def create_article_link_from_source(
        self,
        *,
        product_id: int,
        source_article: object,
        source_name: object,
    ) -> None:
        clean_article = clean_multi_spaces(source_article)
        clean_name = clean_multi_spaces(source_name)

        if clean_article:
            if self.should_create_article_link(clean_article):
                self.create_product_article_if_missing(
                    product_id=product_id,
                    article=clean_article,
                    supplier_name=clean_name or None,
                )
        elif clean_name:
            self.create_product_article_if_missing(
                product_id=product_id,
                article=None,
                supplier_name=clean_name,
            )