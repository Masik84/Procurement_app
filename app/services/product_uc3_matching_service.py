from __future__ import annotations

from dataclasses import dataclass

from sqlalchemy.orm import Session

from app.db.models import Product, ProductArticle
from app.services.product_matching_service import ProductMatchingService
from app.utils.text import clean_multi_spaces, normalize_product_name


@dataclass(frozen=True, slots=True)
class ProductUc3Match:
    product_id: int
    product_name: str
    source: str  # "product" | "article"


class ProductUc3MatchingService:
    """Target uC3 matching with the explicitly agreed priority.

    1) Products.name by normalized name.
    2) ProductArticle.name among names already seen by the application.
    3) GUI manual selection (handled by the page); that manual alias is then
       persisted to ProductArticle with article=NULL.
    """

    def __init__(self, session: Session) -> None:
        self.session = session
        self._product_by_normalized: dict[str, Product] = {}
        self._alias_exact: dict[str, Product] = {}
        self._alias_normalized: dict[str, Product] = {}
        self._loaded = False

    def _load(self) -> None:
        if self._loaded:
            return

        products = (
            self.session.query(Product)
            .filter(Product.name.isnot(None), Product.name != "")
            .order_by(Product.id.asc())
            .all()
        )
        product_by_id: dict[int, Product] = {}
        for product in products:
            product_by_id[int(product.id)] = product
            key = normalize_product_name(product.name)
            if key:
                # Preserve the previous `.first()` behaviour for duplicate
                # normalized names by keeping the lowest product id.
                self._product_by_normalized.setdefault(key, product)

        links = (
            self.session.query(ProductArticle)
            .filter(ProductArticle.name.isnot(None), ProductArticle.name != "")
            .order_by(ProductArticle.id.desc())
            .all()
        )
        for link in links:
            product = product_by_id.get(int(link.product_id))
            if product is None:
                continue
            exact_key = clean_multi_spaces(link.name).upper()
            normalized_key = normalize_product_name(link.name)
            if exact_key:
                self._alias_exact.setdefault(exact_key, product)
            if normalized_key:
                self._alias_normalized.setdefault(normalized_key, product)

        self._loaded = True

    def resolve(self, source_name: object) -> ProductUc3Match | None:
        self._load()
        clean_name = clean_multi_spaces(source_name)
        if not clean_name:
            return None

        normalized_key = normalize_product_name(clean_name)
        if normalized_key:
            product = self._product_by_normalized.get(normalized_key)
            if product is not None:
                return ProductUc3Match(int(product.id), product.name, "product")

        # Only after Products.name failed do we search names remembered in the
        # Product Articles table, as requested for Target uC3.
        product = self._alias_exact.get(clean_name.upper())
        if product is None and normalized_key:
            product = self._alias_normalized.get(normalized_key)
        if product is not None:
            return ProductUc3Match(int(product.id), product.name, "article")
        return None

    def save_manual_alias(self, *, product_id: int, source_name: object) -> None:
        clean_name = clean_multi_spaces(source_name)
        if not clean_name:
            return
        ProductMatchingService(self.session).create_product_article_if_missing(
            product_id=int(product_id),
            article=None,
            supplier_name=clean_name,
        )
