from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Optional
from decimal import Decimal

from sqlalchemy.orm import Session

from app.db.models import Product, ProductArticle
from app.services.qty_in_box_service import default_qty_in_box_for_pack, normalize_qty_in_box
from app.utils.excel_import import excel_text
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
    qty_in_box: int | None = None


class ProductMatchingService:
    BRAND_ALIASES = {
        "VAG": "VAG",
        "VOLKSWAGEN": "VAG",
        "VW": "VAG",
        "AUDI": "VAG",
        "SKODA": "VAG",
        "ŠKODA": "VAG",
        "GM": "GM",
        "GENERAL MOTORS": "GM",
    }

    def __init__(self, session: Session) -> None:
        self.session = session
        self._exact_products_cache: dict[str, Product] | None = None
        self._normalized_products_cache: dict[str, Product] | None = None
        self._exact_products_by_brand_cache: dict[tuple[str, str], Product] | None = None
        self._normalized_products_by_brand_cache: dict[tuple[str, str], Product] | None = None
        self._products_by_id_cache: dict[int, Product] | None = None
        self._product_article_keys_cache: set[tuple[int, str | None, str | None]] | None = None
        self._article_links_cache: dict[str, ProductArticle] | None = None
        self._article_links_by_brand_cache: dict[tuple[str, str], ProductArticle] | None = None
        self._name_links_cache: dict[str, ProductArticle] | None = None
        self._name_links_by_brand_cache: dict[tuple[str, str], ProductArticle] | None = None
        self._normalized_name_links_cache: dict[str, ProductArticle] | None = None
        self._normalized_name_links_by_brand_cache: dict[tuple[str, str], ProductArticle] | None = None

    @staticmethod
    def _brand_key(brand: object) -> str:
        cleaned = clean_multi_spaces(brand).upper()
        return ProductMatchingService.BRAND_ALIASES.get(cleaned, cleaned)

    def _ensure_product_caches(self) -> None:
        if self._exact_products_cache is not None and self._normalized_products_cache is not None:
            return

        exact_cache: dict[str, Product] = {}
        normalized_cache: dict[str, Product] = {}
        exact_by_brand: dict[tuple[str, str], Product] = {}
        normalized_by_brand: dict[tuple[str, str], Product] = {}
        products = (
            self.session.query(Product)
            .filter(Product.name.isnot(None))
            .order_by(Product.id.asc())
            .all()
        )
        for product in products:
            exact_key = clean_multi_spaces(product.name).upper()
            brand_key = self._brand_key(product.brand)
            if exact_key and exact_key not in exact_cache:
                exact_cache[exact_key] = product
            if exact_key and brand_key:
                exact_by_brand.setdefault((exact_key, brand_key), product)
            key = normalize_product_name(product.name)
            if key and key not in normalized_cache:
                normalized_cache[key] = product
            if key and brand_key:
                normalized_by_brand.setdefault((key, brand_key), product)

        self._exact_products_cache = exact_cache
        self._normalized_products_cache = normalized_cache
        self._exact_products_by_brand_cache = exact_by_brand
        self._normalized_products_by_brand_cache = normalized_by_brand
        self._products_by_id_cache = {int(product.id): product for product in products}

    def _build_normalized_products_cache(self) -> dict[str, Product]:
        self._ensure_product_caches()
        return self._normalized_products_cache or {}

    def _build_product_article_keys_cache(self) -> set[tuple[int, str | None, str | None]]:
        keys: set[tuple[int, str | None, str | None]] = set()
        rows = self.session.query(
            ProductArticle.product_id,
            ProductArticle.article,
            ProductArticle.name,
        ).all()
        for product_id, article, name in rows:
            keys.add((
                int(product_id),
                self.canonical_article_text(article) or None,
                clean_multi_spaces(name) or None,
            ))
        return keys

    def _ensure_link_caches(self) -> None:
        if all(
            cache is not None
            for cache in (
                self._article_links_cache,
                self._article_links_by_brand_cache,
                self._name_links_cache,
                self._name_links_by_brand_cache,
                self._normalized_name_links_cache,
                self._normalized_name_links_by_brand_cache,
            )
        ):
            return

        # ProductArticle used to be loaded three times from the network database:
        # once for articles and twice for the two name representations.  One
        # ordered snapshot is enough for all indices.  Iterating it backwards
        # preserves the former DESC priority for name matching.
        # Load products once and keep ProductArticle rows narrow.  Joining every
        # link to the full product record roughly doubles the network payload.
        self._ensure_product_caches()
        products_by_id = self._products_by_id_cache or {}
        links = self.session.query(ProductArticle).order_by(ProductArticle.id.asc()).all()

        article_cache: dict[str, ProductArticle] = {}
        article_by_brand: dict[tuple[str, str], ProductArticle] = {}
        for link in links:
            key = self.normalize_article_key(link.article)
            if key and key not in article_cache:
                article_cache[key] = link
            product = products_by_id.get(int(link.product_id))
            brand_key = self._brand_key(product.brand if product else None)
            if key and brand_key:
                article_by_brand.setdefault((key, brand_key), link)

        name_cache: dict[str, ProductArticle] = {}
        name_by_brand: dict[tuple[str, str], ProductArticle] = {}
        normalized_cache: dict[str, ProductArticle] = {}
        normalized_by_brand: dict[tuple[str, str], ProductArticle] = {}
        for link in reversed(links):
            product = products_by_id.get(int(link.product_id))
            brand_key = self._brand_key(product.brand if product else None)
            for key, cache, by_brand in (
                (clean_multi_spaces(link.name).upper(), name_cache, name_by_brand),
                (normalize_product_name(link.name), normalized_cache, normalized_by_brand),
            ):
                if not key:
                    continue
                current = cache.get(key)
                if current is None or (current.article and not link.article):
                    cache[key] = link
                brand_cache_key = (key, brand_key)
                brand_current = by_brand.get(brand_cache_key)
                if brand_key and (brand_current is None or (brand_current.article and not link.article)):
                    by_brand[brand_cache_key] = link

        self._article_links_cache = article_cache
        self._article_links_by_brand_cache = article_by_brand
        self._name_links_cache = name_cache
        self._name_links_by_brand_cache = name_by_brand
        self._normalized_name_links_cache = normalized_cache
        self._normalized_name_links_by_brand_cache = normalized_by_brand

    def _build_article_links_cache(self) -> dict[str, ProductArticle]:
        self._ensure_link_caches()
        return self._article_links_cache or {}

    def _build_name_links_cache(self) -> dict[str, ProductArticle]:
        self._ensure_link_caches()
        return self._name_links_cache or {}

    def _build_normalized_name_links_cache(self) -> dict[str, ProductArticle]:
        self._ensure_link_caches()
        return self._normalized_name_links_cache or {}

    # =========================================================
    # Article helpers
    # =========================================================

    @staticmethod
    def article_token_normalize(value: object) -> str:
        return ProductMatchingService.canonical_article_text(value)

    @staticmethod
    def canonical_article_text(value: object) -> str:
        return excel_text(value) or ""

    @classmethod
    def normalize_article_key(cls, value: object) -> str:
        return cls.canonical_article_text(value).casefold()

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

            token_key = cls.normalize_article_key(token)
            if token_key and token_key not in seen:
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
        num = parse_loose_number(value)
        if num is None:
            return ""

        if float(num).is_integer():
            return str(int(float(num)))

        s = str(num)
        if "." in s:
            s = s.rstrip("0").rstrip(".")
        return s

    @staticmethod
    def _pack_to_name_format(pack_value: object) -> str:
        s = ProductMatchingService._normalize_pack_text(pack_value)
        if not s:
            return ""

        if "." in s:
            return s.replace(".", ",")

        return s

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
        p_norm = cls._normalize_pack_text(pack_value)
        p_name = cls._pack_to_name_format(pack_value)
        unit_text = clean_multi_spaces(unit_text).upper()

        if not s:
            return ""

        if not p_norm:
            return s

        if re.search(r"([0-9]+(?:[.,][0-9]+)?)\s*(L|KG|л|Л|кг|КГ)\s*$", s, flags=re.IGNORECASE):
            return re.sub(
                r"([0-9]+(?:[.,][0-9]+)?)\s*(L|KG|л|Л|кг|КГ)\s*$",
                rf"\1 {unit_text}",
                s,
                flags=re.IGNORECASE,
            )

        tail_num = cls._get_trailing_number(s)
        if tail_num:
            try:
                if float(tail_num) == float(p_norm):
                    return f"{s} {unit_text}"
            except Exception:
                pass

        return f"{s} {p_name} {unit_text}"
    
    @classmethod
    def build_product_family_from_name(cls, product_name: str, pack_value: object) -> str:
        s_name = clean_multi_spaces(product_name).upper()
        expected_pack = ProductMatchingService._pack_to_name_format(pack_value)

        if not s_name:
            raise ValueError("Не заполнен ProductName.")

        if not expected_pack:
            raise ValueError("Не заполнен Pack.")

        pack_num = parse_loose_number(expected_pack)
        if pack_num is None:
            raise ValueError("Некорректный Pack.")

        matches = list(
            re.finditer(
                r"(?<!\d)([0-9]+(?:[.,][0-9]+)?)\s*(L|KG)\b",
                s_name,
                flags=re.IGNORECASE,
            )
        )

        for match in matches:
            found_num = parse_loose_number(match.group(1))
            if found_num is None:
                continue

            if float(found_num) == float(pack_num):
                family = s_name[:match.start()].strip()
                if not family:
                    raise ValueError(
                        f"Для '{product_name}' не удалось определить family до упаковки."
                    )
                return family

        raise ValueError(
            f"Для '{product_name}' проверь упаковку в названии. "
            f"Ожидается наличие упаковки '{expected_pack}L' или '{expected_pack}KG' "
            f"внутри названия, например: '... {expected_pack}L BIB'."
        )

    # =========================================================
    # Search helpers
    # =========================================================

    def _get_product_by_exact_name(self, product_name: str) -> Optional[Product]:
        self._ensure_product_caches()
        return (self._exact_products_cache or {}).get(clean_multi_spaces(product_name).upper())

    def _get_article_link_by_exact_article(self, article: str, brand: object = None) -> Optional[ProductArticle]:
        key = self.normalize_article_key(article)
        if not key:
            return None
        if self._article_links_cache is None:
            self._article_links_cache = self._build_article_links_cache()
        brand_key = self._brand_key(brand)
        if brand_key:
            return (self._article_links_by_brand_cache or {}).get((key, brand_key))
        return self._article_links_cache.get(key)

    def _get_article_link_by_exact_name(self, supplier_name: str, brand: object = None) -> Optional[ProductArticle]:
        key = clean_multi_spaces(supplier_name).upper()
        if not key:
            return None
        if self._name_links_cache is None:
            self._name_links_cache = self._build_name_links_cache()
        brand_key = self._brand_key(brand)
        if brand_key:
            return (self._name_links_by_brand_cache or {}).get((key, brand_key))
        return self._name_links_cache.get(key)

    def _get_article_link_by_normalized_name(self, supplier_name: str, brand: object = None) -> Optional[ProductArticle]:
        key = normalize_product_name(supplier_name)
        if not key:
            return None
        if self._normalized_name_links_cache is None:
            self._normalized_name_links_cache = self._build_normalized_name_links_cache()
        brand_key = self._brand_key(brand)
        if brand_key:
            return (self._normalized_name_links_by_brand_cache or {}).get((key, brand_key))
        return self._normalized_name_links_cache.get(key)

    def find_by_normalized_product_name(self, source_name: object, brand: object = None) -> Optional[Product]:
        target = normalize_product_name(source_name)
        if not target:
            return None
        if self._normalized_products_cache is None:
            self._normalized_products_cache = self._build_normalized_products_cache()
        brand_key = self._brand_key(brand)
        if brand_key:
            return (self._normalized_products_by_brand_cache or {}).get((target, brand_key))
        return self._normalized_products_cache.get(target)

    def find_product_by_split_articles(self, supplier_article: object, brand: object = None) -> Optional[Product]:
        first_product: Optional[Product] = None

        for token in self.split_article_tokens(supplier_article):
            link = self._get_article_link_by_exact_article(token, brand=brand)
            if link and link.product:
                if first_product is None:
                    first_product = link.product

        return first_product

    def _find_product_by_name_candidate(self, candidate: object, brand: object = None) -> Optional[Product]:
        name = clean_multi_spaces(candidate)
        if not name:
            return None

        link = self._get_article_link_by_exact_name(name, brand=brand)
        if link and link.product:
            return link.product

        link = self._get_article_link_by_normalized_name(name, brand=brand)
        if link and link.product:
            return link.product

        product = self.find_by_normalized_product_name(name, brand=brand)
        if product:
            return product

        return None

    # =========================================================
    # Public matching methods
    # =========================================================

    def find_customer_product(
        self,
        supplier_article: object,
        product_name: object,
        pack: object,
        brand: object = None,
    ) -> Optional[Product]:
        article = clean_multi_spaces(supplier_article)
        name = clean_multi_spaces(product_name)
        normalized_name = normalize_customer_product_name(name)

        # 1) Highest priority: article from ProductArticle.
        # Handles Excel numeric articles like 149610.0 and split articles like A/B.
        if article:
            product = self.find_product_by_split_articles(article, brand=brand)
            if product:
                return product

        # 2) Product name as supplied by the customer/request.
        # Match ProductArticle.name first, then Product.name.
        name_candidates: list[object] = []
        if name:
            name_candidates.append(name)
        if normalized_name and normalized_name != name:
            name_candidates.append(normalized_name)

        seen_names: set[str] = set()
        for candidate in name_candidates:
            candidate_key = normalize_product_name(candidate)
            if not candidate_key or candidate_key in seen_names:
                continue
            seen_names.add(candidate_key)

            product = self._find_product_by_name_candidate(candidate, brand=brand)
            if product:
                return product

        # 3) Product name + pack, only after plain article/name matching failed.
        for candidate in (
            self.build_name_with_pack_unit(name, pack, "L"),
            self.build_name_with_pack_unit(name, pack, "KG"),
        ):
            candidate_key = normalize_product_name(candidate)
            if not candidate_key or candidate_key in seen_names:
                continue
            seen_names.add(candidate_key)

            product = self._find_product_by_name_candidate(candidate, brand=brand)
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
            product = self._find_product_by_name_candidate(name)
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


    def find_is_product(
        self,
        source_article: object,
        source_product_name: object,
    ) -> Optional[Product]:
        article = clean_multi_spaces(source_article)
        name = clean_multi_spaces(source_product_name)

        if article:
            product = self.find_product_by_split_articles(article)
            if product:
                return product

        if name:
            exact = self._get_product_by_exact_name(name)
            if exact:
                return exact

            product = self._find_product_by_name_candidate(name)
            if product:
                return product

        return None


    @staticmethod
    def _format_pack_for_message(pack_value: object) -> str:
        if pack_value is None:
            return ""
        s = str(pack_value).strip()
        return s.replace(".", ",")

    @classmethod
    def validate_new_product_fields(
        cls,
        *,
        product_name: object,
        brand: object,
        pack: object,
        is_excise: object,
        qty_in_box: object = None,
    ) -> None:
        clean_name = clean_multi_spaces(product_name).upper()
        clean_brand = clean_multi_spaces(brand)
        pack_num = parse_loose_number(pack)

        if not clean_name:
            raise ValueError("Не заполнено название нового продукта.")

        if is_excise is None:
            raise ValueError(f"Для '{clean_name}' не заполнено поле 'Акциз'.")

        if not clean_brand:
            raise ValueError(f"Для '{clean_name}' не заполнено поле 'Бренд'.")

        if pack is None or str(pack).strip() == "":
            raise ValueError(f"Для '{clean_name}' не заполнено поле 'Упаковка'.")

        if pack_num is None:
            raise ValueError(
                f"Для '{clean_name}' поле 'Упаковка' должно быть числом."
            )

        normalize_qty_in_box(qty_in_box)

        cls.validate_product_name_pack_format(
            product_name=clean_name,
            pack_value=pack_num,
        )

    @classmethod
    def validate_product_name_pack_format(cls, *, product_name: str, pack_value: object) -> None:
        s_name = clean_multi_spaces(product_name).upper()
        expected_pack = cls._pack_to_name_format(pack_value)

        if not s_name:
            raise ValueError("Не заполнено название нового продукта.")

        if not expected_pack:
            raise ValueError(f"Для '{s_name}' не заполнено поле 'Упаковка'.")

        pack_num = parse_loose_number(expected_pack)
        if pack_num is None:
            raise ValueError(f"Для '{product_name}' некорректное поле 'Упаковка'.")

        matches = list(
            re.finditer(
                r"(?<!\d)([0-9]+(?:[.,][0-9]+)?)\s*(L|KG)\b",
                s_name,
                flags=re.IGNORECASE,
            )
        )

        for match in matches:
            found_num = parse_loose_number(match.group(1))
            if found_num is None:
                continue

            if float(found_num) == float(pack_num):
                return

        raise ValueError(
            f"Для '{product_name}' проверь упаковку в названии. "
            f"Ожидается наличие упаковки '{expected_pack}L' или '{expected_pack}KG' "
            f"внутри названия, например: '... {expected_pack}L BIB'."
        )


    def get_or_create_product(
        self,
        *,
        name: object,
        brand: object,
        pack: object,
        is_excise: object,
        qty_in_box: object = None,
    ) -> Product:
        return self.get_or_create_products_batch([
            ProductCreateData(
                name=name,
                brand=brand,
                pack=pack,
                is_excise=is_excise,
                qty_in_box=qty_in_box,
            )
        ])[0]

    def get_or_create_products_batch(self, items: list[ProductCreateData]) -> list[Product]:
        self._ensure_product_caches()
        exact_cache = self._exact_products_cache or {}
        normalized_cache = self._normalized_products_cache or {}
        exact_by_brand = self._exact_products_by_brand_cache or {}
        normalized_by_brand = self._normalized_products_by_brand_cache or {}

        resolved: list[Product] = []
        created: list[Product] = []

        for item in items:
            clean_name = clean_multi_spaces(item.name).upper()
            clean_brand = clean_multi_spaces(item.brand)
            pack_num = parse_loose_number(item.pack)
            qty_in_box = normalize_qty_in_box(item.qty_in_box)
            self.validate_new_product_fields(
                product_name=clean_name,
                brand=clean_brand,
                pack=pack_num,
                is_excise=item.is_excise,
                qty_in_box=qty_in_box,
            )

            product = exact_cache.get(clean_name)
            if product is None:
                product = normalized_cache.get(normalize_product_name(clean_name))

            if product is None:
                if qty_in_box is None:
                    qty_in_box = default_qty_in_box_for_pack(self.session, pack_num)
                product = Product(
                    name=clean_name,
                    brand=clean_brand,
                    pack=pack_num,
                    qty_in_box=qty_in_box,
                    is_excise=bool(item.is_excise),
                    family=self.build_product_family_from_name(clean_name, pack_num),
                )
                self.session.add(product)
                created.append(product)
                exact_cache[clean_name] = product
                normalized_name = normalize_product_name(clean_name)
                normalized_cache[normalized_name] = product
                brand_key = self._brand_key(clean_brand)
                exact_by_brand[(clean_name, brand_key)] = product
                normalized_by_brand[(normalized_name, brand_key)] = product
            elif product.qty_in_box is None:
                resolved_qty_in_box = qty_in_box
                if resolved_qty_in_box is None:
                    resolved_qty_in_box = default_qty_in_box_for_pack(self.session, product.pack)
                if resolved_qty_in_box is not None:
                    product.qty_in_box = resolved_qty_in_box

            resolved.append(product)

        if created:
            self.session.flush()
            products_by_id = self._products_by_id_cache or {}
            for product in created:
                products_by_id[int(product.id)] = product
            self._products_by_id_cache = products_by_id

        self._exact_products_cache = exact_cache
        self._normalized_products_cache = normalized_cache
        self._exact_products_by_brand_cache = exact_by_brand
        self._normalized_products_by_brand_cache = normalized_by_brand
        return resolved

    def create_product_articles_if_missing_batch(
        self,
        links: list[tuple[int, object, object]],
    ) -> int:
        if self._product_article_keys_cache is None:
            self._product_article_keys_cache = self._build_product_article_keys_cache()

        created: list[ProductArticle] = []
        no_article_links: list[tuple[int, object]] = []
        for product_id, article, supplier_name in links:
            clean_article = self.canonical_article_text(article)
            clean_name = clean_multi_spaces(supplier_name)
            if not clean_article:
                if clean_name:
                    no_article_links.append((int(product_id), clean_name))
                continue
            if not self.should_create_article_link(clean_article):
                continue

            key = (int(product_id), clean_article, clean_name or None)
            if key in self._product_article_keys_cache:
                continue
            self._product_article_keys_cache.add(key)
            created.append(ProductArticle(
                product_id=int(product_id),
                article=clean_article,
                name=clean_name or None,
            ))

        if created:
            self.session.add_all(created)
            self.session.flush()

        for product_id, clean_name in no_article_links:
            self.create_product_article_if_missing(
                product_id=product_id,
                article=None,
                supplier_name=clean_name,
            )

        if created:
            self._article_links_cache = None
            self._article_links_by_brand_cache = None
            self._name_links_cache = None
            self._name_links_by_brand_cache = None
            self._normalized_name_links_cache = None
            self._normalized_name_links_by_brand_cache = None
        return len(created)

    def create_product_article_if_missing(
        self,
        *,
        product_id: int,
        article: object,
        supplier_name: object,
    ) -> Optional[ProductArticle]:
        clean_article = self.canonical_article_text(article)
        clean_name = clean_multi_spaces(supplier_name)

        if not clean_article and not clean_name:
            return None

        # Without Supplier Article, Supplier Product Name is the mapping key.
        # The latest manual selection must replace an older mapping instead of
        # leaving several ambiguous aliases for the same supplier spelling.
        if not clean_article and clean_name:
            name_key = normalize_product_name(clean_name)
            name_links = (
                self.session.query(ProductArticle)
                .filter(
                    (ProductArticle.article.is_(None)) | (ProductArticle.article == ""),
                    ProductArticle.name.isnot(None),
                    ProductArticle.name != "",
                )
                .order_by(ProductArticle.id.desc())
                .all()
            )
            same_name_links = [link for link in name_links if normalize_product_name(link.name) == name_key]

            if same_name_links:
                current = same_name_links[0]
                current.product_id = product_id
                current.name = clean_name

                for duplicate in same_name_links[1:]:
                    self.session.delete(duplicate)

                self.session.flush()
                self._article_links_cache = None
                self._article_links_by_brand_cache = None
                self._name_links_cache = None
                self._name_links_by_brand_cache = None
                self._normalized_name_links_cache = None
                self._normalized_name_links_by_brand_cache = None
                self._product_article_keys_cache = None
                return current

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

        self._article_links_cache = None
        self._article_links_by_brand_cache = None
        self._name_links_cache = None
        self._name_links_by_brand_cache = None
        self._normalized_name_links_cache = None
        self._normalized_name_links_by_brand_cache = None
        self._product_article_keys_cache = None
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
        clean_article = self.canonical_article_text(source_article)
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
