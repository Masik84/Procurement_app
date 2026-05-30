from sqlalchemy import (
    Boolean,
    Column,
    DateTime,
    Date,
    Numeric,
    ForeignKey,
    Integer,
    String,
    Text,
    UniqueConstraint,
    Index,
)
from sqlalchemy.orm import relationship

from app.db.db import Base


# ============================================================
# MASTER TABLES
# ============================================================

class Product(Base):
    __tablename__ = "products"

    id = Column(Integer, primary_key=True, autoincrement=True)
    brand = Column(String(255), nullable=False, index=True)
    name = Column(String(500), nullable=False, unique=True, index=True)
    family = Column(String(500), nullable=True, index=True)
    pack = Column(Numeric, nullable=False)
    is_excise = Column(Boolean, nullable=False, default=False)

    articles = relationship("ProductArticle", back_populates="product", passive_deletes=True)
    price_history = relationship("PriceHistory", back_populates="product", passive_deletes=True)
    current_prices = relationship("CurrentSupplierPrice", back_populates="product", passive_deletes=True)
    price_calculations = relationship("SupplierPriceCalculation", back_populates="product", passive_deletes=True)
    stock = relationship("ProductStock", back_populates="product", uselist=False, passive_deletes=True)

    temp_price_import_rows = relationship("TempPriceImport", back_populates="selected_product", passive_deletes=True)
    temp_customer_cost_rows = relationship("TempCustomerCostImport", back_populates="selected_product", passive_deletes=True)
    temp_customer_cost_options = relationship("TempCustomerCostOption", back_populates="product", passive_deletes=True)
    temp_target_price_rows = relationship("TempTargetPriceImport", back_populates="selected_product", passive_deletes=True)
    temp_target_price_options = relationship("TempTargetPriceOption", back_populates="product", passive_deletes=True)
    temp_stock_import_rows = relationship("TempStockImport", back_populates="selected_product", passive_deletes=True)
    temp_supplier_orders_rows = relationship("TempSupplierOrdersImport", back_populates="selected_product", passive_deletes=True)
    temp_is_rows = relationship("TempIsImport", back_populates="selected_product", passive_deletes=True)
    temp_product_search_rows = relationship("TempProductSearchImport", back_populates="selected_product", passive_deletes=True)
    sales_links = relationship("SalesProductLink", back_populates="product", passive_deletes=True)
    order_planning_calculations = relationship("OrderPlanningCalculation", back_populates="product", passive_deletes=True)

    customer_price_calculations = relationship(
        "CustomerPriceCalculation",
        back_populates="product",
        passive_deletes=True,
    )
    target_price_calculations = relationship(
        "TargetPriceCalculation",
        back_populates="product",
        passive_deletes=True,
    )


class Supplier(Base):
    __tablename__ = "suppliers"

    id = Column(Integer, primary_key=True, autoincrement=True)
    name = Column(String(255), nullable=False, unique=True, index=True)

    base_currency = Column(String(10), nullable=False)
    transport_cost_per_l = Column(Numeric, nullable=False, default=0)
    reexport_percent = Column(Numeric, nullable=False, default=0)
    fx_rate_markup = Column(Numeric, nullable=False, default=0)
    agent_fee = Column(Numeric, nullable=False, default=0)

    is_via_novo = Column(Boolean, nullable=False, default=False)
    has_import_duty = Column(Boolean, nullable=False, default=False)
    rating_calc = Column(Boolean, nullable=False, default=True)
    marks_for_us = Column(Boolean, nullable=False, default=False)
    is_rf = Column(Boolean, nullable=False, default=False)
    country = Column(String(50))

    price_history = relationship("PriceHistory", back_populates="supplier", passive_deletes=True)
    current_prices = relationship("CurrentSupplierPrice", back_populates="supplier", passive_deletes=True)
    price_calculations = relationship("SupplierPriceCalculation", back_populates="supplier", passive_deletes=True)
    temp_price_import_rows = relationship("TempPriceImport", back_populates="supplier", passive_deletes=True)
    temp_customer_cost_options = relationship(
        "TempCustomerCostOption",
        back_populates="supplier",
        passive_deletes=True,
    )
    customer_price_calculations = relationship(
        "CustomerPriceCalculation",
        back_populates="supplier",
        passive_deletes=True,
    )
    temp_target_price_rows = relationship(
        "TempTargetPriceImport",
        back_populates="target_supplier",
        foreign_keys="TempTargetPriceImport.target_supplier_id",
        passive_deletes=True,
    )
    temp_target_price_options = relationship(
        "TempTargetPriceOption",
        back_populates="supplier",
        passive_deletes=True,
    )
    target_price_calculations = relationship(
        "TargetPriceCalculation",
        back_populates="target_supplier",
        foreign_keys="TargetPriceCalculation.target_supplier_id",
        passive_deletes=True,
    )
    target_price_donor_calculations = relationship(
        "TargetPriceCalculation",
        back_populates="donor_supplier",
        foreign_keys="TargetPriceCalculation.donor_supplier_id",
        passive_deletes=True,
    )


class ProductArticle(Base):
    __tablename__ = "product_articles"

    id = Column(Integer, primary_key=True, autoincrement=True)

    product_id = Column(
        Integer,
        ForeignKey("products.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )
    article = Column(String(255), nullable=True, index=True)
    name = Column(String(500), nullable=True, index=True)

    product = relationship("Product", back_populates="articles")

    __table_args__ = (
        UniqueConstraint(
            "product_id",
            "article",
            "name",
            name="uq_prod_articles_main",
        ),
        Index("ix_prod_articles_article_name", "article", "name"),
    )


class ExchangeRate(Base):
    __tablename__ = "exchange_rates"

    currency_code = Column(String(10), primary_key=True)
    rate_to_rub = Column(Numeric, nullable=False)


class FixedCosts(Base):
    __tablename__ = "fixed_costs"

    id = Column(Integer, primary_key=True, autoincrement=True)

    customs_clearance = Column(Numeric, nullable=False, default=0)
    additional_customs = Column(Numeric, nullable=False, default=0)
    excise = Column(Numeric, nullable=False, default=0)
    eco_fee = Column(Numeric, nullable=False, default=0)
    vat = Column(Numeric, nullable=False, default=0)
    customs_fee = Column(Numeric, nullable=False, default=0)
    bank_fee = Column(Numeric, nullable=False, default=0)
    money = Column(Numeric, nullable=False, default=0)
    storage = Column(Numeric, nullable=False, default=0)
    move_novo_tamozh = Column(Numeric, nullable=False, default=0)
    move_tamozh_chekhov = Column(Numeric, nullable=False, default=0)


class PackType(Base):
    __tablename__ = "pack_types"

    id = Column(Integer, primary_key=True, autoincrement=True)
    name = Column(String(100), nullable=False, index=True)
    volume = Column(Numeric, nullable=False, unique=True, index=True)


class MarkingRate(Base):
    __tablename__ = "marking_rates"

    pack_type = Column(String(100), primary_key=True)
    cost_per_l = Column(Numeric, nullable=False)


# ============================================================
# PRICE TABLES
# ============================================================

class PriceHistory(Base):
    __tablename__ = "price_history"

    id = Column(Integer, primary_key=True, autoincrement=True)

    supplier_id = Column(
        Integer,
        ForeignKey("suppliers.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )
    product_id = Column(
        Integer,
        ForeignKey("products.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )

    price_date = Column(DateTime, nullable=False, index=True)
    price = Column(Numeric, nullable=False)
    currency = Column(String(10), nullable=False)

    supplier = relationship("Supplier", back_populates="price_history")
    product = relationship("Product", back_populates="price_history")


class CurrentSupplierPrice(Base):
    __tablename__ = "current_supplier_prices"

    supplier_id = Column(
        Integer,
        ForeignKey("suppliers.id", ondelete="CASCADE"),
        primary_key=True,
    )
    product_id = Column(
        Integer,
        ForeignKey("products.id", ondelete="CASCADE"),
        primary_key=True,
    )

    price = Column(Numeric, nullable=False)
    currency = Column(String(10), nullable=False)
    last_update = Column(DateTime, nullable=False)

    supplier = relationship("Supplier", back_populates="current_prices")
    product = relationship("Product", back_populates="current_prices")


class SupplierPriceCalculation(Base):
    __tablename__ = "supplier_price_calculations"

    id = Column(Integer, primary_key=True, autoincrement=True)

    calc_date = Column(DateTime, nullable=False)
    batch_id = Column(String(64), nullable=False, index=True)
    imported_by = Column(String(255), nullable=False, index=True)
    import_row_no = Column(Integer, nullable=True)

    supplier_id = Column(
        Integer,
        ForeignKey("suppliers.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )
    product_id = Column(
        Integer,
        ForeignKey("products.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )

    supplier_article = Column(String(255), nullable=True)
    supplier_product_name = Column(String(500), nullable=True)

    supplier_price = Column(Numeric, nullable=False)
    cost_novo_wvat = Column(Numeric, nullable=False)
    full_cost_msk = Column(Numeric, nullable=False)

    currency_code = Column(String(10), nullable=False)
    fx_rate_used = Column(Numeric, nullable=False)
    fx_markup_used = Column(Numeric, nullable=False)
    transport_used = Column(Numeric, nullable=False)
    reexport_used = Column(Numeric, nullable=False)
    agent_fee_used = Column(Numeric, nullable=False)

    has_customs_used = Column(Boolean, nullable=False)
    via_novo_used = Column(Boolean, nullable=False)
    bank_fee_used = Column(Numeric, nullable=False)
    customs_fee_used = Column(Numeric, nullable=False)
    move_novo_used = Column(Numeric, nullable=False)
    move_msk_used = Column(Numeric, nullable=False)
    is_excise_used = Column(Boolean, nullable=False)
    additional_customs_used = Column(Numeric, nullable=False)
    storage_used = Column(Numeric, nullable=False)
    marking_used = Column(Numeric, nullable=False)

    supplier = relationship("Supplier", back_populates="price_calculations")
    product = relationship("Product", back_populates="price_calculations")


# ============================================================
# STOCK / ORDERS / IS
# ============================================================

class ProductStock(Base):
    __tablename__ = "product_stock"

    product_id = Column(
        Integer,
        ForeignKey("products.id", ondelete="CASCADE"),
        primary_key=True,
    )

    product_name = Column(String(500), nullable=False)

    stock_update_date = Column(DateTime, nullable=True)
    supplier_orders_update_date = Column(DateTime, nullable=True)
    is_update_date = Column(DateTime, nullable=True)

    stock_qty = Column(Numeric, nullable=False, default=0)
    markdown_qty = Column(Numeric, nullable=False, default=0)
    reserve_qty = Column(Numeric, nullable=False, default=0)
    reserve_ecomm_qty = Column(Numeric, nullable=False, default=0)

    lpc = Column(Numeric, nullable=False, default=0)
    landed_cost = Column(Numeric, nullable=False, default=0)
    distr_price = Column(Numeric, nullable=False, default=0)
    promo_price = Column(Numeric, nullable=False, default=0)

    transit_qty = Column(Numeric, nullable=False, default=0)
    order_qty = Column(Numeric, nullable=False, default=0)

    is_order_qty = Column(Numeric, nullable=False, default=0)
    is_confirmed_order_qty = Column(Numeric, nullable=False, default=0)
    is_stock_qty = Column(Numeric, nullable=False, default=0)

    product = relationship("Product", back_populates="stock")


# ============================================================
# TEMP / STAGING TABLES
# ============================================================

class TempPriceImport(Base):
    __tablename__ = "temp_price_import"

    id = Column(Integer, primary_key=True, autoincrement=True)

    supplier_article = Column(String(255), nullable=True)
    product_name = Column(String(500), nullable=True)

    price = Column(Numeric, nullable=True)
    price_pack = Column(Numeric, nullable=True)
    qty_pcs = Column(Numeric, nullable=True)
    volume_l = Column(Numeric, nullable=True)

    supplier_id = Column(
        Integer,
        ForeignKey("suppliers.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )

    import_date = Column(DateTime, nullable=False)
    batch_id = Column(String(64), nullable=False, index=True)
    imported_by = Column(String(255), nullable=False, index=True)

    selected_product_id = Column(
        Integer,
        ForeignKey("products.id", ondelete="CASCADE"),
        nullable=True,
        index=True,
    )

    import_row_no = Column(Integer, nullable=True)

    new_product_name = Column(String(500), nullable=True)
    new_brand = Column(String(255), nullable=True)
    new_pack = Column(Numeric, nullable=True)
    new_is_excise = Column(Boolean, nullable=True)

    supplier = relationship("Supplier", back_populates="temp_price_import_rows")
    selected_product = relationship("Product", back_populates="temp_price_import_rows")

    __table_args__ = (
        Index("ix_temp_price_batch_user", "batch_id", "imported_by"),
    )


class TempCustomerCostImport(Base):
    __tablename__ = "temp_customer_cost_import"

    id = Column(Integer, primary_key=True, autoincrement=True)

    batch_id = Column(String(64), nullable=False, index=True)
    imported_by = Column(String(255), nullable=False, index=True)
    import_row_no = Column(Integer, nullable=True)
    import_date = Column(DateTime, nullable=False)

    request_date = Column(DateTime, nullable=True)
    manager_name = Column(String(255), nullable=True)
    customer_name = Column(String(255), nullable=True)

    supplier_article = Column(String(255), nullable=True)
    product_name = Column(String(500), nullable=True)

    pack = Column(Numeric, nullable=True)
    qty_pcs = Column(Numeric, nullable=True)
    volume_l = Column(Numeric, nullable=True)

    purchase_type = Column(String(255), nullable=True)
    payment_terms = Column(String(255), nullable=True)
    comments = Column(Text, nullable=True)

    selected_product_id = Column(
        Integer,
        ForeignKey("products.id", ondelete="CASCADE"),
        nullable=True,
        index=True,
    )

    selected_option_id = Column(Integer, nullable=True)

    new_product_name = Column(String(500), nullable=True)
    new_brand = Column(String(255), nullable=True)
    new_pack = Column(Numeric, nullable=True)
    new_is_excise = Column(Boolean, nullable=True)

    selected_product = relationship("Product", back_populates="temp_customer_cost_rows")

    options = relationship(
        "TempCustomerCostOption",
        back_populates="temp_import",
        foreign_keys="TempCustomerCostOption.temp_import_id",
        passive_deletes=True,
    )

    __table_args__ = (
        Index("ix_temp_cc_import_batch_user", "batch_id", "imported_by"),
    )


class TempCustomerCostOption(Base):
    __tablename__ = "temp_customer_cost_options"

    id = Column(Integer, primary_key=True, autoincrement=True)

    temp_import_id = Column(
        Integer,
        ForeignKey("temp_customer_cost_import.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )

    batch_id = Column(String(64), nullable=False, index=True)
    imported_by = Column(String(255), nullable=False, index=True)
    calc_date = Column(DateTime, nullable=False)

    supplier_id = Column(
        Integer,
        ForeignKey("suppliers.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )
    product_id = Column(
        Integer,
        ForeignKey("products.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )

    supplier_name = Column(String(255), nullable=False)
    supplier_article = Column(String(255), nullable=True)
    customer_product_name = Column(String(500), nullable=True)

    supplier_price = Column(Numeric, nullable=False)
    price_date_used = Column(DateTime, nullable=True)

    cost_novo_wvat = Column(Numeric, nullable=False)
    full_cost_msk = Column(Numeric, nullable=False)

    currency_code = Column(String(10), nullable=False)
    fx_rate_used = Column(Numeric, nullable=False)
    fx_markup_used = Column(Numeric, nullable=False)
    transport_used = Column(Numeric, nullable=False)
    reexport_used = Column(Numeric, nullable=False)
    agent_fee_used = Column(Numeric, nullable=False)

    has_customs_used = Column(Boolean, nullable=False)
    via_novo_used = Column(Boolean, nullable=False)
    bank_fee_used = Column(Numeric, nullable=False)
    customs_fee_used = Column(Numeric, nullable=False)
    move_novo_used = Column(Numeric, nullable=False)
    move_msk_used = Column(Numeric, nullable=False)
    is_excise_used = Column(Boolean, nullable=False)
    additional_customs_used = Column(Numeric, nullable=False)
    storage_used = Column(Numeric, nullable=False)
    marking_used = Column(Numeric, nullable=False)

    opt_rank = Column(Integer, nullable=True)

    temp_import = relationship(
        "TempCustomerCostImport",
        back_populates="options",
        foreign_keys=[temp_import_id],
    )
    supplier = relationship("Supplier", back_populates="temp_customer_cost_options")
    product = relationship("Product", back_populates="temp_customer_cost_options")

    __table_args__ = (
        Index("ix_temp_cc_options_batch_user", "batch_id", "imported_by"),
    )


class TempStockImport(Base):
    __tablename__ = "temp_stock_import"

    id = Column(Integer, primary_key=True, autoincrement=True)

    batch_id = Column(String(64), nullable=False, index=True)
    imported_by = Column(String(255), nullable=False, index=True)
    import_date = Column(DateTime, nullable=False)
    import_row_no = Column(Integer, nullable=True)

    source_article = Column(String(255), nullable=True)
    source_sku = Column(String(255), nullable=True)
    source_product_name = Column(String(500), nullable=True)
    source_origin = Column(String(255), nullable=True)
    source_brand_group = Column(String(255), nullable=True)

    lpc = Column(Numeric, nullable=True)
    landed_cost = Column(Numeric, nullable=True)
    distr_price = Column(Numeric, nullable=True)
    promo_price = Column(Numeric, nullable=True)

    stock_qty = Column(Numeric, nullable=False, default=0)
    transit_qty = Column(Numeric, nullable=False, default=0)
    markdown_qty = Column(Numeric, nullable=False, default=0)
    reserve_qty = Column(Numeric, nullable=False, default=0)
    reserve_ecomm_qty = Column(Numeric, nullable=False, default=0)
    
    selected_product_id = Column(
        Integer,
        ForeignKey("products.id", ondelete="CASCADE"),
        nullable=True,
        index=True,
    )

    has_lpc_warning = Column(Boolean, nullable=False, default=False)

    new_product_name = Column(String(500), nullable=True)
    new_brand = Column(String(255), nullable=True)
    new_pack = Column(Numeric, nullable=True)
    new_is_excise = Column(Boolean, nullable=True)

    selected_product = relationship("Product", back_populates="temp_stock_import_rows")

    __table_args__ = (
        Index("ix_temp_stock_batch_user", "batch_id", "imported_by"),
    )


class TempSupplierOrdersImport(Base):
    __tablename__ = "temp_supplier_orders_import"

    id = Column(Integer, primary_key=True, autoincrement=True)

    batch_id = Column(String(64), nullable=False, index=True)
    imported_by = Column(String(255), nullable=False, index=True)
    import_date = Column(DateTime, nullable=False)
    import_row_no = Column(Integer, nullable=True)

    source_article = Column(String(255), nullable=True)
    source_product_name = Column(String(500), nullable=True)

    order_qty = Column(Numeric, nullable=False, default=0)

    selected_product_id = Column(
        Integer,
        ForeignKey("products.id", ondelete="CASCADE"),
        nullable=True,
        index=True,
    )

    new_product_name = Column(String(500), nullable=True)
    new_brand = Column(String(255), nullable=True)
    new_pack = Column(Numeric, nullable=True)
    new_is_excise = Column(Boolean, nullable=True)

    selected_product = relationship("Product", back_populates="temp_supplier_orders_rows")

    __table_args__ = (
        Index("ix_temp_so_batch_user", "batch_id", "imported_by"),
    )


class TempIsImport(Base):
    __tablename__ = "temp_is_import"

    id = Column(Integer, primary_key=True, autoincrement=True)

    batch_id = Column(String(64), nullable=False, index=True)
    imported_by = Column(String(255), nullable=False, index=True)
    import_date = Column(DateTime, nullable=False)
    import_row_no = Column(Integer, nullable=True)

    source_article = Column(String(255), nullable=True)
    source_product_name = Column(String(500), nullable=True)

    confirmed_qty = Column(Numeric, nullable=False, default=0)
    remains_qty = Column(Numeric, nullable=False, default=0)
    stock_qty = Column(Numeric, nullable=False, default=0)

    selected_product_id = Column(
        Integer,
        ForeignKey("products.id", ondelete="CASCADE"),
        nullable=True,
        index=True,
    )

    new_product_name = Column(String(500), nullable=True)
    new_brand = Column(String(255), nullable=True)
    new_pack = Column(Numeric, nullable=True)
    new_is_excise = Column(Boolean, nullable=True)

    selected_product = relationship("Product", back_populates="temp_is_rows")

    __table_args__ = (
        Index("ix_temp_is_batch_user", "batch_id", "imported_by"),
    )


class TempProductSearchImport(Base):
    __tablename__ = "temp_product_search_import"

    id = Column(Integer, primary_key=True, autoincrement=True)

    batch_id = Column(String(64), nullable=False, index=True)
    imported_by = Column(String(255), nullable=False, index=True)
    import_date = Column(DateTime, nullable=False)
    import_row_no = Column(Integer, nullable=True)

    source_article = Column(String(255), nullable=True)
    source_product_name = Column(String(500), nullable=True)

    selected_product_id = Column(
        Integer,
        ForeignKey("products.id", ondelete="CASCADE"),
        nullable=True,
        index=True,
    )

    new_product_name = Column(String(500), nullable=True)
    new_brand = Column(String(255), nullable=True)
    new_pack = Column(Numeric, nullable=True)
    new_is_excise = Column(Boolean, nullable=True)

    selected_product = relationship("Product", back_populates="temp_product_search_rows")

    __table_args__ = (
        Index("ix_temp_product_search_batch_user", "batch_id", "imported_by"),
    )



class TempTargetPriceImport(Base):
    __tablename__ = "temp_target_price_import"

    id = Column(Integer, primary_key=True, autoincrement=True)

    batch_id = Column(String(64), nullable=False, index=True)
    imported_by = Column(String(255), nullable=False, index=True)
    import_row_no = Column(Integer, nullable=True)
    import_date = Column(DateTime, nullable=False)

    target_supplier_id = Column(
        Integer,
        ForeignKey("suppliers.id", ondelete="CASCADE"),
        nullable=True,
        index=True,
    )

    supplier_article = Column(String(255), nullable=True)
    product_name = Column(String(500), nullable=True)

    selected_product_id = Column(
        Integer,
        ForeignKey("products.id", ondelete="CASCADE"),
        nullable=True,
        index=True,
    )
    selected_option_id = Column(Integer, nullable=True)

    new_product_name = Column(String(500), nullable=True)
    new_brand = Column(String(255), nullable=True)
    new_pack = Column(Numeric, nullable=True)
    new_is_excise = Column(Boolean, nullable=True)

    target_supplier = relationship(
        "Supplier",
        back_populates="temp_target_price_rows",
        foreign_keys=[target_supplier_id],
    )
    selected_product = relationship("Product", back_populates="temp_target_price_rows")
    options = relationship(
        "TempTargetPriceOption",
        back_populates="temp_import",
        foreign_keys="TempTargetPriceOption.temp_import_id",
        passive_deletes=True,
    )

    __table_args__ = (
        Index("ix_temp_target_price_batch_user", "batch_id", "imported_by"),
    )


class TempTargetPriceOption(Base):
    __tablename__ = "temp_target_price_options"

    id = Column(Integer, primary_key=True, autoincrement=True)

    temp_import_id = Column(
        Integer,
        ForeignKey("temp_target_price_import.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )
    batch_id = Column(String(64), nullable=False, index=True)
    imported_by = Column(String(255), nullable=False, index=True)
    calc_date = Column(DateTime, nullable=False)

    supplier_id = Column(
        Integer,
        ForeignKey("suppliers.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )
    product_id = Column(
        Integer,
        ForeignKey("products.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )

    supplier_name = Column(String(255), nullable=False)
    supplier_article = Column(String(255), nullable=True)
    supplier_product_name = Column(String(500), nullable=True)

    supplier_price = Column(Numeric, nullable=False)
    price_date_used = Column(DateTime, nullable=True)

    cost_novo_wvat = Column(Numeric, nullable=False)
    full_cost_msk = Column(Numeric, nullable=False)

    currency_code = Column(String(10), nullable=False)
    fx_rate_used = Column(Numeric, nullable=False)
    fx_markup_used = Column(Numeric, nullable=False)
    transport_used = Column(Numeric, nullable=False)
    reexport_used = Column(Numeric, nullable=False)
    agent_fee_used = Column(Numeric, nullable=False)

    has_customs_used = Column(Boolean, nullable=False)
    via_novo_used = Column(Boolean, nullable=False)
    bank_fee_used = Column(Numeric, nullable=False)
    customs_fee_used = Column(Numeric, nullable=False)
    move_novo_used = Column(Numeric, nullable=False)
    move_msk_used = Column(Numeric, nullable=False)
    is_excise_used = Column(Boolean, nullable=False)
    additional_customs_used = Column(Numeric, nullable=False)
    storage_used = Column(Numeric, nullable=False)
    marking_used = Column(Numeric, nullable=False)

    opt_rank = Column(Integer, nullable=True)

    temp_import = relationship(
        "TempTargetPriceImport",
        back_populates="options",
        foreign_keys=[temp_import_id],
    )
    supplier = relationship("Supplier", back_populates="temp_target_price_options")
    product = relationship("Product", back_populates="temp_target_price_options")

    __table_args__ = (
        Index("ix_temp_target_price_options_batch_user", "batch_id", "imported_by"),
    )


class TargetPriceCalculation(Base):
    __tablename__ = "target_price_calculations"

    id = Column(Integer, primary_key=True, autoincrement=True)

    calc_date = Column(DateTime, nullable=False)
    batch_id = Column(String(64), nullable=False, index=True)
    imported_by = Column(String(255), nullable=False, index=True)
    import_row_no = Column(Integer, nullable=True)

    target_supplier_id = Column(
        Integer,
        ForeignKey("suppliers.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )
    donor_supplier_id = Column(
        Integer,
        ForeignKey("suppliers.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )
    product_id = Column(
        Integer,
        ForeignKey("products.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )

    supplier_article = Column(String(255), nullable=True)
    supplier_product_name = Column(String(500), nullable=True)

    target_price_l = Column(Numeric, nullable=False)
    target_price_pack = Column(Numeric, nullable=False)
    currency_code = Column(String(10), nullable=False)
    fx_rate_used = Column(Numeric, nullable=False)
    full_cost_msk_source = Column(Numeric, nullable=False)
    cost_novo_wvat_recalculated = Column(Numeric, nullable=False)

    fx_markup_used = Column(Numeric, nullable=False)
    transport_used = Column(Numeric, nullable=False)
    reexport_used = Column(Numeric, nullable=False)
    agent_fee_used = Column(Numeric, nullable=False)
    has_customs_used = Column(Boolean, nullable=False)
    via_novo_used = Column(Boolean, nullable=False)
    bank_fee_used = Column(Numeric, nullable=False)
    customs_fee_used = Column(Numeric, nullable=False)
    additional_customs_used = Column(Numeric, nullable=False)
    storage_used = Column(Numeric, nullable=False)
    move_novo_used = Column(Numeric, nullable=False)
    move_msk_used = Column(Numeric, nullable=False)
    marking_used = Column(Numeric, nullable=False)
    is_excise_used = Column(Boolean, nullable=False)
    vat_used = Column(Numeric, nullable=False)
    money_used = Column(Numeric, nullable=False)
    price_date_used = Column(DateTime, nullable=True)

    target_supplier = relationship(
        "Supplier",
        back_populates="target_price_calculations",
        foreign_keys=[target_supplier_id],
    )
    donor_supplier = relationship(
        "Supplier",
        back_populates="target_price_donor_calculations",
        foreign_keys=[donor_supplier_id],
    )
    product = relationship("Product", back_populates="target_price_calculations")

    __table_args__ = (
        Index("ix_target_price_calc_batch_user", "batch_id", "imported_by"),
    )


class CustomerPriceCalculation(Base):
    __tablename__ = "customer_price_calculations"

    id = Column(Integer, primary_key=True, autoincrement=True)

    calc_date = Column(DateTime, nullable=False)
    batch_id = Column(String(64), nullable=False, index=True)
    imported_by = Column(String(255), nullable=False, index=True)

    manager_name = Column(String(255), nullable=True)
    customer_name = Column(String(255), nullable=True)
    request_date = Column(DateTime, nullable=True)

    supplier_id = Column(
        Integer,
        ForeignKey("suppliers.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )
    product_id = Column(
        Integer,
        ForeignKey("products.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )

    supplier_article = Column(String(255), nullable=True)
    customer_product_name = Column(String(500), nullable=True)

    pack = Column(Numeric, nullable=True)
    qty_pcs = Column(Numeric, nullable=True)
    volume_l = Column(Numeric, nullable=True)
    comments = Column(Text, nullable=True)

    supplier_price = Column(Numeric, nullable=False)
    cost_novo_wvat = Column(Numeric, nullable=False)
    full_cost_msk = Column(Numeric, nullable=False)

    currency_code = Column(String(10), nullable=False)
    fx_rate_used = Column(Numeric, nullable=False)
    fx_markup_used = Column(Numeric, nullable=False)
    transport_used = Column(Numeric, nullable=False)
    reexport_used = Column(Numeric, nullable=False)
    agent_fee_used = Column(Numeric, nullable=False)

    has_customs_used = Column(Boolean, nullable=False)
    via_novo_used = Column(Boolean, nullable=False)
    bank_fee_used = Column(Numeric, nullable=False)
    customs_fee_used = Column(Numeric, nullable=False)
    additional_customs_used = Column(Numeric, nullable=False)
    storage_used = Column(Numeric, nullable=False)
    move_novo_used = Column(Numeric, nullable=False)
    move_msk_used = Column(Numeric, nullable=False)
    marking_used = Column(Numeric, nullable=False)
    is_excise_used = Column(Boolean, nullable=False)

    price_date_used = Column(DateTime, nullable=True)
    import_row_no = Column(Integer, nullable=True)

    supplier = relationship("Supplier", back_populates="customer_price_calculations")
    product = relationship("Product", back_populates="customer_price_calculations")

    __table_args__ = (
        Index("ix_customer_price_calc_batch_user", "batch_id", "imported_by"),
    )

# ============================================================
# ORDER PLANNING
# ============================================================

class SalesProductLink(Base):
    __tablename__ = "sales_product_links"

    id = Column(Integer, primary_key=True, autoincrement=True)

    # Код продукта из БД продаж. В БД продаж это одновременно id и уникальный код продукта.
    sales_code = Column(String(255), nullable=False, unique=True, index=True)

    product_id = Column(
        Integer,
        ForeignKey("products.id", ondelete="SET NULL"),
        nullable=True,
        index=True,
    )

    sales_article = Column(String(255), nullable=True)
    sales_product_name = Column(String(500), nullable=True)
    sales_pack = Column(Numeric, nullable=True)
    sales_brand = Column(String(255), nullable=True)
    sales_is_excise = Column(Boolean, nullable=True)

    updated_at = Column(DateTime, nullable=False)

    product = relationship("Product", back_populates="sales_links")

    __table_args__ = (
        Index("ix_sales_product_links_product_id", "product_id"),
    )


class OrderPlanningCalculation(Base):
    __tablename__ = "order_planning_calculations"

    id = Column(Integer, primary_key=True, autoincrement=True)

    product_id = Column(
        Integer,
        ForeignKey("products.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )

    period_from = Column(Date, nullable=False, index=True)
    period_to = Column(Date, nullable=False, index=True)

    avg_sales_month = Column(Numeric, nullable=False, default=0)

    product = relationship("Product", back_populates="order_planning_calculations")

    __table_args__ = (
        UniqueConstraint(
            "product_id",
            "period_from",
            "period_to",
            name="uq_order_planning_calc_product_period",
        ),
        Index("ix_order_planning_calc_period", "period_from", "period_to"),
    )

