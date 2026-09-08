# IS legacy archive

Archived from Procurement_app master commit:

`dbea261766f07314616ed139b7ba111c49bcf346` (2026-09-08)

The dedicated IS importer is copied here so the old flow can be restored later.

Current application behaviour after this update:
- the database schema and existing IS values are intentionally retained;
- the separate IS / Stock IS UI mode is disabled and removed from the active interface;
- CORAL `order` + `confirmed` quantities are merged into ordinary `Purchase Order`;
- active order-planning calculations use `Stock + Transit + Purchase Order`;
- `Order IS` and `Stock IS` are blocked from all active Excel exports;
- old IS database values are not zeroed during supplier-order updates.

Legacy implementation still exists in the source history of the commit above, including:
- `app/imports/is_importer.py`
- IS branches in `app/services/product_stock_service.py`
- IS branches in `app/page_functions/product_stock_page.py`
- historical DB model columns and Alembic migrations

The current runtime layer deliberately leaves those DB structures untouched so rollback remains possible.
