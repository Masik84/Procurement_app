from __future__ import annotations

"""Small compatibility fixes for runtime adapters.

These fixes intentionally live outside the page code.  Procurement installs a
few runtime adapters from :mod:`app.no_is_runtime` and shared UI helpers during
package startup.  Keeping their compatibility adjustments here avoids changing
business calculations or database schema merely to satisfy an outdated adapter
assumption.
"""

from typing import Any


_INSTALLED = False


def _install_product_stock_primary_key_fix() -> None:
    """Make the no-IS Supplier Orders adapter use ProductStock.product_id.

    ``ProductStock`` is a one-to-one table whose primary key is ``product_id``;
    it has never exposed a generic ``id`` column.  The old runtime adapter used
    ``row.id``/``ProductStock.id`` while preserving historical IS fields, which
    makes Supplier Orders Save fail before the real service can finish.
    """

    from app import no_is_runtime

    current_patch = no_is_runtime._patch_product_stock_service
    if getattr(current_patch, "_product_stock_pk_fix", False):
        return

    def patch_product_stock_service(service_module: Any) -> None:
        service_cls = getattr(service_module, "ProductStockService", None)
        product_stock_cls = getattr(service_module, "ProductStock", None)
        if service_cls is None or product_stock_cls is None:
            return

        no_is_runtime._patch_supplier_orders_importer(service_module)

        original = service_cls.save_supplier_orders_to_product_stock
        if getattr(original, "_no_is_patch", False):
            return

        def save_supplier_orders_preserve_legacy_is(
            self,
            batch_id: str,
            imported_by: str,
        ) -> int:
            snapshot = {
                int(row.product_id): (
                    row.is_order_qty,
                    row.is_confirmed_order_qty,
                    row.is_stock_qty,
                    row.is_update_date,
                )
                for row in self.session.query(product_stock_cls).all()
                if row.product_id is not None
            }

            result = original(self, batch_id, imported_by)

            if snapshot:
                # The legacy service performs a bulk UPDATE with
                # synchronize_session=False. Restore the retained historical IS
                # values explicitly by the real ProductStock primary key.
                for product_id, old in snapshot.items():
                    (
                        self.session.query(product_stock_cls)
                        .filter(product_stock_cls.product_id == product_id)
                        .update(
                            {
                                product_stock_cls.is_order_qty: old[0],
                                product_stock_cls.is_confirmed_order_qty: old[1],
                                product_stock_cls.is_stock_qty: old[2],
                                product_stock_cls.is_update_date: old[3],
                            },
                            synchronize_session=False,
                        )
                    )
                self.session.flush()
                self.session.expire_all()

            return result

        save_supplier_orders_preserve_legacy_is._no_is_patch = True
        save_supplier_orders_preserve_legacy_is._product_stock_pk_fix = True
        service_cls.save_supplier_orders_to_product_stock = (
            save_supplier_orders_preserve_legacy_is
        )

    patch_product_stock_service._product_stock_pk_fix = True
    no_is_runtime._patch_product_stock_service = patch_product_stock_service


def _install_order_planning_filter_disconnect_fix() -> None:
    """Do not disconnect Order Planning signals that were never connected.

    OrderPlanningPage does not connect the top Brand/Product controls directly
    to ``refresh_current_product_combo``.  The shared selector used to call
    ``Signal.disconnect`` anyway, and PySide6 reports that as a RuntimeWarning.
    Other pages retain the original disconnect behaviour because they may have
    legacy direct connections that genuinely need to be removed.
    """

    from app.utils.product_selection_filter import ProductSelectionFilter

    original = ProductSelectionFilter._remember_and_disconnect_legacy_signals
    if getattr(original, "_order_planning_disconnect_fix", False):
        return

    def remember_and_disconnect_legacy_signals(self) -> None:
        if self.page.__class__.__name__ == "OrderPlanningPage":
            slot = getattr(self.page, "refresh_current_product_combo", None)
            self._legacy_filter_slot = slot
            self._refresh_callback = self._refresh_order_planning_combo
            return
        return original(self)

    remember_and_disconnect_legacy_signals._order_planning_disconnect_fix = True
    ProductSelectionFilter._remember_and_disconnect_legacy_signals = (
        remember_and_disconnect_legacy_signals
    )


def install_runtime_consistency_fixes() -> None:
    global _INSTALLED
    if _INSTALLED:
        return
    _INSTALLED = True

    _install_product_stock_primary_key_fix()
    _install_order_planning_filter_disconnect_fix()
