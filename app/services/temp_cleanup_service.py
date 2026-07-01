from __future__ import annotations

from datetime import date, datetime
from typing import Iterable

from sqlalchemy.orm import Session

from app.db.models import (
    TempCustomerCostImport,
    TempCustomerCostOption,
    TempIsImport,
    TempPriceImport,
    TempProductSearchImport,
    TempStockImport,
    TempSupplierOrdersImport,
    TempTargetPriceImport,
    TempTargetPriceOption,
)


class TempCleanupService:
    """Centralized cleanup for all temp/staging tables.

    Rules:
    - old data (before today by default) is removed globally for all users;
    - after successful save, current user temp data can be removed by imported_by;
    - option tables are deleted before their parent import tables.
    """

    IMPORT_TABLES = (
        TempPriceImport,
        TempCustomerCostImport,
        TempStockImport,
        TempSupplierOrdersImport,
        TempIsImport,
        TempProductSearchImport,
        TempTargetPriceImport,
    )
    OPTION_TABLES = (
        TempCustomerCostOption,
        TempTargetPriceOption,
    )
    # Children first, then parents.
    ALL_TABLES = OPTION_TABLES + IMPORT_TABLES

    def __init__(self, session: Session) -> None:
        self.session = session

    @staticmethod
    def _cutoff_dt(before_date: date | None = None) -> datetime:
        cutoff = before_date or date.today()
        return datetime.combine(cutoff, datetime.min.time())

    def delete_current_batch(self, batch_id: str | None, imported_by: str | None) -> int:
        """Delete only rows of the current user's current batch from every temp table."""
        if not batch_id or not imported_by:
            return 0

        total = 0
        for model in self.ALL_TABLES:
            total += int(
                self.session.query(model)
                .filter(model.batch_id == batch_id, model.imported_by == imported_by)
                .delete(synchronize_session=False)
                or 0
            )
        self.session.flush()
        return total


    def delete_current_user(self, imported_by: str | None, tables: Iterable[type] | None = None) -> int:
        """Delete temp rows of one user.

        Used after successful write to main tables. It never touches other users,
        but it can remove all batches of the current user in selected temp tables.
        """
        if not imported_by:
            return 0

        selected_tables = set(tables or self.ALL_TABLES)
        ordered_tables = [model for model in self.ALL_TABLES if model in selected_tables]

        total = 0
        for model in ordered_tables:
            total += int(
                self.session.query(model)
                .filter(model.imported_by == imported_by)
                .delete(synchronize_session=False)
                or 0
            )

        self.session.flush()
        return total

    def cleanup_old_for_all(self, before_date: date | None = None) -> int:
        """Delete old temp rows for all users.

        This is the safe daily cleanup: only rows older than the cutoff are removed.
        Current-day batches are preserved, so users working today do not affect one
        another.
        """
        cutoff_dt = self._cutoff_dt(before_date)
        total = 0

        for model in self.OPTION_TABLES:
            date_column = getattr(model, "calc_date")
            total += int(
                self.session.query(model)
                .filter(date_column < cutoff_dt)
                .delete(synchronize_session=False)
                or 0
            )

        for model in self.IMPORT_TABLES:
            date_column = getattr(model, "import_date")
            total += int(
                self.session.query(model)
                .filter(date_column < cutoff_dt)
                .delete(synchronize_session=False)
                or 0
            )

        self.session.flush()
        return total

    def cleanup_old_for_user(self, imported_by: str | None, before_date: date | None = None) -> int:
        """Delete old temp rows only for one user, preserving today's batches."""
        if not imported_by:
            return 0

        cutoff_dt = self._cutoff_dt(before_date)
        total = 0

        for model in self.OPTION_TABLES:
            date_column = getattr(model, "calc_date")
            total += int(
                self.session.query(model)
                .filter(model.imported_by == imported_by, date_column < cutoff_dt)
                .delete(synchronize_session=False)
                or 0
            )

        for model in self.IMPORT_TABLES:
            date_column = getattr(model, "import_date")
            total += int(
                self.session.query(model)
                .filter(model.imported_by == imported_by, date_column < cutoff_dt)
                .delete(synchronize_session=False)
                or 0
            )

        self.session.flush()
        return total
