from __future__ import annotations

from datetime import datetime, date
from pathlib import Path

import pandas as pd
from sqlalchemy.orm import Session

from app.db.models import TempProductSearchImport
from app.services.product_matching_service import ProductMatchingService
from app.utils.batch import generate_import_batch_id


class ProductSearchService:
    def __init__(self, session: Session) -> None:
        self.session = session
        self.product_matching_service = ProductMatchingService(session)

    def start_batch(self) -> str:
        return generate_import_batch_id()

    def delete_temp_rows(self, batch_id: str, imported_by: str) -> int:
        deleted_count = (
            self.session.query(TempProductSearchImport)
            .filter(
                TempProductSearchImport.batch_id == batch_id,
                TempProductSearchImport.imported_by == imported_by,
            )
            .delete(synchronize_session=False)
        )
        self.session.flush()
        return int(deleted_count or 0)

    def reset_batch(self, batch_id: str, imported_by: str) -> None:
        self.delete_temp_rows(batch_id, imported_by)
        self.session.flush()

    def cleanup_old_temp_rows(self, imported_by: str, before_date: date | None = None) -> int:
        cutoff = before_date or date.today()
        deleted_count = (
            self.session.query(TempProductSearchImport)
            .filter(
                TempProductSearchImport.imported_by == imported_by,
                TempProductSearchImport.import_date < datetime.combine(cutoff, datetime.min.time()),
            )
            .delete(synchronize_session=False)
        )
        self.session.flush()
        return int(deleted_count or 0)

    def create_empty_temp_row(
        self,
        *,
        batch_id: str,
        imported_by: str,
        import_date: datetime | None = None,
    ) -> TempProductSearchImport:
        if import_date is None:
            import_date = datetime.now()

        last_row_no = (
            self.session.query(TempProductSearchImport.import_row_no)
            .filter(
                TempProductSearchImport.batch_id == batch_id,
                TempProductSearchImport.imported_by == imported_by,
            )
            .order_by(TempProductSearchImport.import_row_no.desc(), TempProductSearchImport.id.desc())
            .first()
        )
        next_row_no = int(last_row_no[0]) + 1 if last_row_no and last_row_no[0] is not None else 1

        row = TempProductSearchImport(
            source_article=None,
            source_product_name=None,
            batch_id=batch_id,
            imported_by=imported_by,
            import_date=import_date,
            import_row_no=next_row_no,
            selected_product_id=None,
            new_product_name=None,
            new_brand=None,
            new_pack=None,
            new_is_excise=False,
        )
        self.session.add(row)
        self.session.flush()
        return row

    def import_rows_to_temp(
        self,
        *,
        batch_id: str,
        imported_by: str,
        rows: list[dict],
        import_date: datetime | None = None,
        replace_existing_batch_rows: bool = True,
    ) -> int:
        if import_date is None:
            import_date = datetime.now()

        if replace_existing_batch_rows:
            self.delete_temp_rows(batch_id, imported_by)

        created_rows = []
        for row in rows:
            created_rows.append(
                TempProductSearchImport(
                    source_article=row.get("source_article") or None,
                    source_product_name=row.get("source_product_name") or None,
                    batch_id=batch_id,
                    imported_by=imported_by,
                    import_date=import_date,
                    import_row_no=row.get("import_row_no"),
                    selected_product_id=None,
                    new_product_name=None,
                    new_brand=None,
                    new_pack=None,
                    new_is_excise=None,
                )
            )

        if created_rows:
            self.session.add_all(created_rows)

        self.session.flush()
        return len(created_rows)

    def automatch_temp_rows(self, batch_id: str, imported_by: str) -> int:
        rows = (
            self.session.query(TempProductSearchImport)
            .filter(
                TempProductSearchImport.batch_id == batch_id,
                TempProductSearchImport.imported_by == imported_by,
                TempProductSearchImport.selected_product_id.is_(None),
            )
            .order_by(TempProductSearchImport.import_row_no.asc(), TempProductSearchImport.id.asc())
            .all()
        )

        matched_count = 0
        for row in rows:
            product = self.product_matching_service.find_price_import_product(
                supplier_article=row.source_article,
                supplier_product_name=row.source_product_name,
            )
            if product is not None:
                row.selected_product_id = product.id
                matched_count += 1

        self.session.flush()
        return matched_count

    def validate_new_products_before_save(self, batch_id: str, imported_by: str) -> None:
        rows = (
            self.session.query(TempProductSearchImport)
            .filter(
                TempProductSearchImport.batch_id == batch_id,
                TempProductSearchImport.imported_by == imported_by,
                TempProductSearchImport.selected_product_id.is_(None),
                TempProductSearchImport.new_product_name.isnot(None),
            )
            .all()
        )

        for row in rows:
            if row.new_product_name is None or not str(row.new_product_name).strip():
                continue

            self.product_matching_service.validate_new_product_fields(
                product_name=row.new_product_name,
                brand=row.new_brand,
                pack=row.new_pack,
                is_excise=row.new_is_excise,
            )

    def create_products_from_temp(self, batch_id: str, imported_by: str) -> int:
        rows = (
            self.session.query(TempProductSearchImport)
            .filter(
                TempProductSearchImport.batch_id == batch_id,
                TempProductSearchImport.imported_by == imported_by,
                TempProductSearchImport.selected_product_id.is_(None),
                TempProductSearchImport.new_product_name.isnot(None),
                TempProductSearchImport.new_brand.isnot(None),
                TempProductSearchImport.new_pack.isnot(None),
                TempProductSearchImport.new_is_excise.isnot(None),
            )
            .order_by(TempProductSearchImport.import_row_no.asc(), TempProductSearchImport.id.asc())
            .all()
        )

        created_count = 0
        for row in rows:
            product = self.product_matching_service.get_or_create_product(
                name=row.new_product_name,
                brand=row.new_brand,
                pack=row.new_pack,
                is_excise=bool(row.new_is_excise),
            )
            row.selected_product_id = product.id
            created_count += 1

        self.session.flush()
        return created_count

    def create_or_update_product_articles(self, batch_id: str, imported_by: str) -> int:
        rows = (
            self.session.query(TempProductSearchImport)
            .filter(
                TempProductSearchImport.batch_id == batch_id,
                TempProductSearchImport.imported_by == imported_by,
                TempProductSearchImport.selected_product_id.isnot(None),
            )
            .order_by(TempProductSearchImport.import_row_no.asc(), TempProductSearchImport.id.asc())
            .all()
        )

        processed_count = 0
        for row in rows:
            self.product_matching_service.save_product_articles_by_split_articles(
                product_id=row.selected_product_id,
                supplier_article=row.source_article,
                supplier_product_name=row.source_product_name,
            )
            processed_count += 1

        self.session.flush()
        return processed_count

    def build_export_dataframe(self, batch_id: str, imported_by: str) -> pd.DataFrame:
        rows = (
            self.session.query(TempProductSearchImport)
            .filter(
                TempProductSearchImport.batch_id == batch_id,
                TempProductSearchImport.imported_by == imported_by,
            )
            .order_by(TempProductSearchImport.import_row_no.asc(), TempProductSearchImport.id.asc())
            .all()
        )

        out_rows = []
        for row in rows:
            final_product_name = ""
            if row.selected_product is not None:
                final_product_name = row.selected_product.name or ""
            elif row.new_product_name:
                final_product_name = row.new_product_name

            out_rows.append(
                {
                    "Article": row.source_article or "",
                    "Product name": row.source_product_name or "",
                    "Our Product Name": final_product_name,
                }
            )

        return pd.DataFrame(out_rows)

    def export_to_excel(self, batch_id: str, imported_by: str, output_path: str | Path) -> Path:
        output_path = Path(output_path)
        if output_path.suffix.lower() != ".xlsx":
            output_path = output_path.with_suffix(".xlsx")

        df = self.build_export_dataframe(batch_id, imported_by)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Sheet1")
        return output_path
