from __future__ import annotations

from pathlib import Path

import pythoncom
import win32com.client as win32
from sqlalchemy.orm import Session

from app.db.models import Product, ProductStock, TargetPriceCalculation, TempTargetPriceImport, TempTargetPriceOption
from app.exports.excel_column_format import (
    apply_standard_worksheet_format,
    apply_target_price_calculated_worksheet_format,
    excel_column_letter,
    excel_value_by_header,
)


class TargetPriceExporter:
    def __init__(self, session: Session):
        self.session = session

    def _create_excel_app(self):
        pythoncom.CoInitialize()
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        return excel

    def _save_workbook(
        self,
        file_path: str | Path,
        sheet_name: str,
        headers: list[str],
        rows: list[dict],
        *,
        freeze_cell: str = "D2",
        format_mode: str = "standard",
    ) -> Path:
        file_path = Path(file_path)
        if file_path.suffix.lower() != ".xlsx":
            file_path = file_path.with_suffix(".xlsx")
        file_path.parent.mkdir(parents=True, exist_ok=True)

        excel = None
        wb = None
        try:
            target_path = file_path.resolve()
            if target_path.exists():
                try:
                    target_path.unlink()
                except PermissionError:
                    raise PermissionError(
                        f"Не удается перезаписать файл:\n{target_path}\n\n"
                        f"Скорее всего, он открыт в Excel. Закрой файл и попробуй снова."
                    )

            excel = self._create_excel_app()
            wb = excel.Workbooks.Add()
            ws = wb.Worksheets(1)
            ws.Name = sheet_name

            for col_index, header in enumerate(headers, start=1):
                ws.Cells(1, col_index).Value = header

            for row_num, row in enumerate(rows, start=2):
                for col_index, header in enumerate(headers, start=1):
                    ws.Cells(row_num, col_index).Value = excel_value_by_header(header, row.get(header))

            if format_mode == "target_price_calculated":
                apply_target_price_calculated_worksheet_format(ws, headers, freeze_cell=freeze_cell)
            else:
                apply_standard_worksheet_format(ws, headers, freeze_cell=freeze_cell)

            # Ensure the filter range includes the full written table, not only the header row.
            if headers:
                last_col = excel_column_letter(len(headers))
                last_row = max(len(rows) + 1, 1)
                ws.Range(f"A1:{last_col}{last_row}").AutoFilter(1)

            wb.SaveAs(str(target_path))
            return target_path
        finally:
            try:
                if wb is not None:
                    wb.Close(SaveChanges=False)
            except Exception:
                pass
            try:
                if excel is not None:
                    excel.Quit()
            except Exception:
                pass
            pythoncom.CoUninitialize()

    def export_template(self, file_path: str | Path) -> Path:
        headers = ["Material number", "Material"]
        return self._save_workbook(file_path, "Sheet1", headers, [], freeze_cell="A2")

    def _collect_calculated_rows(self, batch_id: str, imported_by: str) -> tuple[list[dict], int]:
        rows = self.session.query(TempTargetPriceImport).filter(
            TempTargetPriceImport.batch_id == batch_id,
            TempTargetPriceImport.imported_by == imported_by,
        ).order_by(TempTargetPriceImport.import_row_no.asc(), TempTargetPriceImport.id.asc()).all()
        out: list[dict] = []
        max_opt = 0
        for row in rows:
            product = self.session.query(Product).filter(Product.id == row.selected_product_id).first() if row.selected_product_id else None
            options = self.session.query(TempTargetPriceOption).filter(
                TempTargetPriceOption.batch_id == batch_id,
                TempTargetPriceOption.imported_by == imported_by,
                TempTargetPriceOption.temp_import_id == row.id,
            ).order_by(TempTargetPriceOption.opt_rank.asc(), TempTargetPriceOption.cost_novo_wvat.asc(), TempTargetPriceOption.id.asc()).all()
            max_opt = max(max_opt, len(options))
            stock = self.session.query(ProductStock).filter(ProductStock.product_id == row.selected_product_id).first() if row.selected_product_id else None
            data = {
                "Supplier Article": row.supplier_article,
                "Supplier Product Name": row.product_name,
                "Our Product Name": product.name if product else "",
                "Дистр цена": stock.distr_price if stock else None,
                "Промо цена": stock.promo_price if stock else None,
                "curr LPC": stock.lpc if stock else None,
                "curr Landed cost": stock.landed_cost if stock else None,
            }
            for i, opt in enumerate(options, start=1):
                data[f"Cost Novo with VAT_{i}"] = opt.cost_novo_wvat
                data[f"Full Cost Msk_{i}"] = opt.full_cost_msk
                data[f"Supplier_{i}"] = opt.supplier_name
                data[f"last update_{i}"] = opt.price_date_used
                data[f"Currency_{i}"] = opt.currency_code
            out.append(data)
        for data in out:
            for i in range(1, max_opt + 1):
                data.setdefault(f"Cost Novo with VAT_{i}", None)
                data.setdefault(f"Full Cost Msk_{i}", None)
                data.setdefault(f"Supplier_{i}", None)
                data.setdefault(f"last update_{i}", None)
                data.setdefault(f"Currency_{i}", None)
        return out, max_opt

    def export_calculated(self, batch_id: str, imported_by: str, file_path: str | Path) -> Path:
        rows, max_opt = self._collect_calculated_rows(batch_id, imported_by)
        base_headers = [
            "Supplier Article", "Supplier Product Name", "Our Product Name",
            "Дистр цена", "Промо цена", "curr LPC", "curr Landed cost",
        ]
        dynamic_headers: list[str] = []
        for i in range(1, max_opt + 1):
            dynamic_headers.extend([f"Cost Novo with VAT_{i}", f"Full Cost Msk_{i}", f"Supplier_{i}", f"last update_{i}", f"Currency_{i}"])
        headers = base_headers + dynamic_headers
        return self._save_workbook(
            file_path,
            "Calculated",
            headers,
            rows,
            freeze_cell="D2",
            format_mode="target_price_calculated",
        )

    def _collect_final_rows(self, batch_id: str, imported_by: str) -> list[dict]:
        rows = self.session.query(TargetPriceCalculation).filter(
            TargetPriceCalculation.batch_id == batch_id,
            TargetPriceCalculation.imported_by == imported_by,
        ).order_by(TargetPriceCalculation.import_row_no.asc(), TargetPriceCalculation.id.asc()).all()
        out: list[dict] = []
        for row in rows:
            product = self.session.query(Product).filter(Product.id == row.product_id).first()
            donor = row.donor_supplier.name if row.donor_supplier else ""
            out.append({
                "Supplier Article": row.supplier_article,
                "Supplier Product Name": row.supplier_product_name,
                "Our Product Name": product.name if product else "",
                "Pack": product.pack if product else None,
                "Target Price, L": row.target_price_l,
                "Target Price, pack": row.target_price_pack,
                "Currency": row.currency_code,
                "FX rate": row.fx_rate_used,
                "fin.Supplier for calc": donor,
            })
        return out

    def export_final(self, batch_id: str, imported_by: str, file_path: str | Path) -> Path:
        headers = [
            "Supplier Article", "Supplier Product Name", "Our Product Name", "Pack",
            "Target Price, L", "Target Price, pack", "Currency", "FX rate", "fin.Supplier for calc",
        ]
        rows = self._collect_final_rows(batch_id, imported_by)
        return self._save_workbook(file_path, "Target price", headers, rows, freeze_cell="D2")
