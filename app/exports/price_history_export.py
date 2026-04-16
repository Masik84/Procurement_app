from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font


class PriceHistoryExport:
    def export_rows(self, rows, file_path):
        path = Path(file_path)
        if path.suffix.lower() != ".xlsx":
            path = path.with_suffix(".xlsx")

        wb = Workbook()
        ws = wb.active
        ws.title = "Price history"

        headers = ["Product name", "Supplier name", "Price date", "Price", "Currency"]
        ws.append(headers)

        for cell in ws[1]:
            cell.font = Font(bold=True)

        for row in rows:
            price_date = row.get("price_date")
            if price_date is not None:
                price_date = price_date.strftime("%d.%m.%Y")

            ws.append([
                row.get("product_name", ""),
                row.get("supplier_name", ""),
                price_date or "",
                row.get("price", ""),
                row.get("currency", ""),
            ])

        for column in ws.columns:
            max_len = 0
            col_letter = column[0].column_letter
            for cell in column:
                value = "" if cell.value is None else str(cell.value)
                max_len = max(max_len, len(value))
            ws.column_dimensions[col_letter].width = max(14, min(max_len + 2, 40))

        wb.save(path)
        return path