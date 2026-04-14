from __future__ import annotations

from pathlib import Path
import pandas as pd
from sqlalchemy.orm import Session

from app.db.models import TempCustomerCostImport, TempCustomerCostOption


class CustomerCostExport:
    def __init__(self, session: Session):
        self.session = session

    @staticmethod
    def _safe_name(value: str) -> str:
        s = (value or "").strip() or "NoName"
        for ch in ['\\', '/', ':', '*', '?', '"', '<', '>', '|']:
            s = s.replace(ch, '_')
        return s

    def export_calculated(self, batch_id: str, imported_by: str, file_path: str | Path) -> Path:
        rows = self.session.query(TempCustomerCostImport).filter(
            TempCustomerCostImport.batch_id == batch_id,
            TempCustomerCostImport.imported_by == imported_by,
        ).order_by(TempCustomerCostImport.import_row_no.asc(), TempCustomerCostImport.id.asc()).all()

        out_rows = []
        max_opt = 0
        for row in rows:
            options = self.session.query(TempCustomerCostOption).filter(
                TempCustomerCostOption.batch_id == batch_id,
                TempCustomerCostOption.imported_by == imported_by,
                TempCustomerCostOption.temp_import_id == row.id,
            ).order_by(
                TempCustomerCostOption.opt_rank.asc(),
                TempCustomerCostOption.full_cost_msk.asc(),
                TempCustomerCostOption.id.asc(),
            ).all()
            max_opt = max(max_opt, len(options))
            data = {
                "Дата": row.request_date,
                "Менеджер": row.manager_name,
                "Клиент": row.customer_name,
                "Код продукта": row.supplier_article,
                "Название продукта": row.product_name,
                "Фасовка": row.pack,
                "Количество": row.qty_pcs,
                "Объем л": row.volume_l,
                "Вид закупки": row.purchase_type,
                "Условия оплаты": row.payment_terms,
                "Комментарии": row.comments,
            }
            for i, opt in enumerate(options, start=1):
                data[f"CostNovoWVAT_{i}"] = opt.cost_novo_wvat
                data[f"FullCostMsk_{i}"] = opt.full_cost_msk
                data[f"Supplier_{i}"] = opt.supplier_name
                data[f"last update_{i}"] = opt.price_date_used
                data[f"Currency_{i}"] = opt.currency_code
            out_rows.append(data)

        for row in out_rows:
            for i in range(1, max_opt + 1):
                row.setdefault(f"CostNovoWVAT_{i}", None)
                row.setdefault(f"FullCostMsk_{i}", None)
                row.setdefault(f"Supplier_{i}", None)
                row.setdefault(f"last update_{i}", None)
                row.setdefault(f"Currency_{i}", None)

        df = pd.DataFrame(out_rows)
        file_path = Path(file_path)
        file_path.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Calculated")
        return file_path

    def export_kam_files(self, batch_id: str, imported_by: str, folder_path: str | Path) -> list[Path]:
        folder_path = Path(folder_path)
        folder_path.mkdir(parents=True, exist_ok=True)

        groups = self.session.query(
            TempCustomerCostImport.manager_name,
            TempCustomerCostImport.customer_name,
        ).filter(
            TempCustomerCostImport.batch_id == batch_id,
            TempCustomerCostImport.imported_by == imported_by,
        ).distinct().all()

        result: list[Path] = []
        for manager_name, customer_name in groups:
            file_path = folder_path / f"{self._safe_name(manager_name)}_{self._safe_name(customer_name)}_KAM.xlsx"
            rows = self.session.query(TempCustomerCostImport).filter(
                TempCustomerCostImport.batch_id == batch_id,
                TempCustomerCostImport.imported_by == imported_by,
                TempCustomerCostImport.manager_name == manager_name,
                TempCustomerCostImport.customer_name == customer_name,
            ).order_by(TempCustomerCostImport.import_row_no.asc(), TempCustomerCostImport.id.asc()).all()

            out_rows = []
            for row in rows:
                option = None
                if row.selected_option_id is not None:
                    option = self.session.query(TempCustomerCostOption).filter(
                        TempCustomerCostOption.id == row.selected_option_id
                    ).first()
                out_rows.append({
                    "Дата": row.request_date,
                    "Менеджер": row.manager_name,
                    "Клиент": row.customer_name,
                    "Код продукта": row.supplier_article,
                    "Название продукта": row.product_name,
                    "Фасовка": row.pack,
                    "Количество": row.qty_pcs,
                    "Объем л": row.volume_l,
                    "Вид закупки": row.purchase_type,
                    "Условия оплаты": row.payment_terms,
                    "Комментарии": row.comments,
                    "Кост руб л с НДС": option.cost_novo_wvat if option else None,
                    "Поставщик": option.supplier_name if option else None,
                    "Валюта": option.currency_code if option else None,
                    "Курс": option.fx_rate_used if option else None,
                })

            df = pd.DataFrame(out_rows)
            with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="KAM")
            result.append(file_path)

        return result
