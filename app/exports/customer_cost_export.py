from pathlib import Path
import pandas as pd
from sqlalchemy.orm import Session

from app.db.models import TempCustomerCostImport, TempCustomerCostOption


class CustomerCostExport:
    def __init__(self, session: Session):
        self.session = session

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

            base = {
                "Дата": row.request_date,
                "Менеджер": row.manager_name,
                "Клиент": row.customer_name,
                "Код продукта": row.supplier_article,
                "Название продукта": row.product_name,
                "Фасовка": row.pack,
                "Количество": row.qty_pcs,
                "Объем, л": row.volume_l,
                "Тип закупки": row.purchase_type,
                "Условия оплаты": row.payment_terms,
                "Комментарии": row.comments,
            }

            for i, opt in enumerate(options, start=1):
                base[f"CostNovoWVAT_{i}"] = opt.cost_novo_wvat
                base[f"FullCostMsk_{i}"] = opt.full_cost_msk
                base[f"Supplier_{i}"] = opt.supplier_name
                base[f"PriceDate_{i}"] = opt.price_date_used
                base[f"Currency_{i}"] = opt.currency_code

            out_rows.append(base)

        for row in out_rows:
            for i in range(1, max_opt + 1):
                row.setdefault(f"CostNovoWVAT_{i}", None)
                row.setdefault(f"FullCostMsk_{i}", None)
                row.setdefault(f"Supplier_{i}", None)
                row.setdefault(f"PriceDate_{i}", None)
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

        result = []

        for manager_name, customer_name in groups:
            safe_manager = (manager_name or "NoManager").replace("/", "_").replace("\\", "_")
            safe_customer = (customer_name or "NoCustomer").replace("/", "_").replace("\\", "_")
            file_path = folder_path / f"{safe_manager}_{safe_customer}.xlsx"

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
                    "Объем, л": row.volume_l,
                    "Тип закупки": row.purchase_type,
                    "Условия оплаты": row.payment_terms,
                    "Комментарии": row.comments,
                    "Supplier": option.supplier_name if option else None,
                    "SupplierPrice": option.supplier_price if option else None,
                    "CostNovoWVAT": option.cost_novo_wvat if option else None,
                    "FullCostMsk": option.full_cost_msk if option else None,
                    "Currency": option.currency_code if option else None,
                    "PriceDate": option.price_date_used if option else None,
                })

            df = pd.DataFrame(out_rows)

            with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="KAM")

            result.append(file_path)

        return result