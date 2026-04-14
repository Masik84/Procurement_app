from pathlib import Path
import pandas as pd

from app.utils.parsers import parse_loose_number
from app.utils.text import clean_multi_spaces, normalize_customer_product_name


class CustomerCostImporter:
    sheet_name = "Запрос себестоимости"

    def read_excel(self, file_path: str | Path) -> list[dict]:
        file_path = Path(file_path)

        if not file_path.exists():
            raise FileNotFoundError(f"Файл не найден: {file_path}")

        try:
            df = pd.read_excel(file_path, sheet_name=self.sheet_name, header=0)
        except Exception:
            df = pd.read_excel(file_path, sheet_name=0, header=0)

        if df.shape[1] < 11:
            raise ValueError("Файл customer cost должен содержать минимум 11 колонок.")

        df = df.iloc[:, :11].copy()
        df.columns = [
            "RequestDate",
            "ManagerName",
            "CustomerName",
            "SupplierArticle",
            "ProductName",
            "Pack",
            "QtyPcs",
            "VolumeL",
            "PurchaseType",
            "PaymentTerms",
            "Comments",
        ]

        for col in ["ManagerName", "CustomerName", "SupplierArticle", "ProductName", "PurchaseType", "PaymentTerms", "Comments"]:
            df[col] = df[col].apply(clean_multi_spaces)

        df["ProductName"] = df["ProductName"].apply(lambda x: normalize_customer_product_name(x) or clean_multi_spaces(x))
        df["Pack"] = df["Pack"].apply(parse_loose_number)
        df["QtyPcs"] = df["QtyPcs"].apply(parse_loose_number)
        df["VolumeL"] = df["VolumeL"].apply(parse_loose_number)
        df["RequestDate"] = pd.to_datetime(df["RequestDate"], errors="coerce")

        df = df[df["ProductName"] != ""].copy()
        df = df.reset_index(drop=True)
        df = df.where(pd.notna(df), None)
        df["import_row_no"] = df.index + 2

        return df.to_dict(orient="records")