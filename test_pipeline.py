from app.db import SessionLocal
from app.services.supplier_price_import_service import run_supplier_price_import

def main():
    db = SessionLocal()

    try:
        print("=== START TEST ===")

        run_supplier_price_import(
            db=db,
            file_path="data/test_supplier_price.xlsx",
            supplier_id=1,
            currency="EUR"
        )

        print("=== DONE ===")

    finally:
        db.close()


if __name__ == "__main__":
    main()