from app.db.db import SessionLocal
from app.services.supplier_price_run import SupplierPricePipeline
from app.services.supplier import SupplierService


def main():
    session = SessionLocal()

    try:
        supplier_id = 52  # сюда ставишь id нужного поставщика

        supplier_service = SupplierService(session)
        supplier_data = supplier_service.load_supplier_snapshot(supplier_id)

        pipeline = SupplierPricePipeline(session)

        result = pipeline.run_from_excel(
            file_path="data/test_supplier_price.xlsx",
            imported_by="test_user",
            supplier_id=supplier_id,
            supplier_data=supplier_data,
            import_date=None,
            save_exchange_rate=False,
            explicit_fx_rate=None,   # важно: курс возьмется из БД
        )

        session.commit()
        print(result)

    except Exception as e:
        session.rollback()
        print("ERROR:", e)
    finally:
        session.close()


if __name__ == "__main__":
    main()