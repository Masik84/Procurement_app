import uuid
from sqlalchemy.orm import Session

from app.imports.customer_cost_importer import CustomerCostImporter
from app.services.customer_cost_import import CustomerCostImportService


class CustomerCostImportRun:
    def __init__(self, session: Session):
        self.session = session
        self.importer = CustomerCostImporter()
        self.service = CustomerCostImportService(session)

    def import_from_excel(self, file_path: str, imported_by: str) -> dict:
        batch_id = str(uuid.uuid4())
        rows = self.importer.read_excel(file_path)
        imported_count = self.service.import_rows(rows=rows, batch_id=batch_id, imported_by=imported_by)
        matched_count = self.service.automatch_temp_rows(batch_id=batch_id, imported_by=imported_by)

        return {"batch_id": batch_id, "imported_count": imported_count, "matched_count": matched_count, "file": file_path}

    def calculate(self, batch_id: str, imported_by: str) -> dict:
        return self.service.run_calculation(batch_id=batch_id, imported_by=imported_by)

    def save_history(self, batch_id: str, imported_by: str) -> dict:
        saved_count = self.service.save_calculations(batch_id=batch_id, imported_by=imported_by)
        return {"batch_id": batch_id, "saved_count": saved_count}