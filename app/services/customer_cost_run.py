from __future__ import annotations

import uuid
from pathlib import Path
from sqlalchemy.orm import Session

from app.exports.customer_cost_export import CustomerCostExport
from app.imports.customer_cost_importer import CustomerCostImporter
from app.services.customer_cost_import import CustomerCostImportService


class CustomerCostImportRun:
    def __init__(self, session: Session):
        self.session = session
        self.importer = CustomerCostImporter()
        self.service = CustomerCostImportService(session)
        self.exporter = CustomerCostExport(session)

    def import_from_excel(self, file_path: str, imported_by: str) -> dict:
        batch_id = str(uuid.uuid4())
        rows = self.importer.read_excel(file_path)
        imported_count = self.service.import_rows(rows=rows, batch_id=batch_id, imported_by=imported_by)
        matched_count = self.service.automatch_temp_rows(batch_id=batch_id, imported_by=imported_by)
        return {"batch_id": batch_id, "imported_count": imported_count, "matched_count": matched_count, "file": file_path}

    def calculate(self, batch_id: str, imported_by: str, output_file: str | Path | None = None) -> dict:
        stats = self.service.run_calculation(batch_id=batch_id, imported_by=imported_by)
        export_path = None
        if output_file is not None:
            export_path = self.exporter.export_calculated(batch_id=batch_id, imported_by=imported_by, file_path=output_file)
        return {**stats, "export_file": str(export_path) if export_path else None}

    def save(self, batch_id: str, imported_by: str, kam_folder: str | Path | None = None) -> dict:
        saved_count = self.service.save_calculations(batch_id=batch_id, imported_by=imported_by)
        kam_files = []
        if kam_folder is not None:
            kam_files = self.exporter.export_kam_files(batch_id=batch_id, imported_by=imported_by, folder_path=kam_folder)
        return {
            "batch_id": batch_id,
            "saved_count": saved_count,
            "kam_files": [str(path) for path in kam_files],
        }
