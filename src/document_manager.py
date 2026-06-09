from pathlib import Path

class DocumentManager:
    def __init__(self, base_output_dir: Path | None = None):
        self.base_output_dir = Path(base_output_dir) if base_output_dir else Path("generated_output")
        self.base_output_dir.mkdir(parents=True, exist_ok=True)

    def _clean(self, text):
        return "".join(c if c.isalnum() or c in (" ", "-", "_") else "_" for c in str(text)).strip().replace(" ", "_")

    def get_letter_path(self, police_unit, reference_no):
        d = self.base_output_dir / self._clean(police_unit)
        d.mkdir(parents=True, exist_ok=True)
        return d / f"Letter_{self._clean(reference_no)}.docx"

    def get_invoice_path(self, police_unit, invoice_no):
        d = self.base_output_dir / self._clean(police_unit)
        d.mkdir(parents=True, exist_ok=True)
        return d / f"Invoice_{self._clean(invoice_no)}.docx"#/document_manager
