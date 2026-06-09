#service
from collections import defaultdict
from pathlib import Path

from src.document_manager import DocumentManager
from src.excel_reader import ExcelReader
from src.invoice_generator import InvoiceGenerator
from src.letter_generator import LetterGenerator

class GenerationService:
    def __init__(self, base_output_dir: Path | None = None):
        self.document_manager = DocumentManager(base_output_dir=base_output_dir) if base_output_dir else DocumentManager()
        self.letter_generator = LetterGenerator()
        self.invoice_generator = InvoiceGenerator()

    def group_records_by_unit(self, records):
        grouped = defaultdict(list)
        for record in records:
            grouped[record.police_unit].append(record)
        return grouped

    def generate_from_file(self, file_obj, invoice_no, course_title, training_from, training_to, invoice_date, letter_date):
        records = ExcelReader(file_obj).read_training_records()
        if not records:
            raise ValueError("No valid training records found in the uploaded file.")

        for r in records:
            if course_title:
                r.course_name = course_title
            if training_from:
                r.from_date = training_from
            if training_to:
                r.to_date = training_to
            if not r.duration_days and r.from_date and r.to_date:
                r.duration_days = (r.to_date - r.from_date).days + 1
            r.total_fee = r.duration_days * r.fee_per_day

        grouped_records = self.group_records_by_unit(records)
        generated_files = []
        for police_unit, unit_records in grouped_records.items():
            first = unit_records[0]
            reference_no = first.reference_no or f"CPR-REF-{invoice_no}"
            letter_path = self.document_manager.get_letter_path(police_unit, reference_no)
            invoice_path = self.document_manager.get_invoice_path(police_unit, invoice_no)
            self.letter_generator.generate_letter(unit_records, letter_path, reference_no, letter_date)
            self.invoice_generator.generate_invoice(unit_records, invoice_path, invoice_no, invoice_date)
            generated_files.extend([letter_path, invoice_path])
        return generated_files
