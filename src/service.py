from collections import defaultdict
from pathlib import Path

from src.document_manager import DocumentManager
from src.excel_reader import ExcelReader
from src.invoice_generator import InvoiceGenerator
from src.letter_generator import LetterGenerator


class GenerationService:
    def __init__(self, base_output_dir: Path | None = None):
        self.document_manager = (
            DocumentManager(base_output_dir=base_output_dir)
            if base_output_dir
            else DocumentManager()
        )
        self.letter_generator = LetterGenerator()
        self.invoice_generator = InvoiceGenerator()

    def group_records_by_unit(self, records):
        grouped = defaultdict(list)
        for record in records:
            grouped[record.police_unit].append(record)
        return grouped

    def _split_invoice_no(self, invoice_no: str):
        """
        Split invoice_no into (numeric_prefix, suffix).
        Example: '58/2026' -> (58, '/2026')
                 '123'     -> (123, '')
        If prefix is not numeric, return (None, original_string_as_suffix).
        """
        inv_str = str(invoice_no)
        parts = inv_str.split("/", 1)
        try:
            num = int(parts[0].strip())
            suffix = f"/{parts[1]}" if len(parts) > 1 else ""
            return num, suffix
        except ValueError:
            # Non-numeric prefix: no incrementing possible
            return None, inv_str

    def generate_from_file(
        self,
        file_obj,
        invoice_no,
        course_title,
        training_from,
        training_to,
        invoice_date,
        letter_date,
    ):
        records = ExcelReader(file_obj).read_training_records()
        if not records:
            raise ValueError("No valid training records found in the uploaded file.")

        # Normalize and compute totals
        for r in records:
            if course_title:
                r.course_name = course_title
            if training_from:
                r.from_date = training_from
            if training_to:
                r.to_date = training_to
            if r.from_date and r.to_date:
                r.duration_days = (r.to_date - r.from_date).days + 1
                r.total_fee = r.duration_days * r.fee_per_day

        grouped_records = self.group_records_by_unit(records)
        generated_files = []

        # Prepare unit-wise invoice numbering
        current_num, suffix = self._split_invoice_no(invoice_no)

        for police_unit, unit_records in grouped_records.items():
            # Build this unit's invoice number
            if current_num is not None:
                unit_invoice_no = f"{current_num}{suffix}"
            else:
                # Fallback: use original invoice_no unchanged
                unit_invoice_no = str(invoice_no)

            # Same number in invoice header AND letter's invoice column
            reference_no = unit_invoice_no

            letter_path = self.document_manager.get_letter_path(
                police_unit, reference_no
            )
            invoice_path = self.document_manager.get_invoice_path(
                police_unit, unit_invoice_no
            )

            self.letter_generator.generate_letter(
                unit_records, letter_path, reference_no, letter_date
            )
            self.invoice_generator.generate_invoice(
                unit_records, invoice_path, unit_invoice_no, invoice_date
            )

            generated_files.extend([letter_path, invoice_path])

            # Increment for next unit if numeric prefix exists
            if current_num is not None:
                current_num += 1

        return generated_files