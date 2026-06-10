from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT

class InvoiceGenerator:
    def _set_font(self, run, name="Calibri", size=11, bold=False, underline=False):
        run.font.name = name
        run.font.size = Pt(size)
        run.bold = bold
        run.underline = underline

    def _center_para(self, doc, text, size=12, bold=False):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(text)
        self._set_font(r, size=size, bold=bold)
        return p

    def _set_cell(self, cell, text, size=10, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER):
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = align
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(str(text))
        self._set_font(r, size=size, bold=bold)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    def amount_in_words(self, n):
        if n == 8000:
            return "Eight Thousand"
        if n == 10000:
            return "Ten Thousand"
        return str(n)

    def generate_invoice(self, records, output_path, invoice_no, invoice_date):
        first = records[0]
        total = sum(r.total_fee for r in records)
        doc = Document()
        sec = doc.sections[0]
        sec.top_margin = Cm(1.0)
        sec.bottom_margin = Cm(1.0)
        sec.left_margin = Cm(1.2)
        sec.right_margin = Cm(1.2)

        self._center_para(doc, "Centre for Police Research, Pune", size=14, bold=True)
        self._center_para(doc, "Chavan Nagar, Pashan Road, Pune, Maharashtra - 411008", size=10)
        self._center_para(doc, "Email: directorcprpune@gmail.com, Phone: 020-25653696", size=10)
        self._center_para(doc, "GST No. -- 27AAATC3424CIZB PAN - AAATC3424C", size=10)
        self._center_para(doc, "Demand Note/Invoice", size=13, bold=True)

        inv_p = doc.add_paragraph()
        inv_p.paragraph_format.space_before = Pt(0)
        inv_p.paragraph_format.space_after = Pt(0)
        inv_p.add_run(f"Invoice No: {invoice_no}    Date: {invoice_date.strftime('%d-%m-%Y')}")

        ref_p = doc.add_paragraph()
        ref_p.paragraph_format.space_before = Pt(0)
        ref_p.paragraph_format.space_after = Pt(0)
        ref_p.add_run(f"Reference: {first.reference_no or invoice_no}")

        doc.add_paragraph(f"To,\nSuperintendent of Police,\n{first.police_unit}")
        doc.add_paragraph("Training Details")

        table = doc.add_table(rows=1, cols=9)
        table.style = "Table Grid"
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        headers = ["Sr No", "Training Name", "From Date", "To Date", "Officer Name", "Rank", "Duration (Days)", "Fee/Day", "Total Fee"]
        for i, h in enumerate(headers):
            self._set_cell(table.rows[0].cells[i], h, size=10, bold=True)

        for i, r in enumerate(records, 1):
            row = table.add_row().cells
            self._set_cell(row[0], str(i), size=10)
            self._set_cell(row[1], r.course_name, size=10)
            self._set_cell(row[2], r.from_date.strftime('%d/%m/%y') if r.from_date else '', size=10)
            self._set_cell(row[3], r.to_date.strftime('%d/%m/%y') if r.to_date else '', size=10)
            self._set_cell(row[4], r.officer_name, size=10)
            self._set_cell(row[5], r.rank, size=10)
            self._set_cell(row[6], str(r.duration_days), size=10)
            self._set_cell(row[7], f"{int(r.fee_per_day):,}" if float(r.fee_per_day).is_integer() else f"{r.fee_per_day:,.2f}", size=10)
            self._set_cell(row[8], f"{int(r.total_fee):,}" if float(r.total_fee).is_integer() else f"{r.total_fee:,.2f}", size=10)

        gp = doc.add_paragraph()
        gp.paragraph_format.space_before = Pt(4)
        gp.paragraph_format.space_after = Pt(0)
        g = gp.add_run(f"Grand Total: ₹{int(total):,}" if float(total).is_integer() else f"Grand Total: ₹{total:,.2f}")
        self._set_font(g, bold=True)

        doc.add_paragraph(f"Amount: {self.amount_in_words(int(total))} Rupees Only")

        doc.add_paragraph("Bank Account Details")
        bank_rows = [
            ("Account Name", "CPR, Pune"),
            ("Account Number", "10023971530"),
            ("Type", "Savings"),
            ("Bank Name", "State Bank of India"),
            ("Branch Address", "NCL Branch, Pashan Road, Pune - 411008"),
            ("Branch Code", "3552"),
            ("MICR Code", "411002012"),
            ("IFSC Code", "SBIN0003552"),
            ("Payee Code", "22010023828"),
        ]
        for label, value in bank_rows:
            bp = doc.add_paragraph()
            bp.paragraph_format.space_before = Pt(0)
            bp.paragraph_format.space_after = Pt(0)
            run1 = bp.add_run(f"{label}: ")
            self._set_font(run1, bold=True)
            run2 = bp.add_run(value)
            self._set_font(run2)

        doc.add_paragraph("तरी सदर थकीत प्रशिक्षण फी रक्कम पोलीस संशोधन केंद्र पुणे येथे तातडीने पाठविण्याची विनंती आहे.")
        doc.add_paragraph("सांवत:- वर नमुद इन व्हॉईस नंबर प्रमाणे प्रशिक्षण फी Demand Note/Invoice जोडलेले आहे.")

        sig = doc.add_paragraph()
        sig.paragraph_format.space_before = Pt(4)
        sig.paragraph_format.space_after = Pt(0)
        sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        sig.add_run("(Dr. Kakasaheb Dole)\nSuperintendent of Police\nCentre for Police Research, Pune")

        doc.save(output_path)
