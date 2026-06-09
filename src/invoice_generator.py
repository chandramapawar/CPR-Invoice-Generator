from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT


class InvoiceGenerator:
    def fmt_money(self, x):
        return f"{int(x):,}" if float(x).is_integer() else f"{x:,.2f}"

    def _set_font(self, run, name="Calibri", size=11, bold=False, underline=False):
        run.font.name = name
        run.font.size = Pt(size)
        run.bold = bold
        run.underline = underline

    def _set_para(self, para, text, size=11, bold=False, align=None):
        para.clear()
        if align is not None:
            para.alignment = align
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(0)
        run = para.add_run(str(text))
        self._set_font(run, size=size, bold=bold)
        return para

    def _add_center_para(self, doc, text, size=12, bold=False):
        p = doc.add_paragraph()
        return self._set_para(p, text, size=size, bold=bold, align=WD_ALIGN_PARAGRAPH.CENTER)

    def _set_cell(self, cell, text, size=10, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER):
        cell.text = ""
        p = cell.paragraphs[0]
        self._set_para(p, text, size=size, bold=bold, align=align)
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

        self._add_center_para(doc, "Centre for Police Research, Pune", size=14, bold=True)
        self._add_center_para(doc, "Chavan Nagar, Pashan Road, Pune, Maharashtra - 411008", size=10)
        self._add_center_para(doc, "Email: directorcprpune@gmail.com, Phone: 020-25653696", size=10)
        self._add_center_para(doc, "GST No. -- 27AAATC3424CIZB PAN - AAATC3424C", size=10)
        self._add_center_para(doc, "Demand Note/Invoice", size=13, bold=True)

        meta = doc.add_table(rows=1, cols=2)
        meta.alignment = WD_TABLE_ALIGNMENT.CENTER
        meta.autofit = True
        meta.style = "Table Grid"
        self._set_cell(meta.rows[0].cells[0], f"Invoice No: {invoice_no}   Date: {invoice_date.strftime('%d-%m-%Y')}", size=10, align=WD_ALIGN_PARAGRAPH.LEFT)
        self._set_cell(meta.rows[0].cells[1], f"Reference: {first.reference_no or invoice_no}", size=10, align=WD_ALIGN_PARAGRAPH.RIGHT)

        to_p = doc.add_paragraph()
        to_p.paragraph_format.space_before = Pt(2)
        to_p.paragraph_format.space_after = Pt(0)
        self._set_para(to_p, f"To,\nSuperintendent of Police,\n{first.police_unit}", size=11, align=WD_ALIGN_PARAGRAPH.LEFT)

        doc.add_paragraph("")
        self._add_center_para(doc, "Training Details", size=11, bold=True)

        table = doc.add_table(rows=1, cols=9)
        table.style = "Table Grid"
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        headers = ["Sr No", "Training Name", "From Date", "To Date", "Officer Name", "Rank", "Duration (Days)", "Fee/Day", "Total Fee"]
        for i, h in enumerate(headers):
            self._set_cell(table.rows[0].cells[i], h, size=10, bold=True)

        for i, r in enumerate(records, 1):
            row = table.add_row().cells
            self._set_cell(row[0], str(i), size=10)
            self._set_cell(row[1], r.course_name, size=10, align=WD_ALIGN_PARAGRAPH.LEFT)
            self._set_cell(row[2], r.from_date.strftime('%d/%m/%y') if r.from_date else '', size=10)
            self._set_cell(row[3], r.to_date.strftime('%d/%m/%y') if r.to_date else '', size=10)
            self._set_cell(row[4], r.officer_name, size=10, align=WD_ALIGN_PARAGRAPH.LEFT)
            self._set_cell(row[5], r.rank, size=10)
            self._set_cell(row[6], str(r.duration_days), size=10)
            self._set_cell(row[7], self.fmt_money(r.fee_per_day), size=10, align=WD_ALIGN_PARAGRAPH.RIGHT)
            self._set_cell(row[8], self.fmt_money(r.total_fee), size=10, align=WD_ALIGN_PARAGRAPH.RIGHT)

        gp = doc.add_paragraph()
        gp.paragraph_format.space_before = Pt(4)
        gp.paragraph_format.space_after = Pt(0)
        g = gp.add_run(f"Grand Total: ₹{self.fmt_money(total)}")
        self._set_font(g, bold=True)

        amt_p = doc.add_paragraph()
        amt_p.paragraph_format.space_before = Pt(0)
        amt_p.paragraph_format.space_after = Pt(0)
        self._set_para(amt_p, f"Amount: {self.amount_in_words(int(total))} Rupees Only", size=11, align=WD_ALIGN_PARAGRAPH.LEFT)

        self._add_center_para(doc, "Bank Account Details", size=11, bold=True)

        bank = doc.add_table(rows=1, cols=2)
        bank.style = "Table Grid"
        bank.alignment = WD_TABLE_ALIGNMENT.CENTER
        self._set_cell(bank.rows[0].cells[0], "Field", size=10, bold=True)
        self._set_cell(bank.rows[0].cells[1], "Value", size=10, bold=True)

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

        for field, val in bank_rows:
            row = bank.add_row().cells
            self._set_cell(row[0], field, size=10, align=WD_ALIGN_PARAGRAPH.LEFT)
            self._set_cell(row[1], val, size=10, align=WD_ALIGN_PARAGRAPH.LEFT)

        sig = doc.add_paragraph()
        sig.paragraph_format.space_before = Pt(4)
        sig.paragraph_format.space_after = Pt(0)
        self._set_para(sig, "(Dr. Kakasaheb Dole)\nSuperintendent of Police\nCentre for Police Research, Pune", size=11, align=WD_ALIGN_PARAGRAPH.RIGHT)

        doc.save(output_path)
