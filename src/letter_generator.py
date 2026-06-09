from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT


class LetterGenerator:
    def fmt_money(self, x):
        return f"{int(x):,}" if float(x).is_integer() else f"{x:,.2f}"

    def _set_font(self, run, name="Mangal", size=12, bold=False, underline=False):
        run.font.name = name
        run.font.size = Pt(size)
        run.bold = bold
        run.underline = underline

    def _set_para(self, para, text, size=12, bold=False, underline=False, align=None):
        para.clear()
        if align is not None:
            para.alignment = align
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(0)
        run = para.add_run(str(text))
        self._set_font(run, size=size, bold=bold, underline=underline)
        return para

    def _add_center_para(self, doc, text, size=12, bold=False, underline=False):
        p = doc.add_paragraph()
        return self._set_para(p, text, size=size, bold=bold, underline=underline, align=WD_ALIGN_PARAGRAPH.CENTER)

    def _set_cell(self, cell, text, size=10, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER):
        cell.text = ""
        p = cell.paragraphs[0]
        self._set_para(p, text, size=size, bold=bold, align=align)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    def generate_letter(self, records, output_path, reference_no, letter_date):
        first = records[0]
        doc = Document()

        style = doc.styles["Normal"]
        style.font.name = "Mangal"
        style.font.size = Pt(12)

        sec = doc.sections[0]
        sec.top_margin = Cm(1.0)
        sec.bottom_margin = Cm(1.0)
        sec.left_margin = Cm(1.2)
        sec.right_margin = Cm(1.2)

        self._add_center_para(doc, "पोलीस संशोधन केंद्र, पुणे", size=14, bold=True)
        self._add_center_para(doc, "चव्हाण नगर, पाषाण रोड, पुणे ४११००८", size=11)
        self._add_center_para(doc, "TEL NO- 020-25653696 FAX NO- 020-25653696", size=10)
        self._add_center_para(doc, "E-MAIL- directorcprpune@gmail.com WEBSITE- www.cprpune.org", size=10)

        meta = doc.add_table(rows=1, cols=2)
        meta.style = "Table Grid"
        meta.alignment = WD_TABLE_ALIGNMENT.CENTER
        self._set_cell(meta.rows[0].cells[0], f"जा.क्र.सी.पी.आर./प्रलंबित प्रशिक्षण शुल्क/{reference_no}/२०२६ पुणे", size=10, align=WD_ALIGN_PARAGRAPH.LEFT)
        self._set_cell(meta.rows[0].cells[1], f"दि. {letter_date.strftime('%d/%m/%Y')}", size=10, align=WD_ALIGN_PARAGRAPH.RIGHT)

        doc.add_paragraph("")
        self._set_para(doc.add_paragraph(), "प्रति,", size=12, align=WD_ALIGN_PARAGRAPH.LEFT)
        self._set_para(doc.add_paragraph(), "मा. पोलीस अधिक्षक,", size=12, align=WD_ALIGN_PARAGRAPH.LEFT)
        self._set_para(doc.add_paragraph(), first.police_unit or "", size=12, align=WD_ALIGN_PARAGRAPH.LEFT)

        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        r1 = p.add_run("विषय")
        self._set_font(r1, bold=True, underline=True)
        r2 = p.add_run(" --- पोलीस संशोधन केंद्र, पुणे येथील आर्थिक वर्ष २०२६-२०२७ मधील प्रलंबित प्रशिक्षण शुल्क अदा करणेबाबत.")
        self._set_font(r2)

        self._set_para(doc.add_paragraph(), "महोदय,", size=12, align=WD_ALIGN_PARAGRAPH.LEFT)

        body = doc.add_paragraph()
        body.paragraph_format.space_after = Pt(0)
        body.add_run(
            "उपरोक्त विषय व संदर्भान्वये अनुसरून पोलीस संशोधन केंद्र, पुणे येथे मा. पोलीस महासंचालक, महाराष्ट्र राज्य, मुंबई यांच्या "
            "निर्देशानुसार महाराष्ट्र पोलीस दलातील पोलीस अधीक्षक ते पोलीस उप निरीक्षक दर्जाचे अधिकारी यांचे करीता नियमित सीपीआर येथे "
            "वेगवेगळ्या विषयांवर पोलीस सेवेमधील अधिका-यांसाठी प्रशिक्षण आयोजित होत असतात. आपल्या घटकातील अधिका-यांनी आर्थिक वर्ष "
            "२०२६-२०२७ मधील खालील तक्त्यातील विषयांवर प्रशिक्षण घेतलेले असून त्यांचे प्रशिक्षण रक्कम येणे बाकी आहे."
        )

        table = doc.add_table(rows=1, cols=6)
        table.style = "Table Grid"
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        headers = ["अ.क्र", "प्रशिक्षणाचे नाव व सत्र", "प्रशिक्षणाचा कालावधी", "प्रशिक्षणार्थींचे नाव", "प्रशिक्षण रक्कम", "इनव्हॉइस नंबर"]
        for i, h in enumerate(headers):
            self._set_cell(table.rows[0].cells[i], h, size=10, bold=True)

        for i, r in enumerate(records, 1):
            row = table.add_row().cells
            self._set_cell(row[0], str(i), size=10)
            self._set_cell(row[1], f"{r.course_name} {r.batch_session}".strip(), size=10, align=WD_ALIGN_PARAGRAPH.LEFT)
            self._set_cell(
                row[2],
                f"{r.from_date.strftime('%d/%m/%Y') if r.from_date else ''} To {r.to_date.strftime('%d/%m/%Y') if r.to_date else ''}",
                size=10,
                align=WD_ALIGN_PARAGRAPH.CENTER,
            )
            self._set_cell(row[3], r.officer_name, size=10, align=WD_ALIGN_PARAGRAPH.LEFT)
            self._set_cell(row[4], self.fmt_money(r.total_fee), size=10, align=WD_ALIGN_PARAGRAPH.RIGHT)
            self._set_cell(row[5], reference_no, size=10, align=WD_ALIGN_PARAGRAPH.RIGHT)

        total = sum(r.total_fee for r in records)
        tp = doc.add_paragraph()
        tp.paragraph_format.space_before = Pt(4)
        tp.paragraph_format.space_after = Pt(0)
        rr = tp.add_run("एकूण ")
        self._set_font(rr, bold=True)
        rr2 = tp.add_run(f"{self.fmt_money(total)}/-")
        self._set_font(rr2, bold=True)

        doc.add_paragraph("सदर प्रशिक्षणासाठी आपले घटकाकडून वरील प्रमाणे थकबाकी रक्कम अद्याप पोलीस संशोधन केंद्र पुणे या कार्यालयास प्राप्त झालेली नाही.")
        doc.add_paragraph("प्रती प्रशिक्षणार्थी प्रती दिवस रू.२०००/- या प्रमाणे एकूण रक्कम रू. १०,०००/-या प्रमाणे प्रशिक्षण शुल्क पोलीस संशोधन केंद्र पुणे (CPR Pune) या नावाने धनादेश अथवा खालील बँक खात्यामध्ये RTGS / NEFT ने जमा करण्यास आणि त्याची माहिती directorcprpune@gmail.com या ई-मेलवर पाठविण्याची विनंती आहे.")

        bank = doc.add_table(rows=1, cols=2)
        bank.style = "Table Grid"
        bank.alignment = WD_TABLE_ALIGNMENT.CENTER
        self._set_cell(bank.rows[0].cells[0], "Field", size=10, bold=True)
        self._set_cell(bank.rows[0].cells[1], "Value", size=10, bold=True)

        bank_rows = [
            ("1 Name", "CPR, Pune"),
            ("2 Account No", "10023971530"),
            ("3 Type of Account", "Saving Bank Account"),
            ("4 Bank Name", "State Bank of India"),
            ("5 Bank Address", "NCL Branch, Pashan Road, Pune 411008."),
            ("6 Bank Branch Code", "3552"),
            ("7 MICR Code", "411002012"),
            ("8 IFSC Code", "SBIN0003552"),
            ("9 Mobile No.", "9960631393"),
            ("10 Payee Code", "22010023828"),
        ]
        for field, val in bank_rows:
            row = bank.add_row().cells
            self._set_cell(row[0], field, size=10, align=WD_ALIGN_PARAGRAPH.LEFT)
            self._set_cell(row[1], val, size=10, align=WD_ALIGN_PARAGRAPH.LEFT)

        doc.add_paragraph("तरी सदर थकीत प्रशिक्षण फी रक्कम पोलीस संशोधन केंद्र पुणे येथे तातडीने पाठविण्याची विनंती आहे.")
        doc.add_paragraph("सांवत:- वर नमुद इन व्हॉईस नंबर प्रमाणे प्रशिक्षण फी Demand Note/Invoice जोडलेले आहे.")

        sig = doc.add_paragraph()
        sig.paragraph_format.space_before = Pt(4)
        sig.paragraph_format.space_after = Pt(0)
        self._set_para(sig, "(डॉ. काकासाहेब डोळे)\nपोलीस अधीक्षक,\nपोलीस संशोधन केंद्र पुणे", size=12, align=WD_ALIGN_PARAGRAPH.RIGHT)

        doc.save(output_path)
