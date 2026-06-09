class InvoiceRegistry:
    def get_next_invoice_no(self):
        return 1

    def format_invoice_no(self, invoice_no_int, financial_year):
        return f"{financial_year}-{invoice_no_int:04d}"