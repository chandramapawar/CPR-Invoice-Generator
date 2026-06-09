from io import BytesIO
import pandas as pd
from src.models import TrainingRecord

class ExcelReader:
    def __init__(self, file_obj):
        self.file_obj = file_obj

    def _as_bytes(self):
        if hasattr(self.file_obj, "getvalue"):
            return self.file_obj.getvalue()
        if hasattr(self.file_obj, "read"):
            data = self.file_obj.read()
            if hasattr(self.file_obj, "seek"):
                self.file_obj.seek(0)
            return data
        if isinstance(self.file_obj, (bytes, bytearray)):
            return bytes(self.file_obj)
        with open(self.file_obj, "rb") as f:
            return f.read()

    def _load_dataframe(self):
        data = self._as_bytes()
        xls = pd.ExcelFile(BytesIO(data))
        preferred_sheets = ["Invoice Ready Data", "For Docs", "Form responses 1"]
        sheet_name = next((s for s in preferred_sheets if s in xls.sheet_names), xls.sheet_names[0])
        df = pd.read_excel(BytesIO(data), sheet_name=sheet_name)
        df.columns = [str(c).strip() for c in df.columns]
        df = df.dropna(how="all")
        return df, sheet_name

    def read_training_records(self):
        df, _ = self._load_dataframe()
        if df.empty:
            return []

        records = []
        for _, row in df.iterrows():
            officer_name = str(row.get("Officer Name", "")).strip()
            if not officer_name or officer_name.lower() == "nan":
                first = str(row.get("First Name", "")).strip()
                middle = str(row.get("Middle Name", "")).strip()
                surname = str(row.get("Surname", "")).strip()
                officer_name = " ".join([p for p in [first, middle, surname] if p and p.lower() != "nan"]).strip()

            police_unit = str(row.get("Police Unit Final", row.get("Name of Unit", ""))).strip()
            if not police_unit or police_unit.lower() == "nan":
                police_unit = str(row.get("(If not in the list then) Unit Name", "")).strip()

            rank = str(row.get("Rank", "")).strip()
            sevarth_id = str(row.get("Sevarth ID", row.get("Sevarth ID (Must)", ""))).strip()

            if not officer_name or not police_unit:
                continue

            records.append(TrainingRecord(
                officer_name=officer_name,
                rank=rank,
                buckle_no=sevarth_id,
                police_unit=police_unit,
                district=str(row.get("District/Unit Address", "")).strip(),
                unit_head_designation=str(row.get("Unit Head", "")).strip(),
                course_name=str(row.get("Course Name", row.get("Training", ""))).strip(),
                batch_session=str(row.get("Batch / Session", "")).strip(),
                from_date=pd.to_datetime(row.get("From Date"), dayfirst=True, errors="coerce").date() if pd.notna(row.get("From Date")) else None,
                to_date=pd.to_datetime(row.get("To Date"), dayfirst=True, errors="coerce").date() if pd.notna(row.get("To Date")) else None,
                duration_days=int(row.get("Duration Days", 0)) if pd.notna(row.get("Duration Days")) else 0,
                fee_per_day=float(row.get("Fee Per Day", 0)) if pd.notna(row.get("Fee Per Day")) else 0.0,
                total_fee=float(row.get("Total Fee", 0)) if pd.notna(row.get("Total Fee")) else 0.0,
                reference_no=str(row.get("Reference No", "")).strip(),
                financial_year=str(row.get("Financial Year", "")).strip(),
                additional_notes=str(row.get("Additional Notes", "")).strip(),
            ))
        return records#excel_reader
