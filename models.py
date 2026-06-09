from dataclasses import dataclass
from datetime import date

@dataclass
class TrainingRecord:
    officer_name: str
    rank: str
    buckle_no: str
    police_unit: str
    district: str
    unit_head_designation: str
    course_name: str
    batch_session: str
    from_date: date | None
    to_date: date | None
    duration_days: int
    fee_per_day: float
    total_fee: float
    reference_no: str
    financial_year: str
    additional_notes: str