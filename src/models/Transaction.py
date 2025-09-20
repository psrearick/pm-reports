from dataclasses import dataclass
import datetime

@dataclass
class Transaction:
    property: str
    unit: str
    date: datetime
    credits: float
    fees: float
    debit: float
    security_deposits: float
    explanation: str
    markup_included: bool
    internal_notes: str
