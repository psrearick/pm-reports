from dataclasses import dataclass
import datetime

@dataclass
class Credit:
    property: str
    unit: str
    date: datetime
    amount: float
    category: str
    subcategory: str
    payment_method: str
