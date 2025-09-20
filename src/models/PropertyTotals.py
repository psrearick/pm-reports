from dataclasses import dataclass

@dataclass
class PropertyTotals:
    total_credits: float
    total_fees: float
    total_markup_revenue: float
    total_maf: float
    total_security_deposits: float
    total_debits: float
    due_to_owners: float
    total_to_pm: float
    airbnb_total: float = 0
    airbnb_pm_fee: float = 0
