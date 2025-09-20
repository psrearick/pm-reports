from dataclasses import dataclass

@dataclass
class PropertySummary:
    due_to_owners: float
    total_to_pm: float
    total_fees: float
    new_lease_fees: float
    renewal_fees: float
