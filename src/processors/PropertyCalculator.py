from src.models.PropertyConfig import PropertyConfig
from src.models.PropertyTotals import PropertyTotals
from src.models.Transaction import Transaction


class PropertyCalculator:
    def calculate_totals(self, transactions: list[Transaction],
                        config: PropertyConfig) -> PropertyTotals:
        pass
    def calculate_maf(self, credits_sum: float, unit_count: int,
                     maf_rate: float) -> float:
        pass
    def calculate_markup_revenue(self, transactions: list[Transaction]) -> float:
        pass
