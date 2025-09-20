from src.models.PropertyConfig import PropertyConfig
from src.models.PropertySummary import PropertySummary
from src.models.PropertyTotals import PropertyTotals
from src.models.Transaction import Transaction


class PDFGenerator:
    def generate_income_statement(self, property_name: str,
            transactions: list[Transaction],
            totals: PropertyTotals,
            config: PropertyConfig) -> str:  # returns file path
        pass

    def generate_summary_page(self, summary_data: list[PropertySummary]) -> str:
        pass
