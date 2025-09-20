from src.models.Credit import Credit
from src.models.Transaction import Transaction


class DataMerger:
    def merge_credits_to_transactions(self,
        transactions: list[Transaction],
        credits: list[Credit]
    ) -> list[Transaction]:
        pass

    def detect_duplicates(self) -> list[str]:
        pass
