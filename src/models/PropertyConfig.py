from dataclasses import dataclass

@dataclass
class PropertyConfig:
    name: str
    unit_count: int
    maf_rate: float
    airbnb_enabled: bool = False
