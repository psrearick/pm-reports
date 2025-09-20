# File-Based Property Management Report Generator - Architecture Plan

## File Structure & Data Flow

### Input Files

**1. Transactions CSV (user-created)**
```csv
Property,Unit,Date,Credits,Fees,Debit,Security_Deposits,Debit_Credit_Explanation,Markup_Included,Internal_Notes
Maple Grove,101,2025-09-15,1200.00,0,0,0,Standard rent,FALSE,
Maple Grove,CAM,2025-09-16,0,0,150.00,0,Landscaping,TRUE,Monthly maintenance
Oak Hill,201,2025-09-17,0,25.00,0,0,Late payment fee,FALSE,
```

**2. Credits File (from payment processor)**
- Whatever format the processor exports
- App will map columns to standard fields

**3. Config File (JSON)**
```json
{
  "properties": {
    "Maple Grove": {
      "unit_count": 12,
      "maf_rate": 0.06,
      "property_type": "standard"
    },
    "Park Property": {
      "unit_count": 8,
      "maf_rate": 0.08,
      "property_type": "park_state",
      "airbnb_enabled": true
    }
  },
  "fees": {
    "mailing_fee": 12.00,
    "markup_rate": 0.10,
    "unit_fee": 5.00
  },
  "report_settings": {
    "output_directory": "./reports",
    "date_format": "%m/%d/%Y"
  }
}
```

## Application Architecture

### Core Modules

**1. File Handlers (`file_handlers/`)**
```python
# csv_handler.py
class TransactionCSVHandler:
    def validate_csv(self, file_path) -> ValidationResult
    def load_transactions(self, file_path) -> List[Transaction]
    def get_column_mapping(self) -> Dict[str, str]

# credits_handler.py
class CreditsFileHandler:
    def detect_format(self, file_path) -> str
    def load_credits(self, file_path) -> List[Credit]
    def map_columns(self, raw_data) -> List[Credit]

# config_handler.py
class ConfigHandler:
    def load_config(self, file_path) -> Config
    def validate_config(self, config) -> ValidationResult
    def get_default_config(self) -> Config
```

**2. Data Models (`models/`)**
```python
# transaction.py
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

# credit.py
@dataclass
class Credit:
    property: str
    unit: str
    date: datetime
    amount: float
    category: str
    subcategory: str
    payment_method: str

# property_config.py
@dataclass
class PropertyConfig:
    name: str
    unit_count: int
    maf_rate: float
    property_type: str
    airbnb_enabled: bool = False
```

**3. Data Processing (`processors/`)**
```python
# data_merger.py
class DataMerger:
    def merge_credits_to_transactions(self,
        transactions: List[Transaction],
        credits: List[Credit]) -> List[Transaction]
    def detect_duplicates(self) -> List[str]
    def validate_data_consistency(self) -> ValidationResult

# calculator.py
class PropertyCalculator:
    def calculate_totals(self, transactions: List[Transaction],
                        config: PropertyConfig) -> PropertyTotals
    def calculate_maf(self, credits_sum: float, unit_count: int,
                     maf_rate: float) -> float
    def calculate_markup_revenue(self, transactions: List[Transaction]) -> float

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
```

**4. Report Generation (`report_generator/`)**
```python
# pdf_generator.py
class PDFGenerator:
    def generate_income_statement(self, property_name: str,
        transactions: List[Transaction],
        totals: PropertyTotals,
        config: PropertyConfig) -> str  # returns file path

    def generate_summary_page(self, summary_data: List[PropertySummary]) -> str

# template_engine.py
class ReportTemplate:
    def render_income_statement_html(self, data: dict) -> str
    def render_summary_page_html(self, data: dict) -> str
    def apply_styling(self, html: str) -> str
```

**5. User Interface (`ui/`)**
```python
# main_window.py (using tkinter/customtkinter)
class MainWindow:
    def __init__(self):
        self.setup_ui()
        self.file_paths = {}

    def setup_ui(self):
        # File selection section
        # Progress display
        # Generate button
        # Output directory selection

    def select_transactions_file(self)
    def select_credits_file(self)
    def select_config_file(self)
    def validate_inputs(self) -> bool
    def generate_reports(self)
    def show_progress(self, message: str, percent: int)
```

### Main Application Flow

**1. Application Entry Point (`main.py`)**
```python
class PropertyReportApp:
    def __init__(self):
        self.window = MainWindow()
        self.setup_handlers()

    def run(self):
        # Start the GUI

    def generate_reports_workflow(self, file_paths: dict):
        # 1. Load and validate config
        # 2. Load and validate transactions CSV
        # 3. Load and process credits file
        # 4. Merge data
        # 5. Calculate totals for each property
        # 6. Generate PDF reports
        # 7. Generate summary page
        # 8. Show completion message
```

## Error Handling & Validation

### Validation Layers

**1. File Validation**
- CSV structure (required columns present)
- Data types (dates, numbers)
- Required fields not empty
- Valid property names (match config)

**2. Data Validation**
- Date ranges reasonable
- No negative values where inappropriate
- Credits file matches expected format
- Property names consistent between files

**3. Business Logic Validation**
- Total debits/credits balance makes sense
- MAF calculations within expected ranges
- No duplicate transactions

### Error Reporting
```python
@dataclass
class ValidationResult:
    is_valid: bool
    errors: List[str]
    warnings: List[str]

class ErrorHandler:
    def log_error(self, error: str, context: dict)
    def show_user_error(self, errors: List[str])
    def generate_error_report(self) -> str
```

## User Interface Design

### Main Window Layout
```
┌─────────────────────────────────────────────────┐
│ Property Report Generator                        │
├─────────────────────────────────────────────────┤
│ Input Files:                                    │
│ ┌─────────────────────────────────────────────┐ │
│ │ Transactions CSV: [Browse...] [Validate]   │ │
│ │ Credits File:     [Browse...] [Validate]   │ │
│ │ Config File:      [Browse...] [Optional]   │ │
│ └─────────────────────────────────────────────┘ │
│                                                 │
│ Output Settings:                                │
│ ┌─────────────────────────────────────────────┐ │
│ │ Output Directory: [Browse...]               │ │
│ │ ☑ Open folder when complete                │ │
│ └─────────────────────────────────────────────┘ │
│                                                 │
│ ┌─────────────────────────────────────────────┐ │
│ │ [Validate All Files]  [Generate Reports]   │ │
│ └─────────────────────────────────────────────┘ │
│                                                 │
│ Progress: ████████████████████ 85%             │
│ Status: Generating Maple Grove report...       │
└─────────────────────────────────────────────────┘
```

## Packaging & Distribution

### Development Setup
```
property_reports/
├── main.py
├── requirements.txt
├── config/
│   └── default_config.json
├── ui/
│   ├── __init__.py
│   └── main_window.py
├── models/
│   ├── __init__.py
│   ├── transaction.py
│   ├── credit.py
│   └── property_config.py
├── file_handlers/
│   ├── __init__.py
│   ├── csv_handler.py
│   ├── credits_handler.py
│   └── config_handler.py
├── processors/
│   ├── __init__.py
│   ├── data_merger.py
│   └── calculator.py
├── report_generator/
│   ├── __init__.py
│   ├── pdf_generator.py
│   └── template_engine.py
├── templates/
│   ├── income_statement.html
│   └── summary_page.html
└── tests/
    ├── test_data/
    ├── test_csv_handler.py
    └── test_calculator.py
```

### Key Libraries
```txt
pandas>=1.5.0          # CSV/Excel processing
openpyxl>=3.0.0        # Excel file handling
weasyprint>=57.0       # HTML to PDF conversion
customtkinter>=5.0.0   # Modern tkinter UI
Jinja2>=3.0.0         # Template rendering
python-dateutil>=2.8.0 # Date parsing
```

### Deployment
- Use PyInstaller to create standalone executable
- Include default config template
- Bundle with sample CSV template
- Create simple installer or zip package

This architecture provides a clean separation of concerns, thorough validation, and a user-friendly interface while keeping the complexity manageable for a single-developer project.
