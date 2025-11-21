"""
Configuration file for dental statement automation.

Adjust these settings to match your file names and requirements.
"""

# =============================================================================
# FILE PATHS
# =============================================================================

# Input files
PDF_INPUT_PATH = 'input/PAT_BAL_20251116.pdf'  # PDF from dental software
OLD_EXCEL_PATH = 'input/Statements.xlsx'       # Previous month's tracking sheet

# Output file
OUTPUT_EXCEL_PATH = 'output/Merged_Statements.xlsx'


# =============================================================================
# EXCEL SHEET SETTINGS
# =============================================================================

# Sheet name in the old Excel file
OLD_SHEET_NAME = 'Statement'

# Row number where the COLUMN HEADERS are (1-indexed)
# Example: If row 3 has headers like "Patient Name", "Chart #", etc., set this to 3
# The actual DATA should start in the row immediately after this header row
OLD_SHEET_START_ROW = 3


# =============================================================================
# OUTPUT COLUMNS
# =============================================================================

# These columns will appear in the final Excel output (in this order)
# Columns from PDF: Patient Name, Guarantor, Chart #, dates, balances
# Tracking columns: STATUS, NOTES, Follow-Up Date, etc. (merged from old sheet)
OUTPUT_COLUMNS = [
    'Patient Name',
    'Guarantor',
    'CHART #',
    'LAST PATIENT PMT',
    'LAST VISIT DATE',
    'PEND. CLAIMS',
    'FAMILY BALANCE',
    'PATIENT BALANCE',
    'STATUS',
    'NOTES',
    'Follow-Up Date',
    'Staff Code',
    'Ortho',
    'BALANCE CODE',
    'ST1 DATE',
    'AMOUNT1',
    'ST2 DATE',
    'AMOUNT2',
    'ST3 DATE',
    'AMOUNT3'
]


# =============================================================================
# TRACKING COLUMNS
# =============================================================================

# These columns are merged from the old tracking sheet (not in PDF)
# They contain your manual AR tracking data
TRACKING_COLUMNS = [
    'STATUS',
    'NOTES',
    'Follow-Up Date',
    'Staff Code',
    'Ortho',
    'BALANCE CODE',
    'ST1 DATE',
    'AMOUNT1',
    'ST2 DATE',
    'AMOUNT2',
    'ST3 DATE',
    'AMOUNT3'
]


# =============================================================================
# DATE FORMATTING
# =============================================================================

# Date format for Excel output - MM/DD/YY with no time
DATE_FORMAT = '%m/%d/%y'
