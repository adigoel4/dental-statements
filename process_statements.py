"""
Dental Statement Balance Automation Script

This script automates the monthly workflow of:
1. Parsing PDF patient statements
2. Filtering for outstanding balances
3. Merging with existing tracking data
4. Outputting final Excel file
"""

import pdfplumber
import pandas as pd
import re
import os
from datetime import datetime
from config import (
    PDF_INPUT_PATH, 
    OLD_EXCEL_PATH, 
    OUTPUT_EXCEL_PATH, 
    OLD_SHEET_NAME, 
    OLD_SHEET_START_ROW,
    OUTPUT_COLUMNS,
    TRACKING_COLUMNS,
    DATE_FORMAT
)


# ============================================================================
# STEP 1: PDF PARSING
# ============================================================================

def parse_pdf_statements(pdf_path):
    """
    Parse the dental software PDF and extract patient statement data.
    The PDF contains text lines with space-separated values.
    
    Args:
        pdf_path: Path to the PDF file
        
    Returns:
        DataFrame with all patient records
    """
    print(f"📄 Parsing PDF: {pdf_path}")
    
    all_rows = []
    
    # Open PDF and extract text from each page
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            print(f"   Processing page {page_num}/{len(pdf.pages)}...", end='\r')
            
            # Extract text from page
            text = page.extract_text()
            if not text:
                continue
            
            # Split into lines
            lines = text.split('\n')
            
            # Parse each line
            for line in lines:
                # Skip empty lines
                if not line.strip():
                    continue
                
                # Skip header lines
                if 'PATIENT NAME' in line or 'CHART #' in line:
                    continue
                if 'PATIENT BALANCE REPORT' in line:
                    continue
                if 'Shailaja Singh' in line or 'Date:' in line or 'Page:' in line:
                    continue
                
                # Parse patient data line
                # Format: NAME CHART# (PHONE1) (PHONE2) BT [DATE] [DATE] [YES] BALANCE BALANCE
                parsed_row = parse_patient_line(line)
                if parsed_row:
                    all_rows.append(parsed_row)
    
    print(f"   ✓ Extracted {len(all_rows)} total records")
    
    # Create DataFrame with proper column names
    df = pd.DataFrame(all_rows, columns=[
        'PATIENT NAME',
        'CHART #',
        'HOME PH #',
        'WORK PH #',
        'EXT.',
        'BT',
        'LAST PATIENT PMT',
        'LAST VISIT DATE',
        'PEND. CLAIMS',
        'FAMILY BALANCE',
        'PATIENT BALANCE'
    ])
    
    return df


def parse_patient_line(line):
    """
    Parse a single patient data line from the PDF.
    
    Example lines:
    *Aatlo, Ron 001006 (209)914-6509 (408)813-3034 5 03/07/2025 05/09/2015 0.00 -170.76
    Abcede, Jill 010949 (510)205-1013 ( ) 1 0.00 0.00
    *Achari, Anishkumar 010973 (408)533-2758 ( ) 1 11/08/2025 YES 281.00 152.00
    
    Returns:
        List of values or None if line cannot be parsed
    """
    import re
    
    line = line.strip()
    
    # Must contain a 6-digit chart number to be valid
    if not re.search(r'\d{6}', line):
        return None
    
    # Extract guarantor indicator (*)
    guarantor_star = ''
    if line.startswith('*'):
        guarantor_star = '*'
        line = line[1:]
    elif line.startswith('.'):
        # Some lines start with . instead of *
        line = line[1:]
    
    # Split the line into tokens
    tokens = line.split()
    
    # Find the 6-digit chart number
    chart_idx = None
    for i, token in enumerate(tokens):
        if re.match(r'^\d{6}$', token):
            chart_idx = i
            break
    
    if chart_idx is None:
        return None
    
    # Patient name is everything before chart number
    patient_name = ' '.join(tokens[:chart_idx])
    
    # Get chart number
    chart_num = tokens[chart_idx]
    
    # Rest of the tokens after chart number
    remaining = tokens[chart_idx + 1:]
    
    # Initialize all fields
    home_phone = '( )'
    work_phone = '( )'
    ext = ''
    bt = ''
    last_pmt_date = ''
    last_visit_date = ''
    pend_claims = ''
    family_balance = ''
    patient_balance = ''
    
    # Parse remaining tokens
    # Pattern: (PHONE1) (PHONE2) BT [DATE1] [DATE2] [YES] BALANCE1 BALANCE2
    
    idx = 0
    
    # Extract phone numbers (in parentheses)
    # Handle cases like "( )" which might be split into "(" and ")"
    if idx < len(remaining) and remaining[idx].startswith('('):
        if remaining[idx] == '(' and idx + 1 < len(remaining) and remaining[idx + 1] == ')':
            home_phone = '( )'
            idx += 2
        else:
            home_phone = remaining[idx]
            idx += 1
    
    if idx < len(remaining) and remaining[idx].startswith('('):
        if remaining[idx] == '(' and idx + 1 < len(remaining) and remaining[idx + 1] == ')':
            work_phone = '( )'
            idx += 2
        else:
            work_phone = remaining[idx]
            idx += 1
    
    # BT number (single digit usually)
    if idx < len(remaining) and re.match(r'^\d+$', remaining[idx]) and len(remaining[idx]) <= 2:
        bt = remaining[idx]
        idx += 1
    
    # Look for dates and balances at the end
    # Last two tokens should be balances
    if len(remaining) >= 2:
        patient_balance = remaining[-1]
        family_balance = remaining[-2]
        
        # Remove balances from remaining
        end_tokens = remaining[idx:-2]
        
        # Parse remaining tokens (dates, YES flag)
        for token in end_tokens:
            if re.match(r'\d{2}/\d{2}/\d{4}', token):
                # It's a date
                if not last_pmt_date:
                    last_pmt_date = token
                elif not last_visit_date:
                    last_visit_date = token
            elif token.upper() == 'YES':
                pend_claims = 'YES'
    
    # Add guarantor star back to name
    if guarantor_star:
        patient_name = guarantor_star + patient_name
    
    return [
        patient_name,
        chart_num,
        home_phone,
        work_phone,
        ext,
        bt,
        last_pmt_date,
        last_visit_date,
        pend_claims,
        family_balance,
        patient_balance
    ]


# ============================================================================
# STEP 2: APPLY SPECIAL PARSING RULES
# ============================================================================

def apply_parsing_rules(df):
    """
    Apply special parsing rules to the extracted data:
    1. Extract GUARANTOR from patient name (*)
    2. Clean PEND. CLAIMS field
    3. Convert data types (dates, currency, text)
    
    Args:
        df: DataFrame with raw extracted data
        
    Returns:
        DataFrame with rules applied
    """
    print("\n🔧 Applying parsing rules...")
    
    df = df.copy()
    
    # Rule 1: Create GUARANTOR column
    # If patient name starts with *, they are the guarantor
    df.loc[:, 'GUARANTOR'] = df['PATIENT NAME'].apply(
        lambda x: 'Y' if str(x).startswith('*') else 'N'
    )
    
    # Remove * from patient names
    df.loc[:, 'PATIENT NAME'] = df['PATIENT NAME'].str.replace('*', '', regex=False)
    
    print(f"   ✓ Identified {(df['GUARANTOR'] == 'Y').sum()} guarantors")
    
    # Rule 2: Clean PEND. CLAIMS field
    # Only keep 'YES', everything else becomes blank
    df.loc[:, 'PEND. CLAIMS'] = df['PEND. CLAIMS'].apply(
        lambda x: 'YES' if str(x).strip().upper() == 'YES' else ''
    )
    
    print(f"   ✓ Found {(df['PEND. CLAIMS'] == 'YES').sum()} pending claims")
    
    # Rule 3: Convert data types
    
    # Convert dates (handle various formats gracefully)
    df.loc[:, 'LAST PATIENT PMT'] = pd.to_datetime(df['LAST PATIENT PMT'], errors='coerce')
    df.loc[:, 'LAST VISIT DATE'] = pd.to_datetime(df['LAST VISIT DATE'], errors='coerce')
    
    # Convert balances to numeric (remove $ and , if present)
    df.loc[:, 'FAMILY BALANCE'] = pd.to_numeric(
        df['FAMILY BALANCE'].astype(str).str.replace('$', '').str.replace(',', ''),
        errors='coerce'
    ).fillna(0)
    
    df.loc[:, 'PATIENT BALANCE'] = pd.to_numeric(
        df['PATIENT BALANCE'].astype(str).str.replace('$', '').str.replace(',', ''),
        errors='coerce'
    ).fillna(0)
    
    # Keep as text: CHART #, phone numbers, guarantor
    df.loc[:, 'CHART #'] = df['CHART #'].astype(str)
    df.loc[:, 'HOME PH #'] = df['HOME PH #'].astype(str)
    df.loc[:, 'WORK PH #'] = df['WORK PH #'].astype(str)
    
    print("   ✓ Data types converted")
    
    # Reorder columns to put GUARANTOR after PATIENT NAME
    cols = list(df.columns)
    cols.remove('GUARANTOR')
    cols.insert(1, 'GUARANTOR')
    df = df[cols]
    
    return df


# ============================================================================
# STEP 3: FILTER FOR OUTSTANDING BALANCES
# ============================================================================

def filter_outstanding_balances(df):
    """
    Filter to keep only records with true outstanding AR:
    - PATIENT BALANCE > 0
    - FAMILY BALANCE > 0
    
    Args:
        df: DataFrame with parsed data
        
    Returns:
        Filtered DataFrame with only outstanding balances
    """
    print("\n🔍 Filtering for outstanding balances...")
    
    initial_count = len(df)
    
    # Filter Rule 1: PATIENT BALANCE > 0
    df_filtered = df[df['PATIENT BALANCE'] > 0].copy()
    after_rule1 = len(df_filtered)
    
    print(f"   Rule 1 (PATIENT BALANCE > 0): {initial_count} → {after_rule1} records")
    
    # Filter Rule 2: FAMILY BALANCE > 0
    df_filtered = df_filtered[df_filtered['FAMILY BALANCE'] > 0].copy()
    final_count = len(df_filtered)
    
    print(f"   Rule 2 (FAMILY BALANCE > 0): {after_rule1} → {final_count} records")
    print(f"   ✓ Final outstanding AR: {final_count} accounts")
    
    return df_filtered


# ============================================================================
# STEP 4: LOAD OLD TRACKING SHEET
# ============================================================================

def load_old_tracking_sheet(excel_path, sheet_name='Statement', start_row=3):
    """
    Load the old tracking sheet from Statements.xlsx.
    This contains manual AR tracking data (STATUS, NOTES, Follow-Up Date, etc.)
    
    Args:
        excel_path: Path to the Excel file
        sheet_name: Name of the sheet to load
        start_row: Row to start reading from (1-indexed)
        
    Returns:
        DataFrame with old tracking data
    """
    print(f"\n📊 Loading old tracking sheet: {excel_path}")
    
    try:
        # Read Excel starting from specified row
        # skiprows is 0-indexed, so subtract 1
        # header=0 means first non-skipped row is the header
        df_old = pd.read_excel(
            excel_path,
            sheet_name=sheet_name,
            skiprows=start_row - 1,
            header=0
        )
        
        # Print column names for debugging
        print(f"   Found {len(df_old.columns)} columns: {list(df_old.columns)[:10]}...")
        
        # Standardize column names
        # Rename columns to match what we need
        column_mapping = {
            # Add mappings for any variations in column names
            'Patient Name': 'Patient Name',
            'Guarantor': 'Guarantor',
            'CHART #': 'CHART #',
            'Chart #': 'CHART #',
            'Chart#': 'CHART #',
        }
        
        df_old = df_old.rename(columns=column_mapping)
        
        print(f"   ✓ Loaded {len(df_old)} tracking records with {len(df_old.columns)} columns")
        
        return df_old
        
    except FileNotFoundError:
        print(f"   ⚠️  File not found: {excel_path}")
        print("   Creating empty tracking sheet")
        return pd.DataFrame()
    except Exception as e:
        print(f"   ⚠️  Error loading file: {e}")
        print("   Creating empty tracking sheet")
        return pd.DataFrame()


# ============================================================================
# STEP 5: NORMALIZE CHART # FOR MATCHING
# ============================================================================

def normalize_chart_number(chart_num):
    """
    Normalize CHART # for reliable matching:
    - Convert to text
    - Remove non-digits
    - Pad to 6 digits (e.g., "123" → "000123")
    
    Args:
        chart_num: Chart number (any format)
        
    Returns:
        Normalized 6-digit string
    """
    # Convert to string and remove non-digits
    chart_str = re.sub(r'\D', '', str(chart_num))
    
    # Pad to 6 digits
    return chart_str.zfill(6)


def normalize_chart_numbers(df):
    """
    Apply CHART # normalization to entire DataFrame.
    
    Args:
        df: DataFrame with CHART # column
        
    Returns:
        DataFrame with normalized CHART #
    """
    # Create a proper copy and modify the column
    result = df.copy()
    # Convert to string and normalize chart numbers
    normalized_values = result['CHART #'].astype(str).apply(normalize_chart_number)
    result.loc[:, 'CHART #'] = normalized_values
    return result


# ============================================================================
# STEP 6: MERGE NEW SHEET WITH OLD TRACKING DATA
# ============================================================================

def merge_with_tracking_data(df_new, df_old):
    """
    Merge new statement data with old tracking data.
    Uses LEFT JOIN on CHART # to preserve all new records.
    Uses TRACKING_COLUMNS from config.py to determine which fields to merge.
    
    Args:
        df_new: New filtered statement data
        df_old: Old tracking data
        
    Returns:
        Merged DataFrame with current balances + tracking history
    """
    print("\n🔗 Merging with tracking data...")
    
    # Normalize CHART # in both DataFrames
    df_new = normalize_chart_numbers(df_new)
    
    if not df_old.empty:
        df_old = normalize_chart_numbers(df_old)
        
        # Build list of columns to pull from old sheet
        # Always include CHART # for joining, plus tracking columns from config
        tracking_cols = ['CHART #'] + TRACKING_COLUMNS
        
        # Keep only columns that exist in old sheet
        tracking_cols = [col for col in tracking_cols if col in df_old.columns]
        df_old_tracking = df_old[tracking_cols]
        
        # Remove duplicates in old sheet (keep most recent)
        df_old_tracking = df_old_tracking.drop_duplicates(subset=['CHART #'], keep='last')
        
        # Perform LEFT JOIN on CHART #
        df_merged = df_new.merge(
            df_old_tracking,
            on='CHART #',
            how='left',
            suffixes=('', '_old')
        )
        
        # Count how many records had tracking data
        matched = df_merged['STATUS'].notna().sum() if 'STATUS' in df_merged.columns else 0
        print(f"   ✓ Matched {matched} records with existing tracking data")
        print(f"   ✓ {len(df_merged) - matched} new records without tracking history")
        
    else:
        # No old tracking data - just use new data
        df_merged = df_new
        print("   ℹ️  No old tracking data found - using new data only")
    
    return df_merged


# ============================================================================
# STEP 7: EXPORT TO EXCEL
# ============================================================================

def export_to_excel(df, output_path):
    """
    Export final merged data to Excel file.
    Uses OUTPUT_COLUMNS from config.py to determine which columns to include.
    Formats dates as MM/DD/YY (no time).
    
    Args:
        df: Final merged DataFrame
        output_path: Path for output Excel file
    """
    print(f"\n💾 Exporting to Excel: {output_path}")
    
    # Create a copy to avoid modifying the original
    df = df.copy()
    
    # Rename columns to match expected output names
    df = df.rename(columns={
        'PATIENT NAME': 'Patient Name',
        'GUARANTOR': 'Guarantor'
    })
    
    # Format dates as MM/DD/YY strings (no time)
    date_cols = ['LAST PATIENT PMT', 'LAST VISIT DATE', 'Follow-Up Date', 
                 'ST1 DATE', 'ST2 DATE', 'ST3 DATE']
    
    for col in date_cols:
        if col in df.columns:
            # Convert to datetime
            df.loc[:, col] = pd.to_datetime(df[col], errors='coerce')
            # Format as MM/DD/YY string (removes time)
            df.loc[:, col] = df[col].apply(
                lambda x: x.strftime(DATE_FORMAT) if pd.notna(x) else ''
            )
    
    # Format currency columns as numbers
    currency_cols = ['FAMILY BALANCE', 'PATIENT BALANCE', 
                     'AMOUNT1', 'AMOUNT2', 'AMOUNT3']
    
    for col in currency_cols:
        if col in df.columns:
            df.loc[:, col] = pd.to_numeric(df[col], errors='coerce')
    
    # Add missing columns as empty if they don't exist
    for col in OUTPUT_COLUMNS:
        if col not in df.columns:
            df[col] = ''
    
    # Select only the columns we want (in the order specified)
    df = df[OUTPUT_COLUMNS]
    
    # Write to Excel
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Merged Statements', index=False)
        
        # Auto-adjust column widths
        worksheet = writer.sheets['Merged Statements']
        for idx, col in enumerate(df.columns, 1):
            max_length = max(
                df[col].astype(str).apply(len).max(),
                len(str(col))
            )
            # Use proper Excel column letters (A, B, C... AA, AB, etc.)
            from openpyxl.utils import get_column_letter
            worksheet.column_dimensions[get_column_letter(idx)].width = min(max_length + 2, 50)
    
    print(f"   ✓ Exported {len(df)} records with {len(OUTPUT_COLUMNS)} columns")
    print(f"   ✓ File saved: {output_path}")


# ============================================================================
# MAIN EXECUTION
# ============================================================================

def main():
    """
    Main execution function - runs the complete workflow.
    """
    print("=" * 70)
    print("DENTAL STATEMENT BALANCE AUTOMATION")
    print("=" * 70)
    
    # Configuration loaded from config.py
    PDF_PATH = PDF_INPUT_PATH
    OLD_EXCEL_PATH_VAR = OLD_EXCEL_PATH
    OUTPUT_PATH = OUTPUT_EXCEL_PATH
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(OUTPUT_PATH)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    try:
        # Step 1: Parse PDF
        df_raw = parse_pdf_statements(PDF_PATH)
        
        # Step 2: Apply parsing rules
        df_parsed = apply_parsing_rules(df_raw)
        
        # Step 3: Filter for outstanding balances
        df_filtered = filter_outstanding_balances(df_parsed)
        
        # Step 4: Load old tracking sheet
        df_old = load_old_tracking_sheet(OLD_EXCEL_PATH_VAR)
        
        # Step 5 & 6: Normalize and merge
        df_final = merge_with_tracking_data(df_filtered, df_old)
        
        # Step 7: Export to Excel
        export_to_excel(df_final, OUTPUT_PATH)
        
        print("\n" + "=" * 70)
        print("✅ SUCCESS! Workflow completed.")
        print("=" * 70)
        print(f"\nSummary:")
        print(f"  • Input PDF records: {len(df_raw)}")
        print(f"  • Outstanding balances: {len(df_filtered)}")
        print(f"  • Final merged records: {len(df_final)}")
        print(f"\nOutput file: {OUTPUT_PATH}")
        
    except FileNotFoundError as e:
        print(f"\n❌ ERROR: File not found - {e}")
        print("\nMake sure you have:")
        print(f"  1. PDF file: {PDF_PATH}")
        print(f"  2. Old Excel file: {OLD_EXCEL_PATH}")
        
    except Exception as e:
        print(f"\n❌ ERROR: {e}")
        import traceback
        traceback.print_exc()


if __name__ == '__main__':
    main()

