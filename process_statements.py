"""
Dental Statement Balance Automation

Parses PDF statements, filters outstanding balances, merges with tracking data.
"""

import pdfplumber
import pandas as pd
import re
import os
from openpyxl.utils import get_column_letter
from config import (
    PDF_INPUT_PATH, OLD_EXCEL_PATH, OUTPUT_EXCEL_PATH,
    OLD_SHEET_NAME, OLD_SHEET_START_ROW,
    OUTPUT_COLUMNS, TRACKING_COLUMNS, DATE_FORMAT
)


# =============================================================================
# PDF PARSING
# =============================================================================

def parse_pdf_statements(pdf_path):
    """Extract patient data from PDF."""
    print(f"📄 Parsing PDF: {pdf_path}")
    all_rows = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            print(f"   Processing page {page_num}/{len(pdf.pages)}...", end='\r')
            text = page.extract_text()
            if not text:
                continue
            
            for line in text.split('\n'):
                if not line.strip():
                    continue
                # Skip headers
                if any(x in line for x in ['PATIENT NAME', 'CHART #', 'PATIENT BALANCE REPORT', 'Shailaja Singh', 'Date:', 'Page:']):
                    continue
                
                parsed_row = parse_patient_line(line)
                if parsed_row:
                    all_rows.append(parsed_row)
    
    print(f"   ✓ Extracted {len(all_rows)} total records")
    
    return pd.DataFrame(all_rows, columns=[
        'PATIENT NAME', 'CHART #', 'HOME PH #', 'WORK PH #', 'EXT.', 'BT',
        'LAST PATIENT PMT', 'LAST VISIT DATE', 'PEND. CLAIMS',
        'FAMILY BALANCE', 'PATIENT BALANCE'
    ])


def parse_patient_line(line):
    """Parse a single patient line from PDF."""
    line = line.strip()
    
    # Must contain 6-digit chart number
    if not re.search(r'\d{6}', line):
        return None
    
    # Extract guarantor indicator (*)
    guarantor_star = ''
    if line.startswith('*'):
        guarantor_star = '*'
        line = line[1:]
    elif line.startswith('.'):
        line = line[1:]
    
    tokens = line.split()
    
    # Find chart number
    chart_idx = None
    for i, token in enumerate(tokens):
        if re.match(r'^\d{6}$', token):
            chart_idx = i
            break
    
    if chart_idx is None:
        return None
    
    # Patient name is everything before chart number
    patient_name = ' '.join(tokens[:chart_idx])
    chart_num = tokens[chart_idx]
    remaining = tokens[chart_idx + 1:]
    
    # Initialize fields
    home_phone = work_phone = ext = bt = '( )'
    last_pmt_date = last_visit_date = pend_claims = ''
    family_balance = patient_balance = ''
    
    idx = 0
    
    # Extract phone numbers
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
    
    # BT number
    if idx < len(remaining) and re.match(r'^\d+$', remaining[idx]) and len(remaining[idx]) <= 2:
        bt = remaining[idx]
        idx += 1
    
    # Balances are last two tokens
    if len(remaining) >= 2:
        patient_balance = remaining[-1]
        family_balance = remaining[-2]
        end_tokens = remaining[idx:-2]
        
        # Parse dates and YES flag
        for token in end_tokens:
            if re.match(r'\d{2}/\d{2}/\d{4}', token):
                if not last_pmt_date:
                    last_pmt_date = token
                elif not last_visit_date:
                    last_visit_date = token
            elif token.upper() == 'YES':
                pend_claims = 'YES'
    
    if guarantor_star:
        patient_name = guarantor_star + patient_name
    
    return [patient_name, chart_num, home_phone, work_phone, ext, bt,
            last_pmt_date, last_visit_date, pend_claims, family_balance, patient_balance]


# =============================================================================
# DATA PROCESSING
# =============================================================================

def apply_parsing_rules(df):
    """Extract guarantor, clean claims, convert data types."""
    print("\n🔧 Applying parsing rules...")
    df = df.copy()
    
    # Create GUARANTOR column
    df.loc[:, 'GUARANTOR'] = df['PATIENT NAME'].apply(lambda x: 'Y' if str(x).startswith('*') else 'N')
    df.loc[:, 'PATIENT NAME'] = df['PATIENT NAME'].str.replace('*', '', regex=False)
    print(f"   ✓ Identified {(df['GUARANTOR'] == 'Y').sum()} guarantors")
    
    # Clean PEND. CLAIMS
    df.loc[:, 'PEND. CLAIMS'] = df['PEND. CLAIMS'].apply(lambda x: 'YES' if str(x).strip().upper() == 'YES' else '')
    print(f"   ✓ Found {(df['PEND. CLAIMS'] == 'YES').sum()} pending claims")
    
    # Convert dates
    df.loc[:, 'LAST PATIENT PMT'] = pd.to_datetime(df['LAST PATIENT PMT'], errors='coerce')
    df.loc[:, 'LAST VISIT DATE'] = pd.to_datetime(df['LAST VISIT DATE'], errors='coerce')
    
    # Convert balances
    for col in ['FAMILY BALANCE', 'PATIENT BALANCE']:
        df.loc[:, col] = pd.to_numeric(
            df[col].astype(str).str.replace('$', '').str.replace(',', ''),
            errors='coerce'
        ).fillna(0)
    
    # Ensure text columns
    for col in ['CHART #', 'HOME PH #', 'WORK PH #']:
        df.loc[:, col] = df[col].astype(str)
    
    print("   ✓ Data types converted")
    
    # Reorder columns to put GUARANTOR after PATIENT NAME
    cols = list(df.columns)
    cols.remove('GUARANTOR')
    cols.insert(1, 'GUARANTOR')
    
    return df[cols]


def filter_outstanding_balances(df):
    """Keep only records with PATIENT BALANCE > 0 AND FAMILY BALANCE > 0."""
    print("\n🔍 Filtering for outstanding balances...")
    initial_count = len(df)
    
    df_filtered = df[df['PATIENT BALANCE'] > 0].copy()
    after_rule1 = len(df_filtered)
    print(f"   Rule 1 (PATIENT BALANCE > 0): {initial_count} → {after_rule1} records")
    
    df_filtered = df_filtered[df_filtered['FAMILY BALANCE'] > 0].copy()
    final_count = len(df_filtered)
    print(f"   Rule 2 (FAMILY BALANCE > 0): {after_rule1} → {final_count} records")
    print(f"   ✓ Final outstanding AR: {final_count} accounts")
    
    return df_filtered


def load_old_tracking_sheet(excel_path, sheet_name='Statement', start_row=3):
    """Load previous month's tracking sheet."""
    print(f"\n📊 Loading old tracking sheet: {excel_path}")
    
    try:
        df_old = pd.read_excel(excel_path, sheet_name=sheet_name, skiprows=start_row - 1, header=0)
        print(f"   Found {len(df_old.columns)} columns: {list(df_old.columns)[:10]}...")
        
        # Make a copy to avoid pandas warnings when modifying
        df_old = df_old.copy()
        
        # Standardize column names
        df_old = df_old.rename(columns={
            'Patient Name': 'Patient Name',
            'Guarantor': 'Guarantor',
            'Chart #': 'CHART #',
            'Chart#': 'CHART #',
        })
        
        # Convert CHART # to string
        if 'CHART #' in df_old.columns:
            df_old['CHART #'] = df_old['CHART #'].astype(str)
        
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


# =============================================================================
# CHART NUMBER NORMALIZATION
# =============================================================================

def normalize_chart_number(chart_num):
    """
    Normalize chart number to 6 digits.
    Handles: "1234", "001234", 1234, 1234.0 → "001234"
    """
    chart_str = str(chart_num)
    
    # Remove decimal part (handles 1234.0 → 1234)
    if '.' in chart_str:
        chart_str = chart_str.split('.')[0]
    
    # Remove non-digits
    chart_str = re.sub(r'\D', '', chart_str)
    
    # Pad to 6 digits
    return chart_str.zfill(6)


def normalize_chart_numbers(df):
    """Apply chart number normalization to DataFrame."""
    result = df.copy()
    normalized_values = result['CHART #'].astype(str).apply(normalize_chart_number)
    result.loc[:, 'CHART #'] = normalized_values
    return result


# =============================================================================
# MERGING
# =============================================================================

def merge_with_tracking_data(df_new, df_old):
    """Merge new statement data with old tracking data on CHART #."""
    print("\n🔗 Merging with tracking data...")
    
    # Normalize CHART # in both DataFrames
    df_new = normalize_chart_numbers(df_new)
    
    if not df_old.empty:
        df_old = normalize_chart_numbers(df_old)
        
        # Select tracking columns from old sheet
        tracking_cols = ['CHART #'] + TRACKING_COLUMNS
        tracking_cols = [col for col in tracking_cols if col in df_old.columns]
        df_old_tracking = df_old[tracking_cols]
        
        # Remove duplicates (keep most recent)
        df_old_tracking = df_old_tracking.drop_duplicates(subset=['CHART #'], keep='last')
        
        # LEFT JOIN on CHART #
        df_merged = df_new.merge(df_old_tracking, on='CHART #', how='left', suffixes=('', '_old'))
        
        # Count matches
        matched = df_merged['STATUS'].notna().sum() if 'STATUS' in df_merged.columns else 0
        new_charts = set(df_new['CHART #'].unique())
        old_charts = set(df_old_tracking['CHART #'].unique())
        common_charts = new_charts & old_charts
        
        print(f"   ✓ {len(common_charts)} patients found in old tracking sheet")
        print(f"   ✓ {matched} patients have tracking data (STATUS/NOTES)")
        print(f"   ✓ {len(df_merged) - matched} patients have no tracking history")
        
    else:
        df_merged = df_new
        print("   ℹ️  No old tracking data found - using new data only")
    
    return df_merged


# =============================================================================
# EXCEL EXPORT
# =============================================================================

def export_to_excel(df, output_path):
    """Export final merged data to Excel."""
    print(f"\n💾 Exporting to Excel: {output_path}")
    df = df.copy()
    
    # Rename columns
    df = df.rename(columns={'PATIENT NAME': 'Patient Name', 'GUARANTOR': 'Guarantor'})
    
    # Format dates as MM/DD/YY strings
    date_cols = ['LAST PATIENT PMT', 'LAST VISIT DATE', 'Follow-Up Date', 
                 'ST1 DATE', 'ST2 DATE', 'ST3 DATE']
    for col in date_cols:
        if col in df.columns:
            df.loc[:, col] = pd.to_datetime(df[col], errors='coerce')
            df.loc[:, col] = df[col].apply(lambda x: x.strftime(DATE_FORMAT) if pd.notna(x) else '')
    
    # Format currency columns
    currency_cols = ['FAMILY BALANCE', 'PATIENT BALANCE', 'AMOUNT1', 'AMOUNT2', 'AMOUNT3']
    for col in currency_cols:
        if col in df.columns:
            df.loc[:, col] = pd.to_numeric(df[col], errors='coerce')
    
    # Add missing columns
    for col in OUTPUT_COLUMNS:
        if col not in df.columns:
            df[col] = ''
    
    # Select output columns in order
    df = df[OUTPUT_COLUMNS]
    
    # Write to Excel
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Merged Statements', index=False)
        
        # Auto-adjust column widths
        worksheet = writer.sheets['Merged Statements']
        for idx, col in enumerate(df.columns, 1):
            max_length = max(df[col].astype(str).apply(len).max(), len(str(col)))
            worksheet.column_dimensions[get_column_letter(idx)].width = min(max_length + 2, 50)
    
    print(f"   ✓ Exported {len(df)} records with {len(OUTPUT_COLUMNS)} columns")
    print(f"   ✓ File saved: {output_path}")


# =============================================================================
# MAIN EXECUTION
# =============================================================================

def main():
    """Run the complete workflow."""
    print("=" * 70)
    print("DENTAL STATEMENT BALANCE AUTOMATION")
    print("=" * 70)
    
    # Create output directory if needed
    output_dir = os.path.dirname(OUTPUT_EXCEL_PATH)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    try:
        # Step 1: Parse PDF
        df_raw = parse_pdf_statements(PDF_INPUT_PATH)
        
        # Step 2: Apply parsing rules
        df_parsed = apply_parsing_rules(df_raw)
        
        # Step 3: Filter for outstanding balances
        df_filtered = filter_outstanding_balances(df_parsed)
        
        # Step 4: Load old tracking sheet
        df_old = load_old_tracking_sheet(OLD_EXCEL_PATH, OLD_SHEET_NAME, OLD_SHEET_START_ROW)
        
        # Step 5: Merge data
        df_final = merge_with_tracking_data(df_filtered, df_old)
        
        # Step 6: Export to Excel
        export_to_excel(df_final, OUTPUT_EXCEL_PATH)
        
        print("\n" + "=" * 70)
        print("✅ SUCCESS! Workflow completed.")
        print("=" * 70)
        print(f"\nSummary:")
        print(f"  • Input PDF records: {len(df_raw)}")
        print(f"  • Outstanding balances: {len(df_filtered)}")
        print(f"  • Final merged records: {len(df_final)}")
        print(f"\nOutput file: {OUTPUT_EXCEL_PATH}")
        
    except FileNotFoundError as e:
        print(f"\n❌ ERROR: File not found - {e}")
        print(f"\nMake sure you have:")
        print(f"  1. PDF file: {PDF_INPUT_PATH}")
        print(f"  2. Old Excel file: {OLD_EXCEL_PATH}")
        
    except Exception as e:
        print(f"\n❌ ERROR: {e}")
        import traceback
        traceback.print_exc()


if __name__ == '__main__':
    main()
