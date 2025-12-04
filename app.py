"""
Dental Statement Balance Automation - Streamlit Web App

Upload PDF statements and old tracking sheets to generate merged reports.
"""

import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import pdfplumber
import re
import io
from openpyxl.utils import get_column_letter
from datetime import datetime

# Page configuration
st.set_page_config(
    page_title="Dental Statement Processor",
    page_icon="🦷",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for modern, minimal design with dental office colors
st.markdown("""
<style>
    /* Main content styling */
    .main .block-container {
        padding-top: 1.5rem;
        padding-bottom: 3rem;
    }
    
    /* Headers */
    h1 {
        font-weight: 600;
        font-size: 2.5rem;
        margin-bottom: 0.3rem;
        color: #2275b0;
    }
    
    h2 {
        font-weight: 500;
        font-size: 1.8rem;
        margin-top: 1.5rem;
        margin-bottom: 0.75rem;
        color: #2275b0;
    }
    
    h3 {
        font-weight: 500;
        font-size: 1.3rem;
        margin-top: 1rem;
        margin-bottom: 0.5rem;
        color: #2275b0;
    }
    
    /* Progress indicators */
    .stProgress > div > div > div {
        background-color: #2275b0;
    }
    
    /* Metrics */
    [data-testid="stMetricValue"] {
        font-size: 1.8rem;
        font-weight: 600;
        color: #2275b0;
    }
    
    [data-testid="stMetricLabel"] {
        color: #666;
    }
    
    /* Primary buttons */
    .stButton > button[kind="primary"] {
        background-color: #2275b0;
        color: white;
        border: none;
        width: 100%;
        border-radius: 8px;
        font-weight: 500;
        padding: 0.75rem 1rem;
        transition: background-color 0.2s;
    }
    
    .stButton > button[kind="primary"]:hover {
        background-color: #1a5d8f;
    }
    
    /* Secondary buttons */
    .stButton > button {
        width: 100%;
        border-radius: 6px;
        font-weight: 500;
        padding: 0.5rem 1rem;
    }
    
    /* Download button */
    .stDownloadButton > button {
        background-color: #8cc293;
        color: white;
        border: none;
        width: 100%;
        border-radius: 8px;
        font-weight: 500;
        padding: 0.75rem 1rem;
        transition: background-color 0.2s;
    }
    
    .stDownloadButton > button:hover {
        background-color: #6fa976;
    }
    
    /* File uploader */
    [data-testid="stFileUploader"] {
        background-color: #f8f9fa;
        border-radius: 8px;
        padding: 1rem;
        border: 2px solid #e9ecef;
        transition: border-color 0.2s;
    }
    
    [data-testid="stFileUploader"]:hover {
        border-color: #2275b0;
    }
    
    /* Success messages */
    .stSuccess {
        background-color: #e8f5e9;
        border-left: 4px solid #8cc293;
        border-radius: 6px;
    }
    
    /* Info boxes */
    .stAlert {
        border-radius: 6px;
    }
    
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 2rem;
        border-bottom: 2px solid #e9ecef;
    }
    
    .stTabs [data-baseweb="tab"] {
        padding: 0.75rem 1.5rem;
        font-weight: 500;
        color: #666;
        transition: color 0.2s;
    }
    
    .stTabs [data-baseweb="tab"]:hover {
        color: #2275b0;
    }
    
    .stTabs [data-baseweb="tab"][aria-selected="true"] {
        color: #2275b0;
        border-bottom: 3px solid #2275b0;
    }
    
    /* Dividers */
    hr {
        border-color: #e9ecef;
        margin: 1rem 0;
    }
    
    /* Dataframe styling */
    .stDataFrame {
        border: 1px solid #e9ecef;
        border-radius: 8px;
    }
    
    /* Checkbox accent */
    input[type="checkbox"]:checked {
        accent-color: #2275b0;
    }
    
    /* Text input focus */
    textarea:focus, input:focus {
        border-color: #2275b0 !important;
        box-shadow: 0 0 0 1px #2275b0 !important;
    }
    
    /* Table headers */
    .table-header {
        font-weight: 600;
        font-size: 0.85rem;
        color: #666;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        padding: 0.5rem 0;
        border-bottom: 2px solid #e9ecef;
        margin-bottom: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)


# =============================================================================
# PDF PARSING FUNCTIONS
# =============================================================================

def parse_patient_line(line):
    """Parse a single patient line from PDF. Returns list of fields or None."""
    line = line.strip()
    
    # Must contain 6-digit chart number
    if not re.search(r'\d{6}', line):
        return None
    
    # Extract guarantor indicator
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


def parse_pdf_statements(pdf_file, progress_bar, status_text):
    """Extract patient data from uploaded PDF file."""
    all_rows = []
    
    with pdfplumber.open(pdf_file) as pdf:
        total_pages = len(pdf.pages)
        
        for page_num, page in enumerate(pdf.pages, 1):
            status_text.text(f"Processing page {page_num} of {total_pages}")
            progress_bar.progress(page_num / total_pages)
            
            text = page.extract_text()
            if not text:
                continue
            
            for line in text.split('\n'):
                if not line.strip():
                    continue
                    
                # Skip header lines
                if any(x in line for x in ['PATIENT NAME', 'CHART #', 'PATIENT BALANCE REPORT', 
                                           'Shailaja Singh', 'Date:', 'Page:']):
                    continue
                
                parsed_row = parse_patient_line(line)
                if parsed_row:
                    all_rows.append(parsed_row)
    
    return pd.DataFrame(all_rows, columns=[
        'PATIENT NAME', 'CHART #', 'HOME PH #', 'WORK PH #', 'EXT.', 'BT',
        'LAST PATIENT PMT', 'LAST VISIT DATE', 'PEND. CLAIMS',
        'FAMILY BALANCE', 'PATIENT BALANCE'
    ])


# =============================================================================
# DATA PROCESSING FUNCTIONS
# =============================================================================

def apply_parsing_rules(df):
    """Clean and transform parsed PDF data."""
    df = df.copy()
    
    # Create GUARANTOR column
    df.loc[:, 'GUARANTOR'] = df['PATIENT NAME'].apply(lambda x: 'Y' if str(x).startswith('*') else 'N')
    df.loc[:, 'PATIENT NAME'] = df['PATIENT NAME'].str.replace('*', '', regex=False)
    
    # Clean PEND. CLAIMS
    df.loc[:, 'PEND. CLAIMS'] = df['PEND. CLAIMS'].apply(
        lambda x: 'YES' if str(x).strip().upper() == 'YES' else ''
    )
    
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
    
    # Reorder columns - put GUARANTOR after PATIENT NAME
    cols = list(df.columns)
    cols.remove('GUARANTOR')
    # Rename PATIENT NAME to Patient Name and GUARANTOR to Guarantor
    df = df.rename(columns={'PATIENT NAME': 'Patient Name', 'GUARANTOR': 'Guarantor'})
    cols = ['Patient Name', 'Guarantor'] + [c for c in cols if c not in ['PATIENT NAME', 'Patient Name']]
    
    return df[cols]


def filter_outstanding_balances(df):
    """Keep only records with outstanding balances."""
    df_filtered = df[
        (df['PATIENT BALANCE'] > 0) & 
        (df['FAMILY BALANCE'] > 0)
    ].copy()
    
    return df_filtered


def normalize_chart_number(chart_num):
    """Normalize chart number to 6 digits with leading zeros."""
    chart_str = str(chart_num)
    
    # Remove decimal part
    if '.' in chart_str:
        chart_str = chart_str.split('.')[0]
    
    # Remove non-digits
    chart_str = re.sub(r'\D', '', chart_str)
    
    # Pad to 6 digits
    return chart_str.zfill(6)


def detect_header_row(excel_file):
    """
    Auto-detect the header row in an Excel file.
    Looks for common column names like 'Patient Name', 'CHART #', etc.
    Returns the 1-indexed row number where headers are found.
    """
    # Common header patterns to look for (case-insensitive)
    header_patterns = [
        'patient name', 'chart #', 'chart#', 'patient balance', 
        'family balance', 'guarantor', 'status', 'notes'
    ]
    
    try:
        # Read first 20 rows without header to scan for header row
        df_scan = pd.read_excel(excel_file, sheet_name=0, header=None, nrows=20)
        
        # Check each row for header patterns
        for row_idx in range(len(df_scan)):
            row_values = df_scan.iloc[row_idx].astype(str).str.lower().tolist()
            
            # Count how many header patterns match this row
            matches = sum(1 for pattern in header_patterns 
                         if any(pattern in str(val) for val in row_values))
            
            # If we find 2+ matches, this is likely the header row
            if matches >= 2:
                # Return 1-indexed row number
                return row_idx + 1
        
        # Default to row 1 if no pattern found
        return 1
        
    except Exception:
        return 1


def get_tracking_columns(excel_file, header_row):
    """
    Get columns from old Excel sheet that come after PATIENT BALANCE.
    These are the tracking columns that need to be merged.
    Returns list of column names.
    """
    try:
        # Read with detected header row
        df = pd.read_excel(excel_file, sheet_name=0, skiprows=header_row - 1, header=0)
        columns = list(df.columns)
        
        # Find PATIENT BALANCE column (case-insensitive search)
        patient_balance_idx = None
        for i, col in enumerate(columns):
            if 'patient balance' in str(col).lower():
                patient_balance_idx = i
                break
        
        # If found, return all columns after it
        if patient_balance_idx is not None:
            tracking_cols = columns[patient_balance_idx + 1:]
            # Filter out unnamed columns and empty strings
            tracking_cols = [col for col in tracking_cols 
                           if col and not str(col).startswith('Unnamed')]
            return tracking_cols
        
        # If PATIENT BALANCE not found, return empty list
        return []
        
    except Exception:
        return []


def load_old_tracking_sheet(excel_file, start_row):
    """Load old tracking sheet from uploaded Excel file."""
    try:
        # Read first sheet
        df_old = pd.read_excel(excel_file, sheet_name=0, skiprows=start_row - 1, header=0)
        
        # Standardize column names
        df_old = df_old.rename(columns={
            'Patient Name': 'Patient Name',
            'Guarantor': 'Guarantor',
            'Chart #': 'CHART #',
            'Chart#': 'CHART #',
        })
        
        # Convert CHART # to string and normalize
        if 'CHART #' in df_old.columns:
            df_old = df_old.copy()
            df_old.loc[:, 'CHART #'] = df_old['CHART #'].astype(str).apply(normalize_chart_number)
        
        return df_old
        
    except Exception as e:
        st.error(f"Error loading old tracking sheet: {e}")
        return pd.DataFrame()


def merge_with_tracking_data(df_new, df_old, excel_column_mapping):
    """Merge new statement data with old tracking data."""
    df_new = df_new.copy()
    df_new.loc[:, 'CHART #'] = df_new['CHART #'].apply(normalize_chart_number)
    
    if not df_old.empty:
        # Get columns to include from old sheet
        cols_to_include = [col for col, (include, _) in excel_column_mapping.items() if include]
        
        # Select tracking columns that exist in old sheet
        tracking_cols = ['CHART #'] + [col for col in cols_to_include if col in df_old.columns]
        df_old_subset = df_old[tracking_cols].drop_duplicates(subset=['CHART #'], keep='last')
        
        # LEFT JOIN on CHART #
        return df_new.merge(df_old_subset, on='CHART #', how='left', suffixes=('', '_old'))
    
    return df_new


def prepare_final_output(df, output_columns, date_format='%m/%d/%y'):
    """Prepare final DataFrame for export."""
    df = df.copy()
    
    # Format dates
    date_cols = ['LAST PATIENT PMT', 'LAST VISIT DATE', 'Follow-Up Date', 
                 'ST1 DATE', 'ST2 DATE', 'ST3 DATE']
    for col in date_cols:
        if col in df.columns:
            df.loc[:, col] = pd.to_datetime(df[col], errors='coerce')
            df.loc[:, col] = df[col].apply(lambda x: x.strftime(date_format) if pd.notna(x) else '')
    
    # Ensure numeric format for currency
    currency_cols = ['FAMILY BALANCE', 'PATIENT BALANCE', 'AMOUNT1', 'AMOUNT2', 'AMOUNT3']
    for col in currency_cols:
        if col in df.columns:
            df.loc[:, col] = pd.to_numeric(df[col], errors='coerce')
    
    # Add missing columns
    for col in output_columns:
        if col not in df.columns:
            df[col] = ''
    
    # Select and reorder columns
    df = df[output_columns]
    
    return df


def create_excel_download(df):
    """Create Excel file in memory for download."""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Merged Statements', index=False)
        
        # Auto-adjust column widths
        worksheet = writer.sheets['Merged Statements']
        for idx, col in enumerate(df.columns, 1):
            max_length = max(df[col].astype(str).apply(len).max(), len(str(col)))
            worksheet.column_dimensions[get_column_letter(idx)].width = min(max_length + 2, 50)
    
    output.seek(0)
    return output


# =============================================================================
# STREAMLIT APP UI
# =============================================================================

def main():
    """Main Streamlit application."""
    
    # Initialize session state
    if 'processed' not in st.session_state:
        st.session_state.processed = False
    if 'df_final' not in st.session_state:
        st.session_state.df_final = None
    if 'last_pdf_name' not in st.session_state:
        st.session_state.last_pdf_name = None
    
    # Header
    st.title("Dental Statement Processor")
    st.markdown("Automate your patient balance tracking workflow")
    st.divider()
    
    # Main tabs for workflow
    main_tab1, main_tab2 = st.tabs(["Upload & Process", "Configuration"])
    
    # All columns extracted from PDF (raw)
    all_pdf_columns = ['Patient Name', 'Guarantor', 'CHART #', 'HOME PH #', 'WORK PH #', 
                       'EXT.', 'BT', 'LAST PATIENT PMT', 'LAST VISIT DATE', 'PEND. CLAIMS',
                       'FAMILY BALANCE', 'PATIENT BALANCE']
    
    # Initialize session state for PDF column mapping
    if 'pdf_column_mapping' not in st.session_state:
        # Default PDF columns to include (matches original config.py OUTPUT_COLUMNS)
        st.session_state.pdf_column_mapping = {
            'Patient Name': (True, 'Patient Name'),
            'Guarantor': (True, 'Guarantor'),
            'CHART #': (True, 'CHART #'),
            'HOME PH #': (False, 'HOME PH #'),
            'WORK PH #': (False, 'WORK PH #'),
            'EXT.': (False, 'EXT.'),
            'BT': (False, 'BT'),
            'LAST PATIENT PMT': (True, 'LAST PATIENT PMT'),
            'LAST VISIT DATE': (True, 'LAST VISIT DATE'),
            'PEND. CLAIMS': (True, 'PEND. CLAIMS'),
            'FAMILY BALANCE': (True, 'FAMILY BALANCE'),
            'PATIENT BALANCE': (True, 'PATIENT BALANCE')
        }
    
    # Initialize session state for old Excel column mapping (empty until file uploaded)
    if 'excel_column_mapping' not in st.session_state:
        st.session_state.excel_column_mapping = {}
    
    # Initialize detected columns and header row
    if 'detected_columns' not in st.session_state:
        st.session_state.detected_columns = []
    if 'detected_header_row' not in st.session_state:
        st.session_state.detected_header_row = 1
    if 'config_excel_file' not in st.session_state:
        st.session_state.config_excel_file = None
    
    with main_tab2:
        # Upload Old Excel Sheet section
        st.subheader("Previous Tracking Sheet")
        st.markdown("Upload your old Excel sheet to auto-detect tracking columns.")
        
        config_excel_file = st.file_uploader(
            "Upload old tracking sheet",
            type=['xlsx', 'xls'],
            help="Excel file with your manual tracking columns",
            key="config_excel_uploader"
        )
        
        # When file is uploaded, auto-detect header row and columns
        if config_excel_file is not None:
            # Check if this is a new file
            file_id = f"{config_excel_file.name}_{config_excel_file.size}"
            if st.session_state.get('last_config_excel_id') != file_id:
                # Reset file position for reading
                config_excel_file.seek(0)
                
                # Auto-detect header row
                detected_row = detect_header_row(config_excel_file)
                st.session_state.detected_header_row = detected_row
                
                # Reset file position again
                config_excel_file.seek(0)
                
                # Get tracking columns (columns after PATIENT BALANCE)
                detected_cols = get_tracking_columns(config_excel_file, detected_row)
                st.session_state.detected_columns = detected_cols
                
                # Initialize column mapping with all columns enabled
                st.session_state.excel_column_mapping = {
                    col: (True, col) for col in detected_cols
                }
                
                # Store file reference and ID
                st.session_state.config_excel_file = config_excel_file
                st.session_state.last_config_excel_id = file_id
                
                # Show detection results
                st.toast(f"✓ Detected {len(detected_cols)} tracking columns (header row: {detected_row})", icon="✅")
        
        # Show detected info
        if st.session_state.detected_columns:
            st.success(f"**Header row detected:** Row {st.session_state.detected_header_row}")
            st.info(f"**Found {len(st.session_state.detected_columns)} tracking columns** (columns after PATIENT BALANCE)")
        elif config_excel_file is not None:
            st.warning("No tracking columns found after PATIENT BALANCE column. Check that your Excel has the expected format.")
        
        st.divider()
        
        # Column Mapping section
        st.subheader("Column Selection")
        
        if not st.session_state.detected_columns and config_excel_file is None:
            st.caption("Upload an Excel file above to see available columns.")
        
        # Table header
        col1, col2, col3 = st.columns([0.5, 2.5, 1])
        with col1:
            st.markdown('<div class="table-header">Include</div>', unsafe_allow_html=True)
        with col2:
            st.markdown('<div class="table-header">Column Name</div>', unsafe_allow_html=True)
        with col3:
            st.markdown('<div class="table-header">Source</div>', unsafe_allow_html=True)
        
        # PDF columns section
        updated_pdf_mapping = {}
        for pdf_col in all_pdf_columns:
            current_include, _ = st.session_state.pdf_column_mapping.get(pdf_col, (True, pdf_col))
            
            col1, col2, col3 = st.columns([0.5, 2.5, 1])
            
            with col1:
                include = st.checkbox("", value=current_include, key=f"PDF_inc_{pdf_col}", label_visibility="collapsed")
            
            with col2:
                st.markdown(f"<span style='font-family: monospace; font-size: 0.9rem;'>{pdf_col}</span>", unsafe_allow_html=True)
            
            with col3:
                st.markdown('<span style="color: #2275b0; font-weight: 500;">PDF</span>', unsafe_allow_html=True)
            
            updated_pdf_mapping[pdf_col] = (include, pdf_col)
        
        st.session_state.pdf_column_mapping = updated_pdf_mapping
        
        # Divider between PDF and Excel columns
        if st.session_state.detected_columns:
            st.markdown('<div style="border-bottom: 2px solid #e9ecef; margin: 0.5rem 0;"></div>', unsafe_allow_html=True)
        
        # Excel columns section (from detected columns)
        updated_excel_mapping = {}
        for excel_col in st.session_state.detected_columns:
            current_include, _ = st.session_state.excel_column_mapping.get(excel_col, (True, excel_col))
            
            col1, col2, col3 = st.columns([0.5, 2.5, 1])
            
            with col1:
                include = st.checkbox("", value=current_include, key=f"Excel_inc_{excel_col}", label_visibility="collapsed")
            
            with col2:
                st.markdown(f"<span style='font-family: monospace; font-size: 0.9rem;'>{excel_col}</span>", unsafe_allow_html=True)
            
            with col3:
                st.markdown('<span style="color: #6fa976; font-weight: 500;">Excel</span>', unsafe_allow_html=True)
            
            updated_excel_mapping[excel_col] = (include, excel_col)
        
        st.session_state.excel_column_mapping = updated_excel_mapping
        
        st.divider()
        
        # Advanced Settings section
        st.subheader("Advanced")
        
        date_format = st.selectbox(
            "Date format",
            options=['%m/%d/%y', '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y'],
            index=0,
            help="Format for dates in the output Excel file",
            key="date_format"
        )
    
    with main_tab1:
        # PDF Upload section
        st.subheader("PDF Statement")
        pdf_file = st.file_uploader(
            "Upload patient balance report",
            type=['pdf'],
            help="The PDF export from your dental software"
        )
        
        # Show toast only once per file
        if pdf_file is not None:
            if st.session_state.last_pdf_name != pdf_file.name:
                file_size = len(pdf_file.getvalue()) / (1024 * 1024)
                st.toast(f"✓ {pdf_file.name} loaded ({file_size:.2f} MB)", icon="✅")
                st.session_state.last_pdf_name = pdf_file.name
        
        # Show Excel file status from Configuration tab
        excel_file = st.session_state.get('config_excel_file')
        if excel_file:
            st.success(f"✓ Previous tracking sheet loaded: **{excel_file.name}** ({len(st.session_state.detected_columns)} tracking columns)")
        else:
            st.info("💡 Upload your previous tracking sheet in the **Configuration** tab to merge old data.")
        
        st.divider()
        
        # Preview of output columns
        st.markdown("**Output Columns Preview**")
        
        # Get selected columns
        pdf_column_mapping = st.session_state.get('pdf_column_mapping', {})
        excel_column_mapping = st.session_state.get('excel_column_mapping', {})
        
        selected_pdf = [(col, name) for col, (inc, name) in pdf_column_mapping.items() if inc]
        selected_excel = [(col, name) for col, (inc, name) in excel_column_mapping.items() if inc]
        
        output_preview_columns = [name for _, name in selected_pdf] + [name for _, name in selected_excel]
        
        # Display as horizontally scrollable table
        if output_preview_columns:
            # Build table HTML
            table_cells = "".join([
                f'<th style="padding: 0.5rem 1rem; background-color: #f8f9fa; border-right: 1px solid #e9ecef; '
                f'white-space: nowrap; font-weight: 500; color: #2275b0; font-size: 0.9rem;">{col}</th>'
                for col in output_preview_columns
            ])
            
            st.markdown(
                f"""
                <div style="overflow-x: auto; border: 1px solid #e9ecef; border-radius: 6px;">
                    <table style="width: 100%; border-collapse: collapse; margin: 0;">
                        <thead>
                            <tr>
                                {table_cells}
                            </tr>
                        </thead>
                    </table>
                </div>
                """,
                unsafe_allow_html=True
            )
            st.caption(f"{len(output_preview_columns)} columns total • Configure in Configuration tab")
        else:
            st.caption("No columns selected • Configure in Configuration tab")
        
        st.divider()
        
        # Process button
        process_btn = st.button("Process Statements", type="primary", use_container_width=True)
        
        if process_btn:
            if not pdf_file:
                st.error("Please upload a PDF file to process")
                return
            
            st.session_state.processed = False
            
            # Auto-scroll to processing section
            components.html(
                """
                <script>
                    window.parent.document.querySelector('section.main').scrollTo({
                        top: window.parent.document.querySelector('section.main').scrollHeight,
                        behavior: 'smooth'
                    });
                </script>
                """,
                height=0,
            )
            
            # Get config values from session state
            pdf_column_mapping = st.session_state.get('pdf_column_mapping', {})
            excel_column_mapping = st.session_state.get('excel_column_mapping', {})
            # Use auto-detected header row (defaults to 1)
            old_sheet_start_row = st.session_state.get('detected_header_row', 1)
            date_format = st.session_state.get('date_format', '%m/%d/%y')
            
            # Build output columns from both mappings
            selected_pdf = [(col, name) for col, (inc, name) in pdf_column_mapping.items() if inc]
            selected_excel = [(col, name) for col, (inc, name) in excel_column_mapping.items() if inc]
            output_columns = [name for _, name in selected_pdf] + [name for _, name in selected_excel]
            
            # Processing workflow
            try:
                # Progress tracking
                st.markdown("### Processing")
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                # Step 1: Parse PDF
                status_text.text("Step 1/4: Parsing PDF...")
                progress_bar.progress(0.0)
                df_raw = parse_pdf_statements(pdf_file, progress_bar, status_text)
                
                # Step 2: Apply parsing rules and filter
                progress_bar.progress(0.4)
                status_text.text("Step 2/4: Processing data...")
                df_parsed = apply_parsing_rules(df_raw)
                df_filtered = filter_outstanding_balances(df_parsed)
                
                # Step 3: Load and merge old tracking sheet
                progress_bar.progress(0.6)
                status_text.text("Step 3/4: Merging tracking data...")
                df_old = pd.DataFrame()
                if excel_file:
                    excel_file.seek(0)
                    df_old = load_old_tracking_sheet(excel_file, old_sheet_start_row)
                df_merged = merge_with_tracking_data(df_filtered, df_old, excel_column_mapping)
                
                # Step 4: Prepare final output
                progress_bar.progress(0.8)
                status_text.text("Step 4/4: Preparing output...")
                df_final = prepare_final_output(df_merged, output_columns, date_format)
                
                # Complete
                progress_bar.progress(1.0)
                status_text.text("Processing complete")
                
                # Store in session state
                st.session_state.df_final = df_final
                st.session_state.processed = True
                
                # Auto-scroll to results after processing
                components.html(
                    """
                    <script>
                        setTimeout(function() {
                            window.parent.document.querySelector('section.main').scrollTo({
                                top: window.parent.document.querySelector('section.main').scrollHeight,
                                behavior: 'smooth'
                            });
                        }, 500);
                    </script>
                    """,
                    height=0,
                )
                
            except Exception as e:
                st.error(f"Error during processing: {str(e)}")
                st.exception(e)
                return
        
        # Display results if processed
        if st.session_state.processed and st.session_state.df_final is not None:
            df_final = st.session_state.df_final
            
            st.divider()
            st.markdown("### Results")
            
            # Financial summary
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("Final Output", f"{len(df_final):,}")
            
            with col2:
                total_patient = df_final['PATIENT BALANCE'].sum() if 'PATIENT BALANCE' in df_final.columns else 0
                st.metric("Patient Balance", f"${total_patient:,.2f}")
            
            with col3:
                total_family = df_final['FAMILY BALANCE'].sum() if 'FAMILY BALANCE' in df_final.columns else 0
                st.metric("Family Balance", f"${total_family:,.2f}")
            
            with col4:
                guarantors = len(df_final[df_final['Guarantor'] == 'Y']) if 'Guarantor' in df_final.columns else 0
                st.metric("Guarantors", f"{guarantors:,}")
            
            # Data preview
            st.markdown("### Preview")
            st.dataframe(df_final.head(20), use_container_width=True, height=400)
            st.caption(f"Showing first 20 of {len(df_final)} records")
            
            # Download button
            st.markdown("### Download")
            excel_buffer = create_excel_download(df_final)
            
            st.download_button(
                label="Download Excel File",
                data=excel_buffer,
                file_name=f"Merged_Statements_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary"
            )


if __name__ == '__main__':
    main()

