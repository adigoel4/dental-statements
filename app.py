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
    
    /* Column mapping table */
    .column-row {
        padding: 0.4rem 0;
        border-bottom: 1px solid #f0f0f0;
        display: flex;
        align-items: center;
    }
    
    .column-row:last-child {
        border-bottom: none;
    }
    
    .column-row:hover {
        background-color: #f8f9fa;
        border-radius: 4px;
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
    
    /* Source tags */
    .source-tag {
        display: inline-block;
        padding: 0.2rem 0.6rem;
        border-radius: 12px;
        font-size: 0.75rem;
        font-weight: 500;
    }
    
    .source-pdf {
        background-color: #e3f2fd;
        color: #2275b0;
    }
    
    .source-excel {
        background-color: #e8f5e9;
        color: #2d5f32;
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
    """Merge new statement data with old tracking data using flexible column mapping."""
    # Normalize CHART # in new data
    df_new = df_new.copy()
    df_new.loc[:, 'CHART #'] = df_new['CHART #'].apply(normalize_chart_number)
    
    if not df_old.empty:
        # Get columns to include from old sheet
        old_cols_to_include = [old_col for old_col, (include, _) in excel_column_mapping.items() if include]
        
        # Select tracking columns that exist in old sheet
        tracking_cols = ['CHART #'] + [col for col in old_cols_to_include if col in df_old.columns]
        df_old_tracking = df_old[tracking_cols]
        
        # Remove duplicates
        df_old_tracking = df_old_tracking.drop_duplicates(subset=['CHART #'], keep='last')
        
        # LEFT JOIN
        df_merged = df_new.merge(df_old_tracking, on='CHART #', how='left', suffixes=('', '_old'))
        
        # Rename columns according to mapping
        rename_dict = {}
        for old_col, (include, new_col) in excel_column_mapping.items():
            if include and old_col in df_merged.columns and old_col != new_col:
                rename_dict[old_col] = new_col
        
        if rename_dict:
            df_merged = df_merged.rename(columns=rename_dict)
        
        return df_merged
    else:
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
    if 'pdf_toast_shown' not in st.session_state:
        st.session_state.pdf_toast_shown = False
    if 'excel_toast_shown' not in st.session_state:
        st.session_state.excel_toast_shown = False
    if 'last_pdf_name' not in st.session_state:
        st.session_state.last_pdf_name = None
    if 'last_excel_name' not in st.session_state:
        st.session_state.last_excel_name = None
    
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
    
    # Initialize session state for old Excel column mapping
    if 'excel_column_mapping' not in st.session_state:
        # Default mapping for old Excel tracking columns (matches original config.py TRACKING_COLUMNS)
        st.session_state.excel_column_mapping = {
            'STATUS': (True, 'STATUS'),
            'NOTES': (True, 'NOTES'),
            'Follow-Up Date': (True, 'Follow-Up Date'),
            'Staff Code': (True, 'Staff Code'),
            'Ortho': (True, 'Ortho'),
            'BALANCE CODE': (True, 'BALANCE CODE'),
            'ST1 DATE': (True, 'ST1 DATE'),
            'AMOUNT1': (True, 'AMOUNT1'),
            'ST2 DATE': (True, 'ST2 DATE'),
            'AMOUNT2': (True, 'AMOUNT2'),
            'ST3 DATE': (True, 'ST3 DATE'),
            'AMOUNT3': (True, 'AMOUNT3')
        }
    
    with main_tab2:
        # Column Mapping section
        st.subheader("Column Mapping")
        
        # Input for old Excel columns
        st.markdown("Add columns from old tracking sheet (To be merged with new data):")
        old_columns_text = st.text_area(
            "Old sheet columns",
            value="\n".join(st.session_state.excel_column_mapping.keys()),
            height=120,
            help="Enter column names from your old Excel sheet (one per line)",
            label_visibility="collapsed"
        )
        
        old_columns = [col.strip() for col in old_columns_text.split('\n') if col.strip()]
        
        # Update session state
        current_cols = set(st.session_state.excel_column_mapping.keys())
        new_cols = set(old_columns)
        
        for col in new_cols - current_cols:
            st.session_state.excel_column_mapping[col] = (True, col)
        
        for col in current_cols - new_cols:
            del st.session_state.excel_column_mapping[col]
        
        st.markdown("")
        
        # Table header
        col1, col2, col3, col4 = st.columns([0.6, 1.8, 1.8, 0.8])
        with col1:
            st.markdown("")  # No header for checkbox
        with col2:
            st.markdown('<div class="table-header">Source Column</div>', unsafe_allow_html=True)
        with col3:
            st.markdown('<div class="table-header">Output Name</div>', unsafe_allow_html=True)
        with col4:
            st.markdown('<div class="table-header">From</div>', unsafe_allow_html=True)
        
        # All columns in one unified list
        all_columns = []
        
        # Add PDF columns
        updated_pdf_mapping = {}
        for pdf_col in all_pdf_columns:
            current_include, current_new_name = st.session_state.pdf_column_mapping.get(pdf_col, (True, pdf_col))
            all_columns.append(('PDF', pdf_col, current_include, current_new_name))
        
        # Add Excel columns
        updated_excel_mapping = {}
        for excel_col in old_columns:
            current_include, current_new_name = st.session_state.excel_column_mapping.get(excel_col, (True, excel_col))
            all_columns.append(('Excel', excel_col, current_include, current_new_name))
        
        # Render unified list with styling
        for idx, (source, col_name, current_include, current_new_name) in enumerate(all_columns):
            col1, col2, col3, col4 = st.columns([0.6, 1.8, 1.8, 0.8])
            
            with col1:
                include = st.checkbox("", value=current_include, key=f"{source}_inc_{col_name}", label_visibility="collapsed")
            
            with col2:
                st.markdown(f"<span style='font-family: monospace; font-size: 0.9rem;'>{col_name}</span>", unsafe_allow_html=True)
            
            with col3:
                if include:
                    new_name = st.text_input(
                        "New name",
                        value=current_new_name,
                        key=f"{source}_name_{col_name}",
                        label_visibility="collapsed",
                        placeholder=col_name
                    )
                else:
                    new_name = col_name
                    st.markdown(f"<span style='color: #ccc;'>—</span>", unsafe_allow_html=True)
            
            with col4:
                if source == 'PDF':
                    st.markdown('<span style="color: #2275b0; font-weight: 500;">PDF</span>', unsafe_allow_html=True)
                else:
                    st.markdown('<span style="color: #6fa976; font-weight: 500;">Excel</span>', unsafe_allow_html=True)
            
            # Update mappings
            if source == 'PDF':
                updated_pdf_mapping[col_name] = (include, new_name if new_name else col_name)
            else:
                updated_excel_mapping[col_name] = (include, new_name if new_name else col_name)
            
            # Add subtle separator except for last row
            if idx < len(all_columns) - 1:
                st.markdown('<div style="border-bottom: 1px solid #f5f5f5; margin: 0.2rem 0;"></div>', unsafe_allow_html=True)
        
        st.session_state.pdf_column_mapping = updated_pdf_mapping
        st.session_state.excel_column_mapping = updated_excel_mapping
        
        st.divider()
        
        # Advanced Settings section
        st.subheader("Advanced")
        
        col1, col2 = st.columns(2)
        
        with col1:
            old_sheet_start_row = st.number_input(
                "Header row number in old sheet",
                min_value=1,
                value=st.session_state.get('old_sheet_start_row', 3),
                help="Row number where column headers are located (1-indexed)",
                key="old_sheet_start_row"
            )
        
        with col2:
            date_format = st.selectbox(
                "Date format",
                options=['%m/%d/%y', '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y'],
                index=0,
                help="Format for dates in the output Excel file",
                key="date_format"
            )
    
    with main_tab1:
        # Two-column layout for file uploads
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("PDF Statement")
            pdf_file = st.file_uploader(
                "Upload patient balance report",
                type=['pdf'],
                help="The PDF export from your dental software"
            )
            
            # Show toast only once per file
            if pdf_file is not None:
                if st.session_state.last_pdf_name != pdf_file.name:
                    file_size = len(pdf_file.getvalue()) / (1024 * 1024)  # Convert to MB
                    st.toast(f"✓ {pdf_file.name} loaded ({file_size:.2f} MB)", icon="✅")
                    st.session_state.last_pdf_name = pdf_file.name
                    st.session_state.pdf_toast_shown = True
        
        with col2:
            st.subheader("Previous Tracking Sheet")
            excel_file = st.file_uploader(
                "Upload last month's tracking data (optional)",
                type=['xlsx', 'xls'],
                help="Excel file with your manual tracking columns"
            )
            
            # Show toast only once per file
            if excel_file is not None:
                if st.session_state.last_excel_name != excel_file.name:
                    file_size = len(excel_file.getvalue()) / (1024 * 1024)  # Convert to MB
                    st.toast(f"✓ {excel_file.name} loaded ({file_size:.2f} MB)", icon="✅")
                    st.session_state.last_excel_name = excel_file.name
                    st.session_state.excel_toast_shown = True
        
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
            old_sheet_start_row = st.session_state.get('old_sheet_start_row', 3)
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
                status_text.text("Step 1/5: Parsing PDF...")
                progress_bar.progress(0.0)
                df_raw = parse_pdf_statements(pdf_file, progress_bar, status_text)
                total_pdf_records = len(df_raw)
                
                # Step 2: Apply parsing rules
                progress_bar.progress(0.33)
                status_text.text("Step 2/5: Processing data...")
                df_parsed = apply_parsing_rules(df_raw)
                
                # Step 3: Filter outstanding balances
                progress_bar.progress(0.5)
                status_text.text("Step 3/5: Filtering outstanding balances...")
                df_filtered = filter_outstanding_balances(df_parsed)
                outstanding_records = len(df_filtered)
                
                # Step 4: Load old tracking sheet
                progress_bar.progress(0.66)
                status_text.text("Step 4/5: Loading tracking data...")
                df_old = pd.DataFrame()
                old_records_count = 0
                if excel_file:
                    df_old = load_old_tracking_sheet(excel_file, old_sheet_start_row)
                    old_records_count = len(df_old)
                
                # Step 5: Merge and prepare output
                progress_bar.progress(0.83)
                status_text.text("Step 5/5: Merging data...")
                df_merged = merge_with_tracking_data(df_filtered, df_old, excel_column_mapping)
                
                # Count matched records
                matched_records = 0
                if not df_old.empty and 'STATUS' in df_merged.columns:
                    matched_records = df_merged['STATUS'].notna().sum()
                
                # Apply PDF column mapping (filter and rename)
                pdf_rename_dict = {}
                for pdf_col, (include, new_name) in pdf_column_mapping.items():
                    if pdf_col in df_merged.columns and pdf_col != new_name:
                        pdf_rename_dict[pdf_col] = new_name
                
                if pdf_rename_dict:
                    df_merged = df_merged.rename(columns=pdf_rename_dict)
                
                df_final = prepare_final_output(df_merged, output_columns, date_format)
                
                # Complete
                progress_bar.progress(1.0)
                status_text.text("Processing complete")
                
                # Store in session state with stats
                st.session_state.df_final = df_final
                st.session_state.processed = True
                st.session_state.processing_stats = {
                    'total_pdf_records': total_pdf_records,
                    'outstanding_records': outstanding_records,
                    'old_records_count': old_records_count,
                    'matched_records': matched_records,
                    'final_records': len(df_final)
                }
                
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
            stats = st.session_state.get('processing_stats', {})
            
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



