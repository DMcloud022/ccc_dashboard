import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
import time
import openpyxl
import numpy as np
import re
from dateutil import parser as date_parser

# Column mapping
COLUMN_LABELS = {
    'A': 'ID', 'B': 'Status', 'C': 'Assigned To', 'D': 'Delivery Channels',
    'E': 'Date of Complaint', 'F': 'Resolution', 'G': 'Customer Name',
    'H': 'Customer Email', 'I': 'Location (if NTC)', 'J': 'Narrative',
    'K': 'Service Providers', 'L': 'Complaint Category', 'M': 'Platform',
    'N': 'Complaint Nature', 'O': 'Priority', 'P': 'Date of Resolution',
    'Q': 'Customer Contact No.', 'R': 'Response Time (Hours)',
    'S': 'Resolution Time (Hours)', 'T': 'Agency', 'U': 'DICT UNIT',
    'V': 'Date Responded', 'X': 'CIMS'
}

# Page configuration
st.set_page_config(
    page_title="Complaint Analysis Dashboard - Real-time",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better presentation
st.markdown("""
    <style>
    .metric-card {
        background-color: #f0f2f6;
        padding: 20px;
        border-radius: 10px;
        text-align: center;
    }
    .stPlotlyChart {
        background-color: white;
        border-radius: 10px;
        padding: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .live-indicator {
        display: inline-block;
        width: 10px;
        height: 10px;
        background-color: #00ff00;
        border-radius: 50%;
        animation: pulse 2s infinite;
    }
    @keyframes pulse {
        0% { opacity: 1; }
        50% { opacity: 0.5; }
        100% { opacity: 1; }
    }
    </style>
    """, unsafe_allow_html=True)

def extract_spreadsheet_id(url):
    """Extract spreadsheet ID from various Google Sheets URL formats"""
    # Pattern for spreadsheet ID
    patterns = [
        r'/spreadsheets/d/([a-zA-Z0-9-_]+)',
        r'id=([a-zA-Z0-9-_]+)',
    ]
    
    for pattern in patterns:
        match = re.search(pattern, url)
        if match:
            return match.group(1)
    
    # If no pattern matches, assume the URL itself might be the ID
    if re.match(r'^[a-zA-Z0-9-_]+$', url):
        return url
    
    return None

@st.cache_data(ttl=60)
def load_data_from_public_gsheet(spreadsheet_url, timestamp):
    """Load data from public Google Sheets without authentication"""
    try:
        spreadsheet_id = extract_spreadsheet_id(spreadsheet_url)
        
        if not spreadsheet_id:
            st.error("❌ Invalid Google Sheets URL. Please check the format.")
            return None
        
        # Construct the CSV export URL for public sheets
        csv_export_url = f'https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=csv'
        
        # Try to read as CSV
        try:
            df = pd.read_csv(csv_export_url)
            
            # Remove completely empty rows
            df = df.replace('', np.nan)
            df = df.dropna(how='all')
            
            # Clean column names
            df.columns = df.columns.str.strip()
            
            if len(df) == 0:
                st.warning("⚠️ Sheet appears to be empty")
                return None
            
            st.success(f"✅ Successfully loaded {len(df)} rows from public Google Sheets")
            return df
            
        except Exception as e:
            st.error(f"❌ Cannot access sheet. Please ensure it's set to 'Anyone with the link can view'")
            st.info("To make your sheet public: File → Share → Change to 'Anyone with the link'")
            return None
            
    except Exception as e:
        st.error(f"❌ Error loading public sheet: {str(e)}")
        return None

@st.cache_data(ttl=60)
def load_data_from_gsheet_with_auth(credentials_dict, spreadsheet_url, timestamp):
    """Load data from Google Sheets with service account authentication"""
    try:
        scopes = [
            'https://www.googleapis.com/auth/spreadsheets.readonly',
            'https://www.googleapis.com/auth/drive.readonly'
        ]
        creds = Credentials.from_service_account_info(credentials_dict, scopes=scopes)
        client = gspread.authorize(creds)
        
        # Open spreadsheet by URL or ID
        spreadsheet_id = extract_spreadsheet_id(spreadsheet_url)
        if spreadsheet_id:
            spreadsheet = client.open_by_key(spreadsheet_id)
        else:
            spreadsheet = client.open_by_url(spreadsheet_url)
        
        # Get the first sheet
        sheet = spreadsheet.sheet1
        
        # Get all values including headers
        all_values = sheet.get_all_values()
        
        if not all_values or len(all_values) < 2:
            st.warning("⚠️ Sheet is empty or has no data rows")
            return None
        
        # Create DataFrame with first row as headers
        headers = all_values[0]
        data_rows = all_values[1:]
        
        df = pd.DataFrame(data_rows, columns=headers)
        
        # Remove completely empty rows
        df = df.replace('', np.nan)
        df = df.dropna(how='all')
        
        # Clean column names (strip whitespace)
        df.columns = df.columns.str.strip()
        
        st.success(f"✅ Successfully loaded {len(df)} rows from Google Sheets (authenticated)")
        
        return df
        
    except gspread.exceptions.SpreadsheetNotFound:
        st.error("❌ Spreadsheet not found. Please check the URL and ensure the service account has access.")
        return None
    except gspread.exceptions.APIError as e:
        st.error(f"❌ Google Sheets API Error: {str(e)}")
        return None
    except Exception as e:
        st.error(f"❌ Error loading data from Google Sheets: {str(e)}")
        return None

def load_data_from_excel(file_path):
    """Load data from Excel file with error handling"""
    try:
        df = pd.read_excel(file_path, engine='openpyxl')
        
        # Remove completely empty rows
        df = df.replace('', np.nan)
        df = df.dropna(how='all')
        
        # Clean column names
        df.columns = df.columns.str.strip()
        
        return df
    except FileNotFoundError:
        st.error(f"❌ File not found: {file_path}")
        return None
    except PermissionError:
        st.error(f"❌ Permission denied. File may be open in another program: {file_path}")
        return None
    except Exception as e:
        st.error(f"❌ Error loading Excel: {str(e)}")
        return None

def load_data_from_uploaded_excel(uploaded_file):
    """Load data from uploaded Excel file"""
    try:
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        
        # Remove completely empty rows
        df = df.replace('', np.nan)
        df = df.dropna(how='all')
        
        # Clean column names
        df.columns = df.columns.str.strip()
        
        return df
    except Exception as e:
        st.error(f"❌ Error loading Excel: {str(e)}")
        return None

def find_similar_columns(df, target_column, threshold=0.7):
    """Find similar column names using fuzzy matching"""
    from difflib import SequenceMatcher

    similar = []
    for col in df.columns:
        ratio = SequenceMatcher(None, target_column.lower(), col.lower()).ratio()
        if ratio > threshold:
            similar.append((col, ratio))

    similar.sort(key=lambda x: x[1], reverse=True)
    return [col for col, ratio in similar]

def validate_required_columns(df):
    """Validate and diagnose column issues"""
    if df is None or df.empty:
        return None, []

    required_columns = {
        'Date of Complaint': ['Date of Complaint', 'Complaint Date', 'Date Filed', 'Filing Date'],
        'Complaint Category': ['Complaint Category', 'Category', 'Type', 'Complaint Type'],
        'Complaint Nature': ['Complaint Nature', 'Nature', 'Nature of Complaint'],
        'Service Providers': ['Service Providers', 'Service Provider', 'Provider', 'ISP'],
        'Agency': ['Agency', 'Department', 'Office']
    }

    column_mapping = {}
    missing_columns = []
    suggestions = {}

    for required, alternatives in required_columns.items():
        found = False
        for alt in alternatives:
            if alt in df.columns:
                column_mapping[required] = alt
                found = True
                break

        if not found:
            # Try fuzzy matching
            similar = find_similar_columns(df, required, threshold=0.6)
            if similar:
                suggestions[required] = similar[:3]  # Top 3 suggestions
            missing_columns.append(required)

    return column_mapping, missing_columns, suggestions

def apply_column_mapping(df, column_mapping):
    """Apply column mapping to rename columns to standard names"""
    if not column_mapping:
        return df

    df = df.copy()
    rename_dict = {v: k for k, v in column_mapping.items() if v in df.columns}
    df = df.rename(columns=rename_dict)
    return df

def parse_date_robust(date_value):
    """Robustly parse various date formats"""
    if pd.isna(date_value) or date_value == '' or date_value is None:
        return pd.NaT

    # If already a datetime, return it
    if isinstance(date_value, (pd.Timestamp, datetime)):
        return pd.Timestamp(date_value)

    # Convert to string
    date_str = str(date_value).strip()

    if not date_str or date_str.lower() in ['nan', 'none', 'nat', '']:
        return pd.NaT

    try:
        # Try pandas default parser first
        return pd.to_datetime(date_str, errors='coerce')
    except:
        pass

    try:
        # Try dateutil parser for more flexibility
        return pd.Timestamp(date_parser.parse(date_str, fuzzy=True))
    except:
        pass

    # Common date formats to try
    date_formats = [
        '%Y-%m-%d',
        '%m/%d/%Y',
        '%d/%m/%Y',
        '%Y/%m/%d',
        '%m-%d-%Y',
        '%d-%m-%Y',
        '%B %d, %Y',
        '%b %d, %Y',
        '%d %B %Y',
        '%d %b %Y',
        '%Y%m%d',
        '%m/%d/%y',
        '%d/%m/%y',
    ]

    for fmt in date_formats:
        try:
            return pd.Timestamp(datetime.strptime(date_str, fmt))
        except:
            continue

    return pd.NaT

def prepare_data(df):
    """Prepare and clean the data with robust validation and date parsing"""
    if df is None or df.empty:
        return None

    # Make a copy to avoid modifying original
    df = df.copy()

    # Validate and map columns
    st.info("🔍 Validating column structure...")
    column_mapping, missing_columns, suggestions = validate_required_columns(df)

    # Apply column mapping if found
    if column_mapping:
        df = apply_column_mapping(df, column_mapping)
        mapped_cols = [f"{v} → {k}" for k, v in column_mapping.items()]
        st.success(f"✅ Mapped columns: {', '.join(mapped_cols)}")

    # Show warnings for missing columns with suggestions
    if missing_columns:
        st.warning(f"⚠️ Missing columns: {', '.join(missing_columns)}")
        if suggestions:
            st.info("💡 **Possible column name matches:**")
            for col, similar in suggestions.items():
                if similar:
                    st.write(f"  • For '{col}', did you mean: {', '.join(similar)}?")

    # Convert date columns to datetime with robust parsing
    date_columns = ['Date of Complaint', 'Date of Resolution', 'Date Responded']

    for col in date_columns:
        if col in df.columns:
            st.info(f"📅 Processing {col}...")

            # Apply robust date parsing
            df[col] = df[col].apply(parse_date_robust)

            # Count successful and failed conversions
            valid_dates = df[col].notna().sum()
            total_rows = len(df)

            if valid_dates > 0:
                st.success(f"✅ {col}: {valid_dates}/{total_rows} dates parsed successfully")
            else:
                st.warning(f"⚠️ {col}: No valid dates found")

    # Extract year and month for filtering
    if 'Date of Complaint' in df.columns:
        df['Year'] = df['Date of Complaint'].dt.year
        df['Month'] = df['Date of Complaint'].dt.month

        # Show date range
        valid_dates = df['Date of Complaint'].dropna()
        if len(valid_dates) > 0:
            min_date = valid_dates.min()
            max_date = valid_dates.max()
            st.info(f"📊 Date range: {min_date.strftime('%Y-%m-%d')} to {max_date.strftime('%Y-%m-%d')}")

    # Clean text columns (remove extra whitespace)
    text_columns = ['Agency', 'Service Providers', 'Complaint Category', 'Complaint Nature']
    for col in text_columns:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace(['nan', 'None', '', 'NaN', 'NaT'], np.nan)

    # Remove rows where Date of Complaint is invalid
    if 'Date of Complaint' in df.columns:
        rows_before = len(df)
        df = df[df['Date of Complaint'].notna()]
        rows_after = len(df)

        if rows_before > rows_after:
            st.warning(f"⚠️ Removed {rows_before - rows_after} rows with invalid complaint dates")

    return df

def filter_by_date(df, start_month, start_year=None):
    """Filter dataframe by date range"""
    if df is None or df.empty:
        return pd.DataFrame()
    
    if 'Date of Complaint' not in df.columns:
        return df
    
    if start_year is None:
        start_year = datetime.now().year
    
    try:
        start_date = pd.Timestamp(year=start_year, month=start_month, day=1)
        filtered_df = df[df['Date of Complaint'] >= start_date].copy()
        return filtered_df
    except Exception as e:
        st.error(f"Error filtering by date: {str(e)}")
        return df

def create_bar_chart(data_series, title, color_scale='Blues', height=400):
    """Create a bar chart with proper data handling"""
    if data_series is None or len(data_series) == 0:
        # Return empty figure with message
        fig = go.Figure()
        fig.add_annotation(
            text="No data available",
            xref="paper", yref="paper",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=16)
        )
        fig.update_layout(height=height)
        return fig
    
    # Convert to DataFrame for plotly
    df_plot = pd.DataFrame({
        'Category': data_series.index.astype(str),
        'Count': data_series.values
    })
    
    fig = px.bar(
        df_plot,
        x='Count',
        y='Category',
        orientation='h',
        labels={'Count': 'Count', 'Category': 'Category'},
        color='Count',
        color_continuous_scale=color_scale
    )
    fig.update_layout(
        height=height,
        showlegend=False,
        yaxis={'categoryorder': 'total ascending'},
        margin=dict(l=10, r=10, t=30, b=10)
    )
    return fig

def create_pie_chart(data_series, title, color_scheme=None, height=400):
    """Create a pie chart with proper data handling"""
    if data_series is None or len(data_series) == 0:
        # Return empty figure with message
        fig = go.Figure()
        fig.add_annotation(
            text="No data available",
            xref="paper", yref="paper",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=16)
        )
        fig.update_layout(height=height)
        return fig
    
    # Convert to DataFrame for plotly
    df_plot = pd.DataFrame({
        'Category': data_series.index.astype(str),
        'Count': data_series.values
    })
    
    fig = px.pie(
        df_plot,
        values='Count',
        names='Category',
        hole=0.4,
        color_discrete_sequence=color_scheme
    )
    fig.update_layout(
        height=height,
        margin=dict(l=10, r=10, t=30, b=10)
    )
    return fig

def main():
    # Initialize session state for auto-refresh
    if 'auto_refresh' not in st.session_state:
        st.session_state.auto_refresh = False
    if 'refresh_interval' not in st.session_state:
        st.session_state.refresh_interval = 60
    if 'data_source_type' not in st.session_state:
        st.session_state.data_source_type = None
    if 'file_path' not in st.session_state:
        st.session_state.file_path = None
    if 'gsheet_creds' not in st.session_state:
        st.session_state.gsheet_creds = None
    if 'gsheet_url' not in st.session_state:
        st.session_state.gsheet_url = None
    if 'use_public_sheet' not in st.session_state:
        st.session_state.use_public_sheet = True
    
    # Header with live indicator
    col1, col2 = st.columns([6, 1])
    with col1:
        st.title("📊 Complaint Analysis Dashboard - Real-time")
    with col2:
        if st.session_state.auto_refresh:
            st.markdown('<div class="live-indicator"></div> <span>LIVE</span>', unsafe_allow_html=True)
    
    st.markdown("---")
    
    # Sidebar for data loading and settings
    with st.sidebar:
        st.header("🔧 Data Source")
        data_source = st.radio("Choose data source:", ["Google Sheets (Public)", "Upload Excel", "Excel File Path", "Google Sheets (Private)"])
        
        df = None
        
        if data_source == "Google Sheets (Public)":
            st.info("📋 For public sheets - easiest option!")
            st.markdown("""
            **How to make your sheet public:**
            1. Open your Google Sheet
            2. Click **Share** button
            3. Change to **"Anyone with the link"**
            4. Set permission to **Viewer**
            5. Copy the link and paste below
            """)
            
            spreadsheet_url = st.text_input(
                "Google Sheet URL", 
                value=st.session_state.gsheet_url or "",
                placeholder="https://docs.google.com/spreadsheets/d/..."
            )
            
            if spreadsheet_url:
                st.session_state.gsheet_url = spreadsheet_url
                
                # Use timestamp for cache busting in real-time mode
                if st.session_state.auto_refresh:
                    current_time = datetime.now().timestamp()
                else:
                    current_time = 0
                
                df = load_data_from_public_gsheet(spreadsheet_url, current_time)
                st.session_state.data_source_type = "gsheets_public"
        
        elif data_source == "Upload Excel":
            uploaded_file = st.file_uploader("Upload your Excel file (.xlsx)", type=['xlsx', 'xls'])
            if uploaded_file is not None:
                df = load_data_from_uploaded_excel(uploaded_file)
                st.session_state.data_source_type = "upload"
        
        elif data_source == "Excel File Path":
            file_path = st.text_input("Enter full path to Excel file:", 
                                     value=st.session_state.file_path or "",
                                     placeholder="e.g., C:/Data/complaints.xlsx")
            if file_path:
                st.session_state.file_path = file_path
                df = load_data_from_excel(file_path)
                st.session_state.data_source_type = "filepath"
        
        else:  # Google Sheets (Private)
            st.info("📋 For private sheets - requires authentication")
            st.markdown("""
            **Setup Steps:**
            1. Create service account in Google Cloud
            2. Enable Google Sheets API
            3. Share sheet with service account email
            4. Upload credentials JSON below
            """)
            
            # Check if credentials already in session
            if st.session_state.gsheet_creds is None:
                credentials_file = st.file_uploader("Upload Service Account JSON", type=['json'])
                if credentials_file:
                    import json
                    st.session_state.gsheet_creds = json.load(credentials_file)
                    st.success("✅ Credentials loaded")
            else:
                st.success("✅ Credentials active")
                if st.button("Clear Credentials"):
                    st.session_state.gsheet_creds = None
                    st.rerun()
            
            spreadsheet_url = st.text_input(
                "Google Sheet URL", 
                value=st.session_state.gsheet_url or "",
                placeholder="https://docs.google.com/spreadsheets/d/..."
            )
            
            if spreadsheet_url:
                st.session_state.gsheet_url = spreadsheet_url
            
            if st.session_state.gsheet_creds and st.session_state.gsheet_url:
                # Use timestamp for cache busting
                if st.session_state.auto_refresh:
                    current_time = datetime.now().timestamp()
                else:
                    current_time = 0
                
                df = load_data_from_gsheet_with_auth(
                    st.session_state.gsheet_creds, 
                    st.session_state.gsheet_url,
                    current_time
                )
                st.session_state.data_source_type = "gsheets_private"
        
        if df is not None and not df.empty:
            st.success(f"✅ Data loaded: {len(df)} records")
            st.caption(f"Last updated: {datetime.now().strftime('%H:%M:%S')}")

            # Show column info and data preview
            with st.expander("📋 Column Information & Data Preview"):
                st.write(f"**Total columns:** {len(df.columns)}")
                st.write("**Available columns:**")
                for col in df.columns:
                    st.text(f"  • {col}")

                st.markdown("---")
                st.write("**Data Preview (First 5 rows):**")
                st.dataframe(df.head(), use_container_width=True)

                # Show data types
                st.markdown("---")
                st.write("**Column Data Types:**")
                dtype_df = pd.DataFrame({
                    'Column': df.dtypes.index,
                    'Type': df.dtypes.values.astype(str)
                })
                st.dataframe(dtype_df, use_container_width=True)
            
            st.markdown("---")
            
            # Auto-refresh settings
            st.header("🔄 Real-time Settings")
            auto_refresh = st.checkbox("Enable Auto-refresh", value=st.session_state.auto_refresh)
            st.session_state.auto_refresh = auto_refresh
            
            if auto_refresh:
                refresh_interval = st.selectbox(
                    "Refresh interval (seconds):",
                    [30, 60, 120, 300],
                    index=1
                )
                st.session_state.refresh_interval = refresh_interval
                
                # Manual refresh button
                if st.button("🔄 Refresh Now"):
                    st.cache_data.clear()
                    st.rerun()
                
                st.info(f"⏱️ Auto-refreshing every {refresh_interval}s")
            
            st.markdown("---")
            
            # Date range info
            st.header("📅 Analysis Period")
            current_year = datetime.now().year
            st.info(f"Current Year: {current_year}")
        elif df is not None:
            st.warning("⚠️ Data loaded but empty")
    
    if df is None or df.empty:
        st.info("👈 Please load data from the sidebar to begin analysis")
        st.markdown("""
        ### 📖 Quick Start Guide:
        
        #### 🌐 Option 1: Google Sheets (Public) - **RECOMMENDED**
        - **Easiest for presentations!**
        - Open your Google Sheet
        - Click Share → "Anyone with the link can view"
        - Copy the URL and paste in sidebar
        - ✅ No authentication needed!
        
        #### 📤 Option 2: Upload Excel
        - Export your Google Sheet as .xlsx
        - Upload file directly
        
        #### 📁 Option 3: Excel File Path
        - For real-time monitoring from local file
        - Enter full path to .xlsx file
        
        #### 🔒 Option 4: Google Sheets (Private)
        - For private/restricted sheets
        - Requires service account setup
        
        ### 📊 Required Columns:
        - **Date of Complaint** (Required)
        - Complaint Category
        - Complaint Nature
        - Service Providers
        - Agency
        """)
        return
    
    # Prepare data with progress
    with st.spinner("🔄 Processing data and parsing dates..."):
        df = prepare_data(df)
    
    if df is None or df.empty:
        st.error("❌ No valid data after processing. Please check your date formats.")
        return
    
    # Get current year
    current_year = datetime.now().year
    
    # Filter datasets with validation
    df_jan_present = df[df['Year'] == current_year] if 'Year' in df.columns else df
    df_nov_present = filter_by_date(df, 11, current_year)
    df_sep_present = filter_by_date(df, 9, current_year)
    
    # Summary metrics at the top
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Complaints (Jan-Present)", len(df_jan_present))
    with col2:
        st.metric("Total Complaints (Nov-Present)", len(df_nov_present))
    with col3:
        if 'Agency' in df_jan_present.columns:
            ntc_count = len(df_jan_present[df_jan_present['Agency'].str.contains('NTC', na=False, case=False)])
            st.metric("NTC Complaints", ntc_count)
        else:
            st.metric("NTC Complaints", "N/A")
    with col4:
        if 'Agency' in df_jan_present.columns:
            pemedes_count = len(df_jan_present[df_jan_present['Agency'].str.contains('PEMEDES', na=False, case=False)])
            st.metric("PEMEDES Complaints", pemedes_count)
        else:
            st.metric("PEMEDES Complaints", "N/A")
    
    st.markdown("---")
    
    # Tab layout for organized viewing
    tab1, tab2, tab3 = st.tabs(["📈 Overall Analysis", "🏢 NTC Analysis", "🏢 PEMEDES Analysis"])
    
    with tab1:
        st.header("Overall Complaint Analysis")

        # Row 1: Complaints by Category
        col1, col2 = st.columns(2)

        with col1:
            st.subheader("Complaints by Category (Jan-Present)")
            if 'Complaint Category' in df_jan_present.columns:
                # Filter out null/empty values
                valid_data = df_jan_present['Complaint Category'].dropna()
                valid_data = valid_data[valid_data != '']

                if len(valid_data) > 0:
                    category_counts = valid_data.value_counts().head(10)
                    fig = create_bar_chart(category_counts, "Category", 'Blues')
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info("📊 No category data available in this period")
            else:
                fig = create_bar_chart(pd.Series(), "Category", 'Blues')
                st.plotly_chart(fig, use_container_width=True)
                st.error("❌ 'Complaint Category' column not found. Please check your data structure.")

        with col2:
            st.subheader("Complaints by Category (Nov-Present)")
            if 'Complaint Category' in df_nov_present.columns:
                valid_data = df_nov_present['Complaint Category'].dropna()
                valid_data = valid_data[valid_data != '']

                if len(valid_data) > 0:
                    category_counts = valid_data.value_counts().head(10)
                    fig = create_bar_chart(category_counts, "Category", 'Oranges')
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info("📊 No category data available in this period")
            else:
                fig = create_bar_chart(pd.Series(), "Category", 'Oranges')
                st.plotly_chart(fig, use_container_width=True)
                st.error("❌ 'Complaint Category' column not found. Please check your data structure.")

        # Row 2: Complaints by Nature
        col1, col2 = st.columns(2)

        with col1:
            st.subheader("Complaints by Nature (Jan-Present)")
            if 'Complaint Nature' in df_jan_present.columns:
                valid_data = df_jan_present['Complaint Nature'].dropna()
                valid_data = valid_data[valid_data != '']

                if len(valid_data) > 0:
                    nature_counts = valid_data.value_counts().head(10)
                    fig = create_pie_chart(nature_counts, "Nature")
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info("📊 No nature data available in this period")
            else:
                fig = create_pie_chart(pd.Series(), "Nature")
                st.plotly_chart(fig, use_container_width=True)
                st.error("❌ 'Complaint Nature' column not found. Please check your data structure.")

        with col2:
            st.subheader("Complaints by Nature (Nov-Present)")
            if 'Complaint Nature' in df_nov_present.columns:
                valid_data = df_nov_present['Complaint Nature'].dropna()
                valid_data = valid_data[valid_data != '']

                if len(valid_data) > 0:
                    nature_counts = valid_data.value_counts().head(10)
                    fig = create_pie_chart(nature_counts, "Nature", px.colors.sequential.RdBu)
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info("📊 No nature data available in this period")
            else:
                fig = create_pie_chart(pd.Series(), "Nature", px.colors.sequential.RdBu)
                st.plotly_chart(fig, use_container_width=True)
                st.error("❌ 'Complaint Nature' column not found. Please check your data structure.")

        # Monthly trend
        st.subheader("Monthly Complaint Trend")
        if 'Date of Complaint' in df_jan_present.columns:
            valid_dates = df_jan_present['Date of Complaint'].dropna()
            if len(valid_dates) > 0:
                monthly_data = df_jan_present.groupby(df_jan_present['Date of Complaint'].dt.to_period('M')).size()
                if len(monthly_data) > 0:
                    df_monthly = pd.DataFrame({
                        'Month': monthly_data.index.astype(str),
                        'Count': monthly_data.values
                    })
                    fig = px.line(
                        df_monthly,
                        x='Month',
                        y='Count',
                        markers=True,
                        labels={'Month': 'Month', 'Count': 'Number of Complaints'}
                    )
                    fig.update_layout(height=400)
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info("📊 No monthly data available")
            else:
                st.info("📊 No valid complaint dates found in this period")
        else:
            fig = go.Figure()
            fig.add_annotation(
                text="Date of Complaint column not found",
                xref="paper", yref="paper",
                x=0.5, y=0.5, showarrow=False,
                font=dict(size=16)
            )
            fig.update_layout(height=400)
            st.plotly_chart(fig, use_container_width=True)
            st.error("❌ 'Date of Complaint' column not found. Please check your data structure.")
    
    with tab2:
        st.header("NTC Service Provider Analysis")

        # Filter NTC data
        if 'Agency' in df_jan_present.columns:
            df_ntc_jan = df_jan_present[df_jan_present['Agency'].str.contains('NTC', na=False, case=False)]
            df_ntc_sep = df_sep_present[df_sep_present['Agency'].str.contains('NTC', na=False, case=False)]

            col1, col2 = st.columns(2)

            with col1:
                st.subheader("Service Providers (Jan-Present)")
                st.metric("Total NTC Complaints", len(df_ntc_jan))
                if 'Service Providers' in df_ntc_jan.columns and len(df_ntc_jan) > 0:
                    valid_data = df_ntc_jan['Service Providers'].dropna()
                    valid_data = valid_data[valid_data != '']

                    if len(valid_data) > 0:
                        provider_counts = valid_data.value_counts().head(15)
                        fig = create_bar_chart(provider_counts, "Service Provider", 'Greens', 500)
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.info("📊 No service provider data available in this period")
                else:
                    fig = create_bar_chart(pd.Series(), "Service Provider", 'Greens', 500)
                    st.plotly_chart(fig, use_container_width=True)
                    if 'Service Providers' not in df_ntc_jan.columns:
                        st.error("❌ 'Service Providers' column not found")

            with col2:
                st.subheader("Service Providers (Sep-Present)")
                st.metric("Total NTC Complaints", len(df_ntc_sep))
                if 'Service Providers' in df_ntc_sep.columns and len(df_ntc_sep) > 0:
                    valid_data = df_ntc_sep['Service Providers'].dropna()
                    valid_data = valid_data[valid_data != '']

                    if len(valid_data) > 0:
                        provider_counts = valid_data.value_counts().head(15)
                        fig = create_bar_chart(provider_counts, "Service Provider", 'Teal', 500)
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.info("📊 No service provider data available in this period")
                else:
                    fig = create_bar_chart(pd.Series(), "Service Provider", 'Teal', 500)
                    st.plotly_chart(fig, use_container_width=True)
                    if 'Service Providers' not in df_ntc_sep.columns:
                        st.error("❌ 'Service Providers' column not found")
        else:
            st.error("❌ 'Agency' column not found in data. Cannot filter NTC complaints.")
            st.info("💡 Please ensure your data has an 'Agency' column to filter by organization.")
    
    with tab3:
        st.header("PEMEDES Service Provider Analysis")

        # Filter PEMEDES data
        if 'Agency' in df_jan_present.columns:
            df_pemedes_jan = df_jan_present[df_jan_present['Agency'].str.contains('PEMEDES', na=False, case=False)]
            df_pemedes_sep = df_sep_present[df_sep_present['Agency'].str.contains('PEMEDES', na=False, case=False)]

            col1, col2 = st.columns(2)

            with col1:
                st.subheader("Service Providers (Jan-Present)")
                st.metric("Total PEMEDES Complaints", len(df_pemedes_jan))
                if 'Service Providers' in df_pemedes_jan.columns and len(df_pemedes_jan) > 0:
                    valid_data = df_pemedes_jan['Service Providers'].dropna()
                    valid_data = valid_data[valid_data != '']

                    if len(valid_data) > 0:
                        provider_counts = valid_data.value_counts().head(15)
                        fig = create_bar_chart(provider_counts, "Service Provider", 'Purples', 500)
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.info("📊 No service provider data available in this period")
                else:
                    fig = create_bar_chart(pd.Series(), "Service Provider", 'Purples', 500)
                    st.plotly_chart(fig, use_container_width=True)
                    if 'Service Providers' not in df_pemedes_jan.columns:
                        st.error("❌ 'Service Providers' column not found")

            with col2:
                st.subheader("Service Providers (Sep-Present)")
                st.metric("Total PEMEDES Complaints", len(df_pemedes_sep))
                if 'Service Providers' in df_pemedes_sep.columns and len(df_pemedes_sep) > 0:
                    valid_data = df_pemedes_sep['Service Providers'].dropna()
                    valid_data = valid_data[valid_data != '']

                    if len(valid_data) > 0:
                        provider_counts = valid_data.value_counts().head(15)
                        fig = create_bar_chart(provider_counts, "Service Provider", 'Magenta', 500)
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.info("📊 No service provider data available in this period")
                else:
                    fig = create_bar_chart(pd.Series(), "Service Provider", 'Magenta', 500)
                    st.plotly_chart(fig, use_container_width=True)
                    if 'Service Providers' not in df_pemedes_sep.columns:
                        st.error("❌ 'Service Providers' column not found")
        else:
            st.error("❌ 'Agency' column not found in data. Cannot filter PEMEDES complaints.")
            st.info("💡 Please ensure your data has an 'Agency' column to filter by organization.")
    
    # Footer
    st.markdown("---")
    col1, col2 = st.columns([3, 1])
    with col1:
        st.markdown(f"*Dashboard last updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}*")
    with col2:
        if st.button("🗑️ Clear Cache"):
            st.cache_data.clear()
            st.success("Cache cleared!")
    
    # Auto-refresh mechanism
    if st.session_state.auto_refresh:
        time.sleep(st.session_state.refresh_interval)
        st.cache_data.clear()
        st.rerun()

if __name__ == "__main__":
    main()