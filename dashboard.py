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
    initial_sidebar_state="collapsed"
)

# CSS Styles - Applied in main() function
CUSTOM_CSS = """
    <style>
    /* Import Google Fonts */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    /* Global Styles */
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }

    /* Main container styling */
    .block-container {
        padding-top: 0.5rem;
        padding-bottom: 0rem;
        padding-left: 1rem;
        padding-right: 1rem;
        max-width: 100%;
    }

    /* Header styling */
    h1 {
        padding-top: 0rem;
        margin-top: 0rem;
        margin-bottom: 0.5rem;
        font-weight: 700;
        font-size: 2rem;
        color: #1f2937;
        letter-spacing: -0.5px;
    }

    h2, h3 {
        font-weight: 600;
        color: #374151;
    }

    /* Metric cards enhancement */
    [data-testid="stMetricValue"] {
        font-size: 2.25rem;
        font-weight: 700;
        color: #1f2937;
    }

    [data-testid="stMetricLabel"] {
        font-size: 0.95rem;
        font-weight: 600;
        color: #6b7280;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }

    /* Chart container styling */
    .stPlotlyChart {
        background-color: white;
        border-radius: 8px;
        padding: 4px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1), 0 1px 2px rgba(0,0,0,0.06);
        border: 1px solid #e5e7eb;
        transition: all 0.3s ease;
    }

    .stPlotlyChart:hover {
        box-shadow: 0 4px 6px rgba(0,0,0,0.1), 0 2px 4px rgba(0,0,0,0.06);
    }

    /* Compact subheaders for dashboard */
    h3 {
        font-size: 1.1rem;
        margin-top: 0.25rem;
        margin-bottom: 0.25rem;
        padding-top: 0;
        padding-bottom: 0;
    }

    /* Reduce column gap for more compact layout */
    [data-testid="column"] {
        padding: 0 0.5rem;
    }

    /* Compact styling for markdown headers in detailed view */
    .element-container p strong {
        font-size: 1rem;
        font-weight: 600;
        color: #374151;
    }

    /* Live indicator animation */
    .live-indicator {
        display: inline-block;
        width: 10px;
        height: 10px;
        background-color: #10b981;
        border-radius: 50%;
        animation: pulse 2s infinite;
        margin-right: 8px;
    }

    @keyframes pulse {
        0%, 100% { opacity: 1; box-shadow: 0 0 0 0 rgba(16, 185, 129, 0.7); }
        50% { opacity: 0.7; box-shadow: 0 0 0 6px rgba(16, 185, 129, 0); }
    }

    /* Reduce spacing between elements */
    .element-container {
        margin-bottom: 0.3rem;
    }

    /* Tab styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #f9fafb;
        padding: 8px;
        border-radius: 8px;
    }

    .stTabs [data-baseweb="tab"] {
        height: 50px;
        padding: 0 24px;
        background-color: white;
        border-radius: 6px;
        font-weight: 500;
        border: 1px solid #e5e7eb;
    }

    .stTabs [aria-selected="true"] {
        background-color: #3b82f6;
        color: white;
        border: 1px solid #3b82f6;
    }

    /* Sidebar styling */
    [data-testid="stSidebar"] {
        background-color: #f9fafb;
    }

    /* Button styling */
    .stButton>button {
        border-radius: 8px;
        font-weight: 500;
        border: 1px solid #e5e7eb;
        transition: all 0.2s ease;
    }

    .stButton>button:hover {
        border-color: #3b82f6;
        color: #3b82f6;
    }

    /* Info/Warning boxes */
    .stAlert {
        border-radius: 8px;
        border-left: 4px solid;
    }

    /* Expander styling */
    .streamlit-expanderHeader {
        font-weight: 600;
        border-radius: 8px;
    }

    /* Radio button styling */
    .stRadio > label {
        font-weight: 500;
    }

    /* Divider styling */
    hr {
        margin: 0.75rem 0;
        border: none;
        border-top: 2px solid #e5e7eb;
    }

    /* Fullscreen button */
    .fullscreen-btn {
        position: fixed;
        bottom: 20px;
        right: 20px;
        z-index: 999;
        background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%);
        color: white;
        padding: 12px 24px;
        border-radius: 50px;
        cursor: pointer;
        box-shadow: 0 4px 12px rgba(59, 130, 246, 0.4);
        font-weight: 600;
        transition: all 0.3s ease;
    }

    .fullscreen-btn:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(59, 130, 246, 0.5);
    }
    </style>

    <script>
    function toggleFullscreen() {
        if (!document.fullscreenElement) {
            document.documentElement.requestFullscreen();
        } else {
            if (document.exitFullscreen) {
                document.exitFullscreen();
            }
        }
    }
    </script>
"""

def extract_spreadsheet_id(url):
    """Extract spreadsheet ID from various Google Sheets URL formats with validation"""
    if not url or not isinstance(url, str):
        return None

    # Remove whitespace
    url = url.strip()

    if not url:
        return None

    # Pattern for spreadsheet ID (typical format is 44 characters)
    patterns = [
        r'/spreadsheets/d/([a-zA-Z0-9-_]+)',
        r'id=([a-zA-Z0-9-_]+)',
    ]

    for pattern in patterns:
        match = re.search(pattern, url)
        if match:
            spreadsheet_id = match.group(1)
            # Basic validation: Google Sheets IDs are typically 44 characters
            if len(spreadsheet_id) > 10:  # Minimum reasonable length
                return spreadsheet_id

    # If no pattern matches, assume the URL itself might be the ID
    if re.match(r'^[a-zA-Z0-9-_]{10,}$', url):
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

    # Validate and map columns (silently)
    column_mapping, missing_columns, suggestions = validate_required_columns(df)

    # Apply column mapping if found
    if column_mapping:
        df = apply_column_mapping(df, column_mapping)

    # Only show critical warnings for missing columns
    if missing_columns:
        with st.expander("⚠️ Column Validation Warnings", expanded=False):
            st.warning(f"Missing columns: {', '.join(missing_columns)}")
            if suggestions:
                st.info("**Possible column name matches:**")
                for col, similar in suggestions.items():
                    if similar:
                        st.write(f"  • For '{col}', did you mean: {', '.join(similar)}?")

    # Convert date columns to datetime with robust parsing (silently)
    date_columns = ['Date of Complaint', 'Date of Resolution', 'Date Responded']

    for col in date_columns:
        if col in df.columns:
            # Apply robust date parsing
            df[col] = df[col].apply(parse_date_robust)

    # Extract year and month for filtering
    if 'Date of Complaint' in df.columns:
        df['Year'] = df['Date of Complaint'].dt.year
        df['Month'] = df['Date of Complaint'].dt.month

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

    # Filter out complaints with Resolution = "FLS" (not included in dashboard)
    if 'Resolution' in df.columns:
        rows_before = len(df)
        # Create a mask to exclude rows where Resolution equals "FLS" (case-insensitive)
        # Only filter out actual "FLS" values, preserve NaN/None values
        resolution_upper = df['Resolution'].fillna('').astype(str).str.strip().str.upper()
        df = df[resolution_upper != 'FLS']
        rows_after = len(df)

        if rows_before > rows_after:
            st.info(f"ℹ️ Excluded {rows_before - rows_after} complaints with Resolution = 'FLS' from dashboard")

    return df

def filter_by_date(df, start_month, start_year=None):
    """Filter dataframe by date range with robust error handling"""
    if df is None or df.empty:
        return pd.DataFrame()

    if 'Date of Complaint' not in df.columns:
        return df

    if start_year is None:
        start_year = datetime.now().year

    try:
        # Validate month parameter
        if not isinstance(start_month, int) or start_month < 1 or start_month > 12:
            st.error(f"Invalid month value: {start_month}. Must be between 1 and 12.")
            return df

        # Validate year parameter
        if not isinstance(start_year, int) or start_year < 1900 or start_year > 2100:
            st.error(f"Invalid year value: {start_year}.")
            return df

        start_date = pd.Timestamp(year=start_year, month=start_month, day=1)
        filtered_df = df[df['Date of Complaint'] >= start_date].copy()
        return filtered_df
    except Exception as e:
        st.error(f"Error filtering by date: {str(e)}")
        return df

def create_bar_chart(data_series, title, color_scale='blues', height=400):
    """Create a modern bar chart with enhanced styling"""
    if data_series is None or len(data_series) == 0:
        # Return empty figure with message
        fig = go.Figure()
        fig.add_annotation(
            text="No data available",
            xref="paper", yref="paper",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=16, color='#9ca3af')
        )
        fig.update_layout(
            height=height,
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)'
        )
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
        labels={'Count': 'Count', 'Category': ''},
        color='Count',
        color_continuous_scale=color_scale
    )
    fig.update_layout(
        height=height,
        showlegend=False,
        margin=dict(l=5, r=5, t=5, b=5),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(family='Inter, sans-serif', size=14, color='#374151'),
        xaxis=dict(
            showgrid=True,
            gridcolor='#f3f4f6',
            zeroline=False,
            title_font=dict(size=14)
        ),
        yaxis=dict(
            categoryorder='total ascending',
            showgrid=False,
            tickfont=dict(size=14)
        ),
        coloraxis_showscale=False
    )
    fig.update_traces(
        marker=dict(
            line=dict(width=0)
        ),
        hovertemplate='<b>%{y}</b><br>Count: %{x}<extra></extra>'
    )
    return fig

def create_pie_chart(data_series, title, color_scheme=None, height=400):
    """Create a modern donut chart with enhanced styling"""
    if data_series is None or len(data_series) == 0:
        # Return empty figure with message
        fig = go.Figure()
        fig.add_annotation(
            text="No data available",
            xref="paper", yref="paper",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=16, color='#9ca3af')
        )
        fig.update_layout(
            height=height,
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)'
        )
        return fig

    # Convert to DataFrame for plotly
    df_plot = pd.DataFrame({
        'Category': data_series.index.astype(str),
        'Count': data_series.values
    })

    # Default professional color scheme
    if color_scheme is None:
        color_scheme = ['#3b82f6', '#8b5cf6', '#ec4899', '#f59e0b', '#10b981',
                       '#06b6d4', '#6366f1', '#f43f5e', '#84cc16', '#a855f7']

    fig = px.pie(
        df_plot,
        values='Count',
        names='Category',
        hole=0.45,
        color_discrete_sequence=color_scheme
    )
    fig.update_layout(
        height=height,
        margin=dict(l=5, r=5, t=5, b=5),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(family='Inter, sans-serif', size=14, color='#374151'),
        showlegend=True,
        legend=dict(
            orientation="v",
            yanchor="middle",
            y=0.5,
            xanchor="left",
            x=1.02,
            font=dict(size=13)
        )
    )
    fig.update_traces(
        textposition='inside',
        textinfo='percent',
        textfont=dict(size=14, color='white'),
        hovertemplate='<b>%{label}</b><br>Count: %{value}<br>Percentage: %{percent}<extra></extra>',
        marker=dict(line=dict(color='white', width=2))
    )
    return fig

def create_line_chart(df_monthly, height=400):
    """Create a modern line chart with enhanced styling"""
    fig = px.line(
        df_monthly,
        x='Month',
        y='Count',
        markers=True,
        labels={'Month': 'Month', 'Count': 'Number of Complaints'}
    )
    fig.update_layout(
        height=height,
        margin=dict(l=5, r=5, t=5, b=5),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(family='Inter, sans-serif', size=14, color='#374151'),
        xaxis=dict(
            showgrid=False,
            zeroline=False,
            title=None,
            tickfont=dict(size=13)
        ),
        yaxis=dict(
            showgrid=True,
            gridcolor='#f3f4f6',
            zeroline=False,
            title='Complaints',
            title_font=dict(size=14),
            tickfont=dict(size=14)
        ),
        hovermode='x unified'
    )
    fig.update_traces(
        line=dict(color='#3b82f6', width=3),
        marker=dict(size=8, color='#3b82f6', line=dict(width=2, color='white')),
        hovertemplate='<b>%{y}</b> complaints<extra></extra>'
    )
    return fig

def main():
    # Apply custom CSS styles
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

    # Default Google Sheets URL
    DEFAULT_GSHEET_URL = "https://docs.google.com/spreadsheets/d/1iqgkRJF6HexmWQsCDMwmBLMW--CeAz6YDEnDF9Hof_A/edit?gid=1179220692#gid=1179220692"

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
        st.session_state.gsheet_url = DEFAULT_GSHEET_URL
    if 'use_public_sheet' not in st.session_state:
        st.session_state.use_public_sheet = True
    if 'view_mode' not in st.session_state:
        st.session_state.view_mode = "Dashboard"
    
    # Enhanced header with live indicator
    if st.session_state.auto_refresh:
        st.markdown("""
            <div style="margin-bottom: 1.5rem;">
                <h1 style="display: inline-block; margin: 0;">📊 Complaint Analysis Dashboard</h1>
                <span style="float: right; margin-top: 8px;">
                    <span class="live-indicator" style="display: inline-block; width: 10px; height: 10px; background-color: #10b981; border-radius: 50%; margin-right: 6px;"></span>
                    <span style="font-size: 0.9rem; font-weight: 600; color: #10b981;">LIVE</span>
                </span>
            </div>
        """, unsafe_allow_html=True)
    else:
        st.title("📊 Complaint Analysis Dashboard")
    
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

    # Filter datasets with validation and error handling
    try:
        df_jan_present = df[df['Year'] == current_year].copy() if 'Year' in df.columns else df.copy()
    except Exception as e:
        st.error(f"Error filtering data for current year: {str(e)}")
        df_jan_present = df.copy()

    try:
        df_nov_present = filter_by_date(df, 11, current_year)
    except Exception as e:
        st.error(f"Error filtering data from November: {str(e)}")
        df_nov_present = df.copy()

    try:
        df_sep_present = filter_by_date(df, 9, current_year)
    except Exception as e:
        st.error(f"Error filtering data from September: {str(e)}")
        df_sep_present = df.copy()
    
    # View mode selector - placed in columns with metrics
    view_col1, view_col2 = st.columns([2, 8])
    with view_col1:
        view_mode = st.radio("View:", ["📊 Dashboard", "📋 Detailed"], horizontal=True, label_visibility="collapsed")
        st.session_state.view_mode = view_mode

    # Summary metrics with error handling
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Complaints (Jan-Present)", len(df_jan_present))
    with col2:
        st.metric("Total Complaints (Nov-Present)", len(df_nov_present))
    with col3:
        if 'Agency' in df_jan_present.columns:
            try:
                ntc_count = len(df_jan_present[df_jan_present['Agency'].str.contains('NTC', na=False, case=False)])
                st.metric("NTC Complaints", ntc_count)
            except Exception as e:
                st.metric("NTC Complaints", "Error")
                st.error(f"Error counting NTC complaints: {str(e)}")
        else:
            st.metric("NTC Complaints", "N/A")
    with col4:
        if 'Agency' in df_jan_present.columns:
            try:
                pemedes_count = len(df_jan_present[df_jan_present['Agency'].str.contains('PEMEDES', na=False, case=False)])
                st.metric("PEMEDES Complaints", pemedes_count)
            except Exception as e:
                st.metric("PEMEDES Complaints", "Error")
                st.error(f"Error counting PEMEDES complaints: {str(e)}")
        else:
            st.metric("PEMEDES Complaints", "N/A")

    st.markdown("---")

    # Dashboard View - 4 Charts in 2x2 Grid (Optimized for Fullscreen)
    if view_mode == "📊 Dashboard":
        # Calculate optimal chart height based on screen (assuming ~900px height in fullscreen)
        chart_height = 420

        # Row 1: Category and Nature
        col1, col2 = st.columns(2)

        with col1:
            st.subheader("📊 Complaints by Category")
            if 'Complaint Category' in df_jan_present.columns:
                valid_data = df_jan_present['Complaint Category'].dropna()
                valid_data = valid_data[valid_data != '']
                if len(valid_data) > 0:
                    category_counts = valid_data.value_counts().head(8)
                    fig = create_bar_chart(category_counts, "Category", 'blues', chart_height)
                    st.plotly_chart(fig, use_container_width=True, key="dashboard_category")
                else:
                    st.info("No category data available")
            else:
                st.error("'Complaint Category' column not found")

        with col2:
            st.subheader("📊 Complaints by Nature")
            if 'Complaint Nature' in df_jan_present.columns:
                valid_data = df_jan_present['Complaint Nature'].dropna()
                valid_data = valid_data[valid_data != '']
                if len(valid_data) > 0:
                    nature_counts = valid_data.value_counts().head(8)
                    fig = create_bar_chart(nature_counts, "Nature", 'purples', chart_height)
                    st.plotly_chart(fig, use_container_width=True, key="dashboard_nature")
                else:
                    st.info("No nature data available")
            else:
                st.error("'Complaint Nature' column not found")

        # Row 2: Monthly Trend and Service Providers
        col3, col4 = st.columns(2)

        with col3:
            st.subheader("📈 Monthly Complaint Trend")
            if 'Date of Complaint' in df_jan_present.columns:
                valid_dates = df_jan_present['Date of Complaint'].dropna()
                if len(valid_dates) > 0:
                    monthly_data = df_jan_present.groupby(df_jan_present['Date of Complaint'].dt.to_period('M')).size()
                    if len(monthly_data) > 0:
                        df_monthly = pd.DataFrame({
                            'Month': monthly_data.index.astype(str),
                            'Count': monthly_data.values
                        })
                        fig = create_line_chart(df_monthly, chart_height)
                        st.plotly_chart(fig, use_container_width=True, key="dashboard_monthly_trend")
                    else:
                        st.info("No monthly data available")
                else:
                    st.info("No valid complaint dates found")
            else:
                st.error("'Date of Complaint' column not found")

        with col4:
            st.subheader("📊 Top Service Providers")
            if 'Service Providers' in df_jan_present.columns:
                valid_data = df_jan_present['Service Providers'].dropna()
                valid_data = valid_data[valid_data != '']
                if len(valid_data) > 0:
                    provider_counts = valid_data.value_counts().head(8)
                    fig = create_bar_chart(provider_counts, "Service Provider", 'greens', chart_height)
                    st.plotly_chart(fig, use_container_width=True, key="dashboard_providers")
                else:
                    st.info("No service provider data available")
            else:
                st.error("'Service Providers' column not found")

    # Detailed View - Tab layout for organized viewing
    else:
        tab1, tab2, tab3 = st.tabs(["📈 Overall Analysis", "🏢 NTC Analysis", "🏢 PEMEDES Analysis"])

        with tab1:
            st.markdown("### Overall Complaint Analysis")

            # Row 1: Complaints by Category
            col1, col2 = st.columns(2)

            with col1:
                st.markdown("**Complaints by Category (Jan-Present)**")
                if 'Complaint Category' in df_jan_present.columns:
                    valid_data = df_jan_present['Complaint Category'].dropna()
                    valid_data = valid_data[valid_data != '']

                    if len(valid_data) > 0:
                        category_counts = valid_data.value_counts().head(8)
                        fig = create_bar_chart(category_counts, "Category", 'blues', 360)
                        st.plotly_chart(fig, use_container_width=True, key="detailed_category_jan")
                    else:
                        st.info("📊 No category data available in this period")
                else:
                    fig = create_bar_chart(pd.Series(), "Category", 'blues', 360)
                    st.plotly_chart(fig, use_container_width=True, key="detailed_category_jan_empty")
                    st.error("❌ 'Complaint Category' column not found")

            with col2:
                st.markdown("**Complaints by Category (Nov-Present)**")
                if 'Complaint Category' in df_nov_present.columns:
                    valid_data = df_nov_present['Complaint Category'].dropna()
                    valid_data = valid_data[valid_data != '']

                    if len(valid_data) > 0:
                        category_counts = valid_data.value_counts().head(8)
                        fig = create_bar_chart(category_counts, "Category", 'oranges', 360)
                        st.plotly_chart(fig, use_container_width=True, key="detailed_category_nov")
                    else:
                        st.info("📊 No category data available in this period")
                else:
                    fig = create_bar_chart(pd.Series(), "Category", 'oranges', 360)
                    st.plotly_chart(fig, use_container_width=True, key="detailed_category_nov_empty")
                    st.error("❌ 'Complaint Category' column not found")

            # Row 2: Complaints by Nature
            col1, col2 = st.columns(2)

            with col1:
                st.markdown("**Complaints by Nature (Jan-Present)**")
                if 'Complaint Nature' in df_jan_present.columns:
                    valid_data = df_jan_present['Complaint Nature'].dropna()
                    valid_data = valid_data[valid_data != '']

                    if len(valid_data) > 0:
                        nature_counts = valid_data.value_counts().head(8)
                        fig = create_bar_chart(nature_counts, "Nature", 'purples', 360)
                        st.plotly_chart(fig, use_container_width=True, key="detailed_nature_jan")
                    else:
                        st.info("📊 No nature data available in this period")
                else:
                    fig = create_bar_chart(pd.Series(), "Nature", 'purples', 360)
                    st.plotly_chart(fig, use_container_width=True, key="detailed_nature_jan_empty")
                    st.error("❌ 'Complaint Nature' column not found")

            with col2:
                st.markdown("**Complaints by Nature (Nov-Present)**")
                if 'Complaint Nature' in df_nov_present.columns:
                    valid_data = df_nov_present['Complaint Nature'].dropna()
                    valid_data = valid_data[valid_data != '']

                    if len(valid_data) > 0:
                        nature_counts = valid_data.value_counts().head(8)
                        fig = create_bar_chart(nature_counts, "Nature", 'teal', 360)
                        st.plotly_chart(fig, use_container_width=True, key="detailed_nature_nov")
                    else:
                        st.info("📊 No nature data available in this period")
                else:
                    fig = create_bar_chart(pd.Series(), "Nature", 'teal', 360)
                    st.plotly_chart(fig, use_container_width=True, key="detailed_nature_nov_empty")
                    st.error("❌ 'Complaint Nature' column not found")

            # Monthly trend
            st.markdown("**Monthly Complaint Trend**")
            if 'Date of Complaint' in df_jan_present.columns:
                valid_dates = df_jan_present['Date of Complaint'].dropna()
                if len(valid_dates) > 0:
                    monthly_data = df_jan_present.groupby(df_jan_present['Date of Complaint'].dt.to_period('M')).size()
                    if len(monthly_data) > 0:
                        df_monthly = pd.DataFrame({
                            'Month': monthly_data.index.astype(str),
                            'Count': monthly_data.values
                        })
                        fig = create_line_chart(df_monthly, 320)
                        st.plotly_chart(fig, use_container_width=True, key="detailed_monthly_trend")
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
                fig.update_layout(height=320)
                st.plotly_chart(fig, use_container_width=True, key="detailed_monthly_trend_empty")
                st.error("❌ 'Date of Complaint' column not found")
    
        with tab2:
            st.markdown("### NTC Service Provider Analysis")

            # Filter NTC data with error handling
            if 'Agency' in df_jan_present.columns:
                try:
                    df_ntc_jan = df_jan_present[df_jan_present['Agency'].str.contains('NTC', na=False, case=False)]
                    df_ntc_sep = df_sep_present[df_sep_present['Agency'].str.contains('NTC', na=False, case=False)]
                except Exception as e:
                    st.error(f"Error filtering NTC data: {str(e)}")
                    df_ntc_jan = pd.DataFrame()
                    df_ntc_sep = pd.DataFrame()

                col1, col2 = st.columns(2)

                with col1:
                    st.markdown("**Service Providers (Jan-Present)**")
                    st.metric("Total Complaints", len(df_ntc_jan))
                    if 'Service Providers' in df_ntc_jan.columns and len(df_ntc_jan) > 0:
                        valid_data = df_ntc_jan['Service Providers'].dropna()
                        valid_data = valid_data[valid_data != '']

                        if len(valid_data) > 0:
                            provider_counts = valid_data.value_counts().head(12)
                            fig = create_bar_chart(provider_counts, "Service Provider", 'greens', 450)
                            st.plotly_chart(fig, use_container_width=True, key="ntc_providers_jan")
                        else:
                            st.info("📊 No service provider data available")
                    else:
                        fig = create_bar_chart(pd.Series(), "Service Provider", 'greens', 450)
                        st.plotly_chart(fig, use_container_width=True, key="ntc_providers_jan_empty")
                        if 'Service Providers' not in df_ntc_jan.columns:
                            st.error("❌ 'Service Providers' column not found")

                with col2:
                    st.markdown("**Service Providers (Sep-Present)**")
                    st.metric("Total Complaints", len(df_ntc_sep))
                    if 'Service Providers' in df_ntc_sep.columns and len(df_ntc_sep) > 0:
                        valid_data = df_ntc_sep['Service Providers'].dropna()
                        valid_data = valid_data[valid_data != '']

                        if len(valid_data) > 0:
                            provider_counts = valid_data.value_counts().head(12)
                            fig = create_bar_chart(provider_counts, "Service Provider", 'teal', 450)
                            st.plotly_chart(fig, use_container_width=True, key="ntc_providers_sep")
                        else:
                            st.info("📊 No service provider data available")
                    else:
                        fig = create_bar_chart(pd.Series(), "Service Provider", 'teal', 450)
                        st.plotly_chart(fig, use_container_width=True, key="ntc_providers_sep_empty")
                        if 'Service Providers' not in df_ntc_sep.columns:
                            st.error("❌ 'Service Providers' column not found")
            else:
                st.error("❌ 'Agency' column not found in data. Cannot filter NTC complaints.")
                st.info("💡 Please ensure your data has an 'Agency' column.")
    
        with tab3:
            st.markdown("### PEMEDES Service Provider Analysis")

            # Filter PEMEDES data with error handling
            if 'Agency' in df_jan_present.columns:
                try:
                    df_pemedes_jan = df_jan_present[df_jan_present['Agency'].str.contains('PEMEDES', na=False, case=False)]
                    df_pemedes_sep = df_sep_present[df_sep_present['Agency'].str.contains('PEMEDES', na=False, case=False)]
                except Exception as e:
                    st.error(f"Error filtering PEMEDES data: {str(e)}")
                    df_pemedes_jan = pd.DataFrame()
                    df_pemedes_sep = pd.DataFrame()

                col1, col2 = st.columns(2)

                with col1:
                    st.markdown("**Service Providers (Jan-Present)**")
                    st.metric("Total Complaints", len(df_pemedes_jan))
                    if 'Service Providers' in df_pemedes_jan.columns and len(df_pemedes_jan) > 0:
                        valid_data = df_pemedes_jan['Service Providers'].dropna()
                        valid_data = valid_data[valid_data != '']

                        if len(valid_data) > 0:
                            provider_counts = valid_data.value_counts().head(12)
                            fig = create_bar_chart(provider_counts, "Service Provider", 'purples', 450)
                            st.plotly_chart(fig, use_container_width=True, key="pemedes_providers_jan")
                        else:
                            st.info("📊 No service provider data available")
                    else:
                        fig = create_bar_chart(pd.Series(), "Service Provider", 'purples', 450)
                        st.plotly_chart(fig, use_container_width=True, key="pemedes_providers_jan_empty")
                        if 'Service Providers' not in df_pemedes_jan.columns:
                            st.error("❌ 'Service Providers' column not found")

                with col2:
                    st.markdown("**Service Providers (Sep-Present)**")
                    st.metric("Total Complaints", len(df_pemedes_sep))
                    if 'Service Providers' in df_pemedes_sep.columns and len(df_pemedes_sep) > 0:
                        valid_data = df_pemedes_sep['Service Providers'].dropna()
                        valid_data = valid_data[valid_data != '']

                        if len(valid_data) > 0:
                            provider_counts = valid_data.value_counts().head(12)
                            fig = create_bar_chart(provider_counts, "Service Provider", 'magenta', 450)
                            st.plotly_chart(fig, use_container_width=True, key="pemedes_providers_sep")
                        else:
                            st.info("📊 No service provider data available")
                    else:
                        fig = create_bar_chart(pd.Series(), "Service Provider", 'magenta', 450)
                        st.plotly_chart(fig, use_container_width=True, key="pemedes_providers_sep_empty")
                        if 'Service Providers' not in df_pemedes_sep.columns:
                            st.error("❌ 'Service Providers' column not found")
            else:
                st.error("❌ 'Agency' column not found in data. Cannot filter PEMEDES complaints.")
                st.info("💡 Please ensure your data has an 'Agency' column.")

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