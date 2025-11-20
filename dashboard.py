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

# PEMEDES Service Providers List
PEMEDES_PROVIDERS = [
    "2GO Express, Inc.",
    "Sky Express",
    "3PL Service Provider Philippines, Inc.",
    "AAI Logistics Cargo Express d.b.a. Black Arrow",
    "A-Best Express, Inc.",
    "ACE-REM Messengerial and General Services, Inc.",
    "Airfreight 2100, Inc. d.b.a. Air21",
    "Airspeed International Corp.",
    "Arnold Mabanglo Express (AMEX)",
    "ASAP Courier and Logistics Phils., Inc.",
    "CACS Express and Allied Services, Inc.",
    "Cavatas-MSI Xpress Services Inc.",
    "CBL Freight Forwarder & Courier Express Int'l., Inc.",
    "Chargen Messengerial and General Services",
    "CJ Transnational Logistics Philippines, Inc.",
    "Cloud Panda PH, Inc. (d.b.a.) \"Tok Tok\"",
    "COMET Labor Service Cooperative",
    "DAG Xpress Courier, Inc.",
    "Del Asia Express Delivery & General Services, Inc.",
    "Diar's Assistance, Inc.",
    "Doorbell Technologies, Inc.",
    "El Grande Messengerial Services, Inc.",
    "Electrobill, Inc.",
    "Entrego Express Corporation",
    "EXMER, Inc.",
    "Fastcargo Logistics Corporation",
    "Fastrak Services, Inc.",
    "Fastrust Services, Inc.",
    "FES Business Solutions, Inc.",
    "Flash Express (PH) Co. Ltd., Inc.",
    "Flying High Energy Express Services Corporation",
    "GML Cargo Forwarder & Courier Express Int'l., Inc.",
    "GO21, Inc.",
    "GRABEXPRESS, Inc.",
    "Herald Express, Inc.",
    "Information Express Services, Inc. (INFORMEX)",
    "ICS / Intertraffic Transport Corporation",
    "Intervolve Express Services, Inc.",
    "Jay Messengerial and Manpower Services",
    "JG Manpo Janitorial & Messengerial Services Contractor",
    "Johnny Air Cargo and Delivery Services, Inc.",
    "JRMT Resources Corporation",
    "JRS Business Corporation",
    "LBC Express Corporation",
    "Libcap Super Express Corp.",
    "M.M. Bacarisas Courier Services",
    "Mail Expert Messengerial and General Services, Inc.",
    "Mailouwyz Courier",
    "Mailworld Express Service International, Inc.",
    "Mega Mail Express and General Services, Inc.",
    "MEJBAS Services, Inc.",
    "Mesco Express Service Corp.",
    "Metro Courier, Inc.",
    "Metro Prideco Services Corporation",
    "MMSC Services Corporation",
    "MR Messengerial & General Services",
    "MSPB Courier Services",
    "MTEL Trading & Manpower Services",
    "NGC Enterprises / Nathaniel G Cruz",
    "(Nathaniel G. Cruz Express Service)",
    "Ocean Coast Shipping Corp.",
    "Oceanwave Services, Inc.",
    "Pelican Express, Inc. to Cliqnship",
    "PH Gobal Jet Express, Inc. d.b.a. J&T Express",
    "PPB Messengerial Services, Inc.",
    "PRC Courier and Maintenance Services",
    "Priority Handling Logistics, Inc.",
    "PRO2000 Services, Inc.",
    "Promark Corporation",
    "Pronto Express Distribution, Inc.",
    "Quadx, Inc. d.b.a. GoGo Xpress",
    "Qualitrans Courier and Manpower Services, Inc.",
    "R&H Messengerial & General Services",
    "RAF International Forwarding Philippines, Inc.",
    "Republic Courier Service, Inc.",
    "RGServe Manpower Services",
    "RML Courier Services",
    "Safefreight Services, Inc.",
    "San Gabriel General Messengerial Services and Sales, Inc.",
    "Securetrac, Inc.",
    "Silver Royal General Services",
    "Snappmile Logistics, Inc.",
    "Speedels Services, Inc.",
    "Speedex Courier and Forwarder, Inc.",
    "Speedworks Courier Services Corporation",
    "Spex International Courier Services",
    "SPX Philippines, Inc.",
    "St. Joseph LFS Industrial Corp.",
    "Suremail Courier Services, Inc.",
    "Telexpress, Inc.",
    "TNT Express Deliveries (Phils.), Inc.",
    "Top Dynamics, Inc.",
    "Topserve Worldwide Express, Inc.",
    "Triload Express Systems",
    "UPS-Delbros Transport, Inc.",
    "Virgo Messengerial Services, Inc.",
    "Wall Street Courier Services, Inc. d.b.a. Ninja Van",
    "Wide Wide World Express Corporation",
    "Worklink Services, Inc.",
    "Ximex Delivery Express, Inc.",
    "Xytron International, Inc.",
    "ZIP Business Services, Inc.",
    "SF Express",
    "Mober PH",
    "Litexpress",
    "Lalamove",
    "YTO Express",
    "LEX PH",
    # Common variations and shortened names
    "2GO Express",
    "2GO",
    "J&T Express",
    "J&T",
    "LBC Express",
    "LBC",
    "Ninja Van",
    "Flash Express",
    "Grab Express",
    "GrabExpress",
    "Lalamove",
    "Airspeed",
    "Air21",
    "Black Arrow",
    "GoGo Xpress",
    "SPX",
]

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

    /* ==========================================
       DESIGN SYSTEM - Consistent Color Palette
       ==========================================
       Primary Blue: #3b82f6
       Light Blue: #eff6ff
       Dark Gray: #1f2937
       Medium Gray: #374151
       Light Gray: #6b7280
       Border Gray: #cbd5e1, #d1d5db, #e5e7eb
       Background: #f9fafb, #f8fafc
       White: #ffffff
       ========================================== */

    /* Global Styles */
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }

    /* Main container styling */
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 0rem;
        padding-left: 0.75rem;
        padding-right: 0.75rem;
        max-width: 100%;
    }

    /* Header styling */
    h1 {
        padding-top: 0.5rem;
        margin-top: 0rem;
        margin-bottom: 0.5rem;
        font-weight: 700;
        font-size: 1.75rem;
        color: #1f2937;
        letter-spacing: -0.5px;
        line-height: 1.3;
    }

    h2, h3, h4 {
        font-weight: 600;
        color: #374151;
    }

    h4 {
        font-size: 1.15rem;
        margin-top: 0.2rem;
        margin-bottom: 0.5rem;
        font-weight: 600;
    }

    /* Metric cards enhancement */
    [data-testid="stMetricValue"] {
        font-size: 2.5rem;
        font-weight: 700;
        color: #1f2937;
        line-height: 1.2;
    }

    [data-testid="stMetricLabel"] {
        font-size: 1rem;
        font-weight: 600;
        color: #6b7280;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 0.5rem;
    }

    [data-testid="stMetricDelta"] {
        font-size: 0.95rem;
        font-weight: 500;
        margin-top: 0.5rem;
    }

    /* Specific styling for metric label container */
    [data-testid="stMetricLabel"] > div {
        padding-bottom: 0.25rem;
    }

    /* Metric card container - Enhanced with visible borders */
    [data-testid="metric-container"] {
        background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
        padding: 1.5rem 1.25rem;
        border-radius: 12px;
        border: 2px solid #d1d5db !important;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
        transition: all 0.3s ease;
        margin: 0.25rem 0;
    }

    [data-testid="metric-container"]:hover {
        border-color: #3b82f6 !important;
        box-shadow: 0 10px 15px -3px rgba(59, 130, 246, 0.2), 0 4px 6px -2px rgba(59, 130, 246, 0.1);
        transform: translateY(-3px);
        background: linear-gradient(135deg, #ffffff 0%, #eff6ff 100%);
    }

    /* Make metric cards stand out more - Base styling */
    div[data-testid="stMetric"] {
        background-color: white;
        padding: 1.25rem !important;
        border-radius: 12px !important;
        border: 2px solid #cbd5e1 !important;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1) !important;
        margin: 0.25rem;
    }

    div[data-testid="stMetric"]:hover {
        border-color: #3b82f6 !important;
        box-shadow: 0 8px 12px rgba(59, 130, 246, 0.25) !important;
        transform: translateY(-2px);
        background: linear-gradient(135deg, #ffffff 0%, #eff6ff 100%) !important;
    }

    /* Additional emphasis for metric value and label spacing */
    [data-testid="stMetric"] [data-testid="stMetricValue"] {
        padding: 0.25rem 0;
    }

    /* Chart container styling - Consistent across all views */
    .stPlotlyChart {
        background-color: white;
        border-radius: 8px;
        padding: 2px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1), 0 1px 2px rgba(0,0,0,0.06);
        border: 1px solid #e5e7eb;
        transition: all 0.3s ease;
        margin-bottom: 0.5rem;
    }

    .stPlotlyChart:hover {
        box-shadow: 0 4px 6px rgba(0,0,0,0.1), 0 2px 4px rgba(0,0,0,0.06);
        border-color: #cbd5e1;
    }

    /* Compact subheaders for dashboard - Consistent styling */
    h3 {
        font-size: 1rem;
        margin-top: 0.2rem;
        margin-bottom: 0.5rem;
        padding-top: 0;
        padding-bottom: 0;
        font-weight: 600;
    }

    /* Reduce column gap for more compact layout */
    [data-testid="column"] {
        padding: 0 0.35rem;
    }

    /* Enhanced column styling for metric cards */
    div[data-testid="column"] > div[data-testid="stVerticalBlock"] > div[data-testid="stMetric"] {
        background: white;
        padding: 1.25rem;
        border-radius: 12px;
        border: 2px solid #cbd5e1;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        margin-bottom: 0.5rem;
        transition: all 0.3s ease;
    }

    div[data-testid="column"] > div[data-testid="stVerticalBlock"] > div[data-testid="stMetric"]:hover {
        border-color: #3b82f6;
        box-shadow: 0 8px 12px rgba(59, 130, 246, 0.25);
        transform: translateY(-2px);
        background: linear-gradient(135deg, #ffffff 0%, #eff6ff 100%);
    }

    /* Compact styling for markdown headers in detailed view */
    .element-container p strong {
        font-size: 1rem;
        font-weight: 600;
        color: #374151;
    }

    /* Consistent info/warning message styling */
    .stAlert {
        border-radius: 8px;
        margin: 0.5rem 0;
        font-size: 0.9rem;
    }

    /* Consistent caption styling - Smaller and subtle */
    .stCaption {
        font-size: 0.75rem;
        color: #9ca3af;
        margin: 0.15rem 0;
        line-height: 1.4;
        opacity: 0.85;
    }

    /* Make caption text even more compact */
    [data-testid="stCaptionContainer"] {
        margin-top: 0.25rem;
        margin-bottom: 0.25rem;
    }

    /* Subtle styling for warning/info captions */
    .stCaption p {
        font-size: 0.75rem;
        padding: 0.25rem 0.5rem;
        border-radius: 4px;
        background-color: #f9fafb;
        border-left: 2px solid #d1d5db;
        margin: 0.2rem 0;
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
        margin-bottom: 0.2rem;
    }

    /* Tab styling - Consistent design */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #f9fafb;
        padding: 8px;
        border-radius: 8px;
        margin-bottom: 0.75rem;
    }

    .stTabs [data-baseweb="tab"] {
        height: 50px;
        padding: 0 24px;
        background-color: white;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.95rem;
        border: 2px solid #e5e7eb;
        transition: all 0.3s ease;
    }

    .stTabs [data-baseweb="tab"]:hover {
        border-color: #cbd5e1;
        background-color: #f9fafb;
    }

    .stTabs [aria-selected="true"] {
        background-color: #3b82f6;
        color: white;
        border: 2px solid #3b82f6;
        box-shadow: 0 2px 4px rgba(59, 130, 246, 0.2);
    }

    /* Sidebar styling */
    [data-testid="stSidebar"] {
        background-color: #f9fafb;
    }

    /* Button styling - Consistent design */
    .stButton>button {
        border-radius: 8px;
        font-weight: 600;
        font-size: 0.95rem;
        border: 2px solid #e5e7eb;
        padding: 0.5rem 1rem;
        transition: all 0.3s ease;
        background-color: white;
    }

    .stButton>button:hover {
        border-color: #3b82f6;
        color: #3b82f6;
        background-color: #eff6ff;
        box-shadow: 0 2px 4px rgba(59, 130, 246, 0.2);
    }


    /* Expander styling - Consistent design */
    .streamlit-expanderHeader {
        font-weight: 600;
        font-size: 0.95rem;
        border-radius: 8px;
        padding: 0.75rem;
        background-color: #f9fafb;
        border: 1px solid #e5e7eb;
    }

    /* Radio button styling - Consistent design */
    .stRadio > label {
        font-weight: 600;
        font-size: 0.95rem;
        color: #374151;
    }

    /* Divider styling */
    hr {
        margin: 0.75rem 0;
        border: none;
        border-top: 2px solid #e5e7eb;
    }

    /* KPI Section Background - Consistent styling */
    .element-container:has([data-testid="stHorizontalBlock"]:has([data-testid="column"] [data-testid="stMetric"])) {
        background-color: #f9fafb;
        padding: 1rem;
        border-radius: 12px;
        margin-bottom: 0.75rem;
        border: 1px solid #e5e7eb;
    }

    /* Ensure consistent spacing for all chart sections */
    [data-testid="stHorizontalBlock"] {
        margin-bottom: 0.5rem;
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

def is_ntc_complaint(agency):
    """Check if a complaint belongs to NTC (National Telecommunications Commission)"""
    if pd.isna(agency) or agency == '':
        return False

    # Handle non-string types
    if not isinstance(agency, str):
        agency = str(agency)

    agency_str = agency.strip()

    # Return False for empty strings after stripping
    if not agency_str:
        return False

    agency_lower = agency_str.lower()

    # Check for NTC (case-insensitive)
    # Match whole word to avoid false positives
    return 'ntc' in agency_lower.split() or agency_lower == 'ntc' or 'ntc' in agency_lower

def is_pemedes_provider(service_provider):
    """Check if a service provider is a PEMEDES provider with robust matching

    Matching strategy:
    1. Exact match (case-insensitive)
    2. Provider list name contained in data (e.g., "LBC Express Corporation" contains "LBC Express")
    3. Data contained in provider list (e.g., "LBC" matches "LBC Express Corporation")
    4. Multi-word overlap for complex names
    """
    if pd.isna(service_provider) or service_provider == '':
        return False

    # Handle non-string types
    if not isinstance(service_provider, str):
        service_provider = str(service_provider)

    service_provider_str = service_provider.strip()

    # Return False for empty strings after stripping
    if not service_provider_str:
        return False

    service_provider_lower = service_provider_str.lower()

    # Strategy 1: Exact match (case-insensitive)
    for pemedes_sp in PEMEDES_PROVIDERS:
        if service_provider_lower == pemedes_sp.lower():
            return True

    # Strategy 2 & 3: Containment checks (bidirectional)
    for pemedes_sp in PEMEDES_PROVIDERS:
        pemedes_sp_lower = pemedes_sp.lower()

        # Check if PEMEDES provider name is contained in the data
        # Example: Data="LBC Express Corporation" contains Provider="LBC Express"
        if pemedes_sp_lower in service_provider_lower:
            # Additional safeguard: must match at least 5 characters or be a known short name
            if len(pemedes_sp_lower) >= 5 or pemedes_sp_lower in ['lbc', '2go', 'j&t', 'spx', 'air21']:
                return True

        # Check if data is contained in PEMEDES provider name (reverse check)
        # Example: Data="LBC" is in Provider="LBC Express Corporation"
        if service_provider_lower in pemedes_sp_lower:
            # Additional safeguard: data must be at least 3 characters
            if len(service_provider_lower) >= 3:
                return True

    # Strategy 4: Multi-word name matching for complex cases
    for pemedes_sp in PEMEDES_PROVIDERS:
        pemedes_sp_lower = pemedes_sp.lower()

        # Only apply to multi-word names
        if ' ' in pemedes_sp_lower and ' ' in service_provider_lower:
            words_in_data = set(service_provider_lower.split())
            words_in_pemedes = set(pemedes_sp_lower.split())

            # Remove common words that don't help identify the provider
            common_words = {'inc', 'corp', 'corporation', 'express', 'services', 'service',
                          'courier', 'delivery', 'logistics', 'international', 'phils',
                          'philippines', 'ltd', 'co', 'and', 'the'}
            words_in_data = words_in_data - common_words
            words_in_pemedes = words_in_pemedes - common_words

            # If there's significant word overlap (at least 2 meaningful words)
            if len(words_in_pemedes) >= 2 and len(words_in_data.intersection(words_in_pemedes)) >= 2:
                return True

    return False

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
        return None, []

    # Make a copy to avoid modifying original
    df = df.copy()

    # Store warning messages to display later
    warning_messages = []

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
            warning_messages.append(f"⚠️ Removed {rows_before - rows_after} rows with invalid complaint dates")

    # Filter out complaints with Resolution = "FLS" (not included in dashboard)
    if 'Resolution' in df.columns:
        rows_before = len(df)
        # Create a mask to exclude rows where Resolution equals "FLS" (case-insensitive)
        # Only filter out actual "FLS" values, preserve NaN/None values
        resolution_upper = df['Resolution'].fillna('').astype(str).str.strip().str.upper()
        df = df[resolution_upper != 'FLS']
        rows_after = len(df)

        if rows_before > rows_after:
            warning_messages.append(f"ℹ️ Excluded {rows_before - rows_after} complaints with Resolution = 'FLS' from dashboard")

    return df, warning_messages

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
        color_continuous_scale=color_scale,
        text='Count'  # Add data labels
    )
    fig.update_layout(
        height=height,
        showlegend=False,
        margin=dict(l=3, r=70, t=3, b=3),  # Increased right margin for labels
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
            tickfont=dict(size=13)  # Slightly smaller for better fit
        ),
        coloraxis_showscale=False
    )
    fig.update_traces(
        marker=dict(
            line=dict(width=0)
        ),
        texttemplate='%{text:,}',  # Format numbers with commas
        textposition='auto',  # Auto position labels (inside for long bars, outside for short)
        textfont=dict(size=12, family='Inter, sans-serif', weight='bold'),
        hovertemplate='<b>%{y}</b><br>Count: %{x:,}<extra></extra>',
        cliponaxis=False  # Don't clip labels outside the plot area
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
        margin=dict(l=3, r=3, t=3, b=3),
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
        textposition='auto',  # Auto position to avoid overlap
        textinfo='value+percent',  # Show only count and percentage (labels in legend)
        texttemplate='%{value:,}<br>(%{percent})',  # Format: count (with commas) and percentage
        textfont=dict(size=11, color='white', family='Inter, sans-serif', weight='bold'),
        hovertemplate='<b>%{label}</b><br>Count: %{value:,}<br>Percentage: %{percent}<extra></extra>',
        marker=dict(line=dict(color='white', width=2)),
        pull=[0.02] * len(data_series)  # Slightly pull slices apart for better visibility
    )
    return fig

def create_line_chart(df_monthly, height=400):
    """Create a modern line chart with enhanced styling"""
    fig = px.line(
        df_monthly,
        x='Month',
        y='Count',
        markers=True,
        labels={'Month': 'Month', 'Count': 'Number of Complaints'},
        text='Count'  # Add data labels
    )
    fig.update_layout(
        height=height,
        margin=dict(l=3, r=3, t=35, b=3),  # Increased top margin for labels
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(family='Inter, sans-serif', size=14, color='#374151'),
        xaxis=dict(
            showgrid=False,
            zeroline=False,
            title=None,
            tickfont=dict(size=12)  # Smaller tick font
        ),
        yaxis=dict(
            showgrid=True,
            gridcolor='#f3f4f6',
            zeroline=False,
            title='Complaints',
            title_font=dict(size=13),
            tickfont=dict(size=12)
        ),
        hovermode='x unified'
    )
    fig.update_traces(
        line=dict(color='#3b82f6', width=3),
        marker=dict(size=9, color='#3b82f6', line=dict(width=2, color='white')),
        textposition='top center',  # Position labels above points
        texttemplate='%{text:,}',  # Format numbers with commas
        textfont=dict(size=10, color='#1f2937', family='Inter, sans-serif', weight='bold'),
        hovertemplate='<b>%{y:,}</b> complaints<extra></extra>',
        cliponaxis=False  # Don't clip labels outside the plot area
    )
    return fig

def main():
    # Apply custom CSS styles
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

    # Default Google Sheets URL
    DEFAULT_GSHEET_URL = "https://docs.google.com/spreadsheets/d/1JDd0-4JffW5PB34XKDWaKDfPM-jQ22z1VXeCX1WpKGw/edit?usp=sharing"

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
                st.write("**Available columns:")
                for col in df.columns:
                    st.text(f"  • {col}")

                st.markdown("---")
                st.write("**Data Preview (First 5 rows):")
                st.dataframe(df.head(), use_container_width=True)

                # Show data types
                st.markdown("---")
                st.write("**Column Data Types:")
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
        df, data_warnings = prepare_data(df)

    if df is None or df.empty:
        st.error("❌ No valid data after processing. Please check your date formats.")
        return
    
    # Dynamically detect date range and create flexible filters
    # Get the actual date range from the data
    if 'Date of Complaint' in df.columns:
        valid_dates = df['Date of Complaint'].dropna()
        if len(valid_dates) > 0:
            min_date = valid_dates.min()
            max_date = valid_dates.max()

            # Calculate dynamic date ranges based on actual data
            # Period 1: Year-to-Date (from start of current year to present)
            current_year = max_date.year
            ytd_start = pd.Timestamp(year=current_year, month=1, day=1)

            # Period 2: Last Quarter (3 months from max date)
            last_quarter_start = max_date - pd.DateOffset(months=3)

            # Period 3: Last Month (1 month from max date)
            last_month_start = max_date - pd.DateOffset(months=1)

            # Create filtered datasets
            try:
                df_period1 = df[df['Date of Complaint'] >= ytd_start].copy()
                period1_label = f"{ytd_start.strftime('%b %Y')} - Present"
                period1_short = "YTD"
            except:
                df_period1 = df.copy()
                period1_label = "All Data"
                period1_short = "All"

            try:
                df_period2 = df[df['Date of Complaint'] >= last_quarter_start].copy()
                period2_label = f"Last 3 Months ({last_quarter_start.strftime('%b %Y')} - Present)"
                period2_short = "3M"
            except:
                df_period2 = df.copy()
                period2_label = "All Data"
                period2_short = "All"

            try:
                df_period3 = df[df['Date of Complaint'] >= last_month_start].copy()
                period3_label = f"Last Month ({last_month_start.strftime('%b %Y')} - Present)"
                period3_short = "1M"
            except:
                df_period3 = df.copy()
                period3_label = "All Data"
                period3_short = "All"
        else:
            # No valid dates, use all data
            df_period1 = df.copy()
            df_period2 = df.copy()
            df_period3 = df.copy()
            period1_label = "All Data"
            period2_label = "All Data"
            period3_label = "All Data"
            period1_short = "All"
            period2_short = "All"
            period3_short = "All"
            min_date = None
            max_date = None
    else:
        # No date column, use all data
        df_period1 = df.copy()
        df_period2 = df.copy()
        df_period3 = df.copy()
        period1_label = "All Data"
        period2_label = "All Data"
        period3_label = "All Data"
        period1_short = "All"
        period2_short = "All"
        period3_short = "All"
        min_date = None
        max_date = None
    
    # Display actual date ranges and data validation for verification
    with st.expander("📅 Data Validation & Date Range Verification", expanded=False):
        st.markdown("### Date Ranges")
        col_v1, col_v2, col_v3 = st.columns(3)
        with col_v1:
            st.write(f"**{period1_label}:**")
            if len(df_period1) > 0 and 'Date of Complaint' in df_period1.columns:
                dates = df_period1['Date of Complaint'].dropna()
                if len(dates) > 0:
                    st.write(f"• From: {dates.min().strftime('%Y-%m-%d')}")
                    st.write(f"• To: {dates.max().strftime('%Y-%m-%d')}")
                    st.write(f"• Records: {len(df_period1):,}")
        with col_v2:
            st.write(f"**{period2_label}:**")
            if len(df_period2) > 0 and 'Date of Complaint' in df_period2.columns:
                dates = df_period2['Date of Complaint'].dropna()
                if len(dates) > 0:
                    st.write(f"• From: {dates.min().strftime('%Y-%m-%d')}")
                    st.write(f"• To: {dates.max().strftime('%Y-%m-%d')}")
                    st.write(f"• Records: {len(df_period2):,}")
        with col_v3:
            st.write(f"**{period3_label}:**")
            if len(df_period3) > 0 and 'Date of Complaint' in df_period3.columns:
                dates = df_period3['Date of Complaint'].dropna()
                if len(dates) > 0:
                    st.write(f"• From: {dates.min().strftime('%Y-%m-%d')}")
                    st.write(f"• To: {dates.max().strftime('%Y-%m-%d')}")
                    st.write(f"• Records: {len(df_period3):,}")

        st.markdown("---")
        st.markdown("### Data Quality Metrics")
        qual_col1, qual_col2 = st.columns(2)

        with qual_col1:
            st.write(f"**NTC Validation ({period1_label}):**")
            if 'Agency' in df_period1.columns:
                ntc_filtered = df_period1[df_period1['Agency'].apply(is_ntc_complaint)]
                st.write(f"• Total NTC Complaints: {len(ntc_filtered):,}")
                st.write(f"• Percentage of Total: {(len(ntc_filtered)/len(df_period1)*100):.1f}%")
                if len(ntc_filtered) > 0 and 'Service Providers' in ntc_filtered.columns:
                    providers_with_data = ntc_filtered['Service Providers'].notna().sum()
                    st.write(f"• With Service Provider: {providers_with_data:,}")

        with qual_col2:
            st.write(f"**PEMEDES Validation ({period1_label}):**")
            if 'Service Providers' in df_period1.columns:
                pemedes_filtered = df_period1[df_period1['Service Providers'].apply(is_pemedes_provider)]
                st.write(f"• Total PEMEDES Complaints: {len(pemedes_filtered):,}")
                st.write(f"• Percentage of Total: {(len(pemedes_filtered)/len(df_period1)*100):.1f}%")
                st.write(f"• Unique Providers: {pemedes_filtered['Service Providers'].nunique()}")

                # Show top matched providers for verification
                if len(pemedes_filtered) > 0:
                    top_providers = pemedes_filtered['Service Providers'].value_counts().head(5)
                    st.write("• Top 5 Matched:")
                    for provider, count in top_providers.items():
                        st.write(f"  - {provider}: {count}")

        # Data consistency check
        st.markdown("---")
        st.markdown("### Data Consistency Check")
        cons_col1, cons_col2 = st.columns(2)

        with cons_col1:
            st.write("**Date Range Verification:**")
            jan_sep_diff = len(df_period1) - len(df_period3)
            st.write(f"• {period1_label}: {len(df_period1):,} complaints")
            st.write(f"• {period3_label}: {len(df_period3):,} complaints")
            st.write(f"• Difference: {jan_sep_diff:,}")
            if len(df_period3) > len(df_period1):
                st.error(f"⚠️ {period3_label} has MORE data than {period1_label} - This is incorrect!")

        with cons_col2:
            st.write("**PEMEDES Subset Verification:**")
            if 'Service Providers' in df_period1.columns and 'Service Providers' in df_period3.columns:
                pem_jan = df_period1[df_period1['Service Providers'].apply(is_pemedes_provider)]
                pem_sep = df_period3[df_period3['Service Providers'].apply(is_pemedes_provider)]
                st.write(f"• PEMEDES ({period1_label}): {len(pem_jan):,}")
                st.write(f"• PEMEDES ({period3_label}): {len(pem_sep):,}")
                st.write(f"• Difference: {len(pem_jan) - len(pem_sep):,}")
                if len(pem_sep) > len(pem_jan):
                    st.error("⚠️ Sep PEMEDES has MORE than Jan PEMEDES - This is incorrect!")

    # Summary metrics with error handling and enhanced design
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        jan_count = len(df_period1)
        nov_count = len(df_period2)
        delta_count = jan_count - nov_count
        st.metric(
            label=f"{period1_label}",
            value=f"{jan_count:,}",
            delta=f"{delta_count:+,} vs {period2_short}",
            delta_color="inverse",
            help=f"Total complaints for {period1_label}"
        )

    with col2:
        nov_count = len(df_period2)
        nov_pct = (nov_count / jan_count * 100) if jan_count > 0 else 0
        st.metric(
            label=f"{period2_label}",
            value=f"{nov_count:,}",
            delta=f"{nov_pct:.1f}% of {period1_short}",
            delta_color="off",
            help=f"Total complaints for {period2_label}"
        )

    with col3:
        if 'Agency' in df_period1.columns:
            try:
                ntc_count = len(df_period1[df_period1['Agency'].apply(is_ntc_complaint)])
                ntc_pct = (ntc_count / jan_count * 100) if jan_count > 0 else 0
                st.metric(
                    label="NTC Complaints",
                    value=f"{ntc_count:,}",
                    delta=f"{ntc_pct:.1f}% of total",
                    delta_color="off",
                    help="National Telecommunications Commission complaints"
                )
            except Exception as e:
                st.metric("NTC Complaints", "Error")
                st.error(f"Error counting NTC complaints: {str(e)}")
        else:
            st.metric("NTC Complaints", "N/A")

    with col4:
        if 'Service Providers' in df_period1.columns:
            try:
                pemedes_count = len(df_period1[df_period1['Service Providers'].apply(is_pemedes_provider)])
                pemedes_pct = (pemedes_count / jan_count * 100) if jan_count > 0 else 0
                st.metric(
                    label="PEMEDES Complaints",
                    value=f"{pemedes_count:,}",
                    delta=f"{pemedes_pct:.1f}% of total",
                    delta_color="off",
                    help="Complaints from PEMEDES service providers"
                )
            except Exception as e:
                st.metric("PEMEDES Complaints", "Error")
                st.error(f"Error counting PEMEDES complaints: {str(e)}")
        else:
            st.metric("PEMEDES Complaints", "N/A")

    st.markdown("---")

    # ============================================================================
    # SECTION 1: OVERALL ANALYSIS
    # ============================================================================
    st.markdown("## 📊 Overall Analysis")

    # Overall charts - 4 Charts in 2x2 Grid
    chart_height = 340

    # Row 1: Category and Nature
    col1, col2 = st.columns(2)

    with col1:
        st.markdown(f"#### 📊 Complaints by Category ({period1_label})")
        if 'Complaint Category' in df_period1.columns:
            valid_data = df_period1['Complaint Category'].dropna()
            valid_data = valid_data[valid_data != '']
            if len(valid_data) > 0:
                category_counts = valid_data.value_counts().head(8)
                fig = create_bar_chart(category_counts, "Category", 'blues', chart_height)
                st.plotly_chart(fig, use_container_width=True, key="overall_category")
            else:
                st.info("No category data available")
        else:
            st.error("'Complaint Category' column not found")

    with col2:
        st.markdown(f"#### 📊 Complaints by Nature ({period1_label})")
        if 'Complaint Nature' in df_period1.columns:
            valid_data = df_period1['Complaint Nature'].dropna()
            valid_data = valid_data[valid_data != '']
            if len(valid_data) > 0:
                nature_counts = valid_data.value_counts().head(8)
                fig = create_bar_chart(nature_counts, "Nature", 'purples', chart_height)
                st.plotly_chart(fig, use_container_width=True, key="overall_nature")
            else:
                st.info("No nature data available")
        else:
            st.error("'Complaint Nature' column not found")

    # Monthly Trend
    st.markdown(f"#### 📈 Monthly Complaint Trend ({period1_label})")
    if 'Date of Complaint' in df_period1.columns:
        valid_dates = df_period1['Date of Complaint'].dropna()
        if len(valid_dates) > 0:
            monthly_data = df_period1.groupby(df_period1['Date of Complaint'].dt.to_period('M')).size()
            if len(monthly_data) > 0:
                df_monthly = pd.DataFrame({
                    'Month': monthly_data.index.astype(str),
                    'Count': monthly_data.values
                })
                fig = create_line_chart(df_monthly, 300)
                st.plotly_chart(fig, use_container_width=True, key="overall_monthly_trend")
            else:
                st.info("No monthly data available")
        else:
            st.info("No valid complaint dates found")
    else:
        st.error("'Date of Complaint' column not found")

    st.markdown("---")

    # ============================================================================
    # SECTION 2: NTC ANALYSIS
    # ============================================================================
    st.markdown("## 🏢 NTC Analysis")

    # Filter NTC data with error handling
    if 'Agency' in df_period1.columns:
        try:
            # Use the is_ntc_complaint function for consistent filtering
            df_ntc_jan = df_period1[df_period1['Agency'].apply(is_ntc_complaint)]
            df_ntc_sep = df_period3[df_period3['Agency'].apply(is_ntc_complaint)]

            # Data integrity check
            if len(df_ntc_jan) == 0:
                st.warning(f"⚠️ No NTC complaints found in {period1_label} dataset. This may indicate:")
                st.write("• Agency column doesn't contain 'NTC'")
                st.write("• No NTC-related complaints in this period")
                st.write("• Check the 'Agency' column format")
        except Exception as e:
            st.error(f"❌ Error filtering NTC data: {str(e)}")
            st.info("Please check if the 'Agency' column contains valid text data.")
            df_ntc_jan = pd.DataFrame()
            df_ntc_sep = pd.DataFrame()

        # Enhanced KPI metrics for NTC
        kpi_col1, kpi_col2 = st.columns(2)
        with kpi_col1:
            ntc_jan_count = len(df_ntc_jan)
            ntc_jan_pct = (ntc_jan_count / len(df_period1) * 100) if len(df_period1) > 0 else 0
            st.metric(
                label=f"NTC ({period1_label})",
                value=f"{ntc_jan_count:,}",
                delta=f"{ntc_jan_pct:.1f}% of all complaints",
                delta_color="off",
                help=f"NTC complaints for {period1_label}"
            )
        with kpi_col2:
            ntc_sep_count = len(df_ntc_sep)
            ntc_sep_pct = (ntc_sep_count / len(df_period3) * 100) if len(df_period3) > 0 else 0
            st.metric(
                label=f"NTC ({period3_label})",
                value=f"{ntc_sep_count:,}",
                delta=f"{ntc_sep_pct:.1f}% of all complaints",
                delta_color="off",
                help=f"NTC complaints for {period3_label}"
            )

        # Data accuracy verification
        st.info(f"ℹ️ **Verification:** {period3_label} ({ntc_sep_count:,}) should be ≤ {period1_label} ({ntc_jan_count:,})")
        if ntc_sep_count > ntc_jan_count:
            st.error(f"⚠️ **Data Error:** {period3_label} has MORE NTC complaints than {period1_label}! Please check the date filtering.")

        st.markdown("---")

        col1, col2 = st.columns(2)

        with col1:
            st.markdown(f"#### Service Providers ({period1_label})")
            # Show date range for {period1_label}
            if len(df_ntc_jan) > 0 and 'Date of Complaint' in df_ntc_jan.columns:
                jan_dates = df_ntc_jan['Date of Complaint'].dropna()
                if len(jan_dates) > 0:
                    date_range = f"{jan_dates.min().strftime('%b %Y')} - {jan_dates.max().strftime('%b %Y')}"
                    st.caption(f"📅 {date_range}")

            if 'Service Providers' in df_ntc_jan.columns and len(df_ntc_jan) > 0:
                valid_data = df_ntc_jan['Service Providers'].dropna()
                valid_data = valid_data[valid_data != '']

                if len(valid_data) > 0:
                    provider_counts = valid_data.value_counts().head(12)
                    fig = create_bar_chart(provider_counts, "Service Provider", 'greens', 380)
                    st.plotly_chart(fig, use_container_width=True, key="ntc_providers_jan")
                else:
                    st.info("📊 No service provider data available")
            else:
                fig = create_bar_chart(pd.Series(), "Service Provider", 'greens', 380)
                st.plotly_chart(fig, use_container_width=True, key="ntc_providers_jan_empty")
                if 'Service Providers' not in df_ntc_jan.columns:
                    st.error("❌ 'Service Providers' column not found")

        with col2:
            st.markdown(f"#### Service Providers ({period3_label})")
            # Show date range for {period3_label}
            if len(df_ntc_sep) > 0 and 'Date of Complaint' in df_ntc_sep.columns:
                sep_dates = df_ntc_sep['Date of Complaint'].dropna()
                if len(sep_dates) > 0:
                    date_range = f"{sep_dates.min().strftime('%b %Y')} - {sep_dates.max().strftime('%b %Y')}"
                    st.caption(f"📅 {date_range}")

            if 'Service Providers' in df_ntc_sep.columns and len(df_ntc_sep) > 0:
                valid_data = df_ntc_sep['Service Providers'].dropna()
                valid_data = valid_data[valid_data != '']

                if len(valid_data) > 0:
                    provider_counts = valid_data.value_counts().head(12)
                    fig = create_bar_chart(provider_counts, "Service Provider", 'teal', 380)
                    st.plotly_chart(fig, use_container_width=True, key="ntc_providers_sep")
                else:
                    st.info("📊 No service provider data available")
            else:
                fig = create_bar_chart(pd.Series(), "Service Provider", 'teal', 380)
                st.plotly_chart(fig, use_container_width=True, key="ntc_providers_sep_empty")
                if 'Service Providers' not in df_ntc_sep.columns:
                    st.error("❌ 'Service Providers' column not found")
    else:
        st.error("❌ 'Agency' column not found in data. Cannot filter NTC complaints.")
        st.info("💡 Please ensure your data has an 'Agency' column.")

    st.markdown("---")

    # ============================================================================
    # SECTION 3: PEMEDES ANALYSIS
    # ============================================================================
    st.markdown("## 📦 PEMEDES Analysis")

    # Filter PEMEDES data with error handling
    if 'Service Providers' in df_period1.columns:
        try:
            # Filter by checking if Service Provider is in PEMEDES_PROVIDERS list
            df_pemedes_jan = df_period1[df_period1['Service Providers'].apply(is_pemedes_provider)]
            df_pemedes_sep = df_period3[df_period3['Service Providers'].apply(is_pemedes_provider)]

            # Data integrity check
            if len(df_pemedes_jan) == 0:
                st.warning(f"⚠️ No PEMEDES complaints found in {period1_label} dataset. This may indicate:")
                st.write("• Service provider names don't match the PEMEDES provider list")
                st.write("• No PEMEDES-related complaints in this period")
                st.write("• Check the 'Service Providers' column format")
        except Exception as e:
            st.error(f"❌ Error filtering PEMEDES data: {str(e)}")
            st.info("Please check if the 'Service Providers' column contains valid text data.")
            df_pemedes_jan = pd.DataFrame()
            df_pemedes_sep = pd.DataFrame()

        # Enhanced KPI metrics for PEMEDES
        kpi_col1, kpi_col2 = st.columns(2)
        with kpi_col1:
            pemedes_jan_count = len(df_pemedes_jan)
            pemedes_jan_pct = (pemedes_jan_count / len(df_period1) * 100) if len(df_period1) > 0 else 0
            st.metric(
                label=f"PEMEDES ({period1_label})",
                value=f"{pemedes_jan_count:,}",
                delta=f"{pemedes_jan_pct:.1f}% of all complaints",
                delta_color="off",
                help=f"PEMEDES complaints for {period1_label}"
            )
        with kpi_col2:
            pemedes_sep_count = len(df_pemedes_sep)
            pemedes_sep_pct = (pemedes_sep_count / len(df_period3) * 100) if len(df_period3) > 0 else 0
            st.metric(
                label=f"PEMEDES ({period3_label})",
                value=f"{pemedes_sep_count:,}",
                delta=f"{pemedes_sep_pct:.1f}% of all complaints",
                delta_color="off",
                help=f"PEMEDES complaints for {period3_label}"
            )

        # Data accuracy verification
        st.info(f"ℹ️ **Verification:** {period3_label} ({pemedes_sep_count:,}) should be ≤ {period1_label} ({pemedes_jan_count:,})")
        if pemedes_sep_count > pemedes_jan_count:
            st.error(f"⚠️ **Data Error:** {period3_label} has MORE PEMEDES complaints than {period1_label}! Please check the date filtering.")

        st.markdown("---")

        col1, col2 = st.columns(2)

        with col1:
            st.markdown(f"#### Service Providers ({period1_label})")
            # Show date range for {period1_label}
            if len(df_pemedes_jan) > 0 and 'Date of Complaint' in df_pemedes_jan.columns:
                jan_dates = df_pemedes_jan['Date of Complaint'].dropna()
                if len(jan_dates) > 0:
                    date_range = f"{jan_dates.min().strftime('%b %Y')} - {jan_dates.max().strftime('%b %Y')}"
                    st.caption(f"📅 {date_range}")

            if 'Service Providers' in df_pemedes_jan.columns and len(df_pemedes_jan) > 0:
                valid_data = df_pemedes_jan['Service Providers'].dropna()
                valid_data = valid_data[valid_data != '']

                if len(valid_data) > 0:
                    provider_counts = valid_data.value_counts().head(12)
                    fig = create_bar_chart(provider_counts, "Service Provider", 'purples', 380)
                    st.plotly_chart(fig, use_container_width=True, key="pemedes_providers_jan")
                else:
                    st.info("📊 No service provider data available")
            else:
                fig = create_bar_chart(pd.Series(), "Service Provider", 'purples', 380)
                st.plotly_chart(fig, use_container_width=True, key="pemedes_providers_jan_empty")
                if 'Service Providers' not in df_pemedes_jan.columns:
                    st.error("❌ 'Service Providers' column not found")

        with col2:
            st.markdown(f"#### Service Providers ({period3_label})")
            # Show date range for {period3_label}
            if len(df_pemedes_sep) > 0 and 'Date of Complaint' in df_pemedes_sep.columns:
                sep_dates = df_pemedes_sep['Date of Complaint'].dropna()
                if len(sep_dates) > 0:
                    date_range = f"{sep_dates.min().strftime('%b %Y')} - {sep_dates.max().strftime('%b %Y')}"
                    st.caption(f"📅 {date_range}")

            if 'Service Providers' in df_pemedes_sep.columns and len(df_pemedes_sep) > 0:
                valid_data = df_pemedes_sep['Service Providers'].dropna()
                valid_data = valid_data[valid_data != '']

                if len(valid_data) > 0:
                    provider_counts = valid_data.value_counts().head(12)
                    fig = create_bar_chart(provider_counts, "Service Provider", 'magenta', 380)
                    st.plotly_chart(fig, use_container_width=True, key="pemedes_providers_sep")
                else:
                    st.info("📊 No service provider data available")
            else:
                fig = create_bar_chart(pd.Series(), "Service Provider", 'magenta', 380)
                st.plotly_chart(fig, use_container_width=True, key="pemedes_providers_sep_empty")
                if 'Service Providers' not in df_pemedes_sep.columns:
                    st.error("❌ 'Service Providers' column not found")
    else:
        st.error("❌ 'Service Providers' column not found in data. Cannot filter PEMEDES complaints.")
        st.info("💡 Please ensure your data has a 'Service Providers' column with PEMEDES provider names.")

    # Display data processing warnings at the bottom
    if data_warnings:
        st.markdown("<br>", unsafe_allow_html=True)
        for warning in data_warnings:
            st.caption(warning)

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
