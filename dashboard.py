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
import os
import ai_reports  # Import the new AI module

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
    "Airspeed",
    "Air21",
    "Black Arrow",
    "GoGo Xpress",
    "SPX",
]

# Mapping for normalizing service provider names to handle duplicates
PROVIDER_ALIASES = {
    # SPX
    'spx': 'SPX',
    'spx express': 'SPX',
    'spx philippines': 'SPX',
    'spx philippines, inc.': 'SPX',
    'shopee express': 'SPX',
    
    # J&T
    'j&t': 'J&T',
    'j&t express': 'J&T',
    'j & t': 'J&T',
    'j & t express': 'J&T',
    'ph global jet express': 'J&T',
    'ph gobal jet express, inc. d.b.a. j&t express': 'J&T',
    
    # LBC
    'lbc': 'LBC',
    'lbc express': 'LBC',
    'lbc express corporation': 'LBC',
    
    # 2GO
    '2go': '2GO',
    '2go express': '2GO',
    '2go express, inc.': '2GO',
    
    # Flash
    'flash': 'Flash Express',
    'flash express': 'Flash Express',
    'flash express (ph) co. ltd., inc.': 'Flash Express',
    
    # Ninja Van
    'ninja van': 'Ninja Van',
    'ninjavan': 'Ninja Van',
    'wall street courier services, inc. d.b.a. ninja van': 'Ninja Van',
    
    # Grab
    'grab': 'Grab Express',
    'grab express': 'Grab Express',
    'grabexpress': 'Grab Express',
    'grabexpress, inc.': 'Grab Express',
    
    # Air21
    'air21': 'Air21',
    'air 21': 'Air21',
    'airfreight 2100, inc. d.b.a. air21': 'Air21',
    
    # GoGo Xpress
    'gogo xpress': 'GoGo Xpress',
    'gogoxpress': 'GoGo Xpress',
    'quadx': 'GoGo Xpress',
    'quadx, inc. d.b.a. gogo xpress': 'GoGo Xpress',
}

# Page configuration
st.set_page_config(
    page_title="Complaint Analysis Dashboard - Real-time",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# CSS Styles - Simple white background design, compressed
CUSTOM_CSS = """
    <style>
    /* Import Google Fonts */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    /* Global Styles */
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
        background-color: white;
    }

    /* Main container styling - Compressed */
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 0rem;
        padding-left: 0.5rem;
        padding-right: 0.5rem;
        max-width: 100%;
        background-color: white;
    }

    /* Sidebar specific styling for compression */
    [data-testid="stSidebar"] .block-container {
        padding-top: 1rem;
        padding-left: 1rem;
        padding-right: 1rem;
    }
    
    [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 {
        margin-top: 0;
        padding-top: 0;
        margin-bottom: 0.5rem;
        font-size: 1.1rem !important;
    }

    /* Compact metrics in sidebar */
    [data-testid="stSidebar"] [data-testid="stMetricValue"] {
        font-size: 1.2rem !important;
        padding: 0 !important;
    }
    
    [data-testid="stSidebar"] [data-testid="stMetricLabel"] {
        font-size: 0.8rem !important;
    }
    
    [data-testid="stSidebar"] div[data-testid="stMetric"] {
        padding: 0.5rem !important;
        margin: 0 !important;
        border: none !important;
    }
    
    /* Reduce spacing between elements in sidebar */
    [data-testid="stSidebar"] .stElementContainer {
        margin-bottom: 0.5rem;
    }
    
    [data-testid="stSidebar"] hr {
        margin: 0.5rem 0;
    }

    /* Header styling - Hierarchical sizes */
    h1 {
        padding-top: 0.25rem;
        margin-top: 0rem;
        margin-bottom: 0.75rem;
        font-weight: 700;
        font-size: 2rem !important;
        color: #1f2937;
    }

    h2 {
        font-weight: 600;
        color: #374151;
        margin-top: 0.5rem;
        margin-bottom: 0.5rem;
        font-size: 1.5rem;
    }

    h3 {
        font-weight: 600;
        color: #374151;
        margin-top: 0.5rem;
        margin-bottom: 0.5rem;
        font-size: 1.3rem;
    }

    h4 {
        font-weight: 600;
        color: #374151;
        margin-top: 0.25rem;
        margin-bottom: 0.5rem;
        font-size: 1.1rem;
    }

    /* Metric cards - Compressed */
    [data-testid="stMetricValue"] {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1f2937;
    }

    [data-testid="stMetricLabel"] {
        font-size: 0.9rem;
        font-weight: 600;
        color: #6b7280;
    }

    div[data-testid="stMetric"] {
        background-color: white;
        padding: 1rem !important;
        border-radius: 8px !important;
        border: 1px solid #e5e7eb !important;
        margin: 0.2rem;
    }

    /* Chart container - Compressed */
    .stPlotlyChart {
        background-color: white;
        border-radius: 8px;
        padding: 2px;
        border: 1px solid #e5e7eb;
        margin-bottom: 0.3rem;
    }

    /* Sidebar - Simple white */
    [data-testid="stSidebar"] {
        background-color: white;
    }

    /* Button styling */
    .stButton>button {
        border-radius: 8px;
        font-weight: 600;
        border: 1px solid #e5e7eb;
        padding: 0.5rem 1rem;
        background-color: white;
    }

    /* Divider styling - Compressed */
    hr {
        margin: 0.5rem 0;
        border: none;
        border-top: 1px solid #e5e7eb;
    }
    </style>
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

@st.cache_data(ttl=60, show_spinner=False)
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
            
            return df
            
        except Exception as e:
            st.error(f"❌ Cannot access sheet. Please ensure it's set to 'Anyone with the link can view'")
            st.info("To make your sheet public: File → Share → Change to 'Anyone with the link'")
            return None
            
    except Exception as e:
        st.error(f"❌ Error loading public sheet: {str(e)}")
        return None

@st.cache_data(ttl=60, show_spinner=False)
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

    # Clean up common timezone typos/artifacts
    # "APM" is likely a typo for "PM" or "AM" depending on context, but usually PM if it appears at end
    # We'll replace "APM" with " PM" to be safe
    date_str = date_str.replace('APM', ' PM').replace('A M', ' AM').replace('P M', ' PM')

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

@st.cache_data(show_spinner=False)
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

    # Normalize Service Providers to handle duplicates/variations
    if 'Service Providers' in df.columns:
        # Helper function to normalize provider names
        def normalize_provider(name):
            if pd.isna(name) or name == '':
                return name
            name_str = str(name).strip()
            name_lower = name_str.lower()
            return PROVIDER_ALIASES.get(name_lower, name_lower)

        df['Service Providers'] = df['Service Providers'].apply(normalize_provider)

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
        # Add validation to ensure Date of Complaint column contains valid dates
        if df['Date of Complaint'].dtype != 'datetime64[ns]':
            st.warning("Date of Complaint column contains non-datetime values. Attempting conversion...")
            df['Date of Complaint'] = pd.to_datetime(df['Date of Complaint'], errors='coerce')

        filtered_df = df[df['Date of Complaint'] >= start_date].copy()
        return filtered_df
    except Exception as e:
        st.error(f"Error filtering by date: {str(e)}")
        return df

def create_bar_chart(data_series, title, color_scale='blues', height=500):
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
            title_font=dict(size=14),
            tickangle=0,  # Force numbers horizontal
            tickfont=dict(size=12),
            automargin=True
        ),
        yaxis=dict(
            categoryorder='total ascending',
            showgrid=False,
            tickfont=dict(size=13),
            tickangle=0,  # Force labels horizontal
            automargin=True
        ),
        coloraxis_showscale=False,
        uniformtext_minsize=8,
        uniformtext_mode='hide'
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

def create_pie_chart(data_series, title, color_scheme=None, height=500):
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

def create_line_chart(df_monthly, height=500):
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
            tickfont=dict(size=12),
            tickangle=0,  # Force labels horizontal
            automargin=True
        ),
        yaxis=dict(
            showgrid=True,
            gridcolor='#f3f4f6',
            zeroline=False,
            title='Complaints',
            title_font=dict(size=13),
            tickfont=dict(size=12),
            tickangle=0,  # Force numbers horizontal
            automargin=True
        ),
        hovermode='x unified',
        uniformtext_minsize=8,
        uniformtext_mode='hide'
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

def render_comparison_charts(df_period1, df_period3, period1_label, period3_label, color_theme1, color_theme2, key_prefix):
    """Render side-by-side service provider charts for comparison"""
    col1, col2 = st.columns(2)

    with col1:
        st.markdown(f"#### Service Providers ({period1_label})")
        if len(df_period1) > 0 and 'Date of Complaint' in df_period1.columns:
            period1_dates = df_period1['Date of Complaint'].dropna()
            if len(period1_dates) > 0:
                date_range = f"{period1_dates.min().strftime('%b %Y')} - {period1_dates.max().strftime('%b %Y')}"
                st.caption(f"📅 {date_range}")

        if 'Service Providers' in df_period1.columns and len(df_period1) > 0:
            valid_data = df_period1['Service Providers'].dropna()
            valid_data = valid_data[valid_data != '']

            if len(valid_data) > 0:
                provider_counts = valid_data.value_counts().head(12)
                fig = create_bar_chart(provider_counts, "Service Provider", color_theme1, 400)
                st.plotly_chart(fig, use_container_width=True, key=f"{key_prefix}_period1")
            else:
                st.info("📊 No service provider data available")
        else:
            fig = create_bar_chart(pd.Series(), "Service Provider", color_theme1, 400)
            st.plotly_chart(fig, use_container_width=True, key=f"{key_prefix}_period1_empty")
            if 'Service Providers' not in df_period1.columns:
                st.error("❌ 'Service Providers' column not found")

    with col2:
        st.markdown(f"#### Service Providers ({period3_label})")
        if len(df_period3) > 0 and 'Date of Complaint' in df_period3.columns:
            period3_dates = df_period3['Date of Complaint'].dropna()
            if len(period3_dates) > 0:
                date_range = f"{period3_dates.min().strftime('%b %Y')} - {period3_dates.max().strftime('%b %Y')}"
                st.caption(f"📅 {date_range}")

        if 'Service Providers' in df_period3.columns and len(df_period3) > 0:
            valid_data = df_period3['Service Providers'].dropna()
            valid_data = valid_data[valid_data != '']

            if len(valid_data) > 0:
                provider_counts = valid_data.value_counts().head(12)
                fig = create_bar_chart(provider_counts, "Service Provider", color_theme2, 400)
                st.plotly_chart(fig, use_container_width=True, key=f"{key_prefix}_period3")
            else:
                st.info("📊 No service provider data available")
        else:
            fig = create_bar_chart(pd.Series(), "Service Provider", color_theme2, 400)
            st.plotly_chart(fig, use_container_width=True, key=f"{key_prefix}_period3_empty")
            if 'Service Providers' not in df_period3.columns:
                st.error("❌ 'Service Providers' column not found")

def main():
    # Initialize session state FIRST - before any other code
    if 'auto_refresh' not in st.session_state:
        st.session_state.auto_refresh = True
    if 'refresh_interval' not in st.session_state:
        st.session_state.refresh_interval = 300
    if 'data_source_type' not in st.session_state:
        st.session_state.data_source_type = None
    if 'file_path' not in st.session_state:
        st.session_state.file_path = None
    if 'gsheet_creds' not in st.session_state:
        st.session_state.gsheet_creds = None
    if 'gsheet_url' not in st.session_state:
        st.session_state.gsheet_url = ""
    if 'use_public_sheet' not in st.session_state:
        st.session_state.use_public_sheet = True
    if 'view_mode' not in st.session_state:
        st.session_state.view_mode = "Dashboard"
    if 'last_valid_df' not in st.session_state:
        st.session_state.last_valid_df = None

    # Apply custom CSS styles
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

    # Load default Google Sheets URL from Streamlit secrets or environment variables
    # SECURITY: URL should be stored in .streamlit/secrets.toml (for deployments) or .env (for local)
    DEFAULT_GSHEET_URL = ""
    try:
        # Try Streamlit secrets first (for deployed apps)
        if hasattr(st, 'secrets') and 'DEFAULT_GSHEET_URL' in st.secrets:
            DEFAULT_GSHEET_URL = st.secrets["DEFAULT_GSHEET_URL"]
    except:
        pass

    # Fall back to environment variable (for local development)
    if not DEFAULT_GSHEET_URL:
        DEFAULT_GSHEET_URL = os.getenv("DEFAULT_GSHEET_URL", "")

    DEFAULT_SHEET_AVAILABLE = bool(DEFAULT_GSHEET_URL)

    # Prefill session state with default URL on first load
    if DEFAULT_SHEET_AVAILABLE:
        if 'gsheet_url' not in st.session_state or not st.session_state.gsheet_url:
            st.session_state.gsheet_url = DEFAULT_GSHEET_URL
    
    # Dashboard header
    st.markdown('<h1 style="font-size: 2rem; font-weight: 700; margin-bottom: 0.75rem;">📊 Complaint Analysis Dashboard</h1>', unsafe_allow_html=True)
    
    # Navigation Tabs
    tab_dashboard, tab_ai_report = st.tabs(["Dashboard", "AI Action Plan"])

    # Sidebar for data loading and settings
    with st.sidebar:
        st.markdown("### ⚙️ Settings")
        
        # Simplified Data Source Selection
        st.caption("Data Source")
        data_source = st.selectbox(
            "Data Source", 
            ["Google Sheets (Public)", "Upload Excel", "Excel File Path", "Google Sheets (Private)"],
            index=0,
            label_visibility="collapsed"
        )
        
        df = None
        
        if data_source == "Google Sheets (Public)":
            # Use default URL if available, but don't display it in the input field
            spreadsheet_url = st.text_input(
                "Sheet URL",
                value="",  # Always show blank input field
                placeholder="Paste 'Anyone with the link' URL here...",
                label_visibility="collapsed"
            )

            # Use user input if provided, otherwise use default from env
            active_url = spreadsheet_url if spreadsheet_url else st.session_state.gsheet_url

            if active_url:
                if spreadsheet_url:  # Only update session state if user provides URL
                    st.session_state.gsheet_url = spreadsheet_url
                interval = st.session_state.refresh_interval if st.session_state.auto_refresh else 60
                current_time = int(datetime.now().timestamp() // interval) * interval
                df = load_data_from_public_gsheet(active_url, current_time)
                st.session_state.data_source_type = "gsheets_public"
        
        elif data_source == "Upload Excel":
            uploaded_file = st.file_uploader("Upload .xlsx", type=['xlsx', 'xls'], label_visibility="collapsed")
            if uploaded_file is not None:
                df = load_data_from_uploaded_excel(uploaded_file)
                st.session_state.data_source_type = "upload"
        
        elif data_source == "Excel File Path":
            file_path = st.text_input("File Path", value=st.session_state.file_path or "", placeholder="C:/path/to/file.xlsx", label_visibility="collapsed")
            if file_path:
                st.session_state.file_path = file_path
                df = load_data_from_excel(file_path)
                st.session_state.data_source_type = "filepath"
        
        else:  # Google Sheets (Private)
            if st.session_state.gsheet_creds is None:
                credentials_file = st.file_uploader("Service Account JSON", type=['json'], label_visibility="collapsed")
                if credentials_file:
                    import json
                    st.session_state.gsheet_creds = json.load(credentials_file)
                    st.success("Credentials loaded")
            else:
                if st.button("Clear Credentials", use_container_width=True):
                    st.session_state.gsheet_creds = None
                    st.rerun()
            
            # Use default URL if available, but don't display it in the input field
            spreadsheet_url = st.text_input(
                "Sheet URL",
                value="",  # Always show blank input field
                placeholder="Paste Sheet URL here...",
                label_visibility="collapsed"
            )

            # Use user input if provided, otherwise use default from env
            if spreadsheet_url:
                st.session_state.gsheet_url = spreadsheet_url

            active_url = st.session_state.gsheet_url

            if st.session_state.gsheet_creds and active_url:
                interval = st.session_state.refresh_interval if st.session_state.auto_refresh else 60
                current_time = int(datetime.now().timestamp() // interval) * interval
                df = load_data_from_gsheet_with_auth(st.session_state.gsheet_creds, active_url, current_time)
                st.session_state.data_source_type = "gsheets_private"
        
        if df is not None and not df.empty:
            st.session_state.last_valid_df = df
            st.divider()
            
            # Row 2: Refresh Controls
            st.markdown("**Auto-Refresh**")
            col_r1, col_r2 = st.columns([3, 1])
            with col_r1:
                auto_refresh = st.toggle("Enable", value=st.session_state.auto_refresh)
                st.session_state.auto_refresh = auto_refresh
            with col_r2:
                if st.button("↻", help="Refresh Now"):
                    st.cache_data.clear()
                    st.rerun()
            
            if auto_refresh:
                st.select_slider(
                    "Interval (seconds)", 
                    options=[30, 60, 120, 300], 
                    value=st.session_state.refresh_interval,
                    key="refresh_interval_slider",
                    on_change=lambda: st.session_state.update(refresh_interval=st.session_state.refresh_interval_slider)
                )

            # Compact Data Preview
            with st.expander("📋 View Data", expanded=False):
                st.dataframe(df.head(3), width='stretch', height=150)

    # Determine refresh rate for the fragment
    # If data source is uploaded file, auto-refresh is not needed unless explicitly desired (but file won't change)
    # For now, we respect the user's auto-refresh setting for all sources
    refresh_rate = st.session_state.refresh_interval if st.session_state.auto_refresh else None

    @st.fragment(run_every=refresh_rate)
    def render_dashboard_content(initial_df=None):
        # Re-calculate timestamp for cache busting within the fragment
        interval = st.session_state.refresh_interval if st.session_state.auto_refresh else 60
        current_time = int(datetime.now().timestamp() // interval) * interval
        
        df = None
        source_type = st.session_state.get('data_source_type')
        
        # Logic to determine df
        if source_type == 'upload':
            # For upload, we rely on the dataframe passed from the main script
            df = initial_df
        elif source_type == 'gsheets_public':
            url = st.session_state.get('gsheet_url')
            if url:
                df = load_data_from_public_gsheet(url, current_time)
        elif source_type == 'filepath':
            path = st.session_state.get('file_path')
            if path:
                df = load_data_from_excel(path)
        elif source_type == 'gsheets_private':
            creds = st.session_state.get('gsheet_creds')
            url = st.session_state.get('gsheet_url')
            if creds and url:
                df = load_data_from_gsheet_with_auth(creds, url, current_time)
        
        # Don't use stale data - if reload fails, show error instead
        # Using old data could mislead users about real-time status
        # if df is None and initial_df is not None and source_type != 'upload':
        #    st.warning("⚠️ Unable to refresh data. Please check your connection and data source.")
        #    df = initial_df  # Use cached data but warn user

        # Update last_valid_df if we got new data
        if df is not None and not df.empty:
            st.session_state.last_valid_df = df
        
        # Fallback to last_valid_df if current fetch failed
        if (df is None or df.empty) and st.session_state.last_valid_df is not None:
            df = st.session_state.last_valid_df
            
        # Fallback to initial_df if everything else failed
        if (df is None or df.empty) and initial_df is not None and not initial_df.empty:
             df = initial_df
             st.session_state.last_valid_df = df

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
            """)
            return

        # Prepare data
        # We process the data here to ensure fresh data is always cleaned
        # Use a subtle spinner or none for auto-updates to avoid flickering
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

        # Summary metrics with error handling and enhanced design
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            period1_count = len(df_period1)
            period2_count = len(df_period2)
            delta_count = period1_count - period2_count
            st.metric(
                label=f"{period1_label}",
                value=f"{period1_count:,}",
                delta=f"{delta_count:+,} vs {period2_short}",
                delta_color="inverse",
                help=f"Total complaints for {period1_label}"
            )

        with col2:
            period2_count = len(df_period2)
            period2_pct = (period2_count / period1_count * 100) if period1_count > 0 else 0
            st.metric(
                label=f"{period2_label}",
                value=f"{period2_count:,}",
                delta=f"{period2_pct:.1f}% of {period1_short}",
                delta_color="off",
                help=f"Total complaints for {period2_label}"
            )

        with col3:
            if 'Agency' in df_period1.columns:
                try:
                    ntc_mask = df_period1['Agency'].apply(is_ntc_complaint)
                    if 'Complaint Category' in df_period1.columns:
                        ntc_mask = ntc_mask | (df_period1['Complaint Category'].astype(str).str.strip().str.upper() == "TELCO INTERNET ISSUES")
                    ntc_count = len(df_period1[ntc_mask])
                    ntc_pct = (ntc_count / period1_count * 100) if period1_count > 0 else 0
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
                    # Count PEMEDES complaints (providers OR "Delivery Concerns (SP)" OR Agency="PRD")
                    pemedes_mask = df_period1['Service Providers'].apply(is_pemedes_provider)
                    if 'Complaint Category' in df_period1.columns:
                        pemedes_mask = pemedes_mask | (df_period1['Complaint Category'].astype(str).str.strip().str.upper() == "DELIVERY CONCERNS (SP)")
                    if 'Agency' in df_period1.columns:
                        pemedes_mask = pemedes_mask | (df_period1['Agency'].astype(str).str.strip().str.upper() == "PRD")
                    pemedes_count = len(df_period1[pemedes_mask])
                    pemedes_pct = (pemedes_count / period1_count * 100) if period1_count > 0 else 0
                    st.metric(
                        label="PEMEDES Resolved Complaints",
                        value=f"{pemedes_count:,}",
                        delta=f"{pemedes_pct:.1f}% of total",
                        delta_color="off",
                        help="Complaints from PEMEDES service providers, Delivery Concerns (SP), or Agency PRD"
                    )
                except Exception as e:
                    st.metric("PEMEDES Resolved Complaints", "Error")
                    st.error(f"Error counting PEMEDES resolved complaints: {str(e)}")
            else:
                st.metric("PEMEDES Resolved Complaints", "N/A")

        # Overall charts - Compact design
        chart_height = 350

        # Row 1: Category and Nature
        col1, col2 = st.columns(2)

        with col1:
            st.markdown(f"#### Complaints by Category ({period1_label})")
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
            st.markdown(f"#### Complaints by Nature ({period1_label})")
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

        # Row 2: DICT Unit
        st.markdown(f"#### Complaints by DICT Unit ({period1_label})")
        if 'DICT UNIT' in df_period1.columns:
            valid_data = df_period1['DICT UNIT'].dropna()
            valid_data = valid_data[valid_data != '']
            if len(valid_data) > 0:
                unit_counts = valid_data.value_counts().head(10)
                fig = create_bar_chart(unit_counts, "DICT Unit", 'oranges', chart_height)
                st.plotly_chart(fig, use_container_width=True, key="overall_dict_unit")
            else:
                st.info("No DICT Unit data available")
        else:
            st.info("'DICT UNIT' column not found in data")

        # Monthly Trend
        st.markdown(f"#### Monthly Complaint Trend ({period1_label})")
        if 'Date of Complaint' in df_period1.columns:
            valid_dates = df_period1['Date of Complaint'].dropna()
            if len(valid_dates) > 0:
                monthly_data = df_period1.groupby(df_period1['Date of Complaint'].dt.to_period('M')).size()
                if len(monthly_data) > 0:
                    df_monthly = pd.DataFrame({
                        'Month': monthly_data.index.astype(str),
                        'Count': monthly_data.values
                    })
                    fig = create_line_chart(df_monthly, 350)
                    st.plotly_chart(fig, use_container_width=True, key="overall_monthly_trend")
                else:
                    st.info("No monthly data available")
            else:
                st.info("No valid complaint dates found")
        else:
            st.error("'Date of Complaint' column not found")

        st.markdown("---")

        # NTC Analysis
        st.markdown("### NTC Analysis")

        # Filter NTC data with error handling
        if 'Agency' in df_period1.columns:
            try:
                # Use the is_ntc_complaint function for consistent filtering
                # Also include "Telco Internet Issues" from Complaint Category
                ntc_mask_period1 = df_period1['Agency'].apply(is_ntc_complaint)
                if 'Complaint Category' in df_period1.columns:
                    ntc_mask_period1 = ntc_mask_period1 | (df_period1['Complaint Category'].astype(str).str.strip().str.upper() == "TELCO INTERNET ISSUES")
                df_ntc_period1 = df_period1[ntc_mask_period1]

                ntc_mask_period3 = df_period3['Agency'].apply(is_ntc_complaint)
                if 'Complaint Category' in df_period3.columns:
                    ntc_mask_period3 = ntc_mask_period3 | (df_period3['Complaint Category'].astype(str).str.strip().str.upper() == "TELCO INTERNET ISSUES")
                df_ntc_period3 = df_period3[ntc_mask_period3]

                # Data integrity check
                if len(df_ntc_period1) == 0:
                    st.warning(f"⚠️ No NTC complaints found in {period1_label} dataset. This may indicate:")
                    st.write("• Agency column doesn't contain 'NTC'")
                    st.write("• No NTC-related complaints in this period")
                    st.write("• Check the 'Agency' column format")
            except Exception as e:
                st.error(f"❌ Error filtering NTC data: {str(e)}")
                st.info("Please check if the 'Agency' column contains valid text data.")
                df_ntc_period1 = pd.DataFrame()
                df_ntc_period3 = pd.DataFrame()

            # Enhanced KPI metrics for NTC
            kpi_col1, kpi_col2 = st.columns(2)
            with kpi_col1:
                ntc_period1_count = len(df_ntc_period1)
                ntc_period1_pct = (ntc_period1_count / len(df_period1) * 100) if len(df_period1) > 0 else 0
                st.metric(
                    label=f"NTC ({period1_label})",
                    value=f"{ntc_period1_count:,}",
                    delta=f"{ntc_period1_pct:.1f}% of all complaints",
                    delta_color="off",
                    help=f"NTC complaints for {period1_label}"
                )
            with kpi_col2:
                ntc_period3_count = len(df_ntc_period3)
                ntc_period3_pct = (ntc_period3_count / len(df_period3) * 100) if len(df_period3) > 0 else 0
                st.metric(
                    label=f"NTC ({period3_label})",
                    value=f"{ntc_period3_count:,}",
                    delta=f"{ntc_period3_pct:.1f}% of all complaints",
                    delta_color="off",
                    help=f"NTC complaints for {period3_label}"
                )

            st.markdown("---")

            # Use helper function for charts
            render_comparison_charts(
                df_ntc_period1, df_ntc_period3, 
                period1_label, period3_label, 
                'greens', 'teal', 
                "ntc_providers"
            )
        else:
            st.error("❌ 'Agency' column not found in data. Cannot filter NTC complaints.")
            st.info("💡 Please ensure your data has an 'Agency' column.")

        st.markdown("---")

        # PEMEDES Analysis
        st.markdown("### PEMEDES Analysis")

        # Filter PEMEDES data with error handling
        if 'Service Providers' in df_period1.columns:
            try:
                # Filter by checking if Service Provider is in PEMEDES_PROVIDERS list
                # OR if Complaint Category is "Delivery Concerns (SP)"
                # OR if Agency is "PRD"
                pemedes_mask_period1 = df_period1['Service Providers'].apply(is_pemedes_provider)
                if 'Complaint Category' in df_period1.columns:
                    pemedes_mask_period1 = pemedes_mask_period1 | (df_period1['Complaint Category'].astype(str).str.strip().str.upper() == "DELIVERY CONCERNS (SP)")
                if 'Agency' in df_period1.columns:
                    pemedes_mask_period1 = pemedes_mask_period1 | (df_period1['Agency'].astype(str).str.strip().str.upper() == "PRD")
                df_pemedes_period1 = df_period1[pemedes_mask_period1]

                pemedes_mask_period3 = df_period3['Service Providers'].apply(is_pemedes_provider)
                if 'Complaint Category' in df_period3.columns:
                    pemedes_mask_period3 = pemedes_mask_period3 | (df_period3['Complaint Category'].astype(str).str.strip().str.upper() == "DELIVERY CONCERNS (SP)")
                if 'Agency' in df_period3.columns:
                    pemedes_mask_period3 = pemedes_mask_period3 | (df_period3['Agency'].astype(str).str.strip().str.upper() == "PRD")
                df_pemedes_period3 = df_period3[pemedes_mask_period3]

                # Data integrity check
                if len(df_pemedes_period1) == 0:
                    st.warning(f"⚠️ No PEMEDES resolved complaints found in {period1_label} dataset. This may indicate:")
                    st.write("• Service provider names don't match the PEMEDES provider list")
                    st.write("• No PEMEDES-related complaints in this period")
                    st.write("• Check the 'Service Providers' column format")
            except Exception as e:
                st.error(f"❌ Error filtering PEMEDES data: {str(e)}")
                st.info("Please check if the 'Service Providers' column contains valid text data.")
                df_pemedes_period1 = pd.DataFrame()
                df_pemedes_period3 = pd.DataFrame()

            # Enhanced KPI metrics for PEMEDES
            kpi_col1, kpi_col2 = st.columns(2)
            with kpi_col1:
                pemedes_period1_count = len(df_pemedes_period1)
                pemedes_period1_pct = (pemedes_period1_count / len(df_period1) * 100) if len(df_period1) > 0 else 0
                st.metric(
                    label=f"PEMEDES ({period1_label})",
                    value=f"{pemedes_period1_count:,}",
                    delta=f"{pemedes_period1_pct:.1f}% of all complaints",
                    delta_color="off",
                    help=f"PEMEDES resolved complaints for {period1_label}"
                )
            with kpi_col2:
                pemedes_period3_count = len(df_pemedes_period3)
                pemedes_period3_pct = (pemedes_period3_count / len(df_period3) * 100) if len(df_period3) > 0 else 0
                st.metric(
                    label=f"PEMEDES ({period3_label})",
                    value=f"{pemedes_period3_count:,}",
                    delta=f"{pemedes_period3_pct:.1f}% of all complaints",
                    delta_color="off",
                    help=f"PEMEDES resolved complaints for {period3_label}"
                )

            st.markdown("---")

            # Use helper function for charts
            render_comparison_charts(
                df_pemedes_period1, df_pemedes_period3, 
                period1_label, period3_label, 
                'purples', 'magenta', 
                "pemedes_providers"
            )
        else:
            st.error("❌ 'Service Providers' column not found in data. Cannot filter PEMEDES resolved complaints.")
            st.info("💡 Please ensure your data has a 'Service Providers' column with PEMEDES provider names.")

        # Display data processing warnings at the bottom
        if data_warnings:
            st.markdown("<br>", unsafe_allow_html=True)
            for warning in data_warnings:
                st.caption(warning)

        # Data Validation & Integrity - Moved to bottom
        st.markdown("---")
        with st.expander("📊 Data Validation & Integrity", expanded=False):
            # Period summaries
            col_v1, col_v2, col_v3 = st.columns(3)

            for col, df_period, label in [(col_v1, df_period1, period1_label),
                                        (col_v2, df_period2, period2_label),
                                        (col_v3, df_period3, period3_label)]:
                with col:
                    st.write(f"**{label}**")
                    if len(df_period) > 0 and 'Date of Complaint' in df_period.columns:
                        dates = df_period['Date of Complaint'].dropna()
                        if len(dates) > 0:
                            st.caption(f"{dates.min().strftime('%Y-%m-%d')} to {dates.max().strftime('%Y-%m-%d')}")
                            st.metric("Records", f"{len(df_period):,}", label_visibility="collapsed")

            st.markdown("---")

            # Filter validation and integrity checks
            integrity_col1, integrity_col2 = st.columns(2)

            with integrity_col1:
                st.write("**Filter Validation:**")
                if 'Agency' in df_period1.columns:
                    ntc_mask = df_period1['Agency'].apply(is_ntc_complaint)
                    if 'Complaint Category' in df_period1.columns:
                        ntc_mask = ntc_mask | (df_period1['Complaint Category'].astype(str).str.strip().str.upper() == "TELCO INTERNET ISSUES")
                    ntc_count = len(df_period1[ntc_mask])
                    st.write(f"NTC: {ntc_count:,} ({(ntc_count/len(df_period1)*100):.1f}%)")

                if 'Service Providers' in df_period1.columns:
                    # Use consistent PEMEDES filter
                    pemedes_mask = df_period1['Service Providers'].apply(is_pemedes_provider)
                    if 'Complaint Category' in df_period1.columns:
                        pemedes_mask = pemedes_mask | (df_period1['Complaint Category'].astype(str).str.strip().str.upper() == "DELIVERY CONCERNS (SP)")
                    if 'Agency' in df_period1.columns:
                        pemedes_mask = pemedes_mask | (df_period1['Agency'].astype(str).str.strip().str.upper() == "PRD")
                    pem_count = len(df_period1[pemedes_mask])
                   
                    st.write(f"PEMEDES: {pem_count:,} ({(pem_count/len(df_period1)*100):.1f}%)")

            with integrity_col2:
                st.write("**Integrity Check:**")
                # Verify Period3 ≤ Period1 (subset relationship)
                subset_valid = len(df_period3) <= len(df_period1)
                if subset_valid:
                    st.success(f"✓ {period3_short} ⊆ {period1_short}")
                else:
                    st.error(f"✗ {period3_short} > {period1_short}")

                # Verify Period2 relationship
                period2_valid = len(df_period2) <= len(df_period1)
                if period2_valid:
                    st.success(f"✓ {period2_short} ⊆ {period1_short}")
                else:
                    st.error(f"✗ {period2_short} > {period1_short}")

        # Footer
        st.markdown("---")
        st.markdown(f"*Dashboard last updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}*")

    # Prepare data for both tabs if available
    df_prepared = None
    if df is not None and not df.empty:
        df_prepared, _ = prepare_data(df)

    # Render Dashboard Tab
    with tab_dashboard:
        render_dashboard_content(df)

    # Render AI Action Plan Tab
    with tab_ai_report:
        # Pass the prepared dataframe to the AI report module
        ai_reports.render_weekly_report(df_prepared)

if __name__ == "__main__":
    main()
