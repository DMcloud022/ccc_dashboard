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
    'E': 'Date Received', 'F': 'Resolution', 'G': 'Customer Name',
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
    "Ninja Van",
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

# NTC Service Providers List
NTC_PROVIDERS = [
    "PLDT",
    "Converge",
    "Globe",
    "Smart",
    "DITO",
    "Sky Cable",
    "Cignal",
    "Eastern",
    "DITO Telecommunity",
    "Globe Telecom",
    "Smart Communications",
    "PLDT Inc.",
    "Converge ICT Solutions",
    "Sky Fiber",
    "Royal Cable",
    "Radius Telecoms",
    "PT&T",
    "Now Telecom"
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
    
    # Converge consolidation
    'converge ict': 'Converge',
    'converge': 'Converge',
    
    # Globe consolidation  
    'globe': 'Globe Telecom',
    'globe telecom': 'Globe Telecom',
    
    # TNT Express consolidation
    'tnt express deliveries': 'TNT Express Deliveries (Phils.), Inc.',
    'tnt express deliveries (phils.), inc.': 'TNT Express Deliveries (Phils.), Inc.',
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

    [data-testid="stMetricDelta"] svg {
        display: none;
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
    
    [data-testid="stSidebar"][aria-expanded="true"] {
        min-width: 340px;
        max-width: 340px;
    }
    
    [data-testid="stSidebar"][aria-expanded="true"] > div:first-child {
        width: 340px;
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

def is_ntc_provider(service_provider):
    """Check if a service provider is an NTC provider"""
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

    for ntc_sp in NTC_PROVIDERS:
        ntc_sp_lower = ntc_sp.lower()
        
        # Check for exact match or containment
        if ntc_sp_lower == service_provider_lower:
            return True
        if ntc_sp_lower in service_provider_lower:
            return True
        if service_provider_lower in ntc_sp_lower and len(service_provider_lower) > 3:
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

@st.cache_data(ttl=300, show_spinner=False)
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

@st.cache_data(ttl=300, show_spinner=False)
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
        'Date Received': ['Date Received', 'Date of Complaint', 'Complaint Date', 'Date Filed', 'Filing Date'],
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
    date_columns = ['Date Received', 'Date of Resolution', 'Date Responded']

    for col in date_columns:
        if col in df.columns:
            # Apply robust date parsing
            df[col] = df[col].apply(parse_date_robust)

    # Extract year and month for filtering
    if 'Date Received' in df.columns:
        df['Year'] = df['Date Received'].dt.year
        df['Month'] = df['Date Received'].dt.month

    # Clean text columns (remove extra whitespace) and ensure capitalization
    text_columns = ['Agency', 'Service Providers', 'Complaint Category', 'Complaint Nature', 'DICT UNIT']
    for col in text_columns:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace(['nan', 'None', '', 'NaN', 'NaT'], np.nan)
            
            # Capitalize first letter for specific columns (Sentence case)
            if col in ['Complaint Category', 'Complaint Nature', 'DICT UNIT', 'Agency']:
                # Capitalize first letter, leave rest as is (preserves acronyms like DICT)
                df[col] = df[col].apply(lambda x: str(x)[0].upper() + str(x)[1:] if pd.notna(x) and len(str(x)) > 0 else x)

    # Normalize Service Providers to handle duplicates/variations
    if 'Service Providers' in df.columns:
        # Helper function to normalize provider names
        def normalize_provider(name):
            if pd.isna(name) or name == '':
                return name
            name_str = str(name).strip()
            name_lower = name_str.lower()
            
            # Check aliases first
            if name_lower in PROVIDER_ALIASES:
                return PROVIDER_ALIASES[name_lower]
            
            # If not in aliases, return Title Case to ensure capitalization
            # This handles "unknown courier" -> "Unknown Courier"
            return name_str.title()

        df['Service Providers'] = df['Service Providers'].apply(normalize_provider)

    # Remove rows where Date Received is invalid
    if 'Date Received' in df.columns:
        rows_before = len(df)
        df = df[df['Date Received'].notna()]
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

    if 'Date Received' not in df.columns:
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
        # Add validation to ensure Date Received column contains valid dates
        if df['Date Received'].dtype != 'datetime64[ns]':
            st.warning("Date Received column contains non-datetime values. Attempting conversion...")
            df['Date Received'] = pd.to_datetime(df['Date Received'], errors='coerce')

        filtered_df = df[df['Date Received'] >= start_date].copy()
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
        margin=dict(l=3, r=20, t=3, b=3),  # Optimized margin for better space utilization
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(family='Inter, sans-serif', size=13, color='#374151'),
        xaxis=dict(
            showgrid=True,
            gridcolor='#f3f4f6',
            zeroline=False,
            title_font=dict(size=13),
            tickangle=0,  # Force numbers horizontal
            tickfont=dict(size=13),
            automargin=True,
            fixedrange=True
        ),
        yaxis=dict(
            categoryorder='total ascending',
            showgrid=False,
            tickfont=dict(size=13),
            tickangle=0,  # Force labels horizontal
            automargin=True,
            autorange=True
        ),
        coloraxis_showscale=False,
        uniformtext_minsize=8,
        uniformtext_mode='hide'
    )
    # Calculate percentages and update text
    total = data_series.sum()
    percentages = []
    for value in data_series.values:
        percentage = (value / total * 100) if total > 0 else 0
        percentages.append(f"{percentage:.1f}%")
    
    fig.update_traces(
        marker=dict(
            line=dict(width=0)
        ),
        text=percentages,
        texttemplate='%{text}',  # Use custom text with percentages
        textposition='inside',  # Position labels inside bars to save space
        textfont=dict(size=13, family='Inter, sans-serif', weight='bold'),
        hovertemplate=f'<b>%{{y}}</b><br>Count: %{{x:,}} out of {total:,}<br>Percentage: %{{text}}<extra></extra>',
        cliponaxis=False  # Don't clip labels outside the plot area
    )
    return fig

def create_status_stacked_bar_chart(df, category_col, title, height=500, max_items=8, color_theme='blue'):
    """Create a stacked bar chart split by Status (Closed/Open)"""
    if df is None or df.empty or category_col not in df.columns or 'Status' not in df.columns:
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

    # Filter and clean data
    df_chart = df[[category_col, 'Status']].copy()
    df_chart = df_chart.dropna(subset=[category_col])
    df_chart = df_chart[df_chart[category_col] != '']
    
    # Normalize Status - only include actual Open/Closed statuses
    def normalize_status(status):
        s = str(status).strip().lower()
        if 'closed' in s or 'resolved' in s:
            return 'Closed'
        elif 'open' in s:
            return 'Open'
        else:
            # Return None for statuses like 'Received', 'Pending', etc.
            return None
    
    df_chart['Status_Clean'] = df_chart['Status'].apply(normalize_status)
    
    # Filter out records with non-Open/Closed statuses
    df_chart = df_chart.dropna(subset=['Status_Clean'])
    
    # Get top categories by total count
    top_categories = df_chart[category_col].value_counts().head(max_items).index.tolist()
    df_chart = df_chart[df_chart[category_col].isin(top_categories)]
    
    # Group by Category and Status
    df_grouped = df_chart.groupby([category_col, 'Status_Clean']).size().reset_index(name='Count')
    
    # Calculate total for sorting and hover
    df_totals = df_grouped.groupby(category_col)['Count'].sum().reset_index(name='Total')
    df_grouped = pd.merge(df_grouped, df_totals, on=category_col)
    
    # Sort by total
    df_totals = df_totals.sort_values('Total', ascending=False) # Descending for horizontal bar (top to bottom)
    category_order = df_totals[category_col].tolist()
    
    # Define color maps
    color_maps = {
        'blue': {'Closed': '#1e3a8a', 'Open': '#93c5fd'},    # Dark Blue / Light Blue
        'purple': {'Closed': '#581c87', 'Open': '#d8b4fe'},  # Dark Purple / Light Purple
        'orange': {'Closed': '#c2410c', 'Open': '#fdba74'}   # Dark Orange / Light Orange
    }
    
    selected_colors = color_maps.get(color_theme, color_maps['blue'])
    
    # Calculate percentages for hover and text
    df_grouped['Percentage'] = (df_grouped['Count'] / df_grouped['Total'] * 100)
    
    # Create chart
    fig = px.bar(
        df_grouped,
        y=category_col,
        x='Count',
        color='Status_Clean',
        orientation='h',
        category_orders={category_col: category_order, 'Status_Clean': ['Closed', 'Open']},
        color_discrete_map=selected_colors,
        text='Count',
        custom_data=['Total', 'Percentage']
    )
    
    fig.update_layout(
        height=height,
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
            title=None
        ),
        margin=dict(l=3, r=20, t=30, b=3),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(family='Inter, sans-serif', size=13, color='#374151'),
        xaxis=dict(
            showgrid=True,
            gridcolor='#f3f4f6',
            zeroline=False,
            title_font=dict(size=13),
            tickangle=0,
            tickfont=dict(size=13),
            automargin=True
        ),
        yaxis=dict(
            showgrid=False,
            tickfont=dict(size=13),
            tickangle=0,
            automargin=True,
            autorange=True,
            fixedrange=True
        ),
        uniformtext_minsize=8,
        uniformtext_mode='hide',
        barmode='stack'
    )
    
    # Calculate percentages for each status within each category - only show for 'Closed'
    for i, trace in enumerate(fig.data):
        category_totals = df_grouped.groupby(category_col)['Total'].first().to_dict()
        percentages = []
        trace_name = trace.name if hasattr(trace, 'name') else ''
        
        for j, cat in enumerate(trace.y):
            total = category_totals[cat]
            percentage = (trace.x[j] / total * 100) if total > 0 else 0
            # Only show percentage for 'Closed' status to avoid overlap
            if 'closed' in trace_name.lower():
                percentages.append(f"{percentage:.0f}%")
            else:
                percentages.append("")
        trace.text = percentages
    
    fig.update_traces(
        marker=dict(line=dict(width=0)),
        texttemplate='%{text}',
        textposition='inside',
        textfont=dict(size=13, family='Inter, sans-serif', weight='bold'),
        hovertemplate=(
            '<b>%{y}</b><br><br>' +
            '%{fullData.name}: %{x:,} (%{customdata[1]:.1f}%)<br>' +
            'Total: %{customdata[0]:,}' +
            '<extra></extra>'
        )
    )
    
    return fig

def create_stacked_bar_chart(df, x_col, y_col, color_col, title, height=500):
    """Create a modern stacked bar chart with enhanced styling"""
    # Check if Breakdown exists for custom data
    custom_data = ['Breakdown'] if 'Breakdown' in df.columns else None
    
    fig = px.bar(
        df,
        x=x_col,
        y=y_col,
        color=color_col,
        title=None,
        labels={x_col: 'Month', y_col: 'Number of Complaints', color_col: 'Type'},
        color_discrete_sequence=['#3b82f6', '#ef4444', '#eab308', '#22c55e', '#f97316', '#a855f7'],  # Primary & Secondary colors
        custom_data=custom_data
    )
    
    # Calculate totals for each month to display inside bars
    monthly_totals = df.groupby(x_col)[y_col].sum().reset_index()
    monthly_totals.columns = [x_col, 'Total']
    # Add seamless text annotations showing totals on top of each bar
    for _, row in monthly_totals.iterrows():
        fig.add_annotation(
            x=row[x_col],
            y=row['Total'],
            text=str(int(row['Total'])),
            showarrow=False,
            font=dict(size=11, color='#6b7280', weight='normal'),
            bgcolor='rgba(0,0,0,0)',  # Transparent background
            bordercolor='rgba(0,0,0,0)',  # No border
            borderwidth=0,
            borderpad=0,
            yshift=12  # Position slightly above the bar
        )
    
    fig.update_layout(
        height=height,
        margin=dict(l=3, r=3, t=35, b=3),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(family='Inter, sans-serif', size=14, color='#374151'),
        xaxis=dict(
            showgrid=False,
            zeroline=False,
            title=None,
            tickfont=dict(size=12),
            tickangle=0,
            automargin=True
        ),
        yaxis=dict(
            showgrid=True,
            gridcolor='#f3f4f6',
            zeroline=False,
            title='Complaints',
            title_font=dict(size=13),
            tickfont=dict(size=12),
            tickangle=0,
            automargin=True
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        ),
        hovermode='x unified',
        hoverlabel=dict(
            bgcolor="white",
            font_size=13,
            font_family="Inter, sans-serif",
            align="left"
        ),
        uniformtext_minsize=8,
        uniformtext_mode='hide'
    )
    
    # Update traces with conditional hover template
    if custom_data:
        hovertemplate = '<b>%{fullData.name}</b><br>Count: %{y:,}<br>%{customdata[0]}<extra></extra>'
    else:
        hovertemplate = '<b>%{fullData.name}</b><br>Count: %{y:,}<extra></extra>'
        
    fig.update_traces(
        hovertemplate=hovertemplate
    )
    
    # Add text annotations showing totals for each month
    monthly_totals = df.groupby(x_col)[y_col].sum()
    annotations = []
    
    for month, total in monthly_totals.items():
        annotations.append(
            dict(
                x=month,
                y=total + (total * 0.02),  # Position slightly above the bar
                text=f"{total:,}",
                showarrow=False,
                font=dict(size=13, family='Inter, sans-serif', weight='bold', color='#374151'),
                xanchor='center'
            )
        )
    
    fig.update_layout(annotations=annotations)
    
    # Set the hover template for the x-axis to show only the month and year
    fig.update_xaxes(
        hoverformat='%B %Y'
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
        hovertemplate='<b>%{x}</b><br>Complaints: %{y:,}<extra></extra>',
        cliponaxis=False  # Don't clip labels outside the plot area
    )
    return fig

def create_pie_chart(data, title, height=400, use_ntc_colors=False):
    """Create a modern pie chart with distinct colors and enhanced hover info"""
    # Convert Series to DataFrame if needed
    if isinstance(data, pd.Series):
        df = data.reset_index()
        df.columns = ['Category', 'Count']
    else:
        df = data
    
    # Define custom colors for NTC providers
    if use_ntc_colors:
        # Create color mapping for NTC providers (matching get_ntc_group output)
        ntc_color_map = {
            'PLDT': '#ef4444',      # Red
            'SMART': '#22c55e',     # Green  
            'GLOBE': '#3b82f6',     # Blue
            'CONVERGE': '#7c3aed',  # Darker Purple
            'Others': '#f97316'     # Orange for others
        }
        
        # Create color list based on the categories in the data
        colors = []
        for category in df['Category']:
            if category in ntc_color_map:
                colors.append(ntc_color_map[category])
            else:
                colors.append('#6b7280')  # Default gray for unlisted providers
        
        color_discrete_map = dict(zip(df['Category'], colors))
        
        fig = px.pie(
            df,
            values='Count',
            names='Category',
            title=None,
            hole=0.4,  # Donut chart style
            color='Category',
            color_discrete_map=color_discrete_map
        )
    else:
        fig = px.pie(
            df,
            values='Count',
            names='Category',
            title=None,
            hole=0.4,  # Donut chart style
            color_discrete_sequence=['#3b82f6', '#ef4444', '#eab308', '#22c55e', '#f97316', '#a855f7']  # Primary & Secondary colors
        )
    
    fig.update_layout(
        height=height,
        margin=dict(l=10, r=10, t=20, b=20),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(family='Inter, sans-serif', size=14, color='#374151'),
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.2,
            xanchor="center",
            x=0.5
        )
    )
    
    # Calculate total for hover display
    total_count = df['Count'].sum() if not df.empty else 0
    
    fig.update_traces(
        textposition='inside',
        textinfo='label+percent',
        textfont=dict(size=13, family='Inter, sans-serif', weight='bold'),
        hovertemplate=f'<b>%{{label}}</b><br>Count: %{{value:,}} out of {total_count:,}<br>Percentage: %{{percent}}<extra></extra>',
        marker=dict(line=dict(color='#ffffff', width=2))
    )
    
    return fig

def consolidate_provider_name(provider_name):
    """Consolidate provider names to handle duplicates and variations"""
    if pd.isna(provider_name) or provider_name == '':
        return provider_name
    
    provider_lower = str(provider_name).lower().strip()
    
    # Converge consolidation (Converge ICT -> Converge)
    if 'converge' in provider_lower:
        return 'Converge'
    
    # Globe consolidation (Globe -> Globe Telecom)
    elif 'globe' in provider_lower and 'telecom' not in provider_lower:
        return 'Globe Telecom'
    
    # TNT Express consolidation
    elif 'tnt express' in provider_lower:
        return 'TNT Express Deliveries (Phils.), Inc.'
    
    # Return original if no consolidation needed
    else:
        return provider_name

def get_ntc_group(provider):
    """Group NTC providers into main categories (Case Insensitive)"""
    if not isinstance(provider, str):
        return "Others"
    
    # First consolidate the provider name
    consolidated_name = consolidate_provider_name(provider)
    p = consolidated_name.strip().upper()
    
    # PLDT Group
    if any(x in p for x in ["PLDT", "CIGNAL"]):
        return "PLDT"
        
    # SMART Group
    if any(x in p for x in ["SMART", "SUN", "TNT", "REDFIBER", "RED FIBER"]):
        return "SMART"
        
    # GLOBE Group
    if any(x in p for x in ["GLOBE", "TM", "GOMO"]):
        return "GLOBE"
        
    # CONVERGE Group
    if any(x in p for x in ["CONVERGE", "SURF 2 SAWA", "SURF2SAWA", "BIDA FIBER", "BIDAFIBER", "SKY"]):
        return "CONVERGE"
        
    return "Others"

def render_comparison_charts(df_period1, df_period3, period1_label, period3_label, color_theme1, color_theme2, key_prefix):
    """Render side-by-side service provider charts for comparison"""
    
    # Helper to process data based on key_prefix
    def process_provider_data(df, apply_grouping=True):
        if 'Service Providers' in df.columns and len(df) > 0:
            valid_data = df['Service Providers'].dropna()
            valid_data = valid_data[valid_data != '']
            
            if len(valid_data) > 0:
                # Always consolidate provider names first
                consolidated_data = valid_data.apply(consolidate_provider_name)
                
                # Apply NTC grouping if applicable AND requested
                if key_prefix == "ntc_providers" and apply_grouping:
                    grouped_data = consolidated_data.apply(get_ntc_group)
                    return grouped_data.value_counts()
                else:
                    return consolidated_data.value_counts()
        return pd.Series()

    col1, col2 = st.columns(2)

    with col1:
        st.markdown(f"#### Service Providers ({period1_label})")
        if len(df_period1) > 0 and 'Date Received' in df_period1.columns:
            period1_dates = df_period1['Date Received'].dropna()
            if len(period1_dates) > 0:
                date_range = f"{period1_dates.min().strftime('%b %d, %Y')} - {period1_dates.max().strftime('%b %d, %Y')}"
                # st.caption(f"📅 {date_range}") # Removed for redundancy

        # Apply grouping for the main/total chart (Period 1)
        counts = process_provider_data(df_period1, apply_grouping=True)
        
        if not counts.empty:
            # Top 5 and Others logic (Only applies if NOT NTC grouped, or if NTC groups somehow exceed 5)
            if len(counts) > 5:
                top_5 = counts.head(5)
                others_count = counts.iloc[5:].sum()
                # Create a new series with Others appended
                final_counts = pd.concat([top_5, pd.Series({'Others': others_count})])
            else:
                final_counts = counts

            # Use NTC colors if this is NTC analysis
            use_ntc_colors = (key_prefix == "ntc_providers")
            fig = create_pie_chart(final_counts, "Service Provider", 400, use_ntc_colors)
            st.plotly_chart(fig, use_container_width=True, key=f"{key_prefix}_period1")
        else:
            use_ntc_colors = (key_prefix == "ntc_providers")
            fig = create_pie_chart(pd.Series(), "Service Provider", 400, use_ntc_colors)
            st.plotly_chart(fig, use_container_width=True, key=f"{key_prefix}_period1_empty")
            if 'Service Providers' not in df_period1.columns:
                st.error("❌ 'Service Providers' column not found")
            else:
                st.info("📊 No service provider data available")

    with col2:
        st.markdown(f"#### Service Providers ({period3_label})")
        
        # The new stacked bar chart for the recent period
        if not df_period3.empty and 'Service Providers' in df_period3.columns:
            # Create a copy and consolidate provider names
            df_period3_consolidated = df_period3.copy()
            df_period3_consolidated['Service Providers'] = df_period3_consolidated['Service Providers'].apply(consolidate_provider_name)
            
            # Use the status stacked bar chart function
            fig = create_status_stacked_bar_chart(
                df_period3_consolidated, 
                'Service Providers', 
                "Recent Service Providers", 
                height=400, 
                max_items=12, 
                color_theme=color_theme2
            )
            if fig:
                st.plotly_chart(fig, use_container_width=True, key=f"{key_prefix}_period3_stacked")
            else:
                st.info("📊 No service provider data available for this period.")
        else:
            # Fallback for empty data or missing column
            fig = create_status_stacked_bar_chart(pd.DataFrame(), 'Service Providers', "", 400)
            st.plotly_chart(fig, use_container_width=True, key=f"{key_prefix}_period3_empty")
            if 'Service Providers' not in df_period3.columns:
                st.error("❌ 'Service Providers' column not found")
            else:
                st.info("📊 No service provider data available")

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
                interval = st.session_state.refresh_interval if st.session_state.auto_refresh else 300
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
                interval = st.session_state.refresh_interval if st.session_state.auto_refresh else 300
                current_time = int(datetime.now().timestamp() // interval) * interval
                df = load_data_from_gsheet_with_auth(st.session_state.gsheet_creds, active_url, current_time)
                st.session_state.data_source_type = "gsheets_private"
        
        if df is not None and not df.empty:
            st.session_state.last_valid_df = df
            st.divider()
            
            # Date Range Filter
            st.markdown("**Date Filter**")
            
            # Get available date range from data
            if 'Date Received' in df.columns:
                # Convert to datetime if not already
                temp_dates = pd.to_datetime(df['Date Received'], errors='coerce')
                valid_dates = temp_dates.dropna()
                
                if len(valid_dates) > 0:
                    min_date = valid_dates.min()
                    max_date = valid_dates.max()
                    
                    # Generate year and month options with "All Years" option
                    available_years = sorted(valid_dates.dt.year.unique(), reverse=True)
                    year_options = ["All Years"] + [str(year) for year in available_years]
                    
                    # Initialize session state for date filter
                    if 'filter_year' not in st.session_state:
                        st.session_state.filter_year = max_date.year
                    if 'filter_month' not in st.session_state:
                        st.session_state.filter_month = 0  # 0 means all months
                    
                    col_y, col_m = st.columns(2)
                    with col_y:
                        # Determine current selection index
                        if st.session_state.filter_year == "All Years":
                            current_index = 0
                        elif str(st.session_state.filter_year) in year_options:
                            current_index = year_options.index(str(st.session_state.filter_year))
                        else:
                            current_index = 1 if len(year_options) > 1 else 0
                        
                        selected_year_str = st.selectbox(
                            "Year",
                            options=year_options,
                            index=current_index,
                            key="year_selector"
                        )
                        
                        # Convert back to int or keep as "All Years"
                        if selected_year_str == "All Years":
                            selected_year = "All Years"
                        else:
                            selected_year = int(selected_year_str)
                    
                    with col_m:
                        month_options = [("All Months", 0)] + [(datetime(2020, m, 1).strftime('%b'), m) for m in range(1, 13)]
                        month_labels = [label for label, _ in month_options]
                        month_values = [value for _, value in month_options]
                        
                        current_month_index = month_values.index(st.session_state.filter_month) if st.session_state.filter_month in month_values else 0
                        
                        selected_month_label = st.selectbox(
                            "Month",
                            options=month_labels,
                            index=current_month_index,
                            key="month_selector"
                        )
                        selected_month = month_values[month_labels.index(selected_month_label)]
                    
                    # Update session state
                    st.session_state.filter_year = selected_year
                    st.session_state.filter_month = selected_month
                    
                    # Show selected range
                    if selected_year == "All Years":
                        if selected_month == 0:
                            st.caption(f"📅 Showing: All data")
                        else:
                            st.caption(f"📅 Showing: {selected_month_label} (all years)")
                    else:
                        if selected_month == 0:
                            st.caption(f"📅 Showing: All of {selected_year}")
                        else:
                            st.caption(f"📅 Showing: {selected_month_label} {selected_year}")
            
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
                    options=[60, 180, 300, 600], 
                    value=st.session_state.refresh_interval,
                    key="refresh_interval_slider",
                    on_change=lambda: st.session_state.update(refresh_interval=st.session_state.refresh_interval_slider)
                )

            # Compact Data Preview
            with st.expander("📋 View Data", expanded=False):
                st.dataframe(df.head(3), width='stretch', height=150)

    # Determine refresh rate for the fragment (only for Google Sheets)
    refresh_rate = None
    if st.session_state.auto_refresh and st.session_state.data_source_type in ["gsheets_public", "gsheets_auth"]:
        refresh_rate = st.session_state.refresh_interval

    @st.fragment(run_every=refresh_rate)
    def render_dashboard_content(initial_df=None):
        # Re-calculate timestamp for cache busting within the fragment
        interval = st.session_state.refresh_interval if st.session_state.auto_refresh else 300
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
        
        # Apply user-selected date filter and create labels
        selected_year = st.session_state.get('filter_year')
        selected_month = st.session_state.get('filter_month', 0)
        
        if 'Date Received' in df.columns and 'filter_year' in st.session_state and selected_year != "All Years":
            if selected_month == 0:
                # Filter by year only
                df = df[df['Date Received'].dt.year == selected_year].copy()
            else:
                # Filter by both year and month
                df = df[(df['Date Received'].dt.year == selected_year) & 
                       (df['Date Received'].dt.month == selected_month)].copy()
            
            if df.empty:
                st.warning(f"⚠️ No data found for the selected period.")
                return
        elif 'Date Received' in df.columns and selected_year == "All Years" and selected_month != 0:
            # Filter by month across all years
            df = df[df['Date Received'].dt.month == selected_month].copy()
            
            if df.empty:
                st.warning(f"⚠️ No data found for the selected month.")
                return
        
        # Create labels based on user selection and actual data range
        if 'Date Received' in df.columns:
            valid_dates = df['Date Received'].dropna()
            if len(valid_dates) > 0:
                min_date = valid_dates.min()
                max_date = valid_dates.max()

                # Create user-friendly labels based on filter selection
                if selected_year == "All Years":
                    if selected_month == 0:
                        # All data
                        period1_label = f"{min_date.strftime('%b %d, %Y')} - {max_date.strftime('%b %d, %Y')}"
                        period1_short = "All"
                    else:
                        # Specific month across all years
                        month_name = datetime(2020, selected_month, 1).strftime('%B')
                        period1_label = f"{month_name} (All Years)"
                        period1_short = f"{month_name[:3]} (All)"
                elif selected_month == 0:
                    # Full year selected
                    period1_label = f"{min_date.strftime('%b %d, %Y')} - {max_date.strftime('%b %d, %Y')}"
                    period1_short = str(selected_year)
                else:
                    # Specific month selected
                    month_name = datetime(2020, selected_month, 1).strftime('%B')
                    period1_label = f"{month_name} {selected_year}"
                    period1_short = f"{month_name[:3]} {selected_year}"

                # For period 2 and 3, use same data but create sub-ranges
                # Period 2: Last 3 months of the filtered data
                last_quarter_start = max_date - pd.DateOffset(months=3)
                last_month_start = max_date - pd.DateOffset(months=1)

                try:
                    df_period2 = df[df['Date Received'] >= last_quarter_start].copy()
                    period2_label = f"{last_quarter_start.strftime('%b %d, %Y')} - {max_date.strftime('%b %d, %Y')}"
                    period2_short = "3M"
                except:
                    df_period2 = df.copy()
                    period2_label = period1_label
                    period2_short = "All"

                try:
                    df_period3 = df[df['Date Received'] >= last_month_start].copy()
                    period3_label = f"{last_month_start.strftime('%b %d, %Y')} - {max_date.strftime('%b %d, %Y')}"
                    period3_short = "1M"
                except:
                    df_period3 = df.copy()
                    period3_label = period1_label
                    period3_short = "All"
                
                # Period 1 is the full filtered dataset
                df_period1 = df.copy()
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
                label=f"Total Complaints ({period1_short})",
                value=f"{period1_count:,}",
                delta="100% of Total",
                delta_color="off",
                help=f"Total number of complaints received during {period1_label}"
            )

        with col2:
            period2_count = len(df_period2)
            period2_pct = (period2_count / period1_count * 100) if period1_count > 0 else 0
            st.metric(
                label=f"Recent Complaints ({period2_short})",
                value=f"{period2_count:,}",
                delta=f"{period2_pct:.1f}% of Total",
                delta_color="off",
                help=f"Total number of complaints received during {period2_label}"
            )

        with col3:
            if 'Complaint Category' in df_period1.columns:
                try:
                    # Use ONLY Telco Internet Issues category for exact match with category counts
                    ntc_mask = (df_period1['Complaint Category'].astype(str).str.strip().str.upper() == "TELCO INTERNET ISSUES")
                    ntc_count = len(df_period1[ntc_mask])
                    ntc_pct = (ntc_count / period1_count * 100) if period1_count > 0 else 0
                    st.metric(
                        label=f"NTC Complaints ({period1_short})",
                        value=f"{ntc_count:,}",
                        delta=f"{ntc_pct:.1f}% of Total",
                        delta_color="off",
                        help=f"Includes Telco Internet Issues and NTC Agency complaints during {period1_label}"
                    )
                except Exception as e:
                    st.metric("NTC Complaints", "Error")
                    st.error(f"Error counting NTC complaints: {str(e)}")
            else:
                st.metric("NTC Complaints", "N/A")

        with col4:
            if 'Complaint Category' in df_period1.columns:
                try:
                    # Count PEMEDES complaints (Strictly "Delivery Concerns (SP)")
                    pemedes_mask = (df_period1['Complaint Category'].astype(str).str.strip().str.upper() == "DELIVERY CONCERNS (SP)")
                    
                    pemedes_count = len(df_period1[pemedes_mask])
                    pemedes_pct = (pemedes_count / period1_count * 100) if period1_count > 0 else 0
                    st.metric(
                        label=f"PEMEDES Complaints ({period1_short})",
                        value=f"{pemedes_count:,}",
                        delta=f"{pemedes_pct:.1f}% of Total",
                        delta_color="off",
                        help=f"Complaints categorized as Delivery Concerns (SP) during {period1_label}"
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
                # Use new stacked bar chart function
                fig = create_status_stacked_bar_chart(df_period1, 'Complaint Category', "Category", chart_height, color_theme='blue')
                if fig:
                    st.plotly_chart(fig, use_container_width=True, key="overall_category")
                else:
                    st.info("No category data available")
            else:
                st.error("'Complaint Category' column not found")

        with col2:
            st.markdown(f"#### Complaints by Nature ({period1_label})")
            if 'Complaint Nature' in df_period1.columns:
                # Use new stacked bar chart function
                fig = create_status_stacked_bar_chart(df_period1, 'Complaint Nature', "Nature", chart_height, color_theme='purple')
                if fig:
                    st.plotly_chart(fig, use_container_width=True, key="overall_nature")
                else:
                    st.info("No nature data available")
            else:
                st.error("'Complaint Nature' column not found")

        # Row 2: DICT Unit
        st.markdown(f"#### Complaints by DICT Unit ({period1_label})")
        if 'DICT UNIT' in df_period1.columns:
            # Filter out NTC first
            df_unit = df_period1.copy()
            df_unit = df_unit[~df_unit['DICT UNIT'].astype(str).str.upper().isin(['NTC', 'NATIONAL TELECOMMUNICATIONS COMMISSION'])]
            
            # Use new stacked bar chart function
            fig = create_status_stacked_bar_chart(df_unit, 'DICT UNIT', "DICT Unit", chart_height, max_items=10, color_theme='orange')
            if fig:
                st.plotly_chart(fig, use_container_width=True, key="overall_dict_unit")
            else:
                st.info("No DICT Unit data available")
        else:
            st.info("'DICT UNIT' column not found in data")

        # Monthly Trend
        st.markdown(f"#### Monthly Complaint Trend ({period1_label})")
        if 'Date Received' in df_period1.columns:
            valid_dates = df_period1['Date Received'].dropna()
            if len(valid_dates) > 0:
                # Create a copy for manipulation
                df_trend = df_period1.copy()
                
                # Categorize complaints for stacking
                def get_complaint_type(row):
                    category = str(row.get('Complaint Category', '')).strip().upper()
                    
                    # Check PEMEDES (Strictly Delivery Concerns)
                    if category == "DELIVERY CONCERNS (SP)":
                        return 'PEMEDES'
                        
                    # Check NTC
                    is_ntc = is_ntc_complaint(row.get('Agency', ''))
                    if category == "TELCO INTERNET ISSUES":
                        is_ntc = True
                    if is_ntc:
                        return 'NTC'

                    # Check Cyber-Related
                    if category == "CYBER-RELATED COMPLAINTS":
                        return 'Cyber-Related'
                        
                    # Check EGOV
                    if category == "EGOV SERVICES":
                        return 'EGOV'
                        
                    # Fallback
                    return 'Other'

                df_trend['Type'] = df_trend.apply(get_complaint_type, axis=1)
                
                # Group by Month and Type
                # We need to preserve the Month object for grouping, then convert to string for plotting
                df_trend['MonthPeriod'] = df_trend['Date Received'].dt.to_period('M')
                
                monthly_data = df_trend.groupby(['MonthPeriod', 'Type']).size().reset_index(name='Count')
                monthly_data['Month'] = monthly_data['MonthPeriod'].astype(str)
                
                # Add breakdown info for hover
                def get_breakdown(row):
                    if row['Type'] == 'Other':
                        # Filter for this month and type
                        mask = (df_trend['MonthPeriod'] == row['MonthPeriod']) & (df_trend['Type'] == 'Other')
                        subset = df_trend[mask]
                        if 'Complaint Category' in subset.columns:
                            counts = subset['Complaint Category'].value_counts().head(3)
                            # Format: "Category: Count"
                            items = [f"{cat}: {count}" for cat, count in counts.items()]
                            if items:
                                return "Top Issues:<br>" + "<br>".join(items)
                    return ""

                monthly_data['Breakdown'] = monthly_data.apply(get_breakdown, axis=1)
                
                if len(monthly_data) > 0:
                    fig = create_stacked_bar_chart(monthly_data, 'Month', 'Count', 'Type', "Monthly Trend", 350)
                    st.plotly_chart(fig, use_container_width=True, key="overall_monthly_trend")
                else:
                    st.info("No monthly data available")
            else:
                st.info("No valid complaint dates found")
        else:
            st.error("'Date Received' column not found")

        st.markdown("---")

        # NTC Analysis
        st.markdown("### NTC Analysis")

        # Filter NTC data with error handling
        if 'Complaint Category' in df_period1.columns:
            try:
                # Use ONLY Telco Internet Issues category for exact match with dashboard metrics
                ntc_mask_period1 = (df_period1['Complaint Category'].astype(str).str.strip().str.upper() == "TELCO INTERNET ISSUES")
                df_ntc_period1 = df_period1[ntc_mask_period1]

                ntc_mask_period3 = (df_period3['Complaint Category'].astype(str).str.strip().str.upper() == "TELCO INTERNET ISSUES")
                df_ntc_period3 = df_period3[ntc_mask_period3]

                # Data integrity check
                if len(df_ntc_period1) == 0:
                    st.warning(f"⚠️ No NTC complaints found in {period1_label} dataset. This may indicate:")
                    st.write("• No 'Telco Internet Issues' in Complaint Category")
                    st.write("• No NTC-related complaints in this period")
                    st.write("• Check the 'Complaint Category' column format")
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
                    label=f"NTC Complaints ({period1_short})",
                    value=f"{ntc_period1_count:,}",
                    delta=f"{ntc_period1_pct:.1f}% of Total",
                    delta_color="off",
                    help=f"NTC complaints during {period1_label}"
                )
            with kpi_col2:
                ntc_period3_count = len(df_ntc_period3)
                ntc_period3_pct = (ntc_period3_count / len(df_period3) * 100) if len(df_period3) > 0 else 0
                st.metric(
                    label=f"NTC Complaints ({period3_short})",
                    value=f"{ntc_period3_count:,}",
                    delta=f"{ntc_period3_pct:.1f}% of Total",
                    delta_color="off",
                    help=f"NTC complaints during {period3_label}"
                )

            st.markdown("---")

            # Use helper function for charts
            render_comparison_charts(
                df_ntc_period1, df_ntc_period3, 
                period1_label, period3_label, 
                'greens', 'purple', 
                "ntc_providers"
            )
        else:
            st.error("❌ 'Complaint Category' column not found in data. Cannot filter NTC complaints.")
            st.info("💡 Please ensure your data has a 'Complaint Category' column with 'Telco Internet Issues' entries.")

        st.markdown("---")

        # PEMEDES Analysis
        st.markdown("### PEMEDES Analysis")

        # Filter PEMEDES data with error handling
        if 'Complaint Category' in df_period1.columns:
            try:
                # Primary filter: Complaint Category is "Delivery Concerns (SP)"
                pemedes_mask_period1 = (df_period1['Complaint Category'].astype(str).str.strip().str.upper() == "DELIVERY CONCERNS (SP)")
                df_pemedes_temp1 = df_period1[pemedes_mask_period1]
                
                # Additional filter: Exclude NTC providers that might be miscategorized
                if 'Service Providers' in df_pemedes_temp1.columns:
                    ntc_provider_mask1 = df_pemedes_temp1['Service Providers'].apply(
                        lambda x: not is_ntc_provider(x) if pd.notna(x) else True
                    )
                    df_pemedes_period1 = df_pemedes_temp1[ntc_provider_mask1]
                else:
                    df_pemedes_period1 = df_pemedes_temp1

                # Same filtering for period 3
                pemedes_mask_period3 = (df_period3['Complaint Category'].astype(str).str.strip().str.upper() == "DELIVERY CONCERNS (SP)")
                df_pemedes_temp3 = df_period3[pemedes_mask_period3]
                
                if 'Service Providers' in df_pemedes_temp3.columns:
                    ntc_provider_mask3 = df_pemedes_temp3['Service Providers'].apply(
                        lambda x: not is_ntc_provider(x) if pd.notna(x) else True
                    )
                    df_pemedes_period3 = df_pemedes_temp3[ntc_provider_mask3]
                else:
                    df_pemedes_period3 = df_pemedes_temp3

                # Data integrity check
                if len(df_pemedes_period1) == 0:
                    st.warning(f"⚠️ No PEMEDES resolved complaints found in {period1_label} dataset. This may indicate:")
                    st.write("• No complaints categorized as 'Delivery Concerns (SP)' in this period")
            except Exception as e:
                st.error(f"❌ Error filtering PEMEDES data: {str(e)}")
                df_pemedes_period1 = pd.DataFrame()
                df_pemedes_period3 = pd.DataFrame()

            # Enhanced KPI metrics for PEMEDES
            kpi_col1, kpi_col2 = st.columns(2)
            with kpi_col1:
                pemedes_period1_count = len(df_pemedes_period1)
                pemedes_period1_pct = (pemedes_period1_count / len(df_period1) * 100) if len(df_period1) > 0 else 0
                st.metric(
                    label=f"PEMEDES Complaints ({period1_short})",
                    value=f"{pemedes_period1_count:,}",
                    delta=f"{pemedes_period1_pct:.1f}% of Total",
                    delta_color="off",
                    help=f"PEMEDES resolved complaints during {period1_label}"
                )
            with kpi_col2:
                pemedes_period3_count = len(df_pemedes_period3)
                pemedes_period3_pct = (pemedes_period3_count / len(df_period3) * 100) if len(df_period3) > 0 else 0
                st.metric(
                    label=f"PEMEDES Complaints ({period3_short})",
                    value=f"{pemedes_period3_count:,}",
                    delta=f"{pemedes_period3_pct:.1f}% of Total",
                    delta_color="off",
                    help=f"PEMEDES resolved complaints during {period3_label}"
                )

            st.markdown("---")

            # Use helper function for charts
            render_comparison_charts(
                df_pemedes_period1, df_pemedes_period3, 
                period1_label, period3_label, 
                'purples', 'blue', 
                "pemedes_providers"
            )
        else:
            st.error("❌ 'Complaint Category' column not found in data. Cannot filter PEMEDES resolved complaints.")

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
                    if len(df_period) > 0 and 'Date Received' in df_period.columns:
                        dates = df_period['Date Received'].dropna()
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
                    
                    # Exclude PEMEDES providers from NTC count
                    if 'Service Providers' in df_period1.columns:
                        pemedes_mask_excl = df_period1['Service Providers'].apply(is_pemedes_provider)
                        ntc_mask = ntc_mask & (~pemedes_mask_excl)

                    ntc_count = len(df_period1[ntc_mask])
                    st.write(f"NTC: {ntc_count:,} ({(ntc_count/len(df_period1)*100):.1f}%)")

                if 'Complaint Category' in df_period1.columns:
                    # PEMEDES filter with NTC provider exclusion
                    pemedes_mask = (df_period1['Complaint Category'].astype(str).str.strip().str.upper() == "DELIVERY CONCERNS (SP)")
                    df_pemedes_temp = df_period1[pemedes_mask]
                    
                    # Exclude NTC providers that might be miscategorized
                    if 'Service Providers' in df_pemedes_temp.columns:
                        ntc_provider_mask = df_pemedes_temp['Service Providers'].apply(
                            lambda x: not is_ntc_provider(x) if pd.notna(x) else True
                        )
                        df_pemedes_filtered = df_pemedes_temp[ntc_provider_mask]
                    else:
                        df_pemedes_filtered = df_pemedes_temp
                    
                    pem_count = len(df_pemedes_filtered)
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
        # Pass the prepared dataframe and date filter settings to the AI report module
        filter_year = st.session_state.get('filter_year')
        filter_month = st.session_state.get('filter_month', 0)
        ai_reports.render_weekly_report(df_prepared, filter_year, filter_month)

if __name__ == "__main__":
    main()
