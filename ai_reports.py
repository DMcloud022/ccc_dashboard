"""
AI-Powered Action Plan Report Module for DICT Complaint Dashboard

This module generates strategic action plans based on complaint data analysis.
It is designed to work seamlessly with dashboard.py and receives pre-processed data.

DATA ALIGNMENT WITH DASHBOARD:
- Receives data already processed by dashboard.py's prepare_data() function
- Data has already been filtered to exclude Resolution = "FLS"
- Column mappings have already been applied
- Dates have been parsed and validated
- Service provider names have been normalized
- Invalid/empty rows have been removed

KEY FUNCTIONS:
- get_top_issues(): Identifies top 5 complaint issues from cleaned data
- generate_ai_action_plan(): Uses Gemini AI to create strategic action plans
- export_to_pdf/word(): Generates formatted reports for download
- render_weekly_report(): Main UI rendering function

USAGE:
This module is imported by dashboard.py and called in the AI Action Plan tab.
"""

import streamlit as st
import pandas as pd
import vertexai
from vertexai.generative_models import GenerativeModel
import google.auth
import json
from datetime import datetime
import os
from dotenv import load_dotenv
from io import BytesIO
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

load_dotenv()

# DICT ORGANIZATIONAL STRUCTURE AND ISSUE MAPPING
# This mapping ensures accurate assignment of issues to the correct units and agencies

DICT_UNIT_MAPPING = {
    # 1. DELIVERY UNITS (Within DICT - Internal Services)
    "GDTB": {
        "name": "Government Digital Transformation Bureau",
        "keywords": ["egov", "pbh", "philippine business hub", "government services", "digital transformation"],
        "service_providers": ["GDTB", "eGOV", "NGP"]
    },
    "FPIAP": {
        "name": "Free Public Internet Access Program",
        "keywords": ["free wifi", "free wi-fi", "public internet", "community wifi", "pisonet"],
        "service_providers": ["FPIAP"]
    },
    "ILCDB": {
        "name": "ICT Literacy and Competency Development Bureau",
        "keywords": ["training", "upskilling", "certification", "certificate", "course", "learning", "workshop"],
        "service_providers": ["ILCDB"]
    },
    "AS": {
        "name": "Administrative Service",
        "keywords": ["hr concern", "human resource", "personnel", "employment", "recruitment"],
        "service_providers": ["AS", "HRMD", "HRDD", "GSD"]
    },
    "IMB": {
        "name": "Infrastructure Management Bureau",
        "keywords": ["cloud hosting", "web hosting", "government online", "gosd", "data center"],
        "service_providers": ["IMB", "GOSD", "GWHS", "NGDC"]
    },
    "CSB": {
        "name": "Cybersecurity Bureau",
        "keywords": ["digital certificate", "pki", "certificate authority", "encryption"],
        "service_providers": ["CSB", "PNPKI"]
    },
    "PRD": {
        "name": "Postal Regulation Division",
        "keywords": ["delivery concern", "courier", "logistics", "shipping", "parcel", "package"],
        "service_providers": ["LBC", "Ninja Van", "GO21", "GoGo Xpress", "Flash Express", "J&T", "J&T Express", "2GO", "SPX", "Lalamove"]
    },
    "ROCS": {
        "name": "Regional Operations and Coordination Service",
        "keywords": ["regional office", "regional concern", "regional service"],
        "service_providers": ["ROCS", "Regional Office"]
    },

    # 2. ATTACHED AGENCIES (Separate agencies under DICT supervision)
    "NTC": {
        "name": "National Telecommunications Commission",
        "keywords": ["internet", "telco", "disconnection", "slow connection", "network", "bandwidth", "fiber",
                    "broadband", "mobile data", "signal", "coverage", "technical service", "installation",
                    "unsolicited sms", "spam call", "telecom refund", "billing issue"],
        "service_providers": ["PLDT", "Converge", "Globe", "Smart", "DITO", "Sky Cable", "Cignal", "Eastern"]
    },
    "CICC": {
        "name": "Cybercrime Investigation and Coordinating Center",
        "keywords": ["cybercrime", "hacking", "phishing", "scam", "fraud", "identity theft", "online fraud",
                    "cyber attack", "data breach", "malware", "ransomware"],
        "service_providers": ["CICC"]
    },

    # 3. OTHER AGENCIES (External government bodies)
    "SEC": {
        "name": "Securities and Exchange Commission",
        "keywords": ["harassment", "online lending", "collection", "loan app", "lending app"],
        "service_providers": ["SEC"]
    },
    "DTI": {
        "name": "Department of Trade and Industry",
        "keywords": ["e-commerce", "online shopping", "consumer protection", "refund", "compensation",
                    "product quality", "false advertising", "lazada", "shopee"],
        "service_providers": ["DTI"]
    },
    "DOH": {
        "name": "Department of Health",
        "keywords": ["health", "vax", "vaccine", "vaccination", "health cert", "vax cert",
                    "immunization", "covid", "vaccination certificate"],
        "service_providers": ["DOH"]
    }
}

# Categorize units by organization type
DELIVERY_UNITS = ["GDTB", "FPIAP", "ILCDB", "AS", "IMB", "CSB", "PRD", "ROCS"]
ATTACHED_AGENCIES = ["NTC", "CICC"]
OTHER_AGENCIES = ["SEC", "DTI", "DOH"]

# Units that require service provider breakdown in reports
UNITS_REQUIRING_SP_BREAKDOWN = {
    "PRD": "Delivery Concerns",  # Show courier breakdown
    "NTC": "Telco/Internet Issues"  # Show ISP/telco breakdown
}

# Custom CSS for improved UI - Aligned with dashboard design
AI_REPORT_CSS = """
<style>
/* Import Google Fonts - Match Dashboard */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

/* AI Report Container */
.ai-report-container {
    background: white;
    border-radius: 8px;
    padding: 1.5rem;
    border: 1px solid #e5e7eb;
    margin-bottom: 0.75rem;
    font-family: 'Inter', sans-serif;
}

/* Report Header */
.report-header {
    text-align: center;
    padding-bottom: 1rem;
    border-bottom: 2px solid #3b82f6;
    margin-bottom: 1.5rem;
}

.report-title {
    font-size: 1.75rem;
    font-weight: 700;
    color: #1f2937;
    margin-bottom: 0.5rem;
    font-family: 'Inter', sans-serif;
}

.report-subtitle {
    font-size: 0.95rem;
    color: #6b7280;
    font-weight: 500;
}

/* Info Card */
.info-card {
    background: linear-gradient(135deg, #f0f9ff 0%, #e0f2fe 100%);
    border-left: 4px solid #3b82f6;
    padding: 0.875rem 1rem;
    margin: 0.75rem 0;
    border-radius: 6px;
    font-size: 0.9rem;
    box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05);
}

.info-card strong {
    color: #1f2937;
    font-weight: 600;
}

/* Responsive Table Container */
.responsive-table-container {
    width: 100%;
    overflow-x: auto;
    -webkit-overflow-scrolling: touch;
    border-radius: 8px;
    border: 1px solid #e5e7eb;
    margin: 1rem 0;
    box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05);
}

/* Table Styling */
.action-plan-table {
    width: 100%;
    min-width: 800px;
    border-collapse: collapse;
    font-size: 0.9rem;
    font-family: 'Inter', sans-serif;
}

.action-plan-table th {
    background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%);
    color: white;
    padding: 0.875rem 1rem;
    text-align: left;
    font-weight: 600;
    font-size: 0.875rem;
    text-transform: uppercase;
    letter-spacing: 0.025em;
    position: sticky;
    top: 0;
    z-index: 10;
}

.action-plan-table td {
    padding: 0.875rem 1rem;
    border-bottom: 1px solid #e5e7eb;
    vertical-align: top;
    color: #374151;
    line-height: 1.5;
}

.action-plan-table tbody tr:hover {
    background-color: #f9fafb;
    transition: background-color 0.15s ease;
}

.action-plan-table tbody tr:last-child td {
    border-bottom: none;
}

/* Column Widths */
.col-issue {
    width: 25%;
    min-width: 180px;
    font-weight: 600;
    color: #1f2937;
}

.col-action {
    width: 35%;
    min-width: 260px;
}

.col-unit {
    width: 15%;
    min-width: 120px;
}

.col-remarks {
    width: 15%;
    min-width: 120px;
}

.col-resolution {
    width: 10%;
    min-width: 100px;
    text-align: center;
}

/* Mobile Responsiveness */
@media (max-width: 768px) {
    .ai-report-container {
        padding: 1rem;
    }

    .report-title {
        font-size: 1.5rem;
    }

    .report-subtitle {
        font-size: 0.85rem;
    }

    .action-plan-table {
        font-size: 0.85rem;
    }

    .action-plan-table th,
    .action-plan-table td {
        padding: 0.65rem 0.5rem;
    }

    .info-card {
        font-size: 0.85rem;
        padding: 0.75rem 0.875rem;
    }
}

/* Streamlit Button Overrides for AI Report */
div[data-testid="stButton"] > button {
    border-radius: 6px;
    font-weight: 600;
    font-size: 0.9rem;
    padding: 0.5rem 1.25rem;
    transition: all 0.2s ease;
    font-family: 'Inter', sans-serif;
}

div[data-testid="stButton"] > button[kind="primary"] {
    background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%);
    border: none;
    box-shadow: 0 2px 4px rgba(59, 130, 246, 0.2);
}

div[data-testid="stButton"] > button[kind="primary"]:hover {
    background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%);
    box-shadow: 0 4px 8px rgba(59, 130, 246, 0.3);
    transform: translateY(-1px);
}

div[data-testid="stDownloadButton"] > button {
    border-radius: 6px;
    font-weight: 600;
    font-size: 0.875rem;
    padding: 0.5rem 1rem;
    border: 1px solid #e5e7eb;
    background: white;
    color: #374151;
    transition: all 0.2s ease;
    font-family: 'Inter', sans-serif;
}

div[data-testid="stDownloadButton"] > button:hover {
    background: #f9fafb;
    border-color: #3b82f6;
    color: #3b82f6;
    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);
    transform: translateY(-1px);
}

/* Comprehensive Data Editor Cell Styling - Compact with Proper Text Wrapping */

/* Main container */
div[data-testid="stDataFrame"] {
    font-size: 0.875rem;
    overflow-x: auto !important;
}

div[data-testid="stDataFrame"] div[data-testid="stDataFrameResizable"] {
    min-height: auto !important;
}

/* Glide Data Editor - Core wrapping fix */
.glide-data-editor {
    overflow: visible !important;
}

.glide-data-editor .dvn-scroller {
    overflow-x: auto !important;
    overflow-y: visible !important;
}

/* All cells - proper wrapping and compact sizing */
.glide-data-editor .dvn-underlay > div,
.glide-data-editor .gdg-cell,
div[data-testid="stDataFrame"] .glide-cell,
div[data-testid="stDataFrame"] td,
div[data-testid="stDataFrame"] th {
    white-space: pre-wrap !important;
    word-wrap: break-word !important;
    overflow-wrap: break-word !important;
    word-break: break-word !important;
    line-height: 1.5 !important;
    padding: 8px 6px !important;
    vertical-align: top !important;
    min-height: auto !important;
    max-height: none !important;
    height: auto !important;
    overflow: visible !important;
}

/* Cell content wrapper */
.glide-data-editor .gdg-cell > div,
.glide-data-editor .dvn-underlay > div > div {
    white-space: pre-wrap !important;
    word-wrap: break-word !important;
    overflow-wrap: break-word !important;
    word-break: break-word !important;
    max-width: 100% !important;
    overflow: visible !important;
    line-height: 1.5 !important;
}

/* Text content in cells */
.glide-data-editor .gdg-cell span,
.glide-data-editor .dvn-underlay span {
    white-space: pre-wrap !important;
    word-wrap: break-word !important;
    overflow-wrap: break-word !important;
    word-break: break-word !important;
    display: block !important;
    line-height: 1.5 !important;
}

/* Editable text areas and inputs - compact but usable */
div[data-testid="stDataFrame"] textarea,
div[data-testid="stDataFrame"] input[type="text"],
.glide-data-editor textarea,
.glide-data-editor input[type="text"] {
    min-height: 80px !important;
    line-height: 1.5 !important;
    padding: 8px !important;
    resize: vertical !important;
    white-space: pre-wrap !important;
    word-wrap: break-word !important;
    overflow-wrap: break-word !important;
    font-size: 0.875rem !important;
    width: 100% !important;
    box-sizing: border-box !important;
}

/* Prevent cell content overflow */
.glide-data-editor .gdg-cell,
.glide-data-editor .dvn-underlay > div {
    overflow: visible !important;
    text-overflow: clip !important;
}

/* Column headers - compact */
.glide-data-editor .gdg-header-cell {
    white-space: normal !important;
    word-wrap: break-word !important;
    padding: 8px 6px !important;
}

/* Ensure rows expand to fit content but start compact */
.glide-data-editor .gdg-row {
    min-height: auto !important;
    height: auto !important;
}

/* Fix for data grid canvas - compact */
.glide-data-editor canvas {
    min-height: auto !important;
}

/* Prevent table overflow and ensure proper scrolling */
div[data-testid="stDataFrame"] > div {
    overflow-x: auto !important;
    overflow-y: visible !important;
}

/* Ensure proper spacing between rows - compact */
.glide-data-editor .gdg-growing-entry {
    min-height: auto !important;
    overflow: visible !important;
}

/* Fix cell overlay positioning */
.glide-data-editor .gdg-overlay {
    overflow: visible !important;
}

/* Ensure edit overlay expands properly - compact */
.glide-data-editor .gdg-edit-overlay {
    min-height: 80px !important;
    overflow: visible !important;
}
</style>
"""

def categorize_issue_to_unit(issue_name, issue_type="Category"):
    """
    Intelligently categorize an issue to the appropriate DICT unit or agency

    Args:
        issue_name: The name of the issue (from Complaint Category or Nature)
        issue_type: "Category" or "Nature"

    Returns:
        Tuple of (unit_code, unit_full_name, organization_type)
    """
    # Validate input
    if issue_name is None or issue_name == '':
        return "CICC", "Cybersecurity Investigation and Coordinating Center", "Attached Agency"

    issue_lower = str(issue_name).lower()

    # Score each unit based on keyword matches
    unit_scores = {}

    for unit_code, unit_info in DICT_UNIT_MAPPING.items():
        score = 0

        # Check keyword matches
        for keyword in unit_info["keywords"]:
            if keyword.lower() in issue_lower:
                score += 2  # Keyword match is worth 2 points

        # Check service provider matches (if available in the issue name)
        for provider in unit_info["service_providers"]:
            if provider.lower() in issue_lower:
                score += 3  # Service provider match is worth 3 points

        if score > 0:
            unit_scores[unit_code] = score

    # Return the unit with highest score
    if unit_scores:
        best_unit = max(unit_scores, key=unit_scores.get)

        # Determine organization type
        if best_unit in DELIVERY_UNITS:
            org_type = "Delivery Unit (DICT Internal)"
        elif best_unit in ATTACHED_AGENCIES:
            org_type = "Attached Agency"
        elif best_unit in OTHER_AGENCIES:
            org_type = "External Agency"
        else:
            org_type = "DICT"

        return best_unit, DICT_UNIT_MAPPING[best_unit]["name"], org_type

    # Fallback: categorize based on common patterns
    if any(word in issue_lower for word in ["internet", "connection", "network", "telco", "broadband"]):
        return "NTC", DICT_UNIT_MAPPING["NTC"]["name"], "Attached Agency"
    elif any(word in issue_lower for word in ["delivery", "courier", "shipping", "parcel"]):
        return "PRD", DICT_UNIT_MAPPING["PRD"]["name"], "Delivery Unit (DICT Internal)"
    elif any(word in issue_lower for word in ["cybercrime", "scam", "fraud", "hacking"]):
        return "CICC", DICT_UNIT_MAPPING["CICC"]["name"], "Attached Agency"
    elif any(word in issue_lower for word in ["ecommerce", "e-commerce", "shopping", "consumer"]):
        return "DTI", DICT_UNIT_MAPPING["DTI"]["name"], "External Agency"

    # Default fallback for unmatched issues
    return "CICC", "Cybersecurity Investigation and Coordinating Center", "Attached Agency"

def init_vertex_ai():
    """Initialize Vertex AI with error handling"""
    try:
        # Check if vertexai is properly imported
        if not hasattr(vertexai, 'init'):
            return False, "Vertex AI module not properly loaded. Please check your installation."

        credentials, project_id = google.auth.default()

        if not project_id:
            return False, "Project ID not found. Please set GOOGLE_CLOUD_PROJECT environment variable."

        vertexai.init(project=project_id, location="asia-southeast1")
        return True, None
    except ImportError as e:
        return False, f"Import error: {str(e)}. Please install required packages: pip install google-cloud-aiplatform"
    except Exception as e:
        return False, str(e)

def get_service_provider_breakdown(df, issue_name, issue_type):
    """
    Get service provider breakdown for a specific issue

    Args:
        df: The full complaint dataframe
        issue_name: The name of the issue (e.g., "Delivery Concerns (SP)")
        issue_type: "Category" or "Nature"

    Returns:
        List of dicts with service provider counts and percentages
    """
    # Enhanced validation
    if df is None or df.empty:
        return []

    if 'Service Providers' not in df.columns:
        return []

    # Filter dataframe to only complaints matching this issue
    column_name = 'Complaint Category' if issue_type == "Category" else 'Complaint Nature'

    if column_name not in df.columns:
        return []

    try:
        # Get complaints for this specific issue with safe filtering
        issue_complaints = df[df[column_name] == issue_name]

        if len(issue_complaints) == 0:
            return []

        # Get service provider counts
        sp_counts = issue_complaints['Service Providers'].dropna().value_counts()
    except Exception as e:
        # Silently handle any filtering errors
        return []

    if len(sp_counts) == 0:
        return []

    total_with_sp = sp_counts.sum()

    # Build breakdown list
    breakdown = []
    for provider, count in sp_counts.items():
        if provider and str(provider).strip():  # Skip empty values
            percentage = (count / total_with_sp * 100) if total_with_sp > 0 else 0
            breakdown.append({
                "provider": str(provider),
                "count": int(count),
                "percentage": round(percentage, 1)
            })

    # Sort by count descending and limit to top 5
    breakdown.sort(key=lambda x: x['count'], reverse=True)

    return breakdown[:5]  # Return only top 5 service providers

def get_top_issues(df):
    """Extract top 5 issues based on Category and Nature

    This function aligns with dashboard.py's data structure and ensures
    we're analyzing the same cleaned data that's displayed in the dashboard.
    """
    if df is None or df.empty:
        return []

    issues = []

    try:
        # Analyze Categories (matching dashboard's approach)
        if 'Complaint Category' in df.columns:
            # Filter out NaN and empty values, matching dashboard behavior
            valid_categories = df['Complaint Category'].dropna()
            valid_categories = valid_categories[valid_categories != '']

            if len(valid_categories) > 0:
                top_cats = valid_categories.value_counts().head(5)
                for cat, count in top_cats.items():
                    issues.append({
                        "type": "Category",
                        "name": str(cat),  # Ensure string type
                        "count": int(count)
                    })

        # If we don't have enough categories, look at Nature
        if len(issues) < 5 and 'Complaint Nature' in df.columns:
            # Filter out NaN and empty values
            valid_nature = df['Complaint Nature'].dropna()
            valid_nature = valid_nature[valid_nature != '']

            if len(valid_nature) > 0:
                top_nature = valid_nature.value_counts().head(5 - len(issues))
                for nat, count in top_nature.items():
                    issues.append({
                        "type": "Nature",
                        "name": str(nat),  # Ensure string type
                        "count": int(count)
                    })
    except Exception as e:
        # Return empty list if any error occurs during analysis
        return []

    return issues[:5]

def generate_ai_action_plan(issues, df=None):
    """Generate action plan using Gemini

    This function creates strategic action plans based on the top complaint issues
    identified from the dashboard data (which has already been filtered and cleaned).
    Uses intelligent categorization to recommend the correct DICT unit or agency.

    Args:
        issues: List of top issues
        df: Optional dataframe for service provider analysis
    """
    if not issues or len(issues) == 0:
        return []

    # Validate that issues have required fields
    try:
        # First, categorize each issue to get recommended units and SP breakdown
        enriched_issues = []
        for issue in issues:
            if not isinstance(issue, dict) or 'name' not in issue or 'type' not in issue:
                continue  # Skip invalid issue entries

            unit_code, unit_name, org_type = categorize_issue_to_unit(issue['name'], issue['type'])

            # Get top service provider if applicable
            top_sp = None
            sp_count = 0
            sp_percentage = 0
            if df is not None and unit_code in UNITS_REQUIRING_SP_BREAKDOWN:
                sp_breakdown = get_service_provider_breakdown(df, issue['name'], issue['type'])
                if sp_breakdown and len(sp_breakdown) > 0:
                    top_sp = sp_breakdown[0]['provider']
                    sp_count = sp_breakdown[0]['count']
                    sp_percentage = sp_breakdown[0]['percentage']

            enriched_issues.append({
                **issue,
                "recommended_unit": unit_code,
                "recommended_unit_full": unit_name,
                "org_type": org_type,
                "top_service_provider": top_sp,
                "top_sp_count": sp_count,
                "top_sp_percentage": sp_percentage
            })

        # If no valid issues were enriched, return empty
        if len(enriched_issues) == 0:
            return []

    except Exception as e:
        # If enrichment fails, return empty list
        return []

    try:
        llm_model = os.getenv("LLM_MODEL", "gemini-1.5-flash-001")
        model = GenerativeModel(llm_model)

        system_prompt = os.getenv("SYSTEM_PROMPT", "You are a strategic analyst for the Department of Information and Communications Technology (DICT). Your role is to create actionable, specific, and measurable intervention plans to resolve citizen complaints.")

        # Build comprehensive unit guidelines with action plan templates
        unit_guidelines = """
DICT ORGANIZATIONAL STRUCTURE - UNIT ASSIGNMENT GUIDE WITH ACTION PLAN TEMPLATES:

1. DELIVERY UNITS (DICT Internal Services):
   - GDTB (Government Digital Transformation Bureau): eGov services, PBH, government digital transformation
     Template: "Conduct system audit of [specific service], implement fixes for [issue], and enhance user experience through [specific improvement]"

   - FPIAP (Free Public Internet Access Program): Free Wi-Fi, public internet access
     Template: "Deploy technical team to [location/issue area], restore/upgrade connectivity, and establish monitoring protocol"

   - ILCDB (ICT Literacy and Competency Development Bureau): Training, upskilling, certifications, courses
     Template: "Review [specific program], address [issue], streamline [process], and communicate timeline to affected participants"

   - AS (Administrative Service): HR concerns, personnel, recruitment
     Template: "Investigate [HR issue], implement corrective measures, and update policy/procedure to prevent recurrence"

   - IMB (Infrastructure Management Bureau): Cloud hosting, web hosting, government online services, data centers
     Template: "Conduct infrastructure assessment, resolve [technical issue], and implement redundancy/backup measures"

   - CSB (Cybersecurity Bureau): Digital certificates, PKI, encryption
     Template: "Fast-track certificate issuance/renewal process, resolve [specific issue], and establish expedited processing for backlog"

   - PRD (Postal Regulation Division): Delivery concerns, courier/logistics
     Template: "Escalate to [Top Service Provider] management, demand improved SLA compliance, establish penalty mechanism for delays, and explore alternative couriers"

   - ROCS (Regional Operations): Regional office concerns
     Template: "Coordinate with [specific region], deploy support team, resolve [issue], and strengthen regional coordination protocols"

2. ATTACHED AGENCIES (Under DICT supervision):
   - NTC (National Telecommunications Commission): Internet/telco issues, disconnections, slow connection, technical service
     Template: "Issue compliance directive to [Top Service Provider], mandate service restoration within [timeframe], impose penalties for SLA violations, and monitor resolution progress"

   - CICC (Cybersecurity Investigation and Coordinating Center): Cybercrime, hacking, phishing, scams, fraud
     Template: "Initiate investigation of [scam/fraud type], coordinate with law enforcement, issue public advisory, and pursue legal action against perpetrators"

3. OTHER AGENCIES (External partners):
   - SEC (Securities and Exchange Commission): Harassment, online lending, loan app collections
     Template: "Coordinate referral to SEC, provide complainant documentation support, and follow up on SEC enforcement action"

   - DTI (Department of Trade and Industry): E-commerce, consumer protection, retail refunds
     Template: "Refer to DTI Consumer Protection, facilitate mediation between parties, and support complaint resolution"
"""

        prompt = f"""
        {system_prompt}

        {unit_guidelines}

        Top Complaint Issues (pre-categorized with recommendations and service provider analysis):
        {json.dumps(enriched_issues, indent=2)}

        YOUR TASK: Create specific, actionable intervention plans for each issue.

        REQUIREMENTS FOR EACH ACTION PLAN:
        1. IDENTIFY ROOT CAUSE: What is the underlying problem causing this complaint?
        2. SPECIFIC ACTIONS: What concrete steps must be taken? (use the templates above as guides)
        3. TARGET SERVICE PROVIDER: If "top_service_provider" exists, name them specifically in the action plan
        4. MEASURABLE OUTCOME: What is the expected result?
        5. ACCOUNTABILITY: Who coordinates/implements? (use "recommended_unit" field)

        EXAMPLES OF GOOD ACTION PLANS:

        For NTC + Internet Issues + Top Provider "PLDT":
        "Issue compliance directive to PLDT requiring restoration of service within 48 hours for affected subscribers, impose administrative penalties for repeated SLA violations, establish weekly monitoring of complaint trends, and mandate quarterly service improvement reports."

        For PRD + Delivery Concerns + Top Provider "J&T Express":
        "Escalate to J&T Express regional management demanding immediate improvement in delivery timeframes, implement penalty mechanism for delays exceeding 5 days, require daily tracking updates for DICT shipments, and identify backup courier services to diversify risk."

        For NTC + Unsolicited SMS:
        "Direct all telecommunications providers to strengthen spam filtering mechanisms, issue show-cause orders to identified spam sources, coordinate with NBI Cybercrime Division for prosecution, and launch public awareness campaign on spam reporting procedures."

        For DTI + E-commerce refund issues:
        "Refer to DTI Consumer Protection Group, facilitate mediation between complainant and merchant, provide documentation support for filing formal complaints, and coordinate follow-up on resolution timeline."

        CRITICAL RULES:
        - DO NOT use generic language like "coordinate with" or "address concerns"
        - DO use specific verbs: "Issue directive", "Escalate to", "Implement", "Mandate", "Conduct audit", "Deploy team"
        - ALWAYS mention the top service provider by name if provided in the data
        - Include specific mechanisms (penalties, timelines, monitoring, reporting)
        - Use professional government language suitable for executive review
        - ALWAYS use the "recommended_unit" from the enriched_issues data
        - Keep action plans to 2-3 sentences maximum

        REMARKS FIELD INSTRUCTIONS:
        - Generate SPECIFIC, CONTEXTUAL remarks based on the actual data analysis for each issue
        - Include relevant metrics (complaint count, percentage, top provider if applicable)
        - Highlight priority level, urgency, or special considerations based on the actual issue
        - Examples of good remarks:
          * "Issue affects 234 users (45% of total complaints). Top provider J&T Express accounts for 67% of delivery issues."
          * "Critical priority - complaint volume increased 300% from previous period. Immediate intervention required."
          * "Recurring issue with PLDT services across multiple regions. Coordinate regulatory action with NTC."
          * "Limited to specific region. Monitor for pattern expansion before escalating intervention."
        - DO NOT use generic templates - base remarks on actual issue data provided above

        Return ONLY a valid JSON array with keys: "issue", "action_plan", "unit", "remarks".
        Do not include markdown formatting like ```json.
        """

        response = model.generate_content(prompt)

        # Clean response text to ensure valid JSON
        text = response.text.strip()
        if text.startswith("```json"):
            text = text[7:]
        if text.endswith("```"):
            text = text[:-3]
        text = text.strip()

        # Validate JSON before parsing
        if not text:
            raise ValueError("AI returned empty response")

        ai_plans = json.loads(text)

        # Validate that response is a list
        if not isinstance(ai_plans, list):
            raise ValueError("AI response is not a list of action plans")

        # Validate and correct unit assignments
        validated_plans = []
        for i, plan in enumerate(ai_plans):
            # Ensure plan is a dictionary
            if not isinstance(plan, dict):
                continue

            # Validate required fields exist
            if 'issue' not in plan or 'action_plan' not in plan or 'unit' not in plan:
                # Use enriched issue data if AI response is incomplete
                if i < len(enriched_issues):
                    plan = {
                        "issue": enriched_issues[i]["name"],
                        "action_plan": plan.get("action_plan", "Review and address complaints"),
                        "unit": enriched_issues[i]["recommended_unit"],
                        "remarks": ""
                    }

            # Use the pre-categorized unit if AI didn't assign correctly
            if plan.get("unit") not in DICT_UNIT_MAPPING:
                if i < len(enriched_issues):
                    plan["unit"] = enriched_issues[i]["recommended_unit"]

            # Ensure remarks field exists (AI should have generated this)
            if "remarks" not in plan or not plan["remarks"]:
                # Minimal fallback only if AI failed to generate remarks
                if i < len(enriched_issues):
                    issue_data = enriched_issues[i]
                    count = issue_data.get('count', 0)
                    top_sp = issue_data.get('top_service_provider', '')

                    # Create data-driven remark as minimal fallback
                    remark_parts = [f"Affects {count:,} complaints"]
                    if top_sp:
                        remark_parts.append(f"Top provider: {top_sp}")
                    plan["remarks"] = ". ".join(remark_parts) + ". Requires immediate attention."
                else:
                    plan["remarks"] = "Awaiting detailed analysis and implementation."

            validated_plans.append(plan)

        return validated_plans

    except Exception as e:
        st.error(f"AI Generation Error: {str(e)}")
        # Fallback: use pre-categorized units with specific action plans based on unit type and service provider
        fallback_plans = []
        for issue in enriched_issues:
            unit_code = issue['recommended_unit']
            unit_name = DICT_UNIT_MAPPING[unit_code]['name']
            top_sp = issue.get('top_service_provider')

            # Build specific action plans based on unit type
            if unit_code == "NTC":
                if top_sp:
                    action_plan = f"Issue compliance directive to {top_sp} requiring immediate service improvement, impose penalties for SLA violations, and establish monitoring mechanism for complaint resolution."
                else:
                    action_plan = f"Conduct investigation of telecommunications service quality issues, issue compliance directives to non-compliant providers, and enforce regulatory penalties where applicable."

            elif unit_code == "PRD":
                if top_sp:
                    action_plan = f"Escalate to {top_sp} management demanding improved delivery performance, implement penalty clauses for delays, and evaluate alternative courier services for future contracts."
                else:
                    action_plan = f"Review courier service provider contracts, enforce delivery SLA compliance, and establish performance monitoring system to prevent recurrence."

            elif unit_code == "CICC":
                action_plan = f"Initiate cybercrime investigation, coordinate with law enforcement agencies, issue public advisory on prevention measures, and pursue legal action against identified perpetrators."

            elif unit_code == "DTI":
                action_plan = f"Refer cases to DTI Consumer Protection Group, facilitate merchant-consumer mediation, provide complainants with documentation support, and coordinate follow-up on resolution timeline."

            elif unit_code == "SEC":
                action_plan = f"Coordinate referral to SEC Enforcement Department, assist complainants in filing formal complaints, and monitor SEC's regulatory action against violators."

            elif unit_code == "FPIAP":
                action_plan = f"Deploy technical team to assess connectivity issues, restore or upgrade affected infrastructure, and implement preventive monitoring system."

            elif unit_code == "GDTB":
                action_plan = f"Conduct comprehensive system audit, implement technical fixes for identified issues, and enhance user interface based on feedback analysis."

            elif unit_code == "ILCDB":
                action_plan = f"Review program implementation processes, address identified gaps in service delivery, streamline enrollment/certification procedures, and communicate updated timelines to participants."

            elif unit_code == "IMB":
                action_plan = f"Conduct infrastructure assessment, resolve technical service disruptions, implement system redundancy, and establish improved backup protocols."

            elif unit_code == "CSB":
                action_plan = f"Expedite digital certificate processing, address backlog in certificate issuance, and establish fast-track mechanism for urgent requests."

            else:
                # Generic fallback for any other unit
                action_plan = f"Conduct thorough investigation of complaints, implement corrective measures to address root causes, and establish monitoring system to prevent recurrence."

            # Generate minimal data-driven remarks for fallback (AI failure scenario)
            count = issue.get('count', 0)
            remark_parts = [f"Affects {count:,} complaints"]
            if top_sp:
                remark_parts.append(f"Top provider: {top_sp} ({sp_percentage:.1f}%)")
            remarks = ". ".join(remark_parts) + ". Requires prompt action."

            fallback_plans.append({
                "issue": issue['name'],
                "action_plan": action_plan,
                "unit": unit_code,
                "remarks": remarks
            })

        return fallback_plans

def export_to_pdf(plans_df, top_issues, sp_breakdowns=None):
    """Generate PDF report with service provider breakdowns

    Args:
        plans_df: DataFrame of action plans
        top_issues: List of top issues
        sp_breakdowns: Optional list of service provider breakdown data
    """
    from reportlab.lib.pagesizes import landscape, A4

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), topMargin=0.5*inch, bottomMargin=0.5*inch,
                           leftMargin=0.5*inch, rightMargin=0.5*inch)
    elements = []

    # Styles
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=24,
        textColor=colors.HexColor('#1f2937'),
        spaceAfter=12,
        alignment=TA_CENTER,
        fontName='Helvetica-Bold'
    )

    subtitle_style = ParagraphStyle(
        'CustomSubtitle',
        parent=styles['Normal'],
        fontSize=12,
        textColor=colors.HexColor('#6b7280'),
        spaceAfter=20,
        alignment=TA_CENTER
    )

    heading_style = ParagraphStyle(
        'CustomHeading',
        parent=styles['Heading2'],
        fontSize=14,
        textColor=colors.HexColor('#1f2937'),
        spaceAfter=12,
        spaceBefore=12,
        fontName='Helvetica-Bold'
    )

    # Title
    elements.append(Paragraph("DICT AI Action Plan Report", title_style))
    elements.append(Paragraph(f"Generated: {datetime.now().strftime('%B %d, %Y')}", subtitle_style))
    elements.append(Spacer(1, 0.2*inch))

    # Executive Summary
    body_style = ParagraphStyle(
        'BodyText',
        parent=styles['Normal'],
        fontSize=10,
        textColor=colors.HexColor('#374151'),
        spaceAfter=6,
        leading=14
    )

    total_top5 = sum([issue['count'] for issue in top_issues])
    # Note: total_top5 represents complaints in top 5 issues, not total dataset
    # For full dataset count, would need to pass as parameter

    summary_text = f"""
    <b>Executive Summary:</b><br/>
    This report identifies the top 5 priority complaint issues requiring immediate attention and provides strategic
    action plans for resolution. The analysis covers {total_top5:,} complaints from the top 5 issues, representing the most critical
    areas for intervention. Each issue has been assigned to the appropriate DICT unit or agency based on their
    mandate and expertise. Units are requested to review their assigned action plans, provide implementation remarks,
    and update resolution status as progress is made.
    """
    elements.append(Paragraph(summary_text, body_style))
    elements.append(Spacer(1, 0.2*inch))

    # Top Issues Summary
    elements.append(Paragraph("I. Top 5 Priority Issues", heading_style))

    issues_data = [['#', 'Issue', 'Source', 'Count']]
    for idx, issue in enumerate(top_issues, 1):
        issues_data.append([
            str(idx),
            Paragraph(issue['name'], styles['Normal']),
            issue['type'],  # Shows "Category" or "Nature"
            str(issue['count'])
        ])

    issues_table = Table(issues_data, colWidths=[0.5*inch, 3.5*inch, 1*inch, 0.8*inch])
    issues_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#3b82f6')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 11),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('TOPPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#e5e7eb')),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('TOPPADDING', (0, 1), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 8),
    ]))

    elements.append(issues_table)

    # Add footnote for Source column
    footnote_style = ParagraphStyle(
        'Footnote',
        parent=styles['Normal'],
        fontSize=9,
        textColor=colors.HexColor('#6b7280'),
        spaceAfter=6,
        leftIndent=0.5*inch
    )
    elements.append(Paragraph("Note: 'Source' indicates whether the issue is from 'Category' or 'Nature' field in complaint data.", footnote_style))
    elements.append(Spacer(1, 0.3*inch))

    # Action Plan Table
    elements.append(Paragraph("II. Strategic Action Plan Details", heading_style))

    # Add narrative explanation
    action_plan_narrative = f"""
    The following action plans have been developed for each priority issue. Each plan outlines specific interventions
    required and identifies the responsible DICT unit or agency. The Remarks and Resolution columns are provided for
    units to document their implementation progress, challenges encountered, and final resolution status.
    """
    elements.append(Paragraph(action_plan_narrative, body_style))
    elements.append(Spacer(1, 0.15*inch))

    plan_data = [['Issue', 'Action Plan', 'Assigned Unit', 'Remarks', 'Action Taken by the Unit']]
    for _, row in plans_df.iterrows():
        # Use actual values from edited data (remarks and resolution may be edited)
        remarks_text = str(row.get('remarks', '')) if pd.notna(row.get('remarks', '')) else ''
        resolution_text = str(row.get('resolution', '')) if pd.notna(row.get('resolution', '')) else ''

        plan_data.append([
            Paragraph(str(row['issue']), styles['Normal']),
            Paragraph(str(row['action_plan']), styles['Normal']),
            Paragraph(str(row['unit']), styles['Normal']),
            Paragraph(remarks_text, styles['Normal']),
            Paragraph(resolution_text, styles['Normal'])
        ])

    # Landscape A4 width is about 11 inches, minus margins = ~10 inches available
    plan_table = Table(plan_data, colWidths=[2.0*inch, 3.5*inch, 1.3*inch, 2.0*inch, 1.2*inch])
    plan_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#3b82f6')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('ALIGN', (4, 0), (4, -1), 'CENTER'),  # Center align Resolution column
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('TOPPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#e5e7eb')),
        ('FONTSIZE', (0, 1), (-1, -1), 9),
        ('TOPPADDING', (0, 1), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 8),
    ]))

    elements.append(plan_table)

    # Add note about editable fields
    note_style = ParagraphStyle(
        'Note',
        parent=styles['Normal'],
        fontSize=8,
        textColor=colors.HexColor('#6b7280'),
        spaceAfter=6,
        leftIndent=0.5*inch,
        italic=True
    )
    elements.append(Spacer(1, 0.1*inch))
    elements.append(Paragraph("Note: This report includes any edits made to the Action Plan, Remarks, and Action Taken by the Unit fields before download.", note_style))

    # Service Provider Breakdown Section
    if sp_breakdowns and len(sp_breakdowns) > 0:
        elements.append(Spacer(1, 0.3*inch))
        elements.append(Paragraph("III. Service Provider Analysis", heading_style))

        # Add narrative
        sp_narrative = """
        For issues related to delivery concerns and telecommunications, the following breakdown identifies specific
        service providers responsible for the majority of complaints. This granular analysis enables targeted
        interventions with individual providers to address service quality issues effectively. Each breakdown shows
        the top 5 service providers contributing to the issue, allowing focused engagement with the most problematic providers.
        """
        elements.append(Paragraph(sp_narrative, body_style))
        elements.append(Spacer(1, 0.15*inch))

        body_style_sp = ParagraphStyle(
            'BodySP',
            parent=styles['Normal'],
            fontSize=9,
            textColor=colors.HexColor('#374151')
        )

        for sp_item in sp_breakdowns:
            # Issue header
            issue_header = f"{sp_item['issue']} ({sp_item['unit']}) - {sp_item['total_count']} total complaints"
            elements.append(Paragraph(issue_header, styles['Heading3']))
            elements.append(Spacer(1, 0.1*inch))

            # Add table-specific narrative
            num_providers = len(sp_item['breakdown'])
            top_provider_pct = sp_item['breakdown'][0]['percentage'] if sp_item['breakdown'] else 0

            table_narrative = f"""
            The table below presents the top {num_providers} service providers for this issue category.
            The leading provider accounts for {top_provider_pct:.1f}% of complaints in this category,
            indicating {"a concentrated issue requiring focused intervention" if top_provider_pct > 40 else "a distributed problem across multiple providers"}.
            Coordinating with these providers can directly address {sum([sp['percentage'] for sp in sp_item['breakdown']]):.1f}% of complaints in this category.
            """
            elements.append(Paragraph(table_narrative, body_style_sp))
            elements.append(Spacer(1, 0.1*inch))

            # SP breakdown table
            sp_data = [['Service Provider', 'Complaints', 'Percentage']]
            for sp in sp_item['breakdown']:
                sp_data.append([
                    sp['provider'],
                    str(sp['count']),
                    f"{sp['percentage']}%"
                ])

            sp_table = Table(sp_data, colWidths=[3*inch, 1.2*inch, 1.2*inch])
            sp_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#f3f4f6')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor('#1f2937')),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 9),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                ('TOPPADDING', (0, 0), (-1, 0), 8),
                ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#e5e7eb')),
                ('FONTSIZE', (0, 1), (-1, -1), 8),
                ('TOPPADDING', (0, 1), (-1, -1), 6),
                ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
            ]))

            elements.append(sp_table)
            elements.append(Spacer(1, 0.15*inch))

            # Top provider analysis and recommendation
            if sp_item['breakdown']:
                top_sp = sp_item['breakdown'][0]

                # Generate actionable recommendation based on data
                if top_sp['percentage'] > 50:
                    recommendation = f"Immediate escalation to {top_sp['provider']} management is recommended as they represent the majority of issues."
                elif top_sp['percentage'] > 30:
                    recommendation = f"Priority engagement with {top_sp['provider']} while monitoring other providers is advised."
                else:
                    recommendation = f"A multi-provider approach is recommended given the distributed nature of complaints."

                analysis_text = f"""
                <b>Key Finding:</b> {top_sp['provider']} leads with {top_sp['count']} complaints ({top_sp['percentage']}%).
                <b>Recommended Action:</b> {recommendation}
                """
                elements.append(Paragraph(analysis_text, body_style_sp))
                elements.append(Spacer(1, 0.2*inch))

    # Build PDF
    doc.build(elements)
    buffer.seek(0)
    return buffer

def export_to_word(plans_df, top_issues, sp_breakdowns=None):
    """Generate Word document report with service provider breakdowns

    Args:
        plans_df: DataFrame of action plans
        top_issues: List of top issues
        sp_breakdowns: Optional list of service provider breakdown data
    """
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    def set_cell_background(cell, fill_color):
        """Set cell background color"""
        try:
            shading_elm = OxmlElement('w:shd')
            shading_elm.set(qn('w:fill'), fill_color)
            cell._element.get_or_add_tcPr().append(shading_elm)
        except:
            pass  # Silently fail if shading doesn't work

    doc = Document()

    # Set document margins
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    # Title
    title = doc.add_heading('DICT AI Action Plan Report', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title.runs[0]
    title_run.font.color.rgb = RGBColor(31, 41, 55)

    # Subtitle
    subtitle = doc.add_paragraph(f"Generated: {datetime.now().strftime('%B %d, %Y')}")
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle_run = subtitle.runs[0]
    subtitle_run.font.size = Pt(11)
    subtitle_run.font.color.rgb = RGBColor(107, 114, 128)

    doc.add_paragraph()

    # Top Issues Section
    doc.add_heading('Top 5 Priority Issues', 1)

    issues_table = doc.add_table(rows=1, cols=4)
    issues_table.style = 'Light Grid Accent 1'

    # Header row
    header_cells = issues_table.rows[0].cells
    header_cells[0].text = '#'
    header_cells[1].text = 'Issue'
    header_cells[2].text = 'Source'
    header_cells[3].text = 'Count'

    # Format header
    for cell in header_cells:
        if cell.paragraphs and cell.paragraphs[0].runs:
            cell.paragraphs[0].runs[0].font.bold = True
            cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
        set_cell_background(cell, '3B82F6')

    # Data rows
    for idx, issue in enumerate(top_issues, 1):
        row_cells = issues_table.add_row().cells
        row_cells[0].text = str(idx)
        row_cells[1].text = issue['name']
        row_cells[2].text = issue['type']  # Shows "Category" or "Nature"
        row_cells[3].text = str(issue['count'])

    # Add footnote for Source column
    note = doc.add_paragraph()
    note_run = note.add_run("Note: 'Source' indicates whether the issue is from 'Category' or 'Nature' field in complaint data.")
    note_run.font.size = Pt(9)
    note_run.font.color.rgb = RGBColor(107, 114, 128)
    note_run.italic = True

    doc.add_paragraph()

    # Action Plan Section
    doc.add_heading('Action Plan Details', 1)

    plan_table = doc.add_table(rows=1, cols=5)
    plan_table.style = 'Light Grid Accent 1'

    # Header row
    header_cells = plan_table.rows[0].cells
    header_cells[0].text = 'Issue'
    header_cells[1].text = 'Action Plan'
    header_cells[2].text = 'Assigned Unit'
    header_cells[3].text = 'Remarks'
    header_cells[4].text = 'Action Taken by the Unit'

    # Format header
    for cell in header_cells:
        if cell.paragraphs and cell.paragraphs[0].runs:
            cell.paragraphs[0].runs[0].font.bold = True
            cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
        set_cell_background(cell, '3B82F6')

    # Data rows
    for _, row in plans_df.iterrows():
        # Use actual values from edited data (remarks and resolution may be edited)
        remarks_text = str(row.get('remarks', '')) if pd.notna(row.get('remarks', '')) else ''
        resolution_text = str(row.get('resolution', '')) if pd.notna(row.get('resolution', '')) else ''

        row_cells = plan_table.add_row().cells
        row_cells[0].text = str(row['issue'])
        row_cells[1].text = str(row['action_plan'])
        row_cells[2].text = str(row['unit'])
        row_cells[3].text = remarks_text
        row_cells[4].text = resolution_text
        # Center align the resolution
        if row_cells[4].paragraphs:
            row_cells[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Add note about editable fields
    note = doc.add_paragraph()
    note_run = note.add_run("Note: This report includes any edits made to the Action Plan, Remarks, and Action Taken by the Unit fields before download.")
    note_run.font.size = Pt(8)
    note_run.font.color.rgb = RGBColor(107, 114, 128)
    note_run.italic = True

    # Service Provider Breakdown Section
    if sp_breakdowns and len(sp_breakdowns) > 0:
        doc.add_paragraph()
        doc.add_heading('Service Provider Analysis', 1)

        # Add section narrative
        section_intro = doc.add_paragraph()
        intro_text = (
            "For issues related to delivery concerns and telecommunications, the following breakdown identifies "
            "specific service providers responsible for the majority of complaints. This granular analysis enables "
            "targeted interventions with individual providers to address service quality issues effectively. "
            "Each breakdown shows the top 5 service providers contributing to the issue, allowing focused "
            "engagement with the most problematic providers."
        )
        intro_run = section_intro.add_run(intro_text)
        intro_run.font.size = Pt(10)
        intro_run.font.color.rgb = RGBColor(55, 65, 81)
        doc.add_paragraph()

        for sp_item in sp_breakdowns:
            # Issue subheading
            issue_heading = doc.add_heading(f"{sp_item['issue']} ({sp_item['unit']}) - {sp_item['total_count']} complaints", 2)

            # Add table-specific narrative
            num_providers = len(sp_item['breakdown'])
            top_provider_pct = sp_item['breakdown'][0]['percentage'] if sp_item['breakdown'] else 0

            table_narrative_para = doc.add_paragraph()
            table_narrative_text = (
                f"The table below presents the top {num_providers} service providers for this issue category. "
                f"The leading provider accounts for {top_provider_pct:.1f}% of complaints in this category, "
                f"indicating {'a concentrated issue requiring focused intervention' if top_provider_pct > 40 else 'a distributed problem across multiple providers'}. "
                f"Coordinating with these providers can directly address {sum([sp['percentage'] for sp in sp_item['breakdown']]):.1f}% of complaints in this category."
            )
            narrative_run = table_narrative_para.add_run(table_narrative_text)
            narrative_run.font.size = Pt(10)
            narrative_run.font.color.rgb = RGBColor(55, 65, 81)
            doc.add_paragraph()

            # SP breakdown table
            sp_table = doc.add_table(rows=1, cols=3)
            sp_table.style = 'Light Grid Accent 1'

            # Header
            sp_header_cells = sp_table.rows[0].cells
            sp_header_cells[0].text = 'Service Provider'
            sp_header_cells[1].text = 'Complaints'
            sp_header_cells[2].text = 'Percentage'

            for cell in sp_header_cells:
                if cell.paragraphs and cell.paragraphs[0].runs:
                    cell.paragraphs[0].runs[0].font.bold = True
                    cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
                set_cell_background(cell, 'F3F4F6')

            # Data rows
            for sp in sp_item['breakdown']:
                sp_row_cells = sp_table.add_row().cells
                sp_row_cells[0].text = sp['provider']
                sp_row_cells[1].text = str(sp['count'])
                sp_row_cells[2].text = f"{sp['percentage']}%"

            # Top provider analysis and recommendation
            if sp_item['breakdown']:
                top_sp = sp_item['breakdown'][0]

                # Generate actionable recommendation based on data
                if top_sp['percentage'] > 50:
                    recommendation = f"Immediate escalation to {top_sp['provider']} management is recommended as they represent the majority of issues."
                elif top_sp['percentage'] > 30:
                    recommendation = f"Priority engagement with {top_sp['provider']} while monitoring other providers is advised."
                else:
                    recommendation = f"A multi-provider approach is recommended given the distributed nature of complaints."

                # Key Finding
                finding_para = doc.add_paragraph()
                finding_label = finding_para.add_run("Key Finding: ")
                finding_label.font.bold = True
                finding_label.font.size = Pt(9)
                finding_label.font.color.rgb = RGBColor(31, 41, 55)

                finding_text = finding_para.add_run(f"{top_sp['provider']} leads with {top_sp['count']} complaints ({top_sp['percentage']}%).")
                finding_text.font.size = Pt(9)
                finding_text.font.color.rgb = RGBColor(55, 65, 81)

                # Recommended Action
                action_para = doc.add_paragraph()
                action_label = action_para.add_run("Recommended Action: ")
                action_label.font.bold = True
                action_label.font.size = Pt(9)
                action_label.font.color.rgb = RGBColor(31, 41, 55)

                action_text = action_para.add_run(recommendation)
                action_text.font.size = Pt(9)
                action_text.font.color.rgb = RGBColor(55, 65, 81)

            doc.add_paragraph()

    # Save to buffer
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def render_weekly_report(df):
    """Render the Weekly Report / Action Plan section with improved UI"""

    # Apply custom CSS
    st.markdown(AI_REPORT_CSS, unsafe_allow_html=True)

    # Header
    st.markdown("""
    <div class="report-header">
        <div class="report-title">DICT AI-Powered Action Plan Report</div>
        <div class="report-subtitle">Strategic Analysis and Recommendations for Priority Complaint Issues</div>
    </div>
    """, unsafe_allow_html=True)

    # Date and status info
    col_info1, col_info2 = st.columns([2, 1])
    with col_info1:
        st.caption(f"Report Date: {datetime.now().strftime('%B %d, %Y')}")
    with col_info2:
        if 'report_timestamp' in st.session_state:
            st.caption(f"Last Generated: {st.session_state.report_timestamp.strftime('%I:%M %p')}")

    st.markdown("<br>", unsafe_allow_html=True)

    if df is None or df.empty:
        st.info("No data available to generate report. Please load complaint data from the Dashboard tab first.")
        st.markdown("""
        **Instructions:**
        1. Navigate to the Dashboard tab
        2. Load complaint data using the sidebar options
        3. Return to this tab to generate the AI-powered action plan
        """)
        return

    # IMPORTANT: The df passed here is already prepared by dashboard.py's prepare_data() function
    # which handles:
    # - Column mapping and validation
    # - Date parsing
    # - Service provider normalization
    # - FLS resolution filtering (Resolution != "FLS")
    # - Invalid date removal
    # Therefore, we work with the data as-is without re-processing

    # Data alignment confirmation
    total_records = len(df)
    date_range_info = ""
    if 'Date of Complaint' in df.columns:
        valid_dates = df['Date of Complaint'].dropna()
        if len(valid_dates) > 0:
            min_date = valid_dates.min()
            max_date = valid_dates.max()
            date_range_info = f" | Data Range: {min_date.strftime('%b %Y')} - {max_date.strftime('%b %Y')}"

    st.markdown(f"""
    <div class="info-card">
        <strong>Executive Summary:</strong> This report provides AI-powered analysis of complaint data to identify top priority issues
        and generates strategic action plans with recommended DICT units for resolution.
        <br><br>
        <strong>Data Coverage:</strong> Analysis based on {total_records:,} complaints (excluding FLS resolutions){date_range_info}
    </div>
    """, unsafe_allow_html=True)

    # Initialize Vertex AI
    is_init, error_msg = init_vertex_ai()

    if not is_init:
        st.warning("Vertex AI not initialized. AI features are currently disabled.")
        with st.expander("Error Details"):
            if error_msg:
                st.caption(f"Error: {error_msg}")
                st.caption("Please ensure you have 'Vertex AI User' role and the API is enabled in Google Cloud.")

    # Get Top Issues (always compute to show preview)
    top_issues = get_top_issues(df)

    if not top_issues:
        st.warning("Insufficient data to identify top issues. Please ensure your data has 'Complaint Category' or 'Complaint Nature' columns.")
        return

    # Preview Top Issues with Unit Recommendations - Enhanced Presentation
    st.markdown("### I. Top 5 Priority Issues")

    # Calculate total complaints in top 5
    total_top5 = sum([issue['count'] for issue in top_issues])
    coverage_pct = (total_top5 / total_records * 100) if total_records > 0 else 0

    st.markdown(f"""
    <div class="info-card" style="background: linear-gradient(135deg, #fef3c7 0%, #fde68a 100%); border-left: 4px solid #f59e0b;">
        <strong>Impact Analysis:</strong> The top 5 priority issues represent <strong>{total_top5:,} complaints ({coverage_pct:.1f}%)</strong> of the total dataset.
        Addressing these priorities will have the most significant impact on reducing overall complaint volume.
    </div>
    """, unsafe_allow_html=True)

    # Enrich issues with unit recommendations
    preview_data = []
    for idx, issue in enumerate(top_issues, 1):
        unit_code, unit_name, org_type = categorize_issue_to_unit(issue['name'], issue['type'])
        preview_data.append({
            "#": idx,
            "Issue": issue['name'],
            "Source": issue['type'],
            "Count": issue['count'],
            "Recommended Unit": unit_code,
            "Organization": org_type
        })

    preview_df = pd.DataFrame(preview_data)
    st.dataframe(
        preview_df,
        column_config={
            "#": st.column_config.NumberColumn("Rank", width="small"),
            "Issue": st.column_config.TextColumn("Issue Description", width="large"),
            "Source": st.column_config.TextColumn("Source", width="small"),
            "Count": st.column_config.NumberColumn("Complaints", format="%d", width="small"),
            "Recommended Unit": st.column_config.TextColumn("Assigned Unit", width="medium"),
            "Organization": st.column_config.TextColumn("Organization Type", width="medium")
        },
        hide_index=True,
        use_container_width=True
    )
    st.caption("**Note:** Source indicates whether the issue is from Category or Nature field. Assigned Unit is auto-categorized based on DICT organizational structure.")

    st.markdown("---")

    # Generation Button - Centered and prominent
    col_spacer1, col_btn, col_spacer2 = st.columns([1, 2, 1])

    with col_btn:
        generate_button = st.button(
            "Generate Strategic Action Plan",
            type="primary",
            use_container_width=True,
            help="Analyze top complaints and generate strategic action plans"
        )

    # Handle generation
    if generate_button or ('weekly_action_plan' in st.session_state and st.session_state.get('report_generated', False)):
        if generate_button:
            with st.spinner("Analyzing complaint data and generating strategic action plans..."):
                if is_init:
                    action_plan_data = generate_ai_action_plan(top_issues, df)  # Pass df for SP analysis
                else:
                    # Fallback if no AI - generate concise action plans with minimal data-driven remarks
                    action_plan_data = []
                    for i in top_issues:
                        unit_code, unit_name, org_type = categorize_issue_to_unit(i['name'], i['type'])

                        # Generate minimal data-driven remarks (non-AI scenario)
                        count = i.get('count', 0)
                        remarks = f"Affects {count:,} complaints. Assigned to {unit_code} for resolution."

                        action_plan_data.append({
                            "issue": i['name'],
                            "action_plan": f"Coordinate with {unit_code} to address complaints",
                            "unit": unit_code,
                            "remarks": remarks
                        })
                st.session_state.weekly_action_plan = action_plan_data
                st.session_state.report_generated = True
                st.session_state.report_timestamp = datetime.now()
                st.success("Action plan generated successfully.")

        # Display the action plan
        if 'weekly_action_plan' in st.session_state:
            plans = st.session_state.weekly_action_plan

            # Validate plans data
            if not plans or len(plans) == 0:
                st.warning("No action plans were generated. Please try again.")
                return

            # Convert to DataFrame and add editable columns
            report_df = pd.DataFrame(plans)

            # Validate dataframe is not empty
            if report_df.empty:
                st.warning("Action plan data is empty. Please regenerate the report.")
                return

            # Add resolution column if it doesn't exist (internal name stays 'resolution')
            if 'resolution' not in report_df.columns:
                report_df['resolution'] = ''

            # Rename columns for better display
            display_df = report_df.rename(columns={
                'issue': 'Top Issue',
                'action_plan': 'Action Plan',
                'unit': 'Assigned Unit',
                'remarks': 'Remarks',
                'resolution': 'Action Taken by the Unit'
            })

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("---")
            st.markdown("### II. Strategic Action Plan Details")
            st.caption(f"Generated {len(report_df)} strategic recommendations based on analysis of top complaint patterns")

            # Info box about editing
            st.info("💡 **Fully Editable Table!** All fields can be edited. Make your changes, then scroll down and click 'Save All Changes' to apply.")

            # Initialize edited_action_plan on first load only
            if 'edited_action_plan' not in st.session_state:
                st.session_state.edited_action_plan = report_df.copy()
                st.session_state.data_changed = True

            # Use saved data for display (or original if not yet saved)
            display_source_df = st.session_state.edited_action_plan.copy()

            # Ensure resolution column exists
            if 'resolution' not in display_source_df.columns:
                display_source_df['resolution'] = ''

            # Rename columns for display
            display_df_for_editor = display_source_df.rename(columns={
                'issue': 'Top Issue',
                'action_plan': 'Action Plan',
                'unit': 'Assigned Unit',
                'remarks': 'Remarks',
                'resolution': 'Action Taken by the Unit'
            })

            # Editable data editor - ALL fields are editable with proper wrapping
            edited_df = st.data_editor(
                display_df_for_editor,
                column_config={
                    "Top Issue": st.column_config.TextColumn(
                        "Top Issue",
                        width=200,
                        disabled=False,
                        help="Click to edit issue name"
                    ),
                    "Action Plan": st.column_config.TextColumn(
                        "Action Plan",
                        width=400,
                        disabled=False,
                        help="Click to edit the action plan"
                    ),
                    "Assigned Unit": st.column_config.TextColumn(
                        "Assigned Unit",
                        width=120,
                        disabled=False,
                        help="Click to change assigned unit"
                    ),
                    "Remarks": st.column_config.TextColumn(
                        "Remarks",
                        width=300,
                        disabled=False,
                        help="Click to add or edit remarks"
                    ),
                    "Action Taken by the Unit": st.column_config.TextColumn(
                        "Action Taken by the Unit",
                        width=250,
                        disabled=False,
                        help="Click to update actions taken"
                    )
                },
                hide_index=True,
                use_container_width=True,
                num_rows="fixed",
                key="action_plan_editor"
            )

            # Store edited main table temporarily (local variable only - no state change)
            temp_edited_main_df = edited_df

            # Service Provider Breakdown for PRD and NTC issues
            st.markdown("---")
            st.markdown("### III. Service Provider Analysis")
            st.caption("Detailed breakdown for Delivery Concerns and Telecommunications Issues")

            # Use edited dataframe for further processing
            current_report_df = st.session_state.edited_action_plan

            # Initialize service provider breakdowns in session state if not exists
            if 'sp_breakdowns' not in st.session_state:
                st.session_state.sp_breakdowns = {}

            # Check which issues need SP breakdown
            issues_with_breakdown = []
            temp_sp_edits_local = {}  # Local dictionary to collect SP edits (no state changes)

            for idx, row in current_report_df.iterrows():
                unit = row['unit']
                issue_name = row['issue']

                # Find matching issue from top_issues
                matching_issue = next((i for i in top_issues if i['name'] == issue_name), None)

                if matching_issue and unit in UNITS_REQUIRING_SP_BREAKDOWN:
                    # Use cached breakdown if exists, otherwise fetch new
                    sp_key = f"{issue_name}_{unit}"
                    if sp_key not in st.session_state.sp_breakdowns:
                        sp_breakdown = get_service_provider_breakdown(df, issue_name, matching_issue['type'])
                        if sp_breakdown:
                            st.session_state.sp_breakdowns[sp_key] = sp_breakdown
                    else:
                        sp_breakdown = st.session_state.sp_breakdowns[sp_key]

                    if sp_breakdown:
                        issues_with_breakdown.append({
                            "issue": issue_name,
                            "unit": unit,
                            "unit_label": UNITS_REQUIRING_SP_BREAKDOWN[unit],
                            "total_count": matching_issue['count'],
                            "breakdown": sp_breakdown,
                            "sp_key": sp_key
                        })

            if issues_with_breakdown:
                st.info("💡 Service Provider tables are also editable. Edit as needed, then use 'Save All Changes' button below.")

                for idx, item in enumerate(issues_with_breakdown):
                    with st.expander(f"{item['issue']} ({item['unit']}) - {item['total_count']} complaints", expanded=True):
                        # Create editable breakdown table from session state
                        sp_df = pd.DataFrame(item['breakdown'])

                        edited_sp_df = st.data_editor(
                            sp_df,
                            column_config={
                                "provider": st.column_config.TextColumn(
                                    "Service Provider",
                                    width=300,
                                    disabled=False,
                                    help="Click to edit provider name"
                                ),
                                "count": st.column_config.NumberColumn(
                                    "Complaints",
                                    format="%d",
                                    width=120,
                                    disabled=False,
                                    help="Click to edit count"
                                ),
                                "percentage": st.column_config.NumberColumn(
                                    "Percentage",
                                    format="%.1f%%",
                                    width=120,
                                    disabled=False,
                                    help="Click to edit percentage"
                                )
                            },
                            hide_index=True,
                            use_container_width=True,
                            num_rows="fixed",
                            key=f"sp_breakdown_{idx}",
                            disabled=False
                        )

                        # Store edits in local dictionary (NO session state update - prevents rerun!)
                        new_sp_breakdown = edited_sp_df.to_dict('records')
                        item['breakdown'] = new_sp_breakdown
                        temp_sp_edits_local[item['sp_key']] = new_sp_breakdown

                        # Summary stats
                        st.caption(f"Top Provider: **{item['breakdown'][0]['provider']}** with {item['breakdown'][0]['count']} complaints ({item['breakdown'][0]['percentage']}%)")
                        st.caption(f"Total Providers Identified: {len(item['breakdown'])}")

                        # Add concise explanation
                        st.markdown("---")
                        st.markdown("**Analysis and Recommendations:**")

                        # Calculate insights
                        top_provider = item['breakdown'][0]
                        top_provider_pct = top_provider['percentage']
                        num_providers = len(item['breakdown'])

                        # Generate contextual explanation based on the data
                        if top_provider_pct > 50:
                            concentration = "highly concentrated"
                            recommendation = f"Focus immediate attention on **{top_provider['provider']}** as they account for the majority of issues. Consider escalating to their management team."
                        elif top_provider_pct > 30:
                            concentration = "moderately concentrated"
                            recommendation = f"Prioritize **{top_provider['provider']}** while monitoring other providers. A targeted intervention could significantly reduce complaints."
                        else:
                            concentration = "distributed across multiple providers"
                            recommendation = f"Issues are spread across {num_providers} providers. Consider a broader systemic approach rather than targeting individual providers."

                        st.write(f"Complaint distribution for this issue is **{concentration}**, with the leading provider accounting for **{top_provider_pct:.1f}%** of all complaints in this category. {recommendation}")
            else:
                st.info("No Delivery Concerns or Telecommunications Issues identified in the top 5 priority complaints requiring detailed service provider analysis.")

            st.markdown("---")

            # Save button (after all editable sections)
            st.markdown("### Save All Changes")
            col_save1, col_save2, col_save3 = st.columns([1, 1, 1])
            with col_save2:
                save_button = st.button(
                    "💾 Save All Changes",
                    type="primary",
                    use_container_width=True,
                    help="Save edits from main table and service provider tables, then regenerate export files",
                    key="save_all_changes"
                )

            # Only process and save when save button is clicked
            if save_button:
                # Convert edited main table data back to internal column names
                temp_edited_main = temp_edited_main_df.rename(columns={
                    'Top Issue': 'issue',
                    'Action Plan': 'action_plan',
                    'Assigned Unit': 'unit',
                    'Remarks': 'remarks',
                    'Action Taken by the Unit': 'resolution'
                })

                # Update session state with main table edits
                st.session_state.edited_action_plan = temp_edited_main

                # Apply all SP edits to permanent storage
                for sp_key, sp_data in temp_sp_edits_local.items():
                    st.session_state.sp_breakdowns[sp_key] = sp_data

                # Mark that export cache needs regeneration
                st.session_state.data_changed = True

                st.success("✅ All changes saved successfully! Export files will be updated.")
                st.rerun()

            st.caption("**Note:** Click 'Save All Changes' button above to apply edits and regenerate export files.")

            st.markdown("---")

            # Download Buttons Section
            st.markdown("### IV. Export Options")
            st.caption("Download the strategic action plan in your preferred format (includes your edits)")

            # Use edited dataframe for exports (with safety check)
            if 'edited_action_plan' not in st.session_state:
                st.error("No action plan data available for export. Please generate the report first.")
                return

            export_df = st.session_state.edited_action_plan

            # Cache the export data as bytes to prevent regeneration on download
            if 'cached_pdf_bytes' not in st.session_state or st.session_state.get('data_changed', True):
                try:
                    pdf_buffer = export_to_pdf(export_df, top_issues, issues_with_breakdown)
                    st.session_state.cached_pdf_bytes = pdf_buffer.getvalue()  # Store as bytes
                    st.session_state.pdf_error = None
                except Exception as e:
                    st.session_state.cached_pdf_bytes = None
                    st.session_state.pdf_error = str(e)

            if 'cached_word_bytes' not in st.session_state or st.session_state.get('data_changed', True):
                try:
                    word_buffer = export_to_word(export_df, top_issues, issues_with_breakdown)
                    st.session_state.cached_word_bytes = word_buffer.getvalue()  # Store as bytes
                    st.session_state.word_error = None
                except Exception as e:
                    st.session_state.cached_word_bytes = None
                    st.session_state.word_error = str(e)

            if 'cached_csv_string' not in st.session_state or st.session_state.get('data_changed', True):
                st.session_state.cached_csv_string = export_df.to_csv(index=False)

            # Mark data as cached
            st.session_state.data_changed = False

            col_dl1, col_dl2, col_dl3 = st.columns(3)

            with col_dl1:
                # PDF Download
                if st.session_state.cached_pdf_bytes:
                    st.download_button(
                        label="📄 PDF Document",
                        data=st.session_state.cached_pdf_bytes,
                        file_name=f"DICT_AI_Action_Plan_{datetime.now().strftime('%Y%m%d')}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                        help="Download formatted PDF report with service provider analysis",
                        key="download_pdf_btn"
                    )
                else:
                    st.error(f"PDF Export Error: {st.session_state.get('pdf_error', 'Unknown error')}")

            with col_dl2:
                # Word Download
                if st.session_state.cached_word_bytes:
                    st.download_button(
                        label="📝 Word Document",
                        data=st.session_state.cached_word_bytes,
                        file_name=f"DICT_AI_Action_Plan_{datetime.now().strftime('%Y%m%d')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                        help="Download editable Word document with service provider analysis",
                        key="download_word_btn"
                    )
                else:
                    st.error(f"Word Export Error: {st.session_state.get('word_error', 'Unknown error')}")

            with col_dl3:
                # CSV Download
                st.download_button(
                    label="📊 CSV Spreadsheet",
                    data=st.session_state.cached_csv_string,
                    file_name=f"DICT_AI_Action_Plan_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv",
                    use_container_width=True,
                    help="Download data in CSV format",
                    key="download_csv_btn"
                )

            # Summary Metrics
            st.markdown("---")
            st.markdown("### V. Report Summary")

            m1, m2, m3, m4 = st.columns(4)

            total_complaints = sum([i['count'] for i in top_issues])
            unique_units = export_df['unit'].nunique()

            m1.metric(
                label="Issues Analyzed",
                value=len(export_df),
                help="Number of top issues identified"
            )
            m2.metric(
                label="Total Complaints",
                value=f"{total_complaints:,}",
                help="Total complaints covered by action plans"
            )
            m3.metric(
                label="Units Assigned",
                value=unique_units,
                help="Number of DICT units involved"
            )
            m4.metric(
                label="Coverage",
                value=f"{(total_complaints / len(df) * 100):.1f}%",
                help="Percentage of total complaints addressed"
            )

            # Unit Distribution Analysis with Specific Details
            st.markdown("---")
            st.markdown("### VI. Unit Assignment and Service Provider Details")
            st.caption("Agencies and service providers requiring intervention")

            # Build detailed breakdown with service providers
            unit_details = []
            for idx, row in export_df.iterrows():
                unit = row['unit']
                issue_name = row['issue']

                # Get matching issue from top_issues
                matching_issue = next((i for i in top_issues if i['name'] == issue_name), None)

                # Get top service provider if applicable
                top_provider = None
                provider_count = 0
                if matching_issue and unit in UNITS_REQUIRING_SP_BREAKDOWN:
                    sp_breakdown = get_service_provider_breakdown(df, issue_name, matching_issue['type'])
                    if sp_breakdown and len(sp_breakdown) > 0:
                        top_provider = sp_breakdown[0]['provider']
                        provider_count = sp_breakdown[0]['count']

                # Get full unit name
                unit_full_name = DICT_UNIT_MAPPING.get(unit, {}).get("name", unit)

                # Categorize
                if unit in DELIVERY_UNITS:
                    category = "Delivery Unit (DICT)"
                elif unit in ATTACHED_AGENCIES:
                    category = "Attached Agency"
                elif unit in OTHER_AGENCIES:
                    category = "External Agency"
                else:
                    category = "Unclassified"

                unit_details.append({
                    "Unit Code": unit,
                    "Unit Name": unit_full_name,
                    "Category": category,
                    "Issue": issue_name,
                    "Top Service Provider": top_provider if top_provider else "N/A",
                    "SP Complaints": provider_count if top_provider else 0,
                    "Total Complaints": matching_issue['count'] if matching_issue else 0
                })

            # Display as comprehensive table
            details_df = pd.DataFrame(unit_details)
            st.dataframe(
                details_df,
                column_config={
                    "Unit Code": st.column_config.TextColumn("Unit", width="small"),
                    "Unit Name": st.column_config.TextColumn("Agency/Unit Name", width="large"),
                    "Category": st.column_config.TextColumn("Type", width="medium"),
                    "Issue": st.column_config.TextColumn("Issue Assigned", width="large"),
                    "Top Service Provider": st.column_config.TextColumn("Top Service Provider", width="medium"),
                    "SP Complaints": st.column_config.NumberColumn("SP Count", format="%d", width="small"),
                    "Total Complaints": st.column_config.NumberColumn("Total", format="%d", width="small")
                },
                hide_index=True,
                use_container_width=True
            )

            # Summary by category
            st.markdown("#### A. Summary by Organization Type")
            col_sum1, col_sum2, col_sum3 = st.columns(3)

            delivery_units = [d for d in unit_details if d['Category'] == "Delivery Unit (DICT)"]
            attached_units = [d for d in unit_details if d['Category'] == "Attached Agency"]
            external_units = [d for d in unit_details if d['Category'] == "External Agency"]

            with col_sum1:
                st.metric(
                    label="DICT Delivery Units",
                    value=len(delivery_units),
                    help="Internal DICT units"
                )
                if delivery_units:
                    unique_delivery = list(set([d['Unit Code'] for d in delivery_units]))
                    st.caption(f"Units: {', '.join(unique_delivery)}")

            with col_sum2:
                st.metric(
                    label="Attached Agencies",
                    value=len(attached_units),
                    help="Agencies under DICT"
                )
                if attached_units:
                    unique_attached = list(set([d['Unit Code'] for d in attached_units]))
                    st.caption(f"Agencies: {', '.join(unique_attached)}")

            with col_sum3:
                st.metric(
                    label="External Agencies",
                    value=len(external_units),
                    help="Partner agencies"
                )
                if external_units:
                    unique_external = list(set([d['Unit Code'] for d in external_units]))
                    st.caption(f"Agencies: {', '.join(unique_external)}")

            # Validation status
            unclassified = [d for d in unit_details if d['Category'] == "Unclassified"]
            if unclassified:
                st.warning(f"{len(unclassified)} issue(s) could not be categorized properly")
                for item in unclassified:
                    st.caption(f"• {item['Unit Code']} - {item['Issue']}")
            else:
                st.success("All issues successfully categorized to appropriate units and agencies")

    else:
        # Show placeholder when no report is generated
        st.markdown("<br>", unsafe_allow_html=True)
        st.info("Click the **Generate Strategic Action Plan** button above to analyze your complaint data and create strategic recommendations.")
