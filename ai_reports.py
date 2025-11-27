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
        "name": "Procurement and Records Division",
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
    width: 22%;
    min-width: 160px;
    font-weight: 600;
    color: #1f2937;
}

.col-action {
    width: 38%;
    min-width: 280px;
}

.col-unit {
    width: 18%;
    min-width: 140px;
}

.col-remarks {
    width: 22%;
    min-width: 160px;
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
        credentials, project_id = google.auth.default()

        if not project_id:
            return False, "Project ID not found. Please set GOOGLE_CLOUD_PROJECT environment variable."

        vertexai.init(project=project_id, location="asia-southeast1")
        return True, None
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
    if df is None or df.empty or 'Service Providers' not in df.columns:
        return []

    # Filter dataframe to only complaints matching this issue
    column_name = 'Complaint Category' if issue_type == "Category" else 'Complaint Nature'

    if column_name not in df.columns:
        return []

    # Get complaints for this specific issue
    issue_complaints = df[df[column_name] == issue_name]

    if len(issue_complaints) == 0:
        return []

    # Get service provider counts
    sp_counts = issue_complaints['Service Providers'].dropna().value_counts()

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

    # Sort by count descending
    breakdown.sort(key=lambda x: x['count'], reverse=True)

    return breakdown

def get_top_issues(df):
    """Extract top 5 issues based on Category and Nature

    This function aligns with dashboard.py's data structure and ensures
    we're analyzing the same cleaned data that's displayed in the dashboard.
    """
    if df is None or df.empty:
        return []

    issues = []

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
    if not issues:
        return []

    # First, categorize each issue to get recommended units and SP breakdown
    enriched_issues = []
    for issue in issues:
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

    try:
        llm_model = os.getenv("LLM_MODEL", "gemini-1.5-flash-001")
        model = GenerativeModel(llm_model)

        system_prompt = os.getenv("SYSTEM_PROMPT", "You are a strategic analyst for the Department of Information and Communications Technology (DICT). Analyze the following top complaint issues and propose a concrete action plan for each.")

        # Build comprehensive unit guidelines
        unit_guidelines = """
DICT ORGANIZATIONAL STRUCTURE - UNIT ASSIGNMENT GUIDE:

1. DELIVERY UNITS (DICT Internal Services):
   - GDTB (Government Digital Transformation Bureau): eGov services, PBH, government digital transformation
   - FPIAP (Free Public Internet Access Program): Free Wi-Fi, public internet access
   - ILCDB (ICT Literacy and Competency Development Bureau): Training, upskilling, certifications, courses
   - AS (Administrative Service): HR concerns, personnel, recruitment
   - IMB (Infrastructure Management Bureau): Cloud hosting, web hosting, government online services, data centers
   - CSB (Cybersecurity Bureau): Digital certificates, PKI, encryption
   - PRD (Procurement and Records Division): Delivery concerns, courier/logistics (LBC, Ninja Van, J&T, etc.)
   - ROCS (Regional Operations): Regional office concerns

2. ATTACHED AGENCIES (Under DICT supervision):
   - NTC (National Telecommunications Commission): Internet/telco issues, disconnections, slow connection,
     technical service, unsolicited SMS, telecom billing, refunds (PLDT, Globe, Smart, Converge, DITO, etc.)
   - CICC (Cybersecurity Investigation and Coordinating Center): Cybercrime, hacking, phishing, scams, fraud

3. OTHER AGENCIES (External partners):
   - SEC (Securities and Exchange Commission): Harassment, online lending, loan app collections
   - DTI (Department of Trade and Industry): E-commerce, consumer protection, retail refunds, online shopping
"""

        prompt = f"""
        {system_prompt}

        {unit_guidelines}

        Top Complaint Issues (pre-categorized with recommendations and service provider analysis):
        {json.dumps(enriched_issues, indent=2)}

        For each issue, provide:
        1. A specific, actionable ACTION PLAN (1-2 sentences) that addresses the root cause
        2. The UNIT/AGENCY code from the recommendations provided (use the "recommended_unit" field)
        3. Brief REMARKS explaining why this is a priority, including:
           - Total complaint volume
           - Top service provider if available (from "top_service_provider" field)
           - Impact and urgency

        CRITICAL RULES:
        - ALWAYS use the "recommended_unit" provided for each issue - this has been pre-validated
        - If "top_service_provider" is provided, MENTION it in the remarks to identify the main culprit
        - Be specific about coordination mechanisms and deliverables
        - Reference both total complaints and top provider complaints in remarks
        - Align action plans with the unit's mandate shown in the guide above

        Return ONLY a valid JSON array with keys: "issue", "action_plan", "unit", "remarks".
        Do not include markdown formatting like ```json.

        Example format (use ACTUAL data from the enriched_issues provided):
        [
            {{
                "issue": "{{issue_name}}",
                "action_plan": "{{specific_actionable_plan_based_on_unit_mandate}}",
                "unit": "{{recommended_unit}}",
                "remarks": "{{priority_level}} with {{count}} complaints. {{If top_service_provider exists: Top provider: {{top_service_provider}} ({{top_sp_count}} complaints, {{top_sp_percentage}}%).}} {{Impact_statement}}."
            }}
        ]
        """

        response = model.generate_content(prompt)

        # Clean response text to ensure valid JSON
        text = response.text.strip()
        if text.startswith("```json"):
            text = text[7:]
        if text.endswith("```"):
            text = text[:-3]
        text = text.strip()

        ai_plans = json.loads(text)

        # Validate and correct unit assignments
        validated_plans = []
        for i, plan in enumerate(ai_plans):
            # Use the pre-categorized unit if AI didn't assign correctly
            if plan.get("unit") not in DICT_UNIT_MAPPING:
                plan["unit"] = enriched_issues[i]["recommended_unit"]

            validated_plans.append(plan)

        return validated_plans

    except Exception as e:
        st.error(f"AI Generation Error: {str(e)}")
        # Fallback: use pre-categorized units with SP details
        fallback_plans = []
        for issue in enriched_issues:
            # Build remarks with SP info if available
            remarks = f"Priority issue with {issue['count']} complaints."
            if issue.get('top_service_provider'):
                remarks += f" Top provider: {issue['top_service_provider']} ({issue['top_sp_count']} complaints, {issue['top_sp_percentage']}%)."
            remarks += f" Categorized as {issue['org_type']}. Requires immediate action and monitoring."

            # Build action plan with SP focus if applicable
            action_plan = f"Coordinate with {DICT_UNIT_MAPPING[issue['recommended_unit']]['name']} to investigate and resolve this complaint category"
            if issue.get('top_service_provider'):
                action_plan += f", with priority focus on {issue['top_service_provider']}"
            action_plan += " through targeted interventions and stakeholder engagement."

            fallback_plans.append({
                "issue": issue['name'],
                "action_plan": action_plan,
                "unit": issue['recommended_unit'],
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
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=0.5*inch, bottomMargin=0.5*inch)
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

    # Top Issues Summary
    elements.append(Paragraph("Top 5 Priority Issues", heading_style))

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
    elements.append(Paragraph("Action Plan Details", heading_style))

    plan_data = [['Issue', 'Action Plan', 'Assigned Unit', 'Remarks']]
    for _, row in plans_df.iterrows():
        plan_data.append([
            Paragraph(str(row['issue']), styles['Normal']),
            Paragraph(str(row['action_plan']), styles['Normal']),
            Paragraph(str(row['unit']), styles['Normal']),
            Paragraph(str(row['remarks']), styles['Normal'])
        ])

    plan_table = Table(plan_data, colWidths=[1.5*inch, 2.5*inch, 1.2*inch, 1.5*inch])
    plan_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#3b82f6')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
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

    # Service Provider Breakdown Section
    if sp_breakdowns and len(sp_breakdowns) > 0:
        elements.append(Spacer(1, 0.3*inch))
        elements.append(Paragraph("Service Provider Breakdown", heading_style))

        body_style = ParagraphStyle(
            'Body',
            parent=styles['Normal'],
            fontSize=9,
            textColor=colors.HexColor('#374151')
        )

        for sp_item in sp_breakdowns:
            # Issue header
            issue_header = f"{sp_item['issue']} ({sp_item['unit']}) - {sp_item['total_count']} total complaints"
            elements.append(Paragraph(issue_header, styles['Heading3']))
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

            # Top provider note
            if sp_item['breakdown']:
                top_sp = sp_item['breakdown'][0]
                note_text = f"Top provider: {top_sp['provider']} with {top_sp['count']} complaints ({top_sp['percentage']}%)"
                elements.append(Paragraph(note_text, body_style))
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

    plan_table = doc.add_table(rows=1, cols=4)
    plan_table.style = 'Light Grid Accent 1'

    # Header row
    header_cells = plan_table.rows[0].cells
    header_cells[0].text = 'Issue'
    header_cells[1].text = 'Action Plan'
    header_cells[2].text = 'Assigned Unit'
    header_cells[3].text = 'Remarks'

    # Format header
    for cell in header_cells:
        if cell.paragraphs and cell.paragraphs[0].runs:
            cell.paragraphs[0].runs[0].font.bold = True
            cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
        set_cell_background(cell, '3B82F6')

    # Data rows
    for _, row in plans_df.iterrows():
        row_cells = plan_table.add_row().cells
        row_cells[0].text = str(row['issue'])
        row_cells[1].text = str(row['action_plan'])
        row_cells[2].text = str(row['unit'])
        row_cells[3].text = str(row['remarks'])

    # Service Provider Breakdown Section
    if sp_breakdowns and len(sp_breakdowns) > 0:
        doc.add_paragraph()
        doc.add_heading('Service Provider Breakdown', 1)

        for sp_item in sp_breakdowns:
            # Issue subheading
            issue_heading = doc.add_heading(f"{sp_item['issue']} ({sp_item['unit']}) - {sp_item['total_count']} complaints", 2)

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

            # Top provider note
            if sp_item['breakdown']:
                top_sp = sp_item['breakdown'][0]
                note = doc.add_paragraph()
                note_run = note.add_run(f"Top provider: {top_sp['provider']} with {top_sp['count']} complaints ({top_sp['percentage']}%)")
                note_run.font.size = Pt(9)
                note_run.font.color.rgb = RGBColor(55, 65, 81)
                note_run.italic = True

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
        <div class="report-title">📊 AI Action Plan Report</div>
        <div class="report-subtitle">AI-Powered Complaint Analysis & Strategic Recommendations</div>
    </div>
    """, unsafe_allow_html=True)

    # Date and status info
    col_info1, col_info2 = st.columns([2, 1])
    with col_info1:
        st.caption(f"📅 Report Date: {datetime.now().strftime('%A, %B %d, %Y')}")
    with col_info2:
        if 'report_timestamp' in st.session_state:
            st.caption(f"🕐 Last Generated: {st.session_state.report_timestamp.strftime('%H:%M:%S')}")

    st.markdown("<br>", unsafe_allow_html=True)

    if df is None or df.empty:
        st.info("📊 No data available to generate report. Please load complaint data from the **Dashboard** tab first.")
        st.markdown("""
        ### How to get started:
        1. Go to the **Dashboard** tab
        2. Load your complaint data using the sidebar
        3. Return to this tab to generate AI insights
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
        <strong>ℹ️ About This Report:</strong> This AI-powered tool analyzes your complaint data to identify the top 5
        priority issues and generates strategic action plans with recommended DICT units for resolution.
        <br><br>
        <strong>📊 Data Source:</strong> Using {total_records:,} complaints from the dashboard (excluding FLS resolutions){date_range_info}
    </div>
    """, unsafe_allow_html=True)

    # Initialize Vertex AI
    is_init, error_msg = init_vertex_ai()

    if not is_init:
        st.warning("⚠️ Vertex AI not initialized. AI features disabled.")
        with st.expander("ℹ️ Error Details"):
            if error_msg:
                st.caption(f"Error: {error_msg}")
                st.caption("Ensure you have 'Vertex AI User' role and the API is enabled in Google Cloud.")

    # Get Top Issues (always compute to show preview)
    top_issues = get_top_issues(df)

    if not top_issues:
        st.warning("⚠️ Not enough data to identify top issues. Please ensure your data has 'Complaint Category' or 'Complaint Nature' columns.")
        return

    # Preview Top Issues with Unit Recommendations
    with st.expander("📊 Preview: Top Issues Detected & Recommended Units", expanded=False):
        # Enrich issues with unit recommendations
        preview_data = []
        for issue in top_issues:
            unit_code, unit_name, org_type = categorize_issue_to_unit(issue['name'], issue['type'])
            preview_data.append({
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
                "Issue": st.column_config.TextColumn("Issue Description", width="large"),
                "Source": st.column_config.TextColumn("Source", width="small"),
                "Count": st.column_config.NumberColumn("Complaints", format="%d", width="small"),
                "Recommended Unit": st.column_config.TextColumn("Assigned To", width="medium"),
                "Organization": st.column_config.TextColumn("Type", width="medium")
            },
            hide_index=True,
            use_container_width=True
        )
        st.caption("💡 **Source**: Category/Nature field | **Assigned To**: Auto-categorized DICT unit/agency | **Type**: Organization classification")

    st.markdown("---")

    # Generation Button - Centered and prominent
    col_spacer1, col_btn, col_spacer2 = st.columns([1, 2, 1])

    with col_btn:
        generate_button = st.button(
            "🤖 Generate AI Action Plan",
            type="primary",
            use_container_width=True,
            help="Analyze top complaints and generate strategic action plans"
        )

    # Handle generation
    if generate_button or ('weekly_action_plan' in st.session_state and st.session_state.get('report_generated', False)):
        if generate_button:
            with st.spinner("🤖 AI is analyzing top complaints and generating strategic action plans..."):
                if is_init:
                    action_plan_data = generate_ai_action_plan(top_issues, df)  # Pass df for SP analysis
                else:
                    # Fallback if no AI
                    action_plan_data = [
                        {
                            "issue": i['name'],
                            "action_plan": "Analyze root cause and coordinate with service providers to resolve this issue.",
                            "unit": "Operations Team",
                            "remarks": f"High volume detected: {i['count']} complaints"
                        } for i in top_issues
                    ]
                st.session_state.weekly_action_plan = action_plan_data
                st.session_state.report_generated = True
                st.session_state.report_timestamp = datetime.now()
                st.success("✅ Action plan generated successfully!")

        # Display the action plan
        if 'weekly_action_plan' in st.session_state:
            plans = st.session_state.weekly_action_plan

            # Convert to DataFrame
            report_df = pd.DataFrame(plans)

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("---")
            st.markdown("### 📋 Strategic Action Plan")
            st.caption(f"Generated {len(report_df)} strategic recommendations based on top complaint patterns")

            # Responsive table with horizontal scroll
            import html

            # Create HTML table for better responsiveness
            table_html = '<div class="responsive-table-container">\n'
            table_html += '<table class="action-plan-table">\n'
            table_html += '<thead>\n<tr>\n'
            table_html += '<th class="col-issue">Top Issue</th>\n'
            table_html += '<th class="col-action">Action Plan</th>\n'
            table_html += '<th class="col-unit">Assigned Unit</th>\n'
            table_html += '<th class="col-remarks">Remarks</th>\n'
            table_html += '</tr>\n</thead>\n<tbody>\n'

            for _, row in report_df.iterrows():
                # Escape HTML to prevent rendering issues
                issue = html.escape(str(row['issue']))
                action = html.escape(str(row['action_plan']))
                unit = html.escape(str(row['unit']))
                remarks = html.escape(str(row['remarks']))

                table_html += '<tr>\n'
                table_html += f'<td class="col-issue"><strong>{issue}</strong></td>\n'
                table_html += f'<td class="col-action">{action}</td>\n'
                table_html += f'<td class="col-unit">{unit}</td>\n'
                table_html += f'<td class="col-remarks">{remarks}</td>\n'
                table_html += '</tr>\n'

            table_html += '</tbody>\n</table>\n</div>'

            st.markdown(table_html, unsafe_allow_html=True)

            # Service Provider Breakdown for PRD and NTC issues
            st.markdown("---")
            st.markdown("### 📊 Service Provider Breakdown")
            st.caption("Detailed breakdown for Delivery Concerns and Telco/Internet Issues")

            # Check which issues need SP breakdown
            issues_with_breakdown = []
            for idx, row in report_df.iterrows():
                unit = row['unit']
                issue_name = row['issue']

                # Find matching issue from top_issues
                matching_issue = next((i for i in top_issues if i['name'] == issue_name), None)

                if matching_issue and unit in UNITS_REQUIRING_SP_BREAKDOWN:
                    sp_breakdown = get_service_provider_breakdown(df, issue_name, matching_issue['type'])
                    if sp_breakdown:
                        issues_with_breakdown.append({
                            "issue": issue_name,
                            "unit": unit,
                            "unit_label": UNITS_REQUIRING_SP_BREAKDOWN[unit],
                            "total_count": matching_issue['count'],
                            "breakdown": sp_breakdown
                        })

            if issues_with_breakdown:
                for item in issues_with_breakdown:
                    with st.expander(f"🔍 {item['issue']} ({item['unit']}) - {item['total_count']} complaints", expanded=True):
                        # Create breakdown table
                        sp_df = pd.DataFrame(item['breakdown'])

                        st.dataframe(
                            sp_df,
                            column_config={
                                "provider": st.column_config.TextColumn("Service Provider", width="large"),
                                "count": st.column_config.NumberColumn("Complaints", format="%d", width="small"),
                                "percentage": st.column_config.NumberColumn("Percentage", format="%.1f%%", width="small")
                            },
                            hide_index=True,
                            use_container_width=True
                        )

                        # Summary stats
                        st.caption(f"📌 Top provider: **{item['breakdown'][0]['provider']}** with {item['breakdown'][0]['count']} complaints ({item['breakdown'][0]['percentage']}%)")
                        st.caption(f"📊 Total providers: {len(item['breakdown'])}")
            else:
                st.info("ℹ️ No Delivery Concerns or Telco/Internet Issues in top 5 complaints that require service provider breakdown.")

            st.markdown("---")

            # Download Buttons Section
            st.markdown("### 📥 Export Options")
            st.caption("Download the action plan in your preferred format")

            col_dl1, col_dl2, col_dl3 = st.columns(3)

            with col_dl1:
                # PDF Download
                try:
                    pdf_buffer = export_to_pdf(report_df, top_issues, issues_with_breakdown)
                    st.download_button(
                        label="📄 PDF Document",
                        data=pdf_buffer,
                        file_name=f"DICT_AI_Action_Plan_{datetime.now().strftime('%Y%m%d')}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                        help="Download formatted PDF report with SP breakdown"
                    )
                except Exception as e:
                    st.error(f"❌ PDF Export Error: {str(e)}")

            with col_dl2:
                # Word Download
                try:
                    word_buffer = export_to_word(report_df, top_issues, issues_with_breakdown)
                    st.download_button(
                        label="📝 Word Document",
                        data=word_buffer,
                        file_name=f"DICT_AI_Action_Plan_{datetime.now().strftime('%Y%m%d')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                        help="Download editable Word document with SP breakdown"
                    )
                except Exception as e:
                    st.error(f"❌ Word Export Error: {str(e)}")

            with col_dl3:
                # CSV Download
                csv_buffer = report_df.to_csv(index=False)
                st.download_button(
                    label="📊 CSV Spreadsheet",
                    data=csv_buffer,
                    file_name=f"DICT_AI_Action_Plan_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv",
                    use_container_width=True,
                    help="Download data as CSV file"
                )

            # Summary Metrics
            st.markdown("---")
            st.markdown("### 📈 Report Summary")

            m1, m2, m3, m4 = st.columns(4)

            total_complaints = sum([i['count'] for i in top_issues])
            unique_units = report_df['unit'].nunique()

            m1.metric(
                label="Issues Analyzed",
                value=len(report_df),
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
            st.markdown("### 🏢 Assigned Units & Top Service Providers")
            st.caption("Specific agencies and service providers requiring action")

            # Build detailed breakdown with service providers
            unit_details = []
            for idx, row in report_df.iterrows():
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
            st.markdown("#### Summary by Organization Type")
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
                st.warning(f"⚠️ {len(unclassified)} issue(s) could not be categorized properly")
                for item in unclassified:
                    st.caption(f"• {item['Unit Code']} - {item['Issue']}")
            else:
                st.success("✅ All issues successfully categorized to appropriate units/agencies")

    else:
        # Show placeholder when no report is generated
        st.markdown("<br>", unsafe_allow_html=True)
        st.info("👆 Click the **Generate AI Action Plan** button above to analyze your complaint data and create strategic recommendations.")
