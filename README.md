# Real-Time Dashboard with Google Sheets Integration

A simple, robust, and presentable Streamlit dashboard that connects to Google Sheets in real-time for data analysis and visualization.

## Features

- Real-time connection to Google Sheets
- Automatic data analysis using pandas
- Interactive visualizations with Plotly
- Support for both public and private Google Sheets
- Auto-refresh capability
- Data export functionality
- Clean and presentable UI

## Quick Start

### 1. Setup

Run the setup script to create a virtual environment and install dependencies:

```bash
setup.bat
```

Or manually:

```bash
# Create virtual environment
python -m venv venv

# Activate virtual environment
venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

### 2. Run the Dashboard

```bash
# Make sure venv is activated
venv\Scripts\activate

# Run the app
streamlit run dashboard.py
```

The dashboard will open in your browser at `http://localhost:8501`

## Using with Google Sheets

### ⭐ Option 1: Public Google Sheet (RECOMMENDED - No Setup Required!)

**This is the easiest and fastest method - just paste any public Google Sheets link!**

1. Open your Google Sheet
2. Click **Share** button (top right)
3. Click **"Change to anyone with the link"**
4. Ensure permission is set to **Viewer**
5. Click **Copy link**
6. In the dashboard sidebar:
   - Select **"Google Sheets"**
   - Choose **"Public (Anyone with link)"**
   - Paste your URL
   - Data loads automatically! ✅

**Benefits:**
- ✅ No Google Cloud setup required
- ✅ No service account needed
- ✅ No authentication files
- ✅ Works with any "Anyone with link" sheet
- ✅ Real-time updates with auto-refresh
- ✅ Simple URL-based access

**How it works:**
The dashboard uses Google's public CSV export API to fetch data directly from your sheet. This is fast, reliable, and requires zero configuration!

### Option 2: Private Google Sheet (Advanced - Service Account)

**Use this only if your sheet contains sensitive data that cannot be publicly shared**

1. Create a Google Cloud Project
2. Enable Google Sheets API
3. Create a Service Account and download the JSON credentials
4. Share your Google Sheet with the service account email (found in the JSON file)
5. In the dashboard:
   - Select "Google Sheets"
   - Choose "Private (Service Account)"
   - Upload the JSON credentials file
   - Enter your sheet URL
   - Click **Refresh Data**

**When to use this:**
- For sensitive/confidential data
- For sheets that must remain private
- When you need granular access control

## Dashboard Features

### Overview Metrics
- Total rows and columns
- Number of numeric/categorical columns
- Missing values count

### Visualizations
- Bar charts for categorical vs numeric data
- Pie charts for distribution analysis
- Line charts for trends
- Correlation heatmap for multiple numeric columns

### Data Analysis
- Summary statistics for numeric columns
- Data preview (first 10 rows)
- Export to CSV functionality

## Project Structure

```
ccc_dashboard/
├── dashboard.py          # Main Streamlit application
├── requirements.txt      # Python dependencies
├── setup.bat            # Setup script for Windows
├── README.md            # This file
└── venv/                # Virtual environment (created after setup)
```

## Requirements

- Python 3.8 or higher
- Internet connection for Google Sheets access

## Dependencies

- streamlit: Web app framework
- pandas: Data analysis
- gspread: Google Sheets API
- google-auth: Authentication
- plotly: Interactive visualizations
- openpyxl: Excel file support

## Troubleshooting

### "Error loading data"
- Ensure the Google Sheet is publicly accessible or credentials are correct
- Check that the URL is complete and correct
- Verify the sheet contains data

### "Module not found"
- Make sure you've activated the virtual environment
- Run `pip install -r requirements.txt` again

### Auto-refresh not working
- Check your internet connection
- Ensure the sheet URL hasn't changed
- Try manual refresh first

## Tips

- Use meaningful column names in your Google Sheet
- Keep data organized with headers in the first row
- The dashboard automatically detects numeric vs categorical data
- Large datasets (>10,000 rows) may take longer to load

## License

Open source - feel free to modify and use as needed
