# Column Validation & Error Handling Guide

## Overview

The dashboard has been enhanced with robust column validation and error handling to gracefully manage missing or misnamed columns in your data.

## Key Improvements

### 1. **Automatic Column Detection & Mapping**

The system now automatically detects common variations of column names:

- **Date Received** → Also accepts: "Date of Complaint", "Complaint Date", "Date Filed", "Filing Date"
- **Complaint Category** → Also accepts: "Category", "Type", "Complaint Type"
- **Complaint Nature** → Also accepts: "Nature", "Nature of Complaint"
- **Service Providers** → Also accepts: "Service Provider", "Provider", "ISP"
- **Agency** → Also accepts: "Department", "Office"

### 2. **Fuzzy Column Name Matching**

If exact matches aren't found, the system uses fuzzy matching to suggest similar column names, helping you identify typos or slight variations.

### 3. **Data Preview & Diagnostics**

When you load data, expand the **"📋 Column Information & Data Preview"** section in the sidebar to see:
- All available columns in your dataset
- Preview of the first 5 rows
- Column data types

### 4. **Graceful Error Handling**

Instead of crashing, the dashboard now:
- Shows empty charts with "No data available" messages
- Displays clear error messages indicating which columns are missing
- Provides helpful suggestions for fixing column name issues
- Continues to show visualizations that can be rendered with available data

## Troubleshooting Column Errors

### Error: "Column 'Complaint Category' not found"

**Solution:**
1. Open the sidebar and expand **"📋 Column Information & Data Preview"**
2. Check if your data has a similar column (e.g., "Category", "Type")
3. Look for suggestions provided by the system
4. If needed, rename your column in the source data to match one of the accepted names

### Error: "Column 'Date Received' not found"

**Solution:**
1. Ensure your data has a date column
2. Accepted names: "Date Received", "Date of Complaint", "Complaint Date", "Date Filed", "Filing Date"
3. The column should contain valid dates in formats like:
   - YYYY-MM-DD (2025-01-15)
   - MM/DD/YYYY (01/15/2025)
   - DD/MM/YYYY (15/01/2025)

### Message: "No category data available in this period"

This is **not an error** - it means:
- The column exists, but there's no data for the selected time period
- All values are empty/null for that period
- This is normal if you're filtering to recent months with no complaints

## Data Structure Requirements

### Required Columns (Critical)
- **Date Received** - Required for all time-based filtering and trends

### Optional Columns (Recommended)
- **Complaint Category** - For category analysis
- **Complaint Nature** - For nature analysis
- **Service Providers** - For provider analysis
- **Agency** - For filtering by NTC/PEMEDES

### Column Naming Best Practices

1. **Use exact names** from the list above for best compatibility
2. **Avoid special characters** in column names
3. **Keep names consistent** across all your data files
4. **Remove extra spaces** from column names
5. **Use Title Case** for readability

## Example Data Structure

```csv
Date Received,Complaint Category,Complaint Nature,Service Providers,Agency
2025-01-15,Service Quality,Slow Internet,ISP Provider A,NTC
2025-01-16,Billing,Overcharge,Telco B,PEMEDES
```

## Testing Your Data

1. Load your data file
2. Check the "📋 Column Information & Data Preview" section
3. Verify that key columns are detected (look for green checkmarks)
4. If columns are missing, review the suggestions provided
5. Update your source data file if needed and reload

## Advanced: Manual Column Mapping

If your column names are significantly different, you have two options:

### Option 1: Rename in Source File (Recommended)
Edit your Google Sheet or Excel file to use the standard column names.

### Option 2: Update the Validation Function
Edit `dashboard.py` around line 237 to add your custom column names to the `required_columns` dictionary:

```python
required_columns = {
    'Date Received': ['Date Received', 'Date of Complaint', 'Complaint Date', 'YOUR_CUSTOM_NAME_HERE'],
    # ... other columns
}
```

## Support

If you continue to experience issues:
1. Check the data preview to see what columns are actually loaded
2. Verify your data has the minimum required column (Date Received)
3. Ensure date values are in a recognizable format
4. Check for typos in column names
