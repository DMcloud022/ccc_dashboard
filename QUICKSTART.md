# Quick Start Guide - Real-time Dashboard

## 🚀 Get Started in 3 Minutes!

### Step 1: Install and Run

```bash
# Double-click or run:
setup.bat

# Then activate and run:
venv\Scripts\activate
streamlit run dashboard.py
```

Dashboard opens at `http://localhost:8501`

### Step 2: Prepare Your Google Sheet

**⭐ NEW: Public Sheet Method (EASIEST - No Setup Required!)**

1. Open your Google Sheet
2. Click **Share** button (top right)
3. Click **"Change to anyone with the link"**
4. Set permission to **Viewer**
5. Click **Copy link**

✅ **That's it! No Google Cloud setup, no authentication files needed!**

**Option B: Private Sheet (Advanced)**
- For sensitive data that can't be shared publicly
- Requires Google Cloud service account setup
- Follow detailed instructions in README.md

<!-- Deployment defaults and sample fallback removed — app now uses the fixed default Google Sheet URL. -->

### Step 3: Connect to Dashboard

1. In the sidebar, select **"Google Sheets"**
2. Choose **"Public (Anyone with link)"** (for public sheets)
3. Paste your Google Sheet URL
4. Data loads automatically! ✅
5. Enable **Auto-refresh** for real-time updates

---

## 🆕 What's New?

### Major Improvements:

✅ **Zero Configuration for Public Sheets**
   - Previously: Required complex Google Cloud setup
   - Now: Just paste any "Anyone with link" Google Sheets URL!

✅ **Robust Real-time Data Fetching**
   - Uses Google's CSV export API
   - No authentication needed for public sheets
   - Fast and reliable

✅ **Better Error Handling**
   - Clear, helpful error messages
   - Step-by-step troubleshooting guides
   - Automatic URL format detection

✅ **Flexible Access Methods**
   - Public sheets: No setup required
   - Private sheets: Service account support available

## Example Google Sheet Format

Your sheet should have headers in the first row:

| Product | Sales | Region | Date |
|---------|-------|--------|------|
| Widget A | 1000 | North | 2024-01-01 |
| Widget B | 1500 | South | 2024-01-02 |

The dashboard will automatically:
- Detect numeric columns (Sales)
- Detect categorical columns (Product, Region)
- Create relevant visualizations
- Calculate statistics

## Features You'll See

1. **4 Interactive Charts**: Professional full-screen visualizations in 2x2 grid
2. **Overview Metrics**: Row count, column count, missing values
3. **Presentation Mode**: Clean, dark theme optimized for presentations
4. **Auto-Refresh**: Real-time data updates (configurable 10-300 seconds)
5. **Data Preview**: Collapsible data table
6. **Statistics**: Mean, median, std dev, etc.
7. **Export**: Download processed data as CSV

## 🎥 Presentation Mode

For full-screen presentations:

1. Toggle **"🎥 Presentation Mode"** in the sidebar
2. Press `F11` for full-screen browser mode
3. Enable **Auto-refresh** for real-time updates
4. Use keyboard shortcut `C` to collapse the sidebar

**Presentation Features:**
- 4 charts displayed simultaneously in dark theme
- Compact metrics bar at the top
- Large, readable fonts
- Professional color schemes
- Hidden navigation elements

## 🔧 How It Works (Technical Details)

### Public Sheet Method

When you paste a Google Sheets URL like:
```
https://docs.google.com/spreadsheets/d/ABC123xyz/edit#gid=0
```

The dashboard:
1. **Extracts** sheet ID (`ABC123xyz`) and GID (`0`)
2. **Converts** to CSV export URL:
   ```
   https://docs.google.com/spreadsheets/d/ABC123xyz/export?format=csv&gid=0
   ```
3. **Fetches** data using `pandas.read_csv()` - no auth needed!
4. **Caches** for 60 seconds (configurable)
5. **Auto-refreshes** based on your settings

**Why this works:**
- Google Sheets provides a public CSV export endpoint
- Works for any sheet with "Anyone with link" access
- No API quotas or rate limits for public sheets
- Fast and reliable

---

## 🐛 Troubleshooting

### "Sheet not found" or 404 Error
**Solution:** Make sure your sheet is publicly accessible
1. Open Google Sheet → Share
2. Click "Anyone with the link"
3. Set to "Viewer"
4. Copy the complete URL

### "Access denied" or 403 Error
**Solution:** Sheet is still private
- Double-check sharing settings
- Ensure it says "Anyone with the link can view"
- Don't use "Restricted" access

### Data Not Updating
**Solutions:**
1. Click **"🔄 Refresh Now"** button in sidebar
2. Click **"🗑️ Clear Cache"** button at bottom
3. Check internet connection
4. Verify the sheet URL hasn't changed

### Wrong Tab/Sheet Loading
**Solution:** URL might not include the GID
1. Click on the specific tab you want in Google Sheets
2. Copy the URL (should include `#gid=XXXXX`)
3. Paste the complete URL in dashboard

### Dashboard Won't Start
**Solutions:**
- Activate virtual environment: `venv\Scripts\activate`
- Check Python installed: `python --version`
- Reinstall dependencies: `pip install -r requirements.txt`

### Data Looks Wrong
**Check:**
- First row contains column headers
- Numbers are formatted as numbers (not text)
- No completely empty rows
- Column names are unique

## 💡 Tips for Best Results

### For Public Sheets:
- ✅ Use clear, descriptive column names
- ✅ Put headers in the first row
- ✅ Use consistent data types in each column
- ✅ Remove empty rows/columns
- ✅ Keep sheet size reasonable (<100,000 rows for fast loading)

### For Real-time Updates:
- ✅ Enable **Auto-refresh** in sidebar
- ✅ Set refresh interval to 30-60 seconds
- ✅ Keep dashboard tab active in browser
- ✅ Use public sheet method for fastest performance

### For Large Datasets:
- ✅ Increase cache TTL if data doesn't change often
- ✅ Use specific date ranges to filter data
- ✅ Consider splitting into multiple sheets by category

### For Presentations:
- ✅ Use **Presentation Mode** for clean, full-screen view
- ✅ Enable auto-refresh for live data demos
- ✅ Press F11 for browser full-screen
- ✅ Press C to hide Streamlit sidebar

---

## 🎯 Quick Test

Want to test immediately?

1. **Create test sheet:**
   - Open new Google Sheet
   - Add headers: `Name`, `Age`, `City`
   - Add 2-3 rows of sample data

2. **Share it:**
   - Share → "Anyone with link" → Viewer
   - Copy URL

3. **Load in dashboard:**
   - Paste URL in sidebar
   - Select "Public (Anyone with link)"
   - Watch it load instantly! ✅

4. **Test real-time:**
   - Enable auto-refresh (30 seconds)
   - Edit data in Google Sheet
   - Watch dashboard update automatically!

---

## 📊 Supported URL Formats

All these formats work automatically:
```
https://docs.google.com/spreadsheets/d/ABC123/edit
https://docs.google.com/spreadsheets/d/ABC123/edit#gid=0
https://docs.google.com/spreadsheets/d/ABC123/edit#gid=123456
https://docs.google.com/spreadsheets/d/ABC123/edit?usp=sharing
```

The dashboard extracts the ID and GID from any Google Sheets URL format!

---

## 🚀 Next Steps

- ✅ Enable auto-refresh for live updates
- ✅ Explore different visualizations
- ✅ Filter data by date ranges
- ✅ Export analysis as CSV
- ✅ Share the dashboard with your team
- ✅ Try presentation mode for meetings

## 📚 Additional Resources

- **README.md** - Detailed documentation
- **dashboard.py** - Source code with comments
- **requirements.txt** - All dependencies

---

## 🎉 Summary

**Before:** Complex Google Cloud setup, service accounts, JSON credentials
**Now:** Just paste a link and go! 🚀

**The new public sheet method makes real-time data analysis accessible to everyone - no technical setup required!**

Need help? Check README.md or review error messages in the dashboard (they include helpful troubleshooting tips).
