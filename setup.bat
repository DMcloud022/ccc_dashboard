@echo off
echo Creating virtual environment...
python -m venv venv

echo.
echo Activating virtual environment...
call venv\Scripts\activate.bat

echo.
echo Installing dependencies...
pip install -r requirements.txt

echo.
echo Setup complete! To run the dashboard:
echo 1. Activate the virtual environment: venv\Scripts\activate
echo 2. Run the app: streamlit run dashboard.py
pause
