@echo off
echo ============================================
echo PLI Network Analysis - Setup and Launch
echo ============================================
echo.
echo Installing required packages...
pip install pandas numpy matplotlib scipy openpyxl statsmodels PyQt5 --quiet 2>nul
echo.
echo Starting GUI...
python network_analysis_gui.py
if errorlevel 1 (
    echo.
    echo ERROR: Failed to start GUI
    echo Make sure Python is installed and in PATH
    pause
)
