@echo off
echo ========================================
echo  WE Legal Automation
echo ========================================
echo.

pip install -r requirements.txt

echo.
echo Starting app...
echo.

streamlit run app.py
pause
