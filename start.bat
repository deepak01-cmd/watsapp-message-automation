@echo off
cd /d "%~dp0"
echo Installing / checking required packages (first time only)...
python -m pip install -r requirements.txt
echo.
echo Starting WhatsApp Sender - your browser will open by itself.
python app.py
pause
