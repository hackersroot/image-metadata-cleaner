@echo off
cd /d "%~dp0"
python image_metadata_cleaner.py
if errorlevel 1 (
  echo.
  echo Install dependencies with: python -m pip install -r requirements.txt
  pause
)
