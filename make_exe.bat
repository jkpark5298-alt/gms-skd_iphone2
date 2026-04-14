@echo off
python -m pip install pyinstaller
python -m PyInstaller --noconfirm --onefile --windowed --name AirZeta_Auto airzeta_automation.py
pause