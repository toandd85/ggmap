@echo off
chcp 65001 > nul
echo ===================================================
echo     GOOGLE MAPS SCRAPER - GIAO DIEN TIENG VIET
echo ===================================================
echo Dang khoi dong may chu web tai http://127.0.0.1:5000 ...
start http://127.0.0.1:5000
python viewer_server.py
pause
