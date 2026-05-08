@echo off
cd /d "%~dp0"
start http://localhost:8787/비서앱.html
python -m http.server 8787
