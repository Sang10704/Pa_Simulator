@echo off
chcp 65001 >nul
start "" "http://localhost:8095"
python "%~dp0sandbox_server.py"
