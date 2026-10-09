@echo off
rem Starts the Travel Accommodation Portal with Waitress. Used by the "TravelPortal" scheduled task.
cd /d "%~dp0.."
if not exist logs mkdir logs
".venv\Scripts\python.exe" app.py >> logs\portal.log 2>&1
