@echo off

start "" "C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE"

cd /d D:\Automation\Abbvie-PMF-Automation

git pull

call venv\Scripts\activate.bat

python PowerFlowSummary.py >> execution.log 2>&1

exit