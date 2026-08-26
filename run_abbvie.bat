@echo off

echo Started %date% %time% > task_test.log

cd /d D:\Automation\Abbvie-PMF-Automation

echo Changed directory >> task_test.log

git pull >> task_test.log 2>&1

call venv\Scripts\activate.bat

echo Venv activated >> task_test.log

python PowerFlowSummary.py >> task_test.log 2>&1

echo Finished >> task_test.log

exit