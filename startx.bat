@echo off
echo Installing requirements...
pip install -r requirements.txt

echo.
echo Starting Task Runner...
python task_runner.py