@echo off
cd /d "%~dp0"
set /p olddb=Paste the full path to your existing bam_motor_group.db file: 
python tools\import_existing_database.py "%olddb%"
pause
