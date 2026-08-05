@echo off
cd /d "%~dp0"
if not exist "database\bam_motor_group.db" (
  echo No BAM database was found.
  pause
  exit /b 1
)
for /f "tokens=1-4 delims=/ " %%a in ("%date%") do set datestr=%%d-%%c-%%b
for /f "tokens=1-2 delims=: " %%a in ("%time%") do set timestr=%%a%%b
copy "database\bam_motor_group.db" "backups\bam_motor_group_%datestr%_%timestr%.db"
echo Backup created in the backups folder.
pause
