@echo off
echo.
echo BAM iPhone Address Helper
echo =========================
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /c:"IPv4 Address"') do echo Open Safari and try: http:%%a:5000
echo.
echo Your iPhone must be on the same Wi-Fi network.
pause
