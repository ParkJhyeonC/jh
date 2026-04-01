@echo off
echo ===================================================
echo Starting Emotion and Behavior Analysis Server...
echo ===================================================
echo.
echo Installing dependencies... (This may take a minute)
call npm install
echo.
echo Starting server. Open http://localhost:3000 in your browser.
call npm run dev
pause
