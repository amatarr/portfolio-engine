@echo off
cd /d "%~dp0"
echo Starting live quotes push loop (Excel -> Supabase -> Streamlit).
echo Leave this window open while you want live prices on the website.
echo Closing this window stops the live updates.
echo.

:loop
py push_quotes_supabase.py
echo.
echo Push loop exited unexpectedly -- restarting in 5 seconds...
timeout /t 5 /nobreak >nul
goto loop
