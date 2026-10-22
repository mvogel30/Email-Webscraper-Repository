@echo off
REM Double-click launcher for the expert-commentary scraper.
REM Runs BOTH the web scraper and the vendor email pipeline, then keeps this window open
REM so you can read the result.
title Expert Commentary Scraper
echo Running the expert-commentary scraper (web + vendor emails)...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0run_daily_full_update.ps1"
echo.
echo ============================================================
echo  Done. Shared team workbook (everyone reads this):
echo    Shared Documents\03 Cross-Team Projects\
echo    08 - Geopolitics Sensing ^& Scraping Tools\Webscraping ^& Email Scraping tool\
echo    Expert Commentary Repository (review).xlsx
echo  Your personal quick-look copy: Documents\Expert Commentary Repository\
echo ============================================================
echo.
pause
