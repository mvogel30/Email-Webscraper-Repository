@echo off
REM Double-click launcher for the geopolitics article scraper.
REM Runs BOTH the think-tank web scraper and the Signum Iran email pipeline, then keeps this
REM window open so you can read the result.
title Geopolitics Article Scraper
echo Running the geopolitics article scraper (web + Signum email)...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0run_daily_full_update.ps1"
echo.
echo ============================================================
echo  Done. Shared team workbook (everyone reads this):
echo    Global PP+G - Documents\06 Sensing ^(Inputs^)\
echo    10 - Geopolitics Sensing ^& Scraping Tools\Iran Scraping example\
echo    Geopolitics Article Repository (review).xlsx
echo  Your personal quick-look copy: Documents\Geopolitics Article Repository\
echo ============================================================
echo.
pause
