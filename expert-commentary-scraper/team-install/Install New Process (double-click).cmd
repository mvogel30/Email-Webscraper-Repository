@echo off
REM Double-click installer for the NEW Process & Scope specifically. Asks what time you want the
REM daily run, then does everything else. Do NOT use "Install (double-click).cmd" in this same
REM folder for a research-team handover - that one sets up the OLD Process instead.
title Expert Commentary Scraper - Install (New Process)
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$t = Read-Host 'What time should this run every day (24h, e.g. 08:30, in YOUR machine''s local time)'; & '%~dp0Install-Teammate-NewProcess.ps1' -RunAtTime $t"
echo.
pause
