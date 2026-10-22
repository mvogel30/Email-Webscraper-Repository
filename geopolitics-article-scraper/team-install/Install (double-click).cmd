@echo off
REM Double-click installer. Asks what time you want the daily run, then does everything else.
REM (Use this generic version only if you're not on the US or India team — those have their own
REM  pre-configured installer that needs no typing. See TEAM-INSTALL-README.md.)
title Geopolitics Article Scraper - Install
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$t = Read-Host 'What time should this run every day (24h, e.g. 08:30, in YOUR machine''s local time)'; & '%~dp0Install-Teammate.ps1' -RunAtTime $t"
echo.
pause
