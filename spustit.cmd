@echo off
rem Spustí klirenc.ps1 dvojklikem; systémovou politiku nemění.
start "" powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0klirenc.ps1"
