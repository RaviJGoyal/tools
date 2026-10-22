@echo off
rem Starts PPTXProof without a console window. Presentations dropped onto this file are preselected.
start "" /min powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%~dp0PPTXProof.ps1" %*
