@echo off
SET bat_path=%~dp0

REM Pfade in Anführungszeichen setzen, um Leerzeichen zu behandeln
cscript //NoLogo "%bat_path%xls2csv.vbs" "%~1" "%~2" 	