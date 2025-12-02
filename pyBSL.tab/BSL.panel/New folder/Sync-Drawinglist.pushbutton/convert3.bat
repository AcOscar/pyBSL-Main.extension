rem @echo off
SET bat_path=%~dp0

rem cscript //NoLogo "%bat_path%xlsrange2csv.vbs" "%~1" "%~2" "%~3" 	
"%bat_path%xlsrange2csv.vbs" "%~1" "%~2" "%~3"
