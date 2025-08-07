rem @echo off
SET bat_path=%~dp0

"%bat_path%xlsrange2csv.vbs" "%~1" "%~2" "%~3"
