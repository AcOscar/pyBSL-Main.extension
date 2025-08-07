If WScript.Arguments.Count < 2 Then
    WScript.Echo "Error! Please specify the input file and output file paths."
    Wscript.Quit
End If

Dim excel
Set excel = CreateObject("Excel.Application")
excel.DisplayAlerts = False ' Unterdrückt mögliche Dialoge von Excel

Dim workbook
Set workbook = excel.Workbooks.Open(Wscript.Arguments.Item(0))

' Verwende das erste Blatt
Dim sheet
Set sheet = workbook.Sheets(1)

sheet.Activate

' Speichere die Datei als CSV
workbook.SaveAs WScript.Arguments.Item(1), 6 ' Format 6 entspricht CSV (Windows)

workbook.Close False
excel.Quit

' Objekte freigeben
Set sheet = Nothing
Set workbook = Nothing
Set excel = Nothing
