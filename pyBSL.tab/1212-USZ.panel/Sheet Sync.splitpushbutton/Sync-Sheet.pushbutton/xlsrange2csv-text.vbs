If WScript.Arguments.Count < 3 Then
    WScript.Echo "Error! Please specify the input file, the named range, and the output file paths."
    Wscript.Quit
End If

Dim excel
Set excel = CreateObject("Excel.Application")
excel.DisplayAlerts = False ' Unterdrückt mögliche Dialoge von Excel

Dim workbook
Set workbook = excel.Workbooks.Open(WScript.Arguments.Item(0))

' Hole den benannten Bereich
Dim namedRange
On Error Resume Next
Set namedRange = workbook.Names(WScript.Arguments.Item(1)).RefersToRange
If Err.Number <> 0 Then
    WScript.Echo "Error! Named range '" & WScript.Arguments.Item(1) & "' not found."
    workbook.Close False
    excel.Quit
    WScript.Quit
End If
On Error GoTo 0

' Neue Arbeitsmappe erstellen
Dim tempWorkbook
Set tempWorkbook = excel.Workbooks.Add

' Anzahl Zeilen und Spalten bestimmen
Dim rowCount, colCount
rowCount = namedRange.Rows.Count
colCount = namedRange.Columns.Count

' Zellen einzeln in die neue Mappe übertragen - mit Textformat
Dim r, c
For r = 1 To rowCount
    For c = 1 To colCount
        tempWorkbook.Sheets(1).Cells(r, c).Value = namedRange.Cells(r, c).Text
    Next
Next

' Speichere die temporäre Arbeitsmappe als CSV
tempWorkbook.SaveAs WScript.Arguments.Item(2), 6 ' Format 6 entspricht CSV (Windows)

tempWorkbook.Close False
workbook.Close False
excel.Quit

' Objekte freigeben
Set namedRange = Nothing
Set tempWorkbook = Nothing
Set workbook = Nothing
Set excel = Nothing
