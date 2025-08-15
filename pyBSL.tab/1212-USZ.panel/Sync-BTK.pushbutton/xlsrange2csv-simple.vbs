If WScript.Arguments.Count < 3 Then
    WScript.Echo "Error! Please specify the input file, the named range, and the output file paths."
    Wscript.Quit
End If

Dim excel
Set excel = CreateObject("Excel.Application")
excel.DisplayAlerts = False ' unterdrückt mögliche Dialoge

Dim workbook
Set workbook = excel.Workbooks.Open(WScript.Arguments.Item(0))

' Benannten Bereich holen
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

' Temporäre Arbeitsmappe
Dim tempWorkbook
Set tempWorkbook = excel.Workbooks.Add

' Kopieren ohne PasteSpecial:
namedRange.Copy tempWorkbook.Sheets(1).Range("A1")

' Speichern als CSV
tempWorkbook.SaveAs WScript.Arguments.Item(2), 6 ' 6 = xlCSV (Windows)
tempWorkbook.Close False
workbook.Close False
excel.Quit

' Objekte freigeben
Set namedRange = Nothing
Set tempWorkbook = Nothing
Set workbook = Nothing
Set excel = Nothing
