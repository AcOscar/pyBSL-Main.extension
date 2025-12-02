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

' Erstelle eine temporäre Arbeitsmappe und kopiere den Named Range dorthin
Dim tempWorkbook
Set tempWorkbook = excel.Workbooks.Add
namedRange.Copy
tempWorkbook.Sheets(1).Range("A1").PasteSpecial -4163 ' -4163 entspricht xlPasteValues
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
