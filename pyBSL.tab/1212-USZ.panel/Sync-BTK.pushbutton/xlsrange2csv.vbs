'-------------------------------------------------------------
' Exportiere benannte Range aus Excel (lokal, UNC oder SharePoint)
' Usage:
'   cscript ExportNamedRange.vbs "<PfadOderURL>" "<NameDerRange>" "<ZielCSV>"
'-------------------------------------------------------------

If WScript.Arguments.Count < 3 Then
    WScript.Echo "Error! Bitte Pfad/URL, Named Range und Ziel-CSV angeben."
    WScript.Quit 1
End If

Dim inputPath, rangeName, outputCsv
inputPath  = WScript.Arguments.Item(0)
rangeName  = WScript.Arguments.Item(1)
outputCsv  = WScript.Arguments.Item(2)

' Excel-Instanz starten
Dim excel
Set excel = CreateObject("Excel.Application")
excel.DisplayAlerts = False
excel.AskToUpdateLinks = False
excel.EnableEvents = False

' FSO für lokale Dateien
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

'-------------------------------------------------------------
' Workbook öffnen (lokal/UNC oder SharePoint-URL)
'-------------------------------------------------------------
Dim workbook
If LCase(Left(inputPath, 4)) = "http" Then
    ' SharePoint-URL
    On Error Resume Next
    ' Parameter: Filename, UpdateLinks=0, ReadOnly=True
    Set workbook = excel.Workbooks.Open(inputPath, 0, True)
    If Err.Number <> 0 Then
        WScript.Echo "Fehler beim Öffnen der SharePoint-Datei: " & Err.Description
        excel.Quit
        WScript.Quit 1
    End If
    On Error GoTo 0
Else
    ' Lokale oder UNC-Datei
    If Not fso.FileExists(inputPath) Then
        WScript.Echo "Fehler! Datei '" & inputPath & "' nicht gefunden."
        excel.Quit
        WScript.Quit 1
    End If
    Set workbook = excel.Workbooks.Open(inputPath)
End If

'-------------------------------------------------------------
' Named Range auslesen
'-------------------------------------------------------------
Dim namedRange
On Error Resume Next
Set namedRange = workbook.Names(rangeName).RefersToRange
If Err.Number <> 0 Then
    WScript.Echo "Fehler! Named Range '" & rangeName & "' nicht gefunden."
    workbook.Close False
    excel.Quit
    WScript.Quit 1
End If
On Error GoTo 0

Dim rowCount, colCount
rowCount = namedRange.Rows.Count
colCount = namedRange.Columns.Count

' Werte als Array lesen (Performance!)
Dim rawValues
rawValues = namedRange.Value

'-------------------------------------------------------------
' CSV schreiben
'-------------------------------------------------------------
' True = überschreiben erlaubt, True = Unicode (für Umlaute)
Dim csvFile
Set csvFile = fso.CreateTextFile(outputCsv, True, True)

Dim r, c
For r = 1 To rowCount
    Dim line
    line = ""
    For c = 1 To colCount
        Dim val
        val = rawValues(r, c)
        ' Anführungszeichen maskieren
        val = Replace(val, """", """""""")
        ' In Anführungszeichen setzen
        val = """" & val & """"
        ' Trennzeichen (hier Komma)
        If c > 1 Then line = line & ","
        line = line & val
    Next
    csvFile.WriteLine line
Next

csvFile.Close

'-------------------------------------------------------------
' Aufräumen
'-------------------------------------------------------------
workbook.Close False
excel.Quit

Set csvFile   = Nothing
Set namedRange= Nothing
Set workbook  = Nothing
Set fso       = Nothing
Set excel     = Nothing

'WScript.Echo "Export abgeschlossen: " & outputCsv
