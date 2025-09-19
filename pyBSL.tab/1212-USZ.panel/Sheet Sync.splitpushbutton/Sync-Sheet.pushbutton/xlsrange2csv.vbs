'-------------------------------------------------------------
' Exportiere benannte Range aus Excel (lokal, UNC oder SharePoint) mit Logging
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

' Log-Datei
Dim fso, logPath, log
Set fso = CreateObject("Scripting.FileSystemObject")
logPath = outputCsv & ".log.txt"
Set log = fso.CreateTextFile(logPath, True, True) ' Unicode
log.WriteLine Now & "  START  input='" & inputPath & "'  range='" & rangeName & "'  out='" & outputCsv & "'"

' Excel-Instanz starten
Dim excel
Set excel = CreateObject("Excel.Application")
excel.DisplayAlerts = False
excel.AskToUpdateLinks = False
excel.EnableEvents = False

'-------------------------------------------------------------
' Workbook öffnen (lokal/UNC oder SharePoint-URL)
'-------------------------------------------------------------
Dim workbook
If LCase(Left(inputPath, 4)) = "http" Then
    On Error Resume Next
    Set workbook = excel.Workbooks.Open(inputPath, 0, True)
    If Err.Number <> 0 Then
        log.WriteLine Now & "  ERROR  Beim Öffnen (SharePoint): " & Err.Description
        WScript.Echo "Fehler beim Öffnen der SharePoint-Datei: " & Err.Description
        On Error GoTo 0
        excel.Quit
        log.Close
        WScript.Quit 1
    End If
    On Error GoTo 0
Else
    If Not fso.FileExists(inputPath) Then
        log.WriteLine Now & "  ERROR  Datei nicht gefunden: " & inputPath
        WScript.Echo "Fehler! Datei '" & inputPath & "' nicht gefunden."
        excel.Quit
        log.Close
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
If Err.Number <> 0 Or namedRange Is Nothing Then
    log.WriteLine Now & "  ERROR  Named Range nicht gefunden: " & rangeName & "  Err=" & Err.Description
    WScript.Echo "Fehler! Named Range '" & rangeName & "' nicht gefunden."
    On Error GoTo 0
    workbook.Close False
    excel.Quit
    log.Close
    WScript.Quit 1
End If
On Error GoTo 0

Dim rowCount, colCount
rowCount = namedRange.Rows.Count
colCount = namedRange.Columns.Count

' Werte als Array lesen
Dim rawValues
rawValues = namedRange.Value

'-------------------------------------------------------------
' CSV schreiben
'-------------------------------------------------------------
Dim csvFile
Set csvFile = fso.CreateTextFile(outputCsv, True, True) ' Unicode (UTF-16 LE)

Dim r, c, line, val
For r = 1 To rowCount
    line = ""
    For c = 1 To colCount
        ' Achtung: Bei 1x1-Range liefert Excel einen Skalaren statt Array
        If rowCount = 1 And colCount = 1 Then
            val = rawValues
        Else
            val = rawValues(r, c)
        End If

        Dim cellText
        cellText = ToCsvText(val, r, c, log) ' sichere Konvertierung + Logging bei Problemen

        If c > 1 Then line = line & ","
        line = line & cellText
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
Set excel     = Nothing

log.WriteLine Now & "  DONE"
log.Close
Set log = Nothing
Set fso = Nothing

'WScript.Echo "Export abgeschlossen: " & outputCsv


'=============================================================
' Hilfsfunktionen
'=============================================================
Function ToCsvText(val, r, c, logger)
    On Error Resume Next

    Dim t, v
    t = TypeName(val)

    ' Excel-Fehlerwerte, Null/Empty und Arrays behandeln
    If IsNull(val) Or IsEmpty(val) Then
        v = ""
    ElseIf IsArray(val) Then
        logger.WriteLine Now & "  WARN  Arraywert in Zelle r=" & r & " c=" & c & " (Type=" & t & "). Leerer String verwendet."
        v = ""
    ElseIf VarType(val) = vbError Then
        ' Excel-Fehler (#N/A, #DIV/0!, …)
        logger.WriteLine Now & "  WARN  Excel-Fehlerwert in r=" & r & " c=" & c & " (Type=" & t & "). '#N/A' geschrieben."
        v = "#N/A"
    Else
        ' Sichere Konvertierung
        v = CStr(val)
        If Err.Number <> 0 Then
            logger.WriteLine Now & "  ERROR  CStr fehlgeschlagen in r=" & r & " c=" & c & " (Type=" & t & "): " & Err.Description
            Err.Clear
            v = ""
        End If
    End If

    ' Quotes maskieren
    v = Replace(v, """", """""""")
    If Err.Number <> 0 Then
        logger.WriteLine Now & "  ERROR  Replace fehlgeschlagen in r=" & r & " c=" & c & " (Type=" & t & "): " & Err.Description
        Err.Clear
        v = ""
    End If

    ToCsvText = """" & v & """"
    On Error GoTo 0
End Function
