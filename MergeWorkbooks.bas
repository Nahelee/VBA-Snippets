Sub SammleDateien()

    Dim pfad As String
    Dim datei As String
    Dim wbZiel As Workbook
    Dim wsZiel As Worksheet
    Dim wbQuelle As Workbook
    Dim wsQuelle As Worksheet
    Dim lastRowQuelle As Long
    Dim lastColQuelle As Long
    Dim headerRow As Long
    Dim startRow As Long
    Dim newLastRow As Long
    Dim antw As VbMsgBoxResult
    
    ' Kopfzeilen-Abfrage beim Start
    antw = MsgBox("Befindet sich die Kopfzeile in der ersten Zeile?", vbYesNo + vbQuestion, "Kopfzeile wählen")
    If antw = vbYes Then
        headerRow = 1
    Else
        headerRow = Application.InputBox("Bitte geben Sie die Zeilennummer der Kopfzeile ein:", "Kopfzeile", 1, Type:=1)
        If headerRow < 1 Then
            MsgBox "Abbruch – ungültige Eingabe.", vbExclamation
            Exit Sub
        End If
    End If
    
    Set wbZiel = ThisWorkbook
    Set wsZiel = wbZiel.Sheets(1)
    
    ' Zielblatt komplett leeren (alle Inhalte)
    wsZiel.Cells.Clear
    
    pfad = wbZiel.Path & "\"
    datei = Dir(pfad & "*.xls*")
    
    Application.ScreenUpdating = False
    
    Do While datei <> ""
        If datei <> wbZiel.Name Then
            
            Set wbQuelle = Workbooks.Open(pfad & datei)
            Set wsQuelle = wbQuelle.Sheets(1)
            
            lastRowQuelle = wsQuelle.Cells(wsQuelle.Rows.Count, 1).End(xlUp).Row
            lastColQuelle = wsQuelle.Cells(headerRow, wsQuelle.Columns.Count).End(xlToLeft).Column
            
            ' Wenn Ziel komplett leer, Kopfzeile übernehmen
            If wsZiel.Cells(1, 1).Value = "" Then
                ' Kopfzeile aus der Quelldatei übernehmen
                wsQuelle.Range(wsQuelle.Cells(headerRow, 1), wsQuelle.Cells(headerRow, lastColQuelle)).Copy wsZiel.Cells(1, 1)
                
                ' Überschrift für Dateiname in neuer Spalte setzen
                wsZiel.Cells(1, lastColQuelle + 1).Value = "Dateiname"
                
                ' Autofilter auf Kopfzeile setzen
                wsZiel.Range(wsZiel.Cells(1, 1), wsZiel.Cells(1, lastColQuelle + 1)).AutoFilter
                
                startRow = headerRow + 1 ' Daten ab nächste Zeile nach Kopfzeile
                wsZiel.Calculate
            Else
                startRow = headerRow + 1 ' Für weitere Dateien nur Daten
            End If
            
            If startRow <= lastRowQuelle Then
                ' Letzte Zeile in Ziel ermitteln
                newLastRow = wsZiel.Cells(wsZiel.Rows.Count, 1).End(xlUp).Row
                
                ' Daten kopieren unter bestehende
                wsQuelle.Range(wsQuelle.Cells(startRow, 1), wsQuelle.Cells(lastRowQuelle, lastColQuelle)).Copy _
                    wsZiel.Cells(newLastRow + 1, 1)
                
                ' Dateiname in neue Spalte schreiben
                newLastRow = wsZiel.Cells(wsZiel.Rows.Count, 1).End(xlUp).Row
                wsZiel.Range(wsZiel.Cells(newLastRow - (lastRowQuelle - startRow), lastColQuelle + 1), _
                             wsZiel.Cells(newLastRow, lastColQuelle + 1)).Value = datei
            End If
            
            wbQuelle.Close SaveChanges:=False
        End If
        datei = Dir
    Loop
    
    Application.ScreenUpdating = True
    
    ' Anzahl der befüllten Zeilen (inkl. Kopfzeile)
    Dim anzahlZeilen As Long
    anzahlZeilen = wsZiel.Cells(wsZiel.Rows.Count, 1).End(xlUp).Row
    
    ' Autofilter sicherstellen (optional)
    If Not wsZiel.AutoFilterMode Then
        lastColQuelle = wsZiel.Cells(1, wsZiel.Columns.Count).End(xlToLeft).Column
        wsZiel.Range(wsZiel.Cells(1, 1), wsZiel.Cells(1, lastColQuelle)).AutoFilter
    End If
    
    ' Kopfzeile fixieren
wsZiel.Activate
wsZiel.Rows(2).Select
ActiveWindow.FreezePanes = True
    
    MsgBox "Fertig! Alle Dateien wurden zusammengeführt (Kopfzeile in Zeile " & headerRow & ")." & vbCrLf & _
           "Es sind insgesamt " & anzahlZeilen & " Zeilen befüllt.", vbInformation

End Sub
