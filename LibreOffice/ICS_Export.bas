option VBASupport 1
'-------------------------------------------------------------------------------
' Module Name: ICS_Export
' Description: Creates an .ics from an employee schedule exported from ATOSS
' Licensing: This code is released under the MIT License. For more information, see <https://opensource.org/licenses/MIT>.
' Copyright (c) 2024 Mario Herrmann. All rights reserved.
'-------------------------------------------------------------------------------

Sub ExportToIcs()
    
   'Definition der Variablen
    Dim ws As Object
    Dim lastRow As Long
    Dim oCell As Object
    Dim i As Long
    Dim icsText As String
    Dim TerminName As String
    Dim TerminUhrzeit As String
    Dim StartUhrzeit As String
    Dim EndUhrzeit As String
    Dim Abwesenheit As String
    Dim icsFile As String
    
   'Arbeitsblatt setzen
    ws = ThisComponent.getSheets().getByIndex(0)
    
   'Letzte Zeile ermitteln     
    aData = ws.getCellRangeByPosition(0, 0, 0, 40).getDataArray()
    lastRow = 0
    For i = UBound(aData) To LBound(aData) Step -1   ' 40 → 0
        If Len(Trim(CStr(aData(i)(0)))) > 0 Then
            lastRow = i + 1   ' 1‑basiert
            Exit For
        End If
    Next i
   
   'Header für.ics-Datei erstellen
    icsText = icsText & "BEGIN:VCALENDAR" & vbCrLf
    icsText = icsText & "VERSION:2.0" & vbCrLf
    icsText = icsText & "CALSCALE:GREGORIAN" & vbCrLf
    
   'Daten durchlaufen
    For i = 8 To lastRow
    
       'Termin-Name
        TerminName = ws.getCellByPosition(2, 2).String
        TerminName = Split(TerminName, ",")(1) & " arbeiten"
        
       'Termin-Datum
        TerminDatum = ws.getCellByPosition(0, i).Value
        TerminDatum = Format(TerminDatum, "yyyymmdd")     
        EndTerminDatum = ws.getCellByPosition(0, i).Value
        EndTerminDatum = Format(EndTerminDatum, "yyyymmdd")
              
       'Termin-Uhrzeit
        TerminUhrzeit = ws.getCellByPosition(7, i).String
        Abwesenheit = ws.getCellByPosition(4, i).String
                
        If Abwesenheit <> "" Then
    		TerminName = "Abwesenheit: " & Abwesenheit
    		EndTerminDatum = ws.getCellByPosition(0, i).Value +1
    		EndTerminDatum = Format(EndTerminDatum, "yyyymmdd")
    		TerminUhrzeit = "-"
			StartUhrzeit = Split(TerminUhrzeit, "-")(0)
            EndUhrzeit = Split(TerminUhrzeit, "-")(1)
        Elseif TerminUhrzeit <> "" Then
            StartUhrzeit = Split(TerminUhrzeit, "-")(0)
            EndUhrzeit = Split(TerminUhrzeit, "-")(1)
            TerminDatum = TerminDatum & "T"           
            EndTerminDatum = EndTerminDatum & "T"
       	Else
            GoTo NextEintrag
        End If
                
        '.ics-Format erstellen
        icsText = icsText & "BEGIN:VEVENT" & vbCrLf
        icsText = icsText & "UID:" & TerminName & vbCrLf
        icsText = icsText & "DTSTART;TZID=Europe/Berlin;TZOFFSETFROM=+0100;TZOFFSETTO=+0200:" & TerminDatum & Format(StartUhrzeit, "hhmmss") & vbCrLf
        icsText = icsText & "DTEND;TZID=Europe/Berlin;TZOFFSETFROM=+0100;TZOFFSETTO=+0200:" & EndTerminDatum & Format(EndUhrzeit, "hhmmss") & vbCrLf
        icsText = icsText & "SUMMARY:" & TerminName & vbCrLf
        icsText = icsText & "END:VEVENT" & vbCrLf
        
	NextEintrag:
    Next i
    
   'Footer für.ics-Datei erstellen
    icsText = icsText & "END:VCALENDAR"
    
    '.ics-Datei erstellen
    icsFile = "~/PEP/PEP_" & Split(ws.getCellByPosition(2, 2).String, ",")(0) & "_" & Format(ws.getCellByPosition(0, 8).Value, "MMMM") & ".ics"
    
    '.ics-Datei speichern
    Open icsFile For Output As #1
    Print #1, icsText
    Close #1
    
    MsgBox "Die .ics-Datei wurde erfolgreich erstellt!"
    
End Sub