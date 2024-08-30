Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
'-------------------------------------------------------------------------------
' Module Name: ICS_Export
' Description: Erstellt eine .ics Datei aus einem aus ATOSS exportierten Mitarbeiterdienstplan
' Licensing: This code is released under the MIT License. For more information, see <https://opensource.org/licenses/MIT>.
' Copyright (c) 2024 Mario Herrmann. All rights reserved.
'-------------------------------------------------------------------------------

Sub ExportToIcs()
    
   'Definition der Variablen
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim icsText As String
    Dim TerminName As String
    Dim TerminDatum As Date
    Dim TerminUhrzeit As String
    Dim StartUhrzeit As String
    Dim EndUhrzeit As String
    Dim Abwesenheit As String
    Dim icsFile As String
    
   'Arbeitsblatt setzen
    Set ws = ThisWorkbook.Worksheets("emsche")
    
   'Letzte Zeile ermitteln
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    '.ics-Datei erstellen
    icsFile = "~/PEP.ics"
    
   'Header für.ics-Datei erstellen
    icsText = icsText & "BEGIN:VCALENDAR" & vbCrLf
    icsText = icsText & "VERSION:2.0" & vbCrLf
    icsText = icsText & "CALSCALE:GREGORIAN" & vbCrLf
    
   'Daten durchlaufen
    For i = 9 To lastRow
    
       'Termin-Name
        TerminName = ws.Cells(3, "C").Value
        TerminName = Split(TerminName, " ")(1) & " arbeiten"
        
       'Termin-Datum
        TerminDatum = ws.Cells(i, "A").Value      
             
       'Termin-Uhrzeit
        TerminUhrzeit = ws.Cells(i, "H").Value
        Abwesenheit = ws.Cells(i, "E").Value        
        If Abwesenheit <> "" Then
    		TerminName = "Abwesenheit: " & Abwesenheit
    		TerminUhrzeit = "00:00-23:59"
			StartUhrzeit = Split(TerminUhrzeit, "-")(0)
            EndUhrzeit = Split(TerminUhrzeit, "-")(1)
        ElseIf InStr(TerminUhrzeit, "-") > 0 Then
            StartUhrzeit = Split(TerminUhrzeit, "-")(0)
            EndUhrzeit = Split(TerminUhrzeit, "-")(1) 	
       	Else
            GoTo NextEintrag
        End If
                
        '.ics-Format erstellen
        icsText = icsText & "BEGIN:VEVENT" & vbCrLf
        icsText = icsText & "UID:" & TerminName & vbCrLf
        icsText = icsText & "DTSTART;TZID=Europe/Berlin;TZOFFSETFROM=+0100;TZOFFSETTO=+0200:" & Format(TerminDatum, "yyyymmdd") & "T" & Format(StartUhrzeit, "hhmmss") & vbCrLf
        icsText = icsText & "DTEND;TZID=Europe/Berlin;TZOFFSETFROM=+0100;TZOFFSETTO=+0200:" & Format(TerminDatum, "yyyymmdd") & "T" & Format(EndUhrzeit, "hhmmss") & vbCrLf
        icsText = icsText & "SUMMARY:" & TerminName & vbCrLf
        icsText = icsText & "END:VEVENT" & vbCrLf
        
	NextEintrag:
    	Next i
    
   'Footer für.ics-Datei erstellen
    icsText = icsText & "END:VCALENDAR"
    
    '.ics-Datei speichern
    Open icsFile For Output As #1
    Print #1, icsText
    Close #1
    
    MsgBox "Die.ics-Datei wurde erfolgreich erstellt!"
    
End Sub
