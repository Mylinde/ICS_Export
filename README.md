# ICS_Export

Beschreibung

Dieses Script erstellt eine .ics-Datei aus einem aus ATOSS Staff Efficiency Suite exportierten Mitarbeiterdienstplan. Die .ics-Datei kann dann in einem Kalender-Programm wie Google Kalender oder Apple Kalender importiert werden.

Funktionalität

Liest die zur Terminerstellung relevanten Daten aus dem Excel-Worksheet und erstellt eine .ics-Datei mit den Arbeitszeiten und Abwesenheiten.

Anforderungen

Excel 2013 und höher oder Libre Office <br> ATOSS-Export-Datei im Excel-Format

Öffne das Script in Excel-VBA-Editor, ändere die Variable icsFile auf den gewünschten Dateipfad und -namen, führe das Script aus indem du auf "ExportToIcs" klickst.

Lizenz

Dieses Script ist unter der MIT-Lizenz veröffentlicht. Siehe LICENSE.txt für Details.

Changelog

1.0.0: Erstes Release <br>
1.1.0: Behandlung von Abwesenheiten als ganztägige Termine hinzugefügt <br>
1.2.0: Dateiname beinhaltet Mitarbeitername und Monat
