# ICS_Export

Beschreibung

Dieses Script erstellt eine.ics-Datei aus einem aus ATOSS exportierten Mitarbeiterdienstplan. Die .ics-Datei kann dann in einem Kalender-Programm wie Google Kalender oder Apple Kalender importiert werden.

Funktionalität

Liest den Mitarbeiterdienstplan aus einem Excel-Worksheet, erstellt eine.ics-Datei mit den Terminen und Abwesenheiten, setzt die Terminnamen, -daten und -zeiten korrekt, behandelt Abwesenheiten als ganztägige Termine.

Anforderungen

Excel 2013 oder höher, Libre Office, ATOSS-Export-Datei im Excel-Format, Installation

Öffne das Script in Excel-VBA-Editor, ändere die Variable icsFile auf den gewünschten Dateipfad und -namen führe das Script aus, indem du auf "ExportToIcs" klickst.

Lizenz

Dieses Script ist unter der MIT-Lizenz veröffentlicht. Siehe LICENSE.txt für Details.

Changelog

1.0.0: Erstes Release
1.1.0: Behandlung von Abwesenheiten als ganztägige Termine hinzugefügt
