# Anleitung zur Verwendung der `add_to_registry.bat` Datei und Ausführung der Anwendung

Diese Anleitung beschreibt, wie Sie die `add_to_registry.bat` Datei verwenden, um die Anwendung im Kontextmenü für Excel-Dateien unter Windows hinzuzufügen, und wie Sie die Anwendung anschließend nutzen können.

## Voraussetzungen

- Stellen Sie sicher, dass die ausführbare Datei der Anwendung (`YourApp.exe`) auf Ihrem Computer vorhanden ist.
- Die `add_to_registry.bat` Datei, die zum Hinzufügen der Anwendung zum Kontextmenü verwendet wird.

## Schritte zum Hinzufügen der Anwendung zum Kontextmenü

1. **Öffnen Sie die `add_to_registry.bat` Datei in einem Texteditor**: 
   - Fügen Sie den vollständigen Pfad zur ausführbaren Datei Ihrer Anwendung in der `add_to_registry.bat` Datei ein. Ersetzen Sie `YourAppPath` mit dem tatsächlichen Pfad Ihrer `.exe` Datei. Beachten Sie, dass Sie doppelte Backslashes (`\\`) im Pfad verwenden müssen.

2. **Führen Sie die `add_to_registry.bat` Datei aus**:
   - Doppelklicken Sie auf die `add_to_registry.bat` Datei, um die Änderungen in der Windows-Registrierung vorzunehmen. Bestätigen Sie alle Sicherheitsabfragen, um die Änderungen zu erlauben.

## Verwendung der Anwendung

Nachdem Sie die Anwendung zum Kontextmenü hinzugefügt haben, können Sie sie wie folgt verwenden:

1. **Rechtsklicken Sie auf eine Excel-Datei** (`*.xlsx`), für die Sie die Anwendung verwenden möchten.
2. Wählen Sie **"Process with YourApp"** (oder den von Ihnen gewählten Menütext) aus dem Kontextmenü.
3. Die Anwendung wird gestartet und verarbeitet die ausgewählte Excel-Datei entsprechend Ihrer Logik.

## Entfernen der Anwendung aus dem Kontextmenü

Um die Anwendung aus dem Kontextmenü zu entfernen, müssen Sie die durch die `.reg` Datei vorgenommenen Änderungen rückgängig machen. Dies kann durch eine weitere `.reg` Datei erfolgen, die die entsprechenden Einträge entfernt, oder manuell über den Registrierungseditor.

## Wichtig

- Die Änderung der Windows-Registrierung kann das System beeinflussen. Führen Sie diese Änderungen nur durch, wenn Sie mit den Prozessen vertraut sind.
- Erstellen Sie vor Änderungen an der Registrierung immer eine Sicherung.
