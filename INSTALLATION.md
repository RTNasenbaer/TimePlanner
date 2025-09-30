# Anleitung: Miniconda-Umgebung für das Zeitplanungs-Tool

## Automatische Installation (Empfohlen)

### Windows
1. Doppelklicke auf `setup.bat` oder führe es in der Eingabeaufforderung aus:
   ```
   setup.bat
   ```

### Linux/macOS
1. Mache das Script ausführbar und führe es aus:
   ```
   chmod +x setup.sh
   ./setup.sh
   ```

Das Script wird automatisch:
- Eine neue conda-Umgebung "zeitplan" mit Python 3.11 erstellen
- Alle benötigten Pakete installieren
- Die Installation überprüfen

## Manuelle Installation

Falls du die Installation manuell durchführen möchtest:

1. Öffne die Anaconda Prompt oder dein Terminal.
2. Erstelle eine neue Umgebung:
   ```
   conda create -n zeitplan python=3.11
   ```

3. Aktiviere die Umgebung:
   ```
   conda activate zeitplan
   ```

4. Installiere die benötigten Pakete:
   ```
   conda install pyqt pyqtgraph pandas openpyxl
   pip install -r requirements.txt
   ```

## Programm starten

Nach der Installation kannst du das Programm starten mit:
```
conda activate zeitplan
python main.py
```

Fertig!