# Anleitung: Miniconda-Umgebung für das Zeitplanungs-Tool

1. Öffne die Anaconda Prompt oder dein Terminal.
2. Erstelle eine neue Umgebung:

   conda create -n zeitplan python=3.11

3. Aktiviere die Umgebung:

   conda activate zeitplan


4. Installiere die benötigten Pakete:

   conda install pyqt pyqtgraph pandas openpyxl
   pip install PyQt5 qtawesome

5. Starte das Programm später mit:

   python main.py

Fertig!