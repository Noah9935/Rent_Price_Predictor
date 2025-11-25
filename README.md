# Immobilien Price Predictor

Dieses Projekt ist ein Python-Tool zur Analyse von Mietpreisen.
Es sammelt Daten von Immobilienportalen, speichert diese in Excel-Dateien und
berechnet auf Basis der gesammelten Daten eine Abschätzung des Mietpreises.
Über eine grafische Oberfläche (tkinter) kann der Nutzer
Auswertungen anstoßen.

## Funktionen

- **Web-Scraping**
    - Module `immonet_scraper.py` und `immoscout24_scraper.py` laden Angebote
      von den jeweiligen Immobilienportalen (z.B. Titel, Lage, Miete, Fläche).
    - Die Rohdaten werden als `.xlsx`-Dateien im Ordner `data/raw/` abgelegt.

- **Datenzusammenführung**
    - `datenzusammenfueger.py` liest die verschiedenen Excel-Dateien ein,
      bereinigt sie (z.B. doppelte Einträge, fehlende Werte) und führt sie zu
      einer gemeinsamen Datentabelle zusammen.
    - Basisbibliothek hierfür ist `pandas`.

- **Preisvorhersage**
    - `predictor.py` nutzt die aufbereiteten Daten, um einfache Modelle für
      die Mietpreisabschätzung zu bauen (z.B. auf Basis von Quadratmeterpreis,
      Lage, Zimmeranzahl).
    - Die vorhergesagten Preise können später in der GUI angezeigt oder in
      Tabellen exportiert werden.

- **Grafische Benutzeroberfläche**
    - `benutzeroberflaeche.py` stellt eine tkinter-GUI bereit, über die der
      Nutzer:
        - den Browser/Scraper starten,
        - aktuelle Daten erfassen,
        - Daten zusammenführen,
        - Vorhersagen ausführen
        - und das Programm beenden kann.
