import tkinter as tk
from tkinter import ttk, messagebox
import re
from Immoscout24Scraper import ImmoscoutScraper
from Predictor import Predictor
import pandas as pd

class GUI_Immobilien:
    def __init__(self, url, start_page=1):
        self.url = url
        self.start_page = start_page

        self.scraper = ImmoscoutScraper(self.url, self.start_page)

        self.root = tk.Tk()
        self.root.title("ImmoScout Data Collector")

        ttk.Label(self.root, text="Browser öffnen").pack(pady=5)
        ttk.Button(self.root, text="Browser starten", command=self.start_browser).pack(pady=5)

        ttk.Label(self.root, text="Immoscout24 Expose - Link eingeben:").pack(pady=10)
        self.link_entry = ttk.Entry(self.root, width=50)
        self.link_entry.pack(padx=10, pady=5)

        ttk.Button(self.root, text="Daten erfassen", command=self.extract_current_offer).pack(pady=5)

        ttk.Button(self.root, text="Beenden", command=self.schliessen).pack(pady=10)

        self.root.protocol("WM_DELETE_WINDOW", self.schliessen)
        self.root.mainloop()

    def start_browser(self):

        try:
            self.scraper.start_browser()
            self.scraper.page.goto(self.url)
        except Exception as e:
            messagebox.showerror("Fehler", f"Browser konnte nicht gestartet werden:\n{e}")

    
    def extract_current_offer(self):
        try:

            link = self.link_entry.get()
            print(link)
            match = re.search(r"/expose/(\d+)", link)

            if not match:
                messagebox.showwarning("Kein Exposé", "Bitte in der Browser‑Ansicht ein Exposé öffnen.")
                return
            exp_id = match.group(1)

            data = self.scraper.get_inhalte_from_offer(exp_id)
            (wohnflaeche, zimmer, baujahr, zustand,
             kueche, terrasse, aufzug, garten, keller,
             etage, kalt, warm, energie, bundesland,
             stadt, stadtteil, plz) = data
            
            data_dict = {
                "Wohnfläche": [wohnflaeche],
                "Zimmeranzahl": [zimmer],
                "Baujahr": [baujahr],
                "Zustand": [zustand],
                "Küche": [kueche],
                "Balkon / Terrasse": [terrasse],
                "Aufzug": [aufzug],
                "Garten": [garten],
                "Keller": [keller],
                "Etage": [etage],
                "Kaltmiete": [kalt],
                "Warmmiete": [warm],
                "Energieeffizienzklasse": [energie],
                "Bundesland": [bundesland],
                "Stadt": [stadt],
                "Stadtteil": [stadtteil],
                "PLZ": [plz]
            }
            offer_df = pd.DataFrame(data_dict)
        
            summary = (
                f"Wohnfläche: {wohnflaeche}\n"
                f"Zimmer: {zimmer}\n"
                f"Baujahr: {baujahr}\n"
                f"Zustand: {zustand}\n"
                f"Küche: {kueche}\n"
                f"Balkon / Terrasse: {terrasse}\n"
                f"Aufzug: {aufzug}\n"
                f"Garten: {garten}\n"
                f"Keller: {keller}\n"
                f"Etage: {etage}\n"
                f"Kaltmiete: {kalt}\n"
                f"Warmmiete: {warm}\n"
                f"Energieeffizienzklasse: {energie}\n"
                f"Bundesland: {bundesland}\n"
                f"Stadt: {stadt}\n"
                f"Stadtteil: {stadtteil}\n"
                f"PLZ: {plz}"
            )
            messagebox.showinfo("Daten des Exposés", summary)

            self.vorhersage(offer_df, 3, zielspalte='Kaltmiete', path='C:/temp/____Noah_Ordner/1py_Programme/Programme/Machine_Learning/Code/immobilen_predicter_project_V2/Programm/Tabellen/')

        except Exception as e:
            messagebox.showerror("Extraktionsfehler", str(e))

    def vorhersage(self, immobilien_daten, datensatz_versionsnummer, zielspalte='Kaltmiete', path='C:/temp/...'):
        try:
            datensatz_name = f'V{datensatz_versionsnummer}_Immobilien_Daten_2025_05.xlsx'
            analyse = Predictor(datensatz_name, path)

            # Werte standardisieren und Feature Engineering
            df = analyse.standartisiere_werte(analyse.df)
            df = analyse.feature_engineering(df)

            # Korrelationen entfernen und Modell trainieren
            df = analyse.entferne_korrelationen(df, zielspalte=zielspalte, mind_anforderung=0.05)
            modell = analyse.trainiere_modell(df, zielspalte=zielspalte, test_daten_anteil=0.2)

            # Testdaten und erforderliche Dateien
            Gesellschaft_datei = 'Gesellschafts_Daten_2024.xlsx'
            Bundesland_Daten = 'Bundesland_Daten_2025.xlsx'
            sozial_sheet = 'SozialerUmkreis'
            opnv_sheet = 'OeffentlicherUmkreis'
            gesellschaft_sheet = 'GesellschaftlicherUmkreis'
            output_datei = 'temp_data.xlsx'

            sheets = {
                "sozial": sozial_sheet,
                "opnv": opnv_sheet,
                "gesellschaft": gesellschaft_sheet
            }

            immobilie_bearbeitet = analyse.bearbeite_testdaten(immobilien_daten, Gesellschaft_datei, Bundesland_Daten, sheets, zielspalte)
            #immobilie_standartisiert = analyse.bekomme_daten_standartisieren(immobilie_bearbeitet)
            immobilie_standartisiert = analyse.standartisiere_werte(immobilie_bearbeitet, True)
            immobilie_ende = analyse.feature_engineering(immobilie_standartisiert, True)
            immobilie = immobilie_ende.drop(columns=[col for col in analyse.unwichtige_columns if col in immobilie_ende.columns], errors='ignore')

            
            # Vorhersage
            vorhersage = modell.predict(immobilie)
            print(f"Vorhergesagte Kaltmiete: {vorhersage[0]:.2f} €")
            messagebox.showinfo("Vorhersage", f"Vorhergesagte Kaltmiete: {vorhersage[0]:.2f} €")
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler bei der Vorhersage: {e}")

    def schliessen(self):
        try:
            self.scraper.stop_browser()
        except:
            pass
        self.root.destroy()


if __name__ == "__main__":
    start_page = 1
    url = f"https://www.immobilienscout24.de/Suche/de/wohnung-mieten?pagenumber={start_page}"
    GUI_Immobilien(url, start_page)
