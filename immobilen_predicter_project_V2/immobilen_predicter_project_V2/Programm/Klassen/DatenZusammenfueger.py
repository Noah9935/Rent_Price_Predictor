import pandas as pd
import os
import re
import difflib

RED = '\033[31m'
BLUE = '\033[34m'
YELLOW = '\033[33m'
GREEN = '\033[32m'
RESET = '\033[0m'

class ErstelleImmobilienDaten:
    def __init__(self, path, referenzdatei, output_datei):
        self.path = path
        self.referenzdatei = referenzdatei
        self.referenz_df = pd.DataFrame()

        for datei in referenzdatei:
            dateipfad = os.path.join(path, datei)
            df = pd.read_excel(dateipfad)
            self.referenz_df = pd.concat([self.referenz_df, df], ignore_index=True)
        
        self.referenz_spalten = list(self.referenz_df.columns)
        
        if output_datei:
            output_pfad = os.path.join(self.path, output_datei)
            if os.path.exists(output_pfad):
                self.df = pd.read_excel(output_pfad)
            else:
                self.df = pd.DataFrame()
        else:
            self.df = pd.DataFrame()

    def vereinheitliche_spalten(self, df):

        balkon = df['Balkon'] if 'Balkon' in df.columns else pd.Series([False] * len(df))
        terrasse = df['Terrasse'] if 'Terrasse' in df.columns else pd.Series([False] * len(df))

        balkon_terrasse = (balkon.fillna(False).astype(bool) | terrasse.fillna(False).astype(bool))

        if 'Terrasse' in df.columns:
            insert_pos = df.columns.get_loc('Terrasse') + 1
        elif 'Balkon' in df.columns:
            insert_pos = df.columns.get_loc('Balkon') + 1
        else:
            insert_pos = len(df.columns)

        self.df.insert(insert_pos, 'Balkon/Terrasse', balkon_terrasse)

    
    def start_fusion_dateien(self, datei1, datei2, output_datei):

        print(f"{GREEN}Starte{RESET} Vereinheitlichung")

        df1 = pd.read_excel(os.path.join(self.path, datei1))
        df2 = pd.read_excel(os.path.join(self.path, datei2))

        df1 = self.vereinheitliche_spalten(df1)
        df2 = self.vereinheitliche_spalten(df2)

        gemeinsame_spalten = list(set(df1.columns).intersection(set(df2.columns)))
        df1 = df1[gemeinsame_spalten]
        df2 = df2[gemeinsame_spalten]

        self.df = pd.concat([df1, df2], ignore_index=True)

        self.df.to_excel(os.path.join(self.path, output_datei), index=False)
        print(f"Wohnungsdaten wurde zu output {BLUE}hinzugefügt{RESET}")
    
    def start_datei(self, datei1, output_datei):
                
        print(f"{GREEN}Starte{RESET} übernahme")

        self.df = pd.read_excel(os.path.join(self.path, datei1))

        self.df.to_excel(os.path.join(self.path, output_datei), index=False)

        print(f"Wohnungsdaten wurde zu output {BLUE}hinzugefügt{RESET}")

    def bereinige_stadtname(self, name):
        
        if pd.isna(name):
            return ""
        name = str(name)

        name = re.sub(r'\s*\(.*?\)', '', name)  # Entfernt Text in Klammern
        name = re.sub(r',.*', '', name)  # Entfernt alles nach Komma
        name = re.sub(r'[^a-zA-ZäöüÄÖÜß\s-]', '', name)  # Entfernt Sonderzeichen

        name = name.replace('ä', 'ae').replace('ö', 'oe').replace('ü', 'ue')
        name = name.replace('Ä', 'Ae').replace('Ö', 'Oe').replace('Ü', 'Ue')
        name = name.strip().lower() 
        return name

    def add_arbeitslosenquote(self, output_datei, arbeitslosen_datei, Bundesland_Daten, arbeitslosen_sheet='SozialerUmkreis'):
        # Verbesserungsfähig, Stadt namen richtig ausgleichen

        print(f"{GREEN}Starte{RESET} hinzufügen von der Arbeitslosenquote")

        arbeitslosen_pfad = os.path.join(self.path, arbeitslosen_datei)
        arbeitslosen_bundesland_pfad = os.path.join(self.path, Bundesland_Daten)
        
        quote_df = pd.read_excel(arbeitslosen_pfad, sheet_name=arbeitslosen_sheet)
        bundesland_df = pd.read_excel(arbeitslosen_bundesland_pfad, sheet_name='Tabelle1')
        bundesland_df['Bundesland'] = bundesland_df['Bundesland'].str.strip().str.lower().str.replace('ü', 'ue')

        self.df['Stadt_clean'] = self.df['Stadt'].apply(self.bereinige_stadtname)
        quote_df['Stadt_clean'] = quote_df['Stadt'].apply(self.bereinige_stadtname)
        
        bundesland_mapping = dict(zip(
            bundesland_df['Bundesland'],
            bundesland_df['Arbeitslosenquote in %']
        ))
        quote_map = dict(zip(quote_df['Stadt_clean'], quote_df['Arbeitslosenquote in %']))

        arbeitslosen_liste = []

        for i, stadt in self.df['Stadt_clean'].items():
            if stadt in quote_map:
                arbeitslosen_liste.append(quote_map[stadt])
            else:
                bundesland = self.df['Bundesland'][i]
                if bundesland in bundesland_mapping:
                    arbeitslosen_liste.append(bundesland_mapping[bundesland])
                else:
                    match = difflib.get_close_matches(stadt, quote_map.keys(), n=1, cutoff=0.8)
                    if match:
                        arbeitslosen_liste.append(quote_map[match[0]])
                    else:
                        arbeitslosen_liste.append(None)

        self.df['Arbeitslosenquote in %'] = arbeitslosen_liste

        self.df.drop(columns=['Stadt_clean'], inplace=True)

        self.df.to_excel(os.path.join(self.path, output_datei), index=False)

        print(f"Arbeitslosenquote {BLUE}hinzugefügt{RESET}")

    def add_kaufkraftindex(self, output_datei, kaufkraft_datei):

        print(f"{GREEN}Starte{RESET} hinzufügen von der Kaufkraft (pro Einwohner in €)")

        kaufkraft_pfad = os.path.join(self.path, kaufkraft_datei)
        kaufkraft_df = pd.read_excel(kaufkraft_pfad)


        kaufkraft_df['Bundesland'] = kaufkraft_df['Bundesland'].str.strip().str.lower().str.replace('ü', 'ue')
        self.df['Bundesland'] = self.df['Bundesland'].str.strip().str.lower()


        kaufkraft_mapping = dict(zip(
            kaufkraft_df['Bundesland'],
            kaufkraft_df['Kaufkraft 2025 pro Einwohner in €']
        ))

        self.df['Kaufkraft p. E. in €'] = self.df['Bundesland'].map(kaufkraft_mapping)

        self.df.to_excel(os.path.join(self.path, output_datei), index=False)

        print(f"Kaufkraft Daten {BLUE}hinzugefügt{RESET}")

    def add_oeffentlicher_verkehr_qualitaet(self, output_datei, oeffis_datei, oeffis_sheet='OeffentlicherUmkreis'):
        print(f"{GREEN}Starte{RESET} hinzufügen von der ÖPNV Qualität (jeweilige Stadt)")

        oeffis_pfad = os.path.join(self.path, oeffis_datei)
        
        gemeinde_df = pd.read_excel(oeffis_pfad, sheet_name=oeffis_sheet)

        self.df['Stadt_clean'] = self.df['Stadt'].apply(self.bereinige_stadtname)
        gemeinde_df['Gemeindeverbandsname'] = gemeinde_df['Gemeindeverbandsname'].apply(self.bereinige_stadtname)

        opnv_mapping  = dict(zip(gemeinde_df['Gemeindeverbandsname'], gemeinde_df['ÖPNV (Qualität)']))

        opnv_qualitaet_liste = []

        for stadt in self.df['Stadt_clean']:
            if stadt in opnv_mapping:
                opnv_qualitaet_liste.append(opnv_mapping[stadt])
            else:
                match = difflib.get_close_matches(stadt, opnv_mapping.keys(), n=4, cutoff=0.8)
                if match:
                    opnv_qualitaet_liste.append(opnv_mapping[match[0]])
                else:
                    opnv_qualitaet_liste.append(None)

        self.df['ÖPNV qualität'] = opnv_qualitaet_liste

        self.df.drop(columns=['Stadt_clean'], inplace=True)

        self.df.to_excel(os.path.join(self.path, output_datei), index=False)

        print(f"ÖPNV {BLUE}hinzugefügt{RESET}")

    def add_wohnungsleerstand(self, output_datei, wohnungsleerstand_datei, wohnungsleerstand_sheet='SozialerUmkreis'):

        print(f"{GREEN}Starte{RESET} hinzufügen vom Wohnungsleerstand (jeweilige Stadt)")

        wohnungsleerstands_pfad = os.path.join(self.path, wohnungsleerstand_datei)
        
        wohnungsleerstand_df = pd.read_excel(wohnungsleerstands_pfad, sheet_name=wohnungsleerstand_sheet)

        self.df['Stadt_clean'] = self.df['Stadt'].apply(self.bereinige_stadtname)
        wohnungsleerstand_df['Stadt_clean'] = wohnungsleerstand_df['Stadt'].apply(self.bereinige_stadtname)

        quote_map = dict(zip(wohnungsleerstand_df['Stadt_clean'], wohnungsleerstand_df['Wohnungsleerstand']))

        wohnungsleerstand_df_liste = []

        for stadt in self.df['Stadt_clean']:
            if stadt in quote_map:
                wohnungsleerstand_df_liste.append(quote_map[stadt])
            else:
                match = difflib.get_close_matches(stadt, quote_map.keys(), n=4, cutoff=0.8)
                if match:
                    wohnungsleerstand_df_liste.append(quote_map[match[0]])
                else:
                    wohnungsleerstand_df_liste.append(None)

        self.df['Wohnungsleerstand'] = wohnungsleerstand_df_liste

        self.df.drop(columns=['Stadt_clean'], inplace=True)

        self.df.to_excel(os.path.join(self.path, output_datei), index=False)

        print(f"Wohnungsleerstands Anteil {BLUE}hinzugefügt{RESET}")


print(f"-----------------------------------------------{RED}Starte Programm{RESET}---------------------------------------------------------")

path = 'C:/temp/____Noah_Ordner/1py_Programme/Programme/Machine_Learning/Code/immobilen_predicter_project_V2/Programm/Tabellen'
output_datei = 'V3_Immobilien_Daten_2025_05.xlsx'

referenzdateien = ['immonet_daten_2025_05.xlsx', 'immoscout_daten_2025_05.xlsx']
immobilien_daten = ErstelleImmobilienDaten(path, referenzdateien, output_datei)

Gesellschaft_datei = 'Gesellschafts_Daten_2024.xlsx'
Bundesland_Daten = 'Bundesland_Daten_2025.xlsx'

sozial_sheet = 'SozialerUmkreis'
opnv_sheet = 'OeffentlicherUmkreis'
gesellschaft_sheet = 'GesellschaftlicherUmkreis'
#immobilien_daten.fusion(referenzdateien[0], referenzdateien[1], output_datei)
immobilien_daten.start_datei(referenzdateien[1], output_datei)

immobilien_daten.add_arbeitslosenquote(output_datei, Gesellschaft_datei, Bundesland_Daten, arbeitslosen_sheet=sozial_sheet)
immobilien_daten.add_kaufkraftindex(output_datei, Bundesland_Daten)
immobilien_daten.add_oeffentlicher_verkehr_qualitaet(output_datei, Gesellschaft_datei, oeffis_sheet=opnv_sheet)
immobilien_daten.add_wohnungsleerstand(output_datei, Gesellschaft_datei, wohnungsleerstand_sheet=sozial_sheet)

speicher_path = os.path.join(path, output_datei)
print(f"Datei {output_datei} gespeichert in: {speicher_path}")

print(f"--------------------------------------------{RED}Programm abgeschlossen{RESET}-----------------------------------------------------")
