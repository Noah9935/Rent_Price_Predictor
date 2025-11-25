import pandas as pd
import numpy as np
import os
import re
import math
import warnings
import seaborn as sns
import matplotlib.pyplot as plt

from sklearn.preprocessing import LabelEncoder, StandardScaler
from sklearn.ensemble import RandomForestRegressor
from sklearn.model_selection import train_test_split
from sklearn.metrics import mean_squared_error, mean_absolute_error

warnings.simplefilter("ignore", category=FutureWarning)

RED = '\033[31m'
BLUE = '\033[34m'
YELLOW = '\033[33m'
GREEN = '\033[32m'
RESET = '\033[0m'

class Predictor:
    def __init__(self, datensatz_name, path):
        self.datensatz_name = datensatz_name
        self.df = pd.read_excel(os.path.join(path, datensatz_name))
        self.unwichtige_columns = []
        self.path = path 
        self.laenge_datensatz = len(self.df)
        print(f"Anzahl der eingelesenen Daten: {BLUE}{self.laenge_datensatz}{RESET}")

    def standartisiere_werte(self, df_mit, skip=False):
        print(f"{YELLOW}Starte{RESET} Standardisierung")
        le = LabelEncoder()
        df = df_mit
        if(not skip):
            df = df_mit.dropna(subset=['Wohnfläche', 'Zimmeranzahl', 'Kaltmiete', 'Warmmiete', 'Bundesland', 'Stadt', 'Stadtteil'])
        entfernte_zeilen = len(df_mit) - len(df)

        df = df.drop(columns=['Garten'], errors='ignore')
        if(not skip):
            df['Kaltmiete'] = (df['Kaltmiete'].astype(str)
                            .str.replace('€', '', regex=False)
                            .str.replace('.', '', regex=False)
                            .str.replace(',', '.', regex=False))
            df['Kaltmiete'] = pd.to_numeric(df['Kaltmiete'], errors='coerce')
            df['Kaltmiete'] = df['Kaltmiete'].replace(['Auf Anfrage', ''], np.nan)
        df['Baujahr'] = df['Baujahr'].replace(['unbekannt', ''], np.nan)
        df['ÖPNV qualität'] = df['ÖPNV qualität'].replace(['-9999', ''], np.nan)
        df['Energieeffizienzklasse'] = df['Energieeffizienzklasse'].replace('', 'unbekannt')
        df['Zustand'] = df['Zustand'].replace('', 'unbekannt')

        fill_median = ['Etage', 'Arbeitslosenquote in %', 'Kaufkraft p. E. in €', 'Wohnungsleerstand', 'ÖPNV qualität']
        for col in fill_median:
            if col in df.columns:
                if(not skip):
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(df[col].median())
                else:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        for col in df.select_dtypes(include=['object']).columns:
            df[col] = df[col].str.lower().str.strip().str.replace(',', '.', regex=False).replace({'wahr': True, 'falsch': False})
            try:
                df[col] = pd.to_numeric(df[col], errors='coerce')
            except:
                df[col] = le.fit_transform(df[col].astype(str))

        print(f"Standardisierung {GREEN}Abgeschlossen{RESET}, {RED}{entfernte_zeilen}{RESET} Zeilen entfernt")
        return df

    def feature_engineering(self, df, skip=False):
        print(f"Füge Features {YELLOW}hinzu{RESET}")
        # if(not skip):
        #df['warm_kalt_ratio'] = df['Warmmiete'] / df['Kaltmiete']
        # else:
        #     df['warm_kalt_ratio'] = 0

        df['verhaeltnis_fläche_zimmer'] = df['Wohnfläche'] / df['Zimmeranzahl']
        
        df['ist_neubau'] = (df['Baujahr'] >= 2020).astype(int)
        df['viele_zimmer'] = (df['Zimmeranzahl'] >= 4).astype(int)

        # Fehlende Ausstattungsspalten durch 0 ersetzen
        ausstattung = ['Küche', 'Balkon / Terrasse', 'Aufzug', 'Keller']
        for col in ausstattung:
            if col not in df.columns:
                df[col] = 0
        df['ausstattung_score'] = df[ausstattung].sum(axis=1)

        # df['sozialindex'] = df['Arbeitslosenquote in %'] / df['Kaufkraft p. E. in €']
        # df['stadt_qualtitaet'] = df['Kaufkraft p. E. in €'] * df['ÖPNV qualität'] / df['Arbeitslosenquote in %']

        print(f"Features {GREEN}hinzugefügt{RESET}")
        return df

    def korrelationsmatrix(self):
        plt.figure(figsize=(12, 12))
        sns.heatmap(self.df.corr(numeric_only=True), annot=True, cmap='coolwarm')
        plt.title("Korrelationsmatrix")
        plt.show()

    def entferne_korrelationen(self, df, zielspalte='Kaltmiete', mind_anforderung=0.1):
        print(f"{YELLOW}Analysiere{RESET} Korrelationen zu {zielspalte}")
        
        korrelation = df.corr(numeric_only=True)[zielspalte].abs()
        self.unwichtige_columns = korrelation[korrelation < mind_anforderung].index.tolist()
        if zielspalte in self.unwichtige_columns:
            self.unwichtige_columns.remove(zielspalte)
        print(f"{RED}Entferne{RESET}: {self.unwichtige_columns}")
        df = df.drop(columns=self.unwichtige_columns, errors='ignore')

        return df

    def entferne_ausreisser(self, df, zielspalte='Kaltmiete'):
        q1, q3 = df[zielspalte].quantile([0.25, 0.75])
        iqr = q3 - q1
        df = df[(df[zielspalte] >= q1 - 1.5 * iqr) & (df[zielspalte] <= q3 + 1.5 * iqr)]
        return df

    def trainiere_modell(self, df, zielspalte='Kaltmiete', test_daten_anteil=0.2):
        print(f"{YELLOW}Trainiere Modell für{RESET} {zielspalte}")
        df = self.entferne_ausreisser(df, zielspalte)

        #X = df.drop(columns=[zielspalte])
        X = df.drop(columns=[zielspalte, 'Warmmiete'], errors='ignore')
        y = df[zielspalte]

        X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=test_daten_anteil, random_state=42)

        modell = RandomForestRegressor(n_estimators=100, max_depth=30, random_state=42, n_jobs=-1)
        modell.fit(X_train, y_train)
        y_pred = modell.predict(X_test)

        mae = mean_absolute_error(y_test, y_pred)
        mse = mean_squared_error(y_test, y_pred)
        rmse = math.sqrt(mse)

        print(f"{GREEN}Modell erfolgreich trainiert{RESET}")
        print(f"{BLUE}MAE:{RESET} {mae:.2f}")
        print(f"{BLUE}RMSE:{RESET} {rmse:.2f}")
        print(f"{BLUE}MSE:{RESET} {mse:.2f}")

        return modell

    def bereinige_stadtname(self, name):
        if pd.isna(name): return ""
        name = re.sub(r'\s*\(.*?\)', '', str(name))
        name = re.sub(r',.*', '', name)
        name = re.sub(r'[^a-zA-ZäöüÄÖÜß\s-]', '', name)
        return name.lower().replace('ä', 'ae').replace('ö', 'oe').replace('ü', 'ue').strip()

    # Refactored add_ methods to operate on any DataFrame:
    def add_kaufkraft(self, df):
        print("Füge Kaufkraft hinzu...")
        kaufkraft = pd.read_excel(os.path.join(self.path, "Bundesland_Daten_2025.xlsx"))
        df["Bundesland"] = df["Bundesland"].apply(self.bereinige_stadtname)
        kaufkraft["Bundesland"] = kaufkraft["Bundesland"].apply(self.bereinige_stadtname)
        df =  df.merge(kaufkraft[['Bundesland', 'Kaufkraft 2025 pro Einwohner in €']], on="Bundesland", how="left")
        return df.rename(columns={'Kaufkraft 2025 pro Einwohner in €': 'Kaufkraft p. E. in €'})

    def add_arbeitslosenquote(self, df):
        print("Füge Arbeitslosenquote hinzu...")
        quote_stadt = pd.read_excel(os.path.join(self.path, "Gesellschafts_Daten_2024.xlsx"), sheet_name="SozialerUmkreis")
        quote_bundesland = pd.read_excel(os.path.join(self.path, "Bundesland_Daten_2025.xlsx"))
        quote_stadt["Stadt"] = quote_stadt["Stadt"].apply(self.bereinige_stadtname)
        df_merged = df.merge(quote_stadt[['Stadt', 'Arbeitslosenquote in %']], on="Stadt", how="left")
        
        quote_bundesland["Bundesland"] = quote_bundesland["Bundesland"].apply(self.bereinige_stadtname)
        df_merged = df_merged.merge(quote_bundesland[['Bundesland', 'Arbeitslosenquote in %']], on="Bundesland", how="left", suffixes=('', '_bl'))

        df_merged["Arbeitslosenquote in %"] = df_merged["Arbeitslosenquote in %"].fillna(df_merged["Arbeitslosenquote in %_bl"])
        df_merged = df_merged.drop(columns=["Arbeitslosenquote in %_bl"], errors='ignore')
        return df_merged

    def add_wohnungsleerstand(self, df):
        print("Füge Wohnungsleerstand hinzu...")
        wl = pd.read_excel(os.path.join(self.path, "Gesellschafts_Daten_2024.xlsx"), sheet_name="SozialerUmkreis")
        wl["Stadt"] = wl["Stadt"].apply(self.bereinige_stadtname)
        df["Stadt"] = df["Stadt"].apply(self.bereinige_stadtname)

        return df.merge(wl[['Stadt', 'Wohnungsleerstand']], on="Stadt", how="left")

    def add_oepnv(self, df):
        print("Füge ÖPNV-Qualität hinzu...")
        oepnv = pd.read_excel(os.path.join(self.path, "Gesellschafts_Daten_2024.xlsx"), sheet_name="OeffentlicherUmkreis")
        oepnv = oepnv.rename(columns={"Gemeindeverbandsname": "Stadt", "ÖPNV (Qualität)": "ÖPNV qualität"})
        oepnv["Stadt"] = oepnv["Stadt"].apply(self.bereinige_stadtname)
        df["Stadt"] = df["Stadt"].apply(self.bereinige_stadtname)

        return df.merge(oepnv[['Stadt', 'ÖPNV qualität']], on="Stadt", how="left")

    def bearbeite_testdaten(self, immobilien_daten, gesellschaft_datei, bundesland_datei, sheets, zielspalte):
        """
        Bearbeitet Testdaten, fügt externe Daten hinzu und bereinigt sie.
        Entfernt anschließend die Zielspalte aus den Features für die Vorhersage.
        """
        print(f"Verarbeite Testdaten...")
        df = immobilien_daten.copy()
        df = df.drop(columns=[zielspalte, 'Warmmiete'], errors='ignore')
        # Stadtspalten bereinigen, bevor gemerged wird
        df["Stadt"] = df["Stadt"].apply(self.bereinige_stadtname)
        df["Bundesland"] = df["Bundesland"].apply(self.bereinige_stadtname)

        # Externe Daten hinzufügen
        df = self.add_kaufkraft(df)
        df = self.add_arbeitslosenquote(df)
        df = self.add_wohnungsleerstand(df)
        df = self.add_oepnv(df)

        # Sicherstellen, dass Zielspalte keine fehlenden Werte enthält (optional)
        # if zielspalte in df.columns:
        #     df[zielspalte] = pd.to_numeric(df[zielspalte], errors='coerce')
        #     df[zielspalte] = df[zielspalte].fillna(df[zielspalte].median())

        # Zielspalte entfernen, da sie nicht in die Vorhersage einfließen soll
        features = df.drop(columns=[zielspalte], errors='ignore')

        return features

    def bekomme_daten_standartisieren(self, immobilie_bearbeitet):
        print(f"Standardisiere Testdaten...")

        le = LabelEncoder()

        for col in immobilie_bearbeitet.select_dtypes(include=['object']).columns:
            try:
                immobilie_bearbeitet[col] = pd.to_numeric(immobilie_bearbeitet[col])
            except:
                immobilie_bearbeitet[col] = le.fit_transform(immobilie_bearbeitet[col].astype(str))

        immobilie_bearbeitet = immobilie_bearbeitet.fillna(0)

        scaler = StandardScaler()
        scaled = scaler.fit_transform(immobilie_bearbeitet)

        scaled_df = pd.DataFrame(scaled, columns=immobilie_bearbeitet.columns)

        return scaled_df