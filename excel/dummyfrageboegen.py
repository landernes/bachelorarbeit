import os
import pandas as pd
import random


def sim_receiving_boegen(source_folder, target_folder):
    # Sicherstellen, dass der Zielordner existiert
    if not os.path.exists(target_folder):
        os.makedirs(target_folder)

    # Durchlaufen aller Dateien im Quellordner
    for filename in os.listdir(source_folder):
        if filename.endswith(".xlsx") or filename.endswith(".xls"):
            file_path = os.path.join(source_folder, filename)

            try:
                # Einlesen der Excel-Datei
                excel_file = pd.ExcelFile(file_path)

                # Neuer Dictionary für die bearbeiteten DataFrames
                writer = pd.ExcelWriter(os.path.join(target_folder, filename), engine='openpyxl')

                # Durchlaufen aller Blätter in der Excel-Datei
                for sheet_name in excel_file.sheet_names:
                    df = pd.read_excel(file_path, sheet_name=sheet_name)

                    # Überprüfen, ob die Spalte "Bewertung" existiert, andernfalls hinzufügen
                    if 'Bewertung' not in df.columns:
                        df['Bewertung'] = None

                    # Füllen der "Bewertung"-Spalte basierend auf der "Unterkriterien"-Spalte
                    df['Bewertung'] = df['Unterkriterien'].apply(
                        lambda x: random.randint(1, 5) if pd.notna(x) and x != '' else None)

                    # Speichern des bearbeiteten Blattes in die neue Excel-Datei
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

                # Schließen des ExcelWriters
                writer.close()

            except Exception as e:
                print(f"Fehler beim Verarbeiten der Datei {filename}: {e}")



#
# # Beispielaufruf der Funktion
# source_folder = "../files/bewertungsBoegen"
# target_folder = "../files/erwarteteBoegen"
# fill_ratings_in_excel(source_folder, target_folder)
