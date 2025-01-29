import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import json
import pandas as pd
import os
import re
import excel.erzeuge_boegen as ee
import mail.mail as mail
import mail.mail_lieferant as mail_lieferant
import excel.bewerten as eb
import excel.dummyfrageboegen as ed
import excel.bewertung_hinzu as ebh
import excel.lieferantenscoresliste as el


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Liste und Tabelle")

        # Menüleiste erstellen
        menubar = tk.Menu(root)
        root.config(menu=menubar)

        # Datei-Menü
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Datei", menu=file_menu)
        file_menu.add_command(label="Schließen", command=self.quit)

        # Bewertung-Menü
        bewertung_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Bewertung", menu=bewertung_menu)
        bewertung_menu.add_command(label="Bewertung starten", command=self.start_bewertung)

        # Rahmen für die linke Seite
        left_frame = tk.Frame(root)
        left_frame.pack(side=tk.LEFT, fill=tk.Y)

        # Liste
        self.listbox = tk.Listbox(left_frame)
        self.listbox.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.listbox.bind('<<ListboxSelect>>', self.on_listbox_select)

        # Aktualisieren-Knopf
        refresh_button = tk.Button(left_frame, text="Aktualisieren", command=self.load_data)
        refresh_button.pack(side=tk.TOP, fill=tk.X)

        # Rahmen für die rechte Seite
        right_frame = tk.Frame(root)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # Tabelle
        self.table = ttk.Treeview(right_frame,
                                  columns=("E-Mail", "Ansprechpartner", "Bewertungsscore", "Letzte Bewertung"),
                                  show='headings')
        self.table.heading("E-Mail", text="E-Mail")
        self.table.heading("Ansprechpartner", text="Ansprechpartner")
        self.table.heading("Bewertungsscore", text="Bewertungsscore")
        self.table.heading("Letzte Bewertung", text="Letzte Bewertung")
        self.table.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # Excel-Tabelle
        self.excel_table = ttk.Treeview(right_frame)
        self.excel_table.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # Daten laden
        self.data = []
        self.excel_data = None
        self.excel_directory = './files/bewertungen'
        self.load_data()

    def load_data(self):
        # JSON-Daten laden
        with open('./files/lieferanten.json', 'r') as file:
            self.data = json.load(file)

        # Listbox aktualisieren
        self.listbox.delete(0, tk.END)
        for item in self.data:
            self.listbox.insert(tk.END, item['Name'])

    def on_listbox_select(self, event):
        selected_index = self.listbox.curselection()
        if selected_index:
            selected_item = self.data[selected_index[0]]
            self.update_table(selected_item)
            self.load_excel_data(selected_item['Name'])

    def update_table(self, item):
        # Löschen der alten Daten in der Tabelle
        for row in self.table.get_children():
            self.table.delete(row)

        # Neue Daten in die Tabelle einfügen
        self.table.insert('', tk.END, values=(
        item["E-Mail"], item["Ansprechpartner"], item["Bewertungsscore"], item["LetzteBewertung"]))

    def update_excel_table(self):
        if self.excel_data is not None:
            # Löschen der alten Daten in der Excel-Tabelle
            for row in self.excel_table.get_children():
                self.excel_table.delete(row)

            # Spaltenüberschriften setzen, letzte Spaltenüberschrift ausblenden
            columns = list(self.excel_data.columns)
            self.excel_table["columns"] = columns[:-1]
            self.excel_table["show"] = "headings"
            for col in self.excel_table["columns"]:
                self.excel_table.heading(col, text=col)
                self.excel_table.column(col, anchor="center")

            # Neue Daten in die Excel-Tabelle einfügen
            for _, row in self.excel_data.iterrows():
                values = ["" if pd.isna(value) else value for value in row[:-1]]
                self.excel_table.insert("", "end", values=values)

    def load_excel_data(self, name):
        # Namen umwandeln
        file_name = re.sub(r'\W+', '', name.lower()) + ".xlsx"
        file_path = os.path.join(self.excel_directory, file_name)

        if os.path.exists(file_path):
            xls = pd.ExcelFile(file_path)
            sheet_name = max(xls.sheet_names, key=lambda s: pd.to_datetime(s, errors='coerce'))
            self.excel_data = pd.read_excel(xls, sheet_name=sheet_name)
            self.update_excel_table()
        else:
            messagebox.showerror("Fehler", f"Die Datei {file_name} wurde nicht im Verzeichnis gefunden.")
            self.excel_data = None
            self.update_excel_table()

    def quit(self):
        self.root.quit()

    def start_bewertung(self):
        # Popup-Fenster erstellen
        self.popup = tk.Toplevel(self.root)
        self.popup.title("Bewertung starten")

        # Scrollbar hinzufügen
        scrollbar = tk.Scrollbar(self.popup)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Canvas hinzufügen
        canvas = tk.Canvas(self.popup, yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=canvas.yview)

        # Frame für Checkboxen
        checkbox_frame = tk.Frame(canvas)
        canvas.create_window((0, 0), window=checkbox_frame, anchor='nw')

        # Variable für Checkboxen
        self.check_vars = {}
        for item in self.data:
            var = tk.BooleanVar()
            chk = tk.Checkbutton(checkbox_frame, text=item['Name'], variable=var)
            chk.pack(anchor='w')
            self.check_vars[item['Name']] = var

        # Funktion zum Anpassen der Canvas-Größe
        def on_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        checkbox_frame.bind("<Configure>", on_configure)

        # OK-Button hinzufügen
        ok_button = tk.Button(self.popup, text="OK", command=self.evaluate)
        ok_button.pack(side=tk.BOTTOM)

    def evaluate(self):
        selected_items = [key for key, var in self.check_vars.items() if var.get()]
        messagebox.showinfo("Ausgewählte Elemente",
                            f"Sie haben folgende Elemente ausgewählt:\n{', '.join(selected_items)}")
        ee.create_excel_files('./files/kriterienKatalog.xlsx', selected_items)
        ed.sim_receiving_boegen('./files/bewertungsBoegen', './files/erwarteteBoegen')
        mail.send_mail_intern()
        eb.evaluate('./files/ausführendeBewertung.xlsx', './files/erwarteteBoegen')
        el.update_lieferanten_json('./files/ausführendeBewertung.xlsx', './files/lieferanten.json')
        mail_lieferant.send_mail_lieferanten(selected_items)
        ebh.update_bewertungen('./files/ausführendeBewertung.xlsx', './files/bewertungen')
        self.popup.destroy()


