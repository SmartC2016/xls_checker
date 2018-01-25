#!/Library/Frameworks/Python.framework/Versions/3.6/bin/python3
# -*- coding: utf8 -*-


"""
Dieses Programm durchsucht mehrere in einer Datei angegebenen Dateiein nach definierten Fehlern
Die Dateien können als xlsb angegeben werden und werden dann für die Untersuchung temporär 
in xlsx umgewandelt. Die Fehler werden sowohl als Übersicht angegeben, alsauch im Detail.

Am 25.01.18 als GUI neu aufgesetzt.
"""

__version__ = "0.003b - 25.01.2018"
__author__ = "Christian Hetmann"

# todo xlsb-Datei mit Excel einlesen, in xlsx konvertieren und dann speichern
# todo xlsx mit openpyxl einlesen
# todo letzte relevante Zeile ermitteln
# todo alle Spalten mit Formeln identifizieren
# todo alle Zeilen durch gehen und nach Fehlern suchen
# todo cell.data_type BEACHTEN ... s=string, n=none, f=formula, e=error
# todo GUI zur Bedienung
# todo Text-Datei einlesen mit absoluten Pfaden der einzulesenden und zu prüfenden Dateien ...
# todo ermitteln, ob die Datei in der Liste eine xlsb oder xlsx ist, ggf. umwandeln
# todo checken, ob die "Dateiliste.txt" bereits existiert.

import tkinter as tk
from tkinter import ttk
from tkinter import scrolledtext
import os

LARGE_FONT = ("Verdana", 12)

class Klaus_App(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.fenster = parent
        self.fenster.title("Klaus' Excel-Checker")

        # Bildschirmbreite ermitteln und Fensterposition bestimmen
        x_pos = int((self.fenster.winfo_screenwidth() - self.fenster.winfo_reqwidth()) / 3)
        y_pos = int((self.fenster.winfo_screenheight() - self.fenster.winfo_reqheight()) / 3)
        self.fenster.geometry(f'{x_pos}x{2*y_pos}+{x_pos}+{y_pos}')

        # Allgemeine Variablen definieren
        self.dateiliste = []  # Hier werden später alle Datei-Namen gespeichert

        # Tab-Control erstellen
        self.tab_control = ttk.Notebook(self.fenster)
        self.tab_control.grid(row=0, column=0, columnspan=50, rowspan=50, sticky='NESW')

        # Tabs erstellen
        self.tab1 = ttk.Frame(self.tab_control)
        self.tab2 = ttk.Frame(self.tab_control)
        self.tab3 = ttk.Frame(self.tab_control)
        self.tab4 = ttk.Frame(self.tab_control)


        zeile = 0
        while zeile < 50:
            self.tab1.rowconfigure(zeile, weight=1)
            self.tab1.columnconfigure(zeile, weight=1)
            zeile += 1


        # Tabs benamen
        self.tab_control.add(self.tab1, text='Status')
        self.tab_control.add(self.tab2, text='Zusammenfassung')
        self.tab_control.add(self.tab3, text='Details')
        self.tab_control.add(self.tab4, text='Dateiliste')

        # Tab 1 - Status
        # todo checken ob Dateiliste existiert
        # todo checken ob Dateiliste Files enthält
        # todo Duplikate und leere Zeilen löschen
        # todo checken of files auch wirklich da sind
        # todo checken ob files xlsb oder xlsx sind
        # todo Buttons erstellen

        self.label1 = ttk.Label(self.tab1, text='STATUS')
        self.label1.grid(row=0, column=0)

        self.tb1 = scrolledtext.ScrolledText(self.tab1, wrap=tk.WORD)
        self.tb1.grid(row=1, column=0, columnspan=49, sticky='NSEW')
        #self.tb1.insert(tk.END, 10*__doc__)

        self.btn1 = ttk.Button(self.tab1, text='Excel-Check starten!', default='active', command=dummy)
        self.btn1.grid(row=48, column=10, padx=5, pady=5)
        #self.btn1.focus()
        self.btn2 = ttk.Button(self.tab1, text='Programm beenden', command=quit)
        self.btn2.grid(row=48, column=40, padx=5, pady=5)


        # Tab 2 - Zusammenfassung
        self.lbl2 = tk.Label(self.tab2, text='My Label 2')
        self.lbl2.grid(row=0, column=0)

        self.combo = ttk.Combobox(self.tab2)
        self.combo['values'] = (1, 2, 3, 4, 5, "Text")
        self.combo.current(1)  # set the pre-selected item
        self.combo.grid(row=2, column=1)

        # Tab 3 - Details

        # Tab 4 - Dateiliste


        # Tab - Control
        self.existiert_dateiliste()
        return

    def existiert_dateiliste(self):
        if os.path.exists('Dateiliste.txt') == True:
            print('Dateiliste.txt', "existiert.")
            self.tb1.insert(tk.END, 'Die Datei "Dateiliste.txt" existiert.\n')
            self.dateiliste_einlesen()
        else:
            print('Dateiliste.txt', "existiert nicht.")
            self.tb1.insert(tk.END, 'Die Datei "Dateiliste.txt" existiert NICHT.')
        return

    def dateiliste_einlesen(self):
        tmp_dateiliste = []
        try:
            with open('Dateiliste.txt', 'r') as file:
                for line in file:
                    tmp_dateiliste.append(line.strip())
            print(tmp_dateiliste)
            self.dateiliste = tmp_dateiliste
            msg = f'Datei erfolgreich eingelesen. {len(self.dateiliste)} Dateinamen gefunden:\n'
            self.tb1.insert(tk.END, msg)
            for zeile in self.dateiliste:
                self.tb1.insert(tk.END, zeile+'\n')
        except IOError:
            msg = 'FEHLER beim Lesen der Datei! Erstelle eine neue "Dateiliste.txt"!'
            print(msg)
            self.tb1.insert(tk.END, msg)
        return

    def checke_alle_dateien(self):
        pass
        return

def dummy():
    print('Do nothing!')
    return





if __name__ == "__main__":
    root = tk.Tk()
    rows = 0
    while rows < 50:
        root.rowconfigure(rows, weight=1)
        root.columnconfigure(rows, weight=1)
        rows += 1
    Klaus_App(root).grid(sticky='NSEW')
    root.mainloop()
