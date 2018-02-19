#!/Library/Frameworks/Python.framework/Versions/3.6/bin/python3
# -*- coding: utf8 -*-

"""
Dieses Programm liest eine 'Dateiliste.txt' ein.
Diese Dateiliste enthält eine oder mehrere Excel-Files. Die Excel-Files können als xlsx vorliegen oder als
xlsb - Dateien. Die xlsb - Dateien werden - nur unter Windows mit installiertem Excel - in xlsx - Dateien
umgewandelt und dann verarbeitet.

Dieses Programm durchsucht die jeweilige Datei nach 'definierten' Fehlern.
Fehler können zum Beispiel sein:
* In einer Zelle, in der sich eine Formel befinden sollte, befindet sich keine Formel.
* In einer Formelzelle steht ein Fehlerwert.

Als Ergebnis wird ein Report in Excel erstellt. Dieser Report enthält pro getestete / in der Dateiliste.txt angegebene
Datei sowohl eine Zusammenfassung als auch im Detail, in welcher Zelle sich ein Problem befindet.
Man kann von den spezifischen / angegebenen Fehlern DIREKT zu der betroffenen Zelle in der Original-Datei springen.

Das Programm wurde nochmal komplett neu geschrieben, um dem Design Ansatz von 'Model-View-Controler' zu entsprechen.
Am 27.01.18 als GUI neu aufgesetzt.
"""

__version__ = "0.01b - 19.02.2018"
__author__ = "Christian Hetmann"

import tkinter as tk
from tkinter import ttk
from tkinter import scrolledtext
from tkinter.filedialog import askopenfilename
from tkinter import messagebox
import os
import sys
import time
import locale

LARGE_FONT = ("Verdana", 12)
# Breite der Buttons
BWIDTH = 21

class Model(object):
    pass


class View(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.fenster = parent
        self.fenster.title("Klaus' Excel-Checker")

        # Bildschirmbreite ermitteln und Fensterposition bestimmen
        x_pos = int((self.fenster.winfo_screenwidth() - self.fenster.winfo_reqwidth()) / 3)
        y_pos = int((self.fenster.winfo_screenheight() - self.fenster.winfo_reqheight()) / 3)
        self.fenster.geometry(f'{x_pos}x{2*y_pos}+{x_pos}+{y_pos}')

        # Überschreibt die Betriebssystem Funktion, wenn auf das 'x' zum schliessen gedrückt wird.
        # So kann noch was anderes vorher geprüft / gemacht werden, bevor das Programm endet.
        #self.fenster.protocol('WM_DELETE_WINDOW', self.click_beenden)
        #self.fenster.protocol('WM_DELETE_WINDOW', self.dummy)

        # Allgemeine Variablen definieren
        self.dateiliste = []  # Hier werden später alle Datei-Namen gespeichert
        self.fehler = False
        self.gespeichert = True
        self.Liste_Fehler = []

        # Tab-Control erstellen
        self.tab_control = ttk.Notebook(self.fenster)
        self.tab_control.grid(row=0, column=0, columnspan=50, rowspan=50, sticky='NESW')

        # Tabs erstellen
        self.tab1 = ttk.Frame(self.tab_control)
        self.tab2 = ttk.Frame(self.tab_control)
        self.tab3 = ttk.Frame(self.tab_control)
        self.tab4 = ttk.Frame(self.tab_control)

        # Beim "Grid" Window Manager sollte man einstellen, wieviel Spalten und Zeilen das Fenster hat
        tabs = [self.tab1, self.tab2, self.tab3, self.tab4]
        self.definiere_zeilen(self.fenster)  # Erst für das Fenster im allgemeinen
        for tab in tabs:  # dann für alle einzelnen Tabs
            self.definiere_zeilen(tab)

        # Tabs benennen
        self.tab_control.add(self.tab1, text='Status')
        self.tab_control.add(self.tab2, text='Zusammenfassung')
        self.tab_control.add(self.tab3, text='Details')
        self.tab_control.add(self.tab4, text='Dateiliste')

        # Tab 1 - Status
        # todo Buttons erstellen
        self.label1 = ttk.Label(self.tab1, text='STATUS')
        self.label1.grid(row=0, column=0, sticky='NSEW')

        self.tb1 = scrolledtext.ScrolledText(self.tab1, wrap=tk.WORD, height=25)
        self.tb1.grid(row=1, column=0, columnspan=39, sticky='NSEW')

        # self.btn_excel = ttk.Button(self.tab1, text='Excel-Check starten',
        #                             default='active', command=self.excel_focus, width=BWIDTH)
        self.btn_excel = ttk.Button(self.tab1, text='Excel-Check starten',
                                    default='active', command=self.dummy, width=BWIDTH)
        self.btn_excel.grid(row=29, column=0, padx=2, pady=2)

        # self.btn_dl_lesen = ttk.Button(self.tab1, text='Dateiliste neu lesen',
        #                                command=self.existiert_dateiliste, width=BWIDTH)
        self.btn_dl_lesen = ttk.Button(self.tab1, text='Dateiliste neu lesen',
                                       command=self.dummy, width=BWIDTH)
        self.btn_dl_lesen.grid(row=29, column=15, padx=2, pady=2)

        #self.btn_quit_t1 = ttk.Button(self.tab1, text='Beenden', command=self.click_beenden, width=BWIDTH)
        self.btn_quit_t1 = ttk.Button(self.tab1, text='Beenden', command=self.dummy, width=BWIDTH)
        self.btn_quit_t1.grid(row=29, column=30, padx=2, pady=2)

        # Tab 2 - Zusammenfassung
        # todo checken ob files xlsb oder xlsx sind
        self.label2 = ttk.Label(self.tab2, text='EXCEL-DATEI BEARBEITEN')
        self.label2.grid(row=0, column=0, sticky='NSEW')

        self.tb2 = scrolledtext.ScrolledText(self.tab2, wrap=tk.WORD, height=25)
        self.tb2.grid(row=1, column=0, columnspan=39, sticky='NSEW')

        self.btn_excel_t2 = ttk.Button(self.tab2, text='Excel-Check starten', command=self.dummy, width=BWIDTH)
        self.btn_excel_t2.grid(row=29, column=0, padx=2, pady=2)
        #self.btn_quit_t2 = ttk.Button(self.tab2, text='Beenden', command=self.click_beenden, width=BWIDTH)
        self.btn_quit_t2 = ttk.Button(self.tab2, text='Beenden', command=self.dummy, width=BWIDTH)
        self.btn_quit_t2.grid(row=29, column=30, padx=2, pady=2)

        # Tab 3 - Details
        self.label3 = ttk.Label(self.tab3, text='DETAILS')
        self.label3.grid(row=0, column=0, sticky='NSEW')

        self.tb3 = scrolledtext.ScrolledText(self.tab3, wrap=tk.WORD, height=25)
        self.tb3.grid(row=1, column=0, columnspan=39, sticky='NSEW')

        # self.btn_report_oeffnen = ttk.Button(self.tab3, text='Excel-Report öffnen', default='active',
        #                                      command=self.starte_excel_report, width=BWIDTH)
        self.btn_report_oeffnen = ttk.Button(self.tab3, text='Excel-Report öffnen', default='active',
                                             command=self.dummy, width=BWIDTH)
        self.btn_report_oeffnen.grid(row=29, column=0, padx=2, pady=2)

        #self.btn_quit_t3 = ttk.Button(self.tab3, text='Beenden', command=self.click_beenden, width=BWIDTH)
        self.btn_quit_t3 = ttk.Button(self.tab3, text='Beenden', command=self.dummy, width=BWIDTH)
        self.btn_quit_t3.grid(row=29, column=30, padx=2, pady=2)


        # Tab 4 - Dateiliste
        # todo Duplikate und leere Zeilen löschen
        self.label4 = ttk.Label(self.tab4, text='EINGELESENE DATEIEN')
        self.label4.grid(row=0, column=0, sticky='NSEW')

        self.tb4 = scrolledtext.ScrolledText(self.tab4, relief=tk.SUNKEN, wrap=tk.WORD, height=25)
        self.tb4.grid(row=1, column=0, columnspan=39, sticky='NSEW')
        self.tb4.insert(tk.END, 'Hier stehen nach der Auswahl die Dateien (inkl. Pfade) ...')

        # self.btn_waehlen = ttk.Button(self.tab4, text='Datei wählen',
        #                               default='active', command=self.click_datei_waehlen, width=BWIDTH)
        self.btn_waehlen = ttk.Button(self.tab4, text='Datei wählen',
                                      default='active', command=self.dummy, width=BWIDTH)
        self.btn_waehlen.grid(row=29, column=0, padx=2, pady=2)

        # self.btn_speichern = ttk.Button(self.tab4, text='Liste speichern',
        #                                 command=self.click_speicher_dateinamen, width=BWIDTH)
        self.btn_speichern = ttk.Button(self.tab4, text='Liste speichern',
                                        command=self.dummy, width=BWIDTH)
        self.btn_speichern.grid(row=29, column=15, padx=2, pady=2)

        #self.btn_quit_t4 = ttk.Button(self.tab4, text='Beenden', command=self.click_beenden, width=BWIDTH)
        self.btn_quit_t4 = ttk.Button(self.tab4, text='Beenden', command=self.dummy, width=BWIDTH)
        self.btn_quit_t4.grid(row=29, column=30, padx=2, pady=2)

        self.update()
        # Tab - Control
        #self.existiert_dateiliste()
        return


    def definiere_zeilen(self, element):
        zeile = 0
        while zeile < 30:
            element.rowconfigure(zeile, weight=1)
            element.columnconfigure(zeile, weight=1)
            zeile += 1
        return

    def dummy(self):
        pass

class Controller(object):

    def __init__(self, model, view):
        self.model = model
        self.view = view
        self.view.mainloop()
        return


if __name__ == '__main__':
    root = tk.Tk()
    #view = View(root).grid(sticky='NSEW')
    view = View(root)
    model = Model()
    controller = Controller(model, view)
    # controller.show_items()
    # controller.show_item_information('cheese')



