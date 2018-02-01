#!/Library/Frameworks/Python.framework/Versions/3.6/bin/python3
# -*- coding: utf8 -*-


"""
Dieses Programm durchsucht mehrere in einer Datei angegebenen Dateiein nach definierten Fehlern
Die Dateien können als xlsb angegeben werden und werden dann für die Untersuchung temporär 
in xlsx umgewandelt. Die Fehler werden sowohl als Übersicht angegeben, alsauch im Detail.

Am 27.01.18 als GUI neu aufgesetzt.
"""

__version__ = "0.003b - 27.01.2018"
__author__ = "Christian Hetmann"

# todo xlsb-Datei mit Excel einlesen, in xlsx konvertieren und dann speichern
# todo xlsx mit openpyxl einlesen
# todo letzte relevante Zeile ermitteln
# todo alle Spalten mit Formeln identifizieren
# todo alle Zeilen durch gehen und nach Fehlern suchen
# todo cell.data_type BEACHTEN ... s=string, n=none, f=formula, e=error
# todo GUI zur Bedienung
# todo ermitteln, ob die Datei in der Liste eine xlsb oder xlsx ist, ggf. umwandeln

import tkinter as tk
from tkinter import ttk
from tkinter import scrolledtext
from tkinter.filedialog import askopenfilename
from tkinter import messagebox
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

        # Überschreibt die Betriebssystem Funktion, wenn auf das 'x' zum schliessen gedrückt wird.
        # So kann noch was anderes vorher geprüft / gemacht werden, bevor das Programm endet.
        self.fenster.protocol('WM_DELETE_WINDOW', self.click_beenden)


        # Allgemeine Variablen definieren
        self.dateiliste = []  # Hier werden später alle Datei-Namen gespeichert
        self.fehler = False
        self.gespeichert = True

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
        zeile = 0
        while zeile < 30:
            # self.tab1.rowconfigure(zeile, weight=1)
            # self.tab1.columnconfigure(zeile, weight=1)
            for t in tabs:
                t.rowconfigure(zeile, weight=1)
                t.columnconfigure(zeile, weight=1)
            zeile += 1

        # Tabs benamen
        self.tab_control.add(self.tab1, text='Status')
        self.tab_control.add(self.tab2, text='Zusammenfassung')
        self.tab_control.add(self.tab3, text='Details')
        self.tab_control.add(self.tab4, text='Dateiliste')

        # Tab 1 - Status
        # todo Buttons erstellen
        self.label1 = ttk.Label(self.tab1, text='STATUS')
        self.label1.grid(row=0, column=0, sticky='NSEW')

        self.tb1 = scrolledtext.ScrolledText(self.tab1, wrap=tk.WORD)
        self.tb1.grid(row=1, column=0, columnspan=39, sticky='NSEW')

        self.btn_excel = ttk.Button(self.tab1, text='Excel-Check starten',
                                    default='active', command=self.excel_focus, width=16)
        self.btn_excel.grid(row=29, column=0, padx=2, pady=2)
        self.btn_dl_lesen = ttk.Button(self.tab1, text='Dateiliste neu lesen',
                                       command=self.existiert_dateiliste, width=16)
        self.btn_dl_lesen.grid(row=29, column=15, padx=2, pady=2)
        self.btn_quit_t1 = ttk.Button(self.tab1, text='Beenden', command=self.click_beenden, width=16)
        self.btn_quit_t1.grid(row=29, column=30, padx=2, pady=2)

        # Tab 2 - Zusammenfassung
        # todo checken ob files xlsb oder xlsx sind
        self.label2 = ttk.Label(self.tab2, text='EXCEL-DATEI BEARBEITEN')
        self.label2.grid(row=0, column=0, sticky='NSEW')

        self.tb2 = scrolledtext.ScrolledText(self.tab2, wrap=tk.WORD, height=45)
        self.tb2.grid(row=1, column=0, columnspan=39, sticky='NSEW')

        self.btn_excel_t2 = ttk.Button(self.tab2, text='Excel-Check starten', command=dummy, width=16)
        self.btn_excel_t2.grid(row=29, column=0, padx=2, pady=2)

        self.btn_report = ttk.Button(self.tab2, text='Report speichern', command=dummy, width=16)
        self.btn_report.grid(row=29, column=15, padx=2, pady=2)

        self.btn_quit_t2 = ttk.Button(self.tab2, text='Beenden', command=self.click_beenden, width=16)
        self.btn_quit_t2.grid(row=29, column=30, padx=2, pady=2)

        # msg = 'Die Datei "Dateiliste.txt" existiert N I C H T !!!.\n' \
        #       'Du bla bla !\n\n'
        # self.tb2.insert(tk.END, msg)

        # Tab 3 - Details

        # Tab 4 - Dateiliste
        # todo Duplikate und leere Zeilen löschen
        self.label4 = ttk.Label(self.tab4, text='EINGELESENE DATEIEN')
        self.label4.grid(row=0, column=0)

        self.tb4 = scrolledtext.ScrolledText(self.tab4, relief=tk.SUNKEN, wrap=tk.WORD)
        self.tb4.grid(row=1, column=0, columnspan=39, sticky='NSEW')
        self.tb4.insert(tk.END, 'Hier stehen nach der Auswahl die Dateien (inkl. Pfade) ...')

        self.btn_waehlen = ttk.Button(self.tab4, text='Datei wählen',
                                      default='active', command=self.click_datei_waehlen, width=16)
        self.btn_waehlen.grid(row=29, column=0, padx=2, pady=2)
        self.btn_speichern = ttk.Button(self.tab4, text='Liste speichern',
                                        command=self.click_speicher_dateinamen, width=16)
        self.btn_speichern.grid(row=29, column=15, padx=2, pady=2)
        self.btn_quit_t4 = ttk.Button(self.tab4, text='Beenden', command=self.click_beenden, width=16)
        self.btn_quit_t4.grid(row=29, column=30, padx=2, pady=2)

        # Tab - Control
        self.existiert_dateiliste()
        return

    def existiert_dateiliste(self):
        self.tb1.delete(1.0, tk.END)
        if os.path.exists('Dateiliste.txt') == True:
            msg = 'Die Datei "Dateiliste.txt" existiert.\n\n'
            print(msg)
            self.tb1.insert(tk.END, msg)
            self.dateiliste_einlesen()
        else:
            msg = 'Die Datei "Dateiliste.txt" existiert N I C H T !!!.\n' \
                  'Du musst die "Dateiliste" im Tab Dateiliste neu erstellen!\n\n'
            print(msg)
            self.tb1.insert(tk.END, msg)
        return

    def click_datei_waehlen(self, event=None):
        # Checken, ob es schon eine Dateiliste gibt und erfolgreich eingelesen wurde
        def lese_datei_ein():
            filename = askopenfilename()
            self.dateiliste.append(filename)
            self.gespeichert = False
            self.fuelle_tb2()
            return
        if os.path.exists('Dateiliste.txt'):
            titel = '"Dateiliste.txt" existiert'
            ergebnis = messagebox.askyesno(titel, 'Es existiert bereits eine Dateiliste! Löschen und neue anlegen?')
            if ergebnis:
                os.remove('Dateiliste.txt')
                self.dateiliste = []
                lese_datei_ein()
        else:
            lese_datei_ein()
        self.tb4.focus()
        return

    def click_speicher_dateinamen(self, event=None):
        # Liste bereinigen: Duplikate löschen und leere Zeilen
        if len(self.dateiliste) > 0:
            self.liste_bereinigen()
            self.fuelle_tb2()
            print('Die Liste ist länger als 0 - es wird gespeichert!')
            dateiname = 'Dateiliste.txt'
            #import os.path
            def schreibe_datei(dname, dliste):
                with open(dname, 'w') as writefile:
                    for line in dliste:
                        writefile.write(line + '\n')
                self.gespeichert = True
                return
            if os.path.isfile(dateiname):
                ergebnis = messagebox.askyesno("Datei existiert!",
                                               "Die Datei existiert bereits! Wollen Sie sie überschreiben?")
                if ergebnis:
                    schreibe_datei(dateiname, self.dateiliste)
                else:
                    return
            else:
                schreibe_datei(dateiname, self.dateiliste)

        else:
            messagebox.showinfo("Speichern?", "Gibt nichts zu speichern!")
            self.tab4.focus()
        return

    def liste_bereinigen(self):
        if len(self.dateiliste) > 0:
            # Das ist ein Trick, die Liste in ein Set umzuwandeln, denn Sets haben keine Duplikate und dann
            # wieder zurück in eine Liste
            self.dateiliste = list(set(self.dateiliste))
            self.dateiliste.sort()
            loesch_index = []
            for ind, line in enumerate(self.dateiliste):
                if line == '':
                    loesch_index.append(ind)
            for i in sorted(loesch_index, reverse=True):
                del loesch_index[i]
        return

    def click_beenden(self):
        if self.gespeichert:
            self.fenster.destroy()
        else:
            ergebnis = messagebox.askyesno('Speichern?',
                                           'Die "Dateiliste.txt" wurde noch nicht gespeichert! Wirklich beenden?')
            if ergebnis:
                self.fenster.destroy()
            else:
                self.tab4.focus()
        return

    def excel_bearbeiten(self):
        # todo checke welche datei (xlsx od. xlsb)
        # todo wenn xlsb, konvertiere zu xlsx
        # todo wenn konvertieren, dann checke ob Excel noch auf ist ?!
        # todo wenn excel zu und konvertieren, dann konvertieren
        # todo öffne/lade xlsx (2x, daten, formeln)
        # todo ermittel die formel spalten
        # todo checke alle formel-spalten, ob fehler da sind
        # todo checke alle formel-spalten, ob formeln fehlen
        # todo schreibe zusammenfassung
        # todo schreibe detail report (evtl. html?)

        # Solange die Bearbeitung läuft, wird der Button disabled
        self.btn_excel_t2.configure(state='disabled')

        msg = 'Die Dateiliste enthält folgende Dateien:\n'
        for line in self.dateiliste:
            msg = msg + line + '\n'
        self.tb2.insert(tk.END, msg+'\n')

        # Alle Dateien nach einander abarbeiten
        for datei in self.dateiliste:
            self.ist_xlsb_oder_xlsx(datei)
        return

    def ist_xlsb_oder_xlsx(self, dateiname):
        """
        Diese Funktion tested, ob es sich um xlsx oder xlsb handelt
        :param dateiname: der pfad der datei inklusive dateinamen
        :return: True, wenn xlsx ist, False wenn es xlsb ist, sonst Datei-Liste neu!
        """
        if dateiname[-4:] == 'xlsx':
            print('xlsx')
        elif dateiname[-4:] == 'xlsb':
            print('xlsb')
        else:
            print('nix')
        return


    def excel_focus(self):
        self.tab_control.select(self.tab2)
        self.excel_bearbeiten()
        return


    def fuelle_tb2(self):
        self.tb4.delete(1.0, tk.END)
        for zeile in self.dateiliste:
            self.tb4.insert(tk.END, zeile + '\n')
        return


    def dateiliste_einlesen(self):
        tmp_dateiliste = []
        try:
            with open('Dateiliste.txt', 'r') as file:
                for line in file:
                    tmp_dateiliste.append(line.strip())
            print(tmp_dateiliste)
            self.dateiliste = tmp_dateiliste
            msg = f'Die Dateiliste wurde erfolgreich eingelesen. {len(self.dateiliste)} Dateinamen gefunden:\n'
            self.tb1.insert(tk.END, msg)
            self.aktiviere_deaktiviere_elemente(self.tab2, 'enable')
            self.aktiviere_deaktiviere_elemente(self.tab3, 'enable')
            for zeile in self.dateiliste:
                self.tb1.insert(tk.END, zeile+'\n')
            self.checke_alle_dateien()
        except IOError:
            msg = 'FEHLER beim Lesen der Datei! Erstelle eine neue "Dateiliste.txt"!'
            print(msg)
            self.tb1.insert(tk.END, msg)
            self.aktiviere_deaktiviere_elemente(self.tab2, 'disabled')
            self.aktiviere_deaktiviere_elemente(self.tab3, 'disabled')
        return

    def checke_alle_dateien(self):
        """
        Diese Funktion testet ob alle Datein in der Datei-Liste auch tatsächlich existieren
        :return: kein Rückgabewert
        """
        self.fehler = False
        for ind, datei in enumerate(self.dateiliste):
            if os.path.exists(datei):
                msg = f'\nDie {ind+1}. Datei "{datei}" existiert.\n'
                print(msg)
                self.tb1.insert(tk.END, msg)
            else:
                msg = f'\nDie {ind+1}. Datei "{datei}" existiert N I C H T !!!.\n'
                print(msg)
                self.tb1.insert(tk.END, msg)
                self.fehler = True
        if self.fehler:
            # Elemente deaktivieren
            msg = '\nDa es ein Problem mit einer Datei in der "Dateiliste.txt" gibt, sind einige Elemente ' \
                  'deaktiviert. Du musst die "Dateiliste.txt" in dem Tab Dateiliste neu erstellen.\n'
            print(msg)
            self.tb1.insert(tk.END, msg)
            self.aktiviere_deaktiviere_elemente(self.tab2, 'disabled')
            self.aktiviere_deaktiviere_elemente(self.tab3, 'disabled')
            self.btn_excel.configure(state='disabled')
            self.btn_quit_t1.configure(default='active')
        return

    def aktiviere_deaktiviere_elemente(self, element, state):
        for ele in element.winfo_children():
            try:
                ele.configure(state=state)
            except:
                print(f'Dieses Element kann nicht deaktiviert werden: {ele}')
        return

def dummy():
    print('Do nothing!')
    return

if __name__ == "__main__":
    root = tk.Tk()
    rows = 0
    while rows < 30:
        root.rowconfigure(rows, weight=1)
        root.columnconfigure(rows, weight=1)
        rows += 1
    Klaus_App(root).grid(sticky='NSEW')
    root.mainloop()
