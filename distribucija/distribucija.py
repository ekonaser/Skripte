import sys
import os
import csv
import time

sys.path.append(os.getcwd() + '\\' + 'knjiznice')

from PyQt6.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, QMainWindow,
    QLabel, QFileDialog, QMessageBox, QTableWidget, QTableWidgetItem, QScrollArea, QHBoxLayout
)
from PyQt6.QtGui import QAction, QPixmap, QGuiApplication

from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas, NavigationToolbar2QT
import matplotlib.pyplot as plt
import numpy as np

from moduli import *

with open(os.getcwd() + '\\viri\\odstopi.csv','r',encoding='utf-8-sig') as dat:
    for vr in dat:
        mno_odstopov.add(vr.replace('\n',''))

class MplCanvas(FigureCanvas):

    def __init__(self, parent=None, width=5, height=4, dpi=200):
        fig, self.ax = plt.subplots(figsize=(width, height), dpi=dpi)
        super().__init__(fig)

class mojaAplikacija(QMainWindow, shraniXLSX, shraniTXT, delitevMnozic):
    
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle('Distribucija')
        self.setGeometry(500, 500, 500, 500)
        #self.resize(500, 500)
        self.center_zaslona()
        
        # postavitev
        postavitev = QVBoxLayout()
        
        # graf
        self.okvir_grafa = MplCanvas(self, width=5, height=4, dpi=85)
        orodja_grafa = NavigationToolbar2QT(self.okvir_grafa, self)
        
        # menu
        datoteka = self.menuBar()
        
        datotecni_meni1 = datoteka.addMenu("&Datoteka")
        deklaranti = QAction("Izberi deklarante", self)
        izhod = QAction("Izhod", self)
        uvozi_csv = QAction("Uvozi CSV/txt", self)
        dodaj_deklarante = QAction("Dodaj deklarante", self)
        odstrani_deklarante = QAction("Odstrani deklarante", self)
        shrani_podatke_kot = QAction("Shrani CSV/txt kot", self)
        shrani_kot_txt = QAction("Shrani kot TXT", self)
        shrani_kot_excel = QAction("Shrani kot xslx", self)
        datotecni_meni1.addAction(deklaranti)
        datotecni_meni1.addAction(dodaj_deklarante)
        datotecni_meni1.addAction(odstrani_deklarante)
        datotecni_meni1.addAction(uvozi_csv)
        datotecni_meni1.addAction(shrani_podatke_kot)
        datotecni_meni1.addSeparator()
        datotecni_meni1.addAction(shrani_kot_txt)
        datotecni_meni1.addAction(shrani_kot_excel)
        datotecni_meni1.addSeparator()
        datotecni_meni1.addAction(izhod)
        
        datotecni_meni2 = datoteka.addMenu("&Pomoč")
        o_programu = QAction("O Programu", self)
        datotecni_meni2.addAction(o_programu)
        
        # dodamo sliko
        # postavitev.addWidget(self.image_label)
        # gumbi
        self.izracunaj = QPushButton("Naredi distribucijo")
        self.izracunaj.clicked.connect(self.distribucija)
        postavitev.addWidget(self.izracunaj)
        
        self.razdeli_deklarantom = QPushButton("Razdeli med izbrane deklarante")
        # zato da gumb pravilno deluje moramo najprej preveriti mnozice
        self.razdeli_deklarantom.clicked.connect(self.preveri_mnozice)
        postavitev.addWidget(self.razdeli_deklarantom)
        
        # menu ukazi
        izhod.triggered.connect(self._izhod)
        deklaranti.triggered.connect(self.izberi_deklarante)
        dodaj_deklarante.triggered.connect(self.dodaj_deklaranta)
        odstrani_deklarante.triggered.connect(self._odstrani_deklarante)
        uvozi_csv.triggered.connect(self.izberi_podatke)
        shrani_podatke_kot.triggered.connect(self.shrani_CSV_podatke_kot)
        shrani_kot_txt.triggered.connect(self.shrani_TXT)
        shrani_kot_excel.triggered.connect(self.shrani_XLSX)
        o_programu.triggered.connect(self.oprogramu)
        
        postavitev.addWidget(orodja_grafa)
        postavitev.addWidget(self.okvir_grafa)
        
        # Create a central widget and set the layout
        central_widget = QWidget()
        central_widget.setLayout(postavitev)
        self.setCentralWidget(central_widget)
               
    # metode
    def _izhod(self):
        self.close()
    
    def shrani_XLSX(self):
        self.shrani_kot_XLSX(slovar_deklarantov, slovar_deklarantov_odstopi)
    
    def shrani_TXT(self):
        self.shrani_kot_TXT(slovar_deklarantov, slovar_deklarantov_odstopi)
    
    def izberi_podatke(self):
        datoteka = QFileDialog.getOpenFileName(self, 'Odpri datoteko', os.getcwd(), 'Vrsta datoteke (*.txt *.csv)')
        if datoteka[0] != '':
            self.dodaj_podatke(datoteka[0])
    
    def dodaj_podatke(self, datoteka):
        with open(datoteka, 'r', encoding='utf-8-sig') as dat:
            reader = csv.reader(dat, delimiter=';')
            ttb = zip(*[zamenjaj_z_none(row) for row in reader])
        
        self.uvozno_okno = uvoznoOknoCSV(ttb)
        self.uvozno_okno.show()
    
    def izberi_deklarante(self, datoteka): 
        self.uvozno_okno = izberiDeklarante(preberi_deklarante())
        self.uvozno_okno.show()
        
    def _odstrani_deklarante(self, datoteka): 
        self.izbirno_okno = odstraniDeklarante(preberi_deklarante())
        self.izbirno_okno.show()
    
    def shrani_CSV_podatke_kot(self):
        datoteka = QFileDialog.getSaveFileName(self, 'Shrani kot', os.getcwd(), 'Vrsta datoteke (*.txt *.csv)')
        # transponiramo nazaj podatke zapisane v slovarju
        if datoteka[0] != '':
            dat = open(datoteka[0], 'w', encoding='utf-8')
            tab = zip(*slovar_CSV_podatkov.values())
            for nab in tab:
                dat.write(';'.join(nab).replace('None','') + '\n')
            dat.close()
    
    def distribucija(self):
        """Metoda nam naredi distribucijo"""
        if not slovar_deklarantov:
            sporocilo = QMessageBox()
            sporocilo.setWindowTitle("Pozor!")
            sporocilo.setText("Slovar deklarantov je prazen")
            sporocilo.exec()
        elif not slovar_CSV_podatkov:
            sporocilo = QMessageBox()
            sporocilo.setWindowTitle("Pozor!")
            sporocilo.setText("Slovar podatkov je prazen")
            sporocilo.exec()
        else:
            # to postane seznam na pomnilniku in ga ne spreminjaj!
            # seznam je poln naborov
            nova_tab1 = []
            nova_tab2 = []
            tab = list(zip(*slovar_CSV_podatkov.values()))
            for nab in tab:
                if filter_odstopov(nab, mno_odstopov):
                    nova_tab1.append(nab)
                else:
                    nova_tab2.append(nab)
            mno_odstopi.update(nova_tab1)
            self.uvozno_okno = izberiCSV(nova_tab2)
            self.uvozno_okno.show()
    
    def preveri_mnozice(self):
        """Metoda zazene podedovano metodo ter nam izrise graf"""
        if mno_odstopi or mno_fiz or mno_pod:
            self.razdeli_med_deklarante(slovar_deklarantov, slovar_deklarantov_odstopi, mno_fiz, mno_pod, mno_odstopi)
            # tukaj resetiramo nas graf
            self.okvir_grafa.ax.cla()
            # Mnozice
            mno_H7_fiz = set([nab[0] for nab in mno_fiz if nab[2] == 'H7'])
            mno_FORMAL_fiz = set([nab[0] for nab in mno_fiz if nab[2] == 'FORMAL'])
            mno_H7IOSS_fiz = set([nab[0] for nab in mno_fiz if nab[2] == 'H7IOSS'])
            
            mno_H7_pod = set([nab[0] for nab in mno_pod if nab[2] == 'H7'])
            mno_FORMAL_pod = set([nab[0] for nab in mno_pod if nab[2] == 'FORMAL'])
            mno_H7IOSS_pod = set([nab[0] for nab in mno_pod if nab[2] == 'H7IOSS'])
            
            mno_ODSTOPI = set([nab[0] for nab in mno_odstopi])
            slovar = {k:v for k,v in slovar_deklarantov.items()} | {k:v for k,v in slovar_deklarantov_odstopi.items()}
            slo_mno = {k:set(v[2]) for k,v in slovar.items()} # glavni slovar
            imena = [k for k,v in slo_mno.items()]
            skupaj = [len(v) for k,v in slo_mno.items()]
            
            h7_fiz = [len(v & mno_H7_fiz) for k,v in slo_mno.items()]
            formal_fiz = [len(v & mno_FORMAL_fiz) for k,v in slo_mno.items()]
            h7ioss_fiz = [len(v & mno_H7IOSS_fiz) for k,v in slo_mno.items()]
            
            h7_pod = [len(v & mno_H7_pod) for k,v in slo_mno.items()]
            formal_pod = [len(v & mno_FORMAL_pod) for k,v in slo_mno.items()]
            h7ioss_pod = [len(v & mno_H7IOSS_pod) for k,v in slo_mno.items()]
            
            odstopi = [len(v & mno_ODSTOPI) for k,v in slo_mno.items()]
            
            tab_vseh = [h7_fiz, formal_fiz, h7ioss_fiz, h7_pod, formal_pod, h7ioss_pod, odstopi]
            tab_imen = ['H7 F', 'Formal F', 'H7IOSS F', 'H7 P', 'Formal P', 'H7IOSS P', 'Odstopi']
            
            n_groups = len(imena)
            bar_width = 0.30
            index = np.arange(len(imena))
            
            fig, ax = plt.subplots()
            
            self.okvir_grafa.ax.bar(index, h7_fiz, bar_width, label='H7 F')
        
            # Stacking the rest of the bars on top of the first set
            bottom = np.array(h7_fiz)
            for i, tab in enumerate(tab_vseh[1:], start=1):
                self.okvir_grafa.ax.bar(index, tab, bar_width, bottom=bottom, label=tab_imen[i])
                bottom += np.array(tab)
            
            self.okvir_grafa.ax.set_ylabel('Število pošiljk')
            self.okvir_grafa.ax.set_title('Distribucija')
            self.okvir_grafa.ax.set_xticks(index)
            self.okvir_grafa.ax.set_xticklabels(imena)
            self.okvir_grafa.ax.set_ylim(0, max(bottom) * 1.2)
            self.okvir_grafa.ax.legend()

            self.okvir_grafa.draw()
            
        else:
            sporocilo = QMessageBox()
            sporocilo.setWindowTitle("Pozor!")
            sporocilo.setText("Niste dodali podatkov")
            sporocilo.exec()
    
    def dodaj_deklaranta(self):
        self.dod_dek = dodajDeklaranta()
        self.dod_dek.show()
        
    def oprogramu(self):
        self.prikaz_o_programu = oProgramu()
        self.prikaz_o_programu.show()
    
    def center_zaslona(self):
        screen = QGuiApplication.primaryScreen().availableGeometry().center()
        frame = self.frameGeometry()
        frame.moveCenter(screen)
        self.move(frame.topLeft())
    
distribucija = QApplication([])
glavno_okno = mojaAplikacija()
glavno_okno.show()
distribucija.exec()