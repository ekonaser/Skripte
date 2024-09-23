import sys
import os
import csv

from PyQt6.QtWidgets import (
    QApplication, QWidget, QLineEdit, QPushButton, QTextEdit, QVBoxLayout, QToolBar, QMainWindow,
    QLabel, QStatusBar, QCheckBox, QFileDialog, QMessageBox
)
from PyQt6.QtGui import QAction
from PyQt6.QtCore import Qt

slovar_deklarantov = {}
slovar_CSV_podatkov = {}

def zamenjaj_z_none(row):
    return ['None' if not value else value for value in row]

class uvoznoOknoCSV(QWidget):
    def __init__(self, ttb):
        super().__init__()
        self.setWindowTitle("Uvozni podatki")
        self.setGeometry(300, 300, 300, 300)
        
        postavitev = QVBoxLayout()
        
        self.slovar_podatkov = {}
        self.potr_polja = []
        
        for nab in ttb:
            potrditveno_polje = QCheckBox(nab[0])
            self.slovar_podatkov[nab[0]] = nab
            self.potr_polja.append(potrditveno_polje)
            postavitev.addWidget(potrditveno_polje)
        
        self.uvozi = QPushButton("Uvozi")
        self.uvozi.clicked.connect(self._uvozi)
        
        self.zapri = QPushButton("Zapri")
        self.zapri.clicked.connect(self._zapri)
        
        postavitev.addWidget(self.uvozi)
        postavitev.addWidget(self.zapri)
        
        self.setLayout(postavitev)
    
    def _uvozi(self):
        """Metoda nam uvozi podatke, ki so obkljukani"""
        slovar_CSV_podatkov.clear()
        for polje in self.potr_polja:
            if polje.isChecked():
                slovar_CSV_podatkov[polje.text()] = self.slovar_podatkov[polje.text()]
                
        sporocilo = QMessageBox()
        sporocilo.setWindowTitle("Izbrani podatki")
        sporocilo.setText(', '.join(slovar_CSV_podatkov))
        sporocilo.exec()
        
    def _zapri(self):
        self.close()

class uvoznoOkno(QWidget):
    def __init__(self, tab):
        super().__init__()
        self.setWindowTitle("Uvozni podatki")
        self.setGeometry(300, 300, 300, 300)
        
        postavitev = QVBoxLayout()
        
        self.slovar = dict(tab)
        self.potr_polja = []
        for par in tab:
            potrditveno_polje = QCheckBox(par[0])
            self.potr_polja.append(potrditveno_polje)
            postavitev.addWidget(potrditveno_polje)
        
        self.uvozi = QPushButton("Uvozi")
        self.uvozi.clicked.connect(self._uvozi)
        
        self.zapri = QPushButton("Zapri")
        self.zapri.clicked.connect(self._zapri)
        
        postavitev.addWidget(self.uvozi)
        postavitev.addWidget(self.zapri)
        
        self.setLayout(postavitev)
    
    def _uvozi(self):
        """Metoda nam uvozi podatke, ki so obkljukani"""
        slovar_deklarantov.clear()
        for polje in self.potr_polja:
            if polje.isChecked():
                slovar_deklarantov[polje.text()] = self.slovar[polje.text()]
        
        sporocilo = QMessageBox()
        sporocilo.setWindowTitle("Izbrani deklarantje")
        sporocilo.setText(', '.join(slovar_deklarantov))
        sporocilo.exec()
    
    def _zapri(self):
        self.close()

class mojaAplikacija(QMainWindow):
    
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle('Distribucija')
        self.setGeometry(500, 500, 500, 500)
        
        # postavitev
        postavitev = QVBoxLayout()
        
        # menu
        datoteka = self.menuBar()
        
        datotecni_meni1 = datoteka.addMenu("&Datoteka")
        deklaranti = QAction("Dodaj deklarante", self)
        izhod = QAction("Izhod", self)
        uvozi_csv = QAction("Uvozi CSV/txt", self)
        shrani_podatke = QAction("Shrani podatke", self)
        shrani_kot_txt = QAction("Shrani kot TXT", self)
        shrani_kot_excel = QAction("Shrani kot xslx", self)
        datotecni_meni1.addAction(deklaranti)
        datotecni_meni1.addAction(uvozi_csv)
        datotecni_meni1.addAction(shrani_podatke)
        datotecni_meni1.addSeparator()
        datotecni_meni1.addAction(shrani_kot_txt)
        datotecni_meni1.addAction(shrani_kot_excel)
        datotecni_meni1.addSeparator()
        datotecni_meni1.addAction(izhod)
        
        datotecni_meni2 = datoteka.addMenu("&Pomoč")
        o_programu = QAction("O Programu", self)
        datotecni_meni2.addAction(o_programu)
        
        # menu ukazi
        izhod.triggered.connect(self._izhod)
        deklaranti.triggered.connect(self.izberi_deklarante)
        
        uvozi_csv.triggered.connect(self.izberi_podatke)
        
        shrani_podatke.triggered.connect(self.shrani_CSV_podatke)
        
    # metode
    def _izhod(self):
        self.close()
        
    def izberi_deklarante(self):
        datoteka = QFileDialog.getOpenFileName(self, 'Odpri datoteko', os.getcwd(), 'Vrsta datoteke (*.txt *.csv)')
        if datoteka[0] != '':
            self.dodaj_deklarante(datoteka[0])
    
    def izberi_podatke(self):
        datoteka = QFileDialog.getOpenFileName(self, 'Odpri datoteko', os.getcwd(), 'Vrsta datoteke (*.txt *.csv)')
        if datoteka[0] != '':
            self.dodaj_podatke(datoteka[0])
    
    def dodaj_deklarante(self, datoteka):
        tab = []
        with open(datoteka, 'r', encoding='utf-8-sig') as dat:
            for vr in dat:
                tab.append(vr.replace('\n','').split(sep=';'))
        
        self.uvozno_okno = uvoznoOkno(tab)
        self.uvozno_okno.show()
    
    def dodaj_podatke(self, datoteka):
        with open(datoteka, 'r', encoding='utf-8-sig') as dat:
            reader = csv.reader(dat, delimiter=';')
            ttb = zip(*[zamenjaj_z_none(row) for row in reader])
        
        self.uvozno_okno = uvoznoOknoCSV(ttb)
        self.uvozno_okno.show()
    
    def shrani_CSV_podatke(self):
        datoteka = QFileDialog.getSaveFileName(self, 'Shrani kot', os.getcwd(), 'Vrsta datoteke (*.txt *.csv)')
        # transponiramo nazaj podatke zapisane v slovarju
        if datoteka[0] != '':
            dat = open(datoteka[0], 'w', encoding='utf-8')
            tab = zip(*slovar_CSV_podatkov.values())
            for nab in tab:
                dat.write(';'.join(nab).replace('None','') + '\n')
        
distribucija = QApplication([])

glavno_okno = mojaAplikacija()

glavno_okno.show()

distribucija.exec()