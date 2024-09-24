import sys
import os
import csv
import xlsxwriter
import itertools
import time

from PyQt6.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, QMainWindow,
    QLabel, QCheckBox, QFileDialog, QMessageBox, QTableWidget, QTableWidgetItem, QScrollArea, QHBoxLayout
)
from PyQt6.QtGui import QAction, QPixmap
from PyQt6.QtCore import Qt

# mno_odstopov je mnozica ki jo uporabljamo za primerjanje
mno_odstopov = set()
slovar_deklarantov = {}
slovar_CSV_podatkov = {}
slovar_deklarantov_odstopi = {}
# tri mnozice za razvrscanje
mno_odstopi = set()
mno_fiz = set()
mno_pod = set()

with open('odstopi.csv','r',encoding='utf-8-sig') as dat:
    for vr in dat:
        mno_odstopov.add(vr.replace('\n',''))

def zamenjaj_z_none(row):
    return ['None' if not value else value for value in row]

def filter_odstopov(x: tuple, y: set):
    """Funkcija nam filtrira odstope s pomocjo mnozic"""
    if set(x).intersection(y):
        return True
    return False

class prikaziDistribucijo(QWidget):
    def __init__(self, slovar):
        super().__init__()
        self.setWindowTitle("Delitev")
        self.setGeometry(300, 300, 300, 300)
        
        self.tabela = QTableWidget()
        
        nov_slo = {k: v[2] for k, v in slovar.items()}
        
        glave = list(nov_slo.keys())
        vrstice = len(next(iter(nov_slo.values())))
        
        self.tabela.setColumnCount(len(nov_slo))
        # +2 ker dodame se stevilo vseh deklaracij dodeljenih deklarantu
        # v prvo vrstico
        self.tabela.setRowCount(vrstice+2)
        self.tabela.setHorizontalHeaderLabels(glave)
        
        st = 0
        for k, v in nov_slo.items():
            self.tabela.setItem(0, st, QTableWidgetItem(str(len(v))))
            st += 1
        
        for stolpec, (glava, vrednost) in enumerate(nov_slo.items()):
            for vrstica, vrednost in enumerate(vrednost, 1):
                self.tabela.setItem(vrstica, stolpec, QTableWidgetItem(vrednost))
        
        layout = QVBoxLayout()
        layout.addWidget(self.tabela)
        self.setLayout(layout)


class izberiCSV(QMainWindow):
    def __init__(self, tab_nab):
        super().__init__()
        self.setWindowTitle("Podatki")
        self.setGeometry(100, 100, 600, 400)
        
        self.selected_rows = set()
        self.prvotna_mno = set()
        
        self.tabela = QTableWidget()
        # po 'dogovoru' bomo stevilo vrstic zmanjsali za 1, saj
        # prvo vrstico uporabi za naslove stolpcev
        self.tabela.setRowCount(len(tab_nab)-1)
        # ker bomo dodali se dva gumba dodamo + 2
        self.tabela.setColumnCount(len(tab_nab[0]) + 2)
        self.tabela.setHorizontalHeaderLabels(tab_nab.pop(0) + ("Dodaj", "Odstrani"))
        # sele ko odstranimo prvi element tabele z nabori
        # posodobimo mnozico
        self.prvotna_mno.update(tab_nab)
        
        for vrstica, nab in enumerate(tab_nab):
            for x in range(len(nab)):
                self.tabela.setItem(vrstica, x, QTableWidgetItem(nab[x]))
            button1 = QPushButton("Dodaj")
            button1.clicked.connect(lambda _, row=vrstica: self.add_to_set(row))
            button_widget1 = QWidget()
            button_layout1 = QHBoxLayout(button_widget1)
            button_layout1.addWidget(button1)
            button_layout1.setAlignment(Qt.AlignmentFlag.AlignCenter)
            button_layout1.setContentsMargins(0, 0, 0, 0)
            self.tabela.setCellWidget(vrstica, x+1, button_widget1)
            
            button2 = QPushButton("Odstrani")
            button2.clicked.connect(lambda _, row=vrstica: self.remove_from_set(row))
            button_widget2 = QWidget()
            button_layout2 = QHBoxLayout(button_widget2)
            button_layout2.addWidget(button2)
            button_layout2.setAlignment(Qt.AlignmentFlag.AlignCenter)
            button_layout2.setContentsMargins(0, 0, 0, 0)
            self.tabela.setCellWidget(vrstica, x+2, button_widget2)
        
        drsenje = QScrollArea()
        drsenje.setWidgetResizable(True)
        drsenje.setWidget(self.tabela)
        
        postavitev = QVBoxLayout()
        postavitev.addWidget(drsenje)
        
        self.razdeli_gumb = QPushButton("Razdeli")
        self.razdeli_gumb.clicked.connect(self.razdeli)
        postavitev.addWidget(self.razdeli_gumb)
        
        self.zapri = QPushButton("Zapri")
        self.zapri.clicked.connect(self._zapri)
        postavitev.addWidget(self.zapri)
        
        vsebina = QWidget()
        vsebina.setLayout(postavitev)
        self.setCentralWidget(vsebina)
    
    def add_to_set(self, row):
        row_data = tuple(self.tabela.item(row, col).text() for col in range(self.tabela.columnCount() - 2))
        self.selected_rows.add(row_data)
        
        button_widget = self.tabela.cellWidget(row, self.tabela.columnCount() - 2)
        button_widget.setStyleSheet("background-color: rgba(77, 20, 140, 64);")
    
    def remove_from_set(self, row):
        row_data = tuple(self.tabela.item(row, col).text() for col in range(self.tabela.columnCount() - 2))
        self.selected_rows.discard(row_data)
        
        button_widget = self.tabela.cellWidget(row, self.tabela.columnCount() - 2)
        button_widget.setStyleSheet("background-color: none;")
    
    def razdeli(self):
        """Metoda nam razdeli podatke in posodobi dve globalni mnozici"""
        mno_fiz.clear()
        mno_pod.clear()
        
        mno_fiz.update(self.selected_rows)
        mno_pod.update(self.prvotna_mno - self.selected_rows)
    
    def _zapri(self):
        self.close()

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
        
        drsenje = QScrollArea()
        drsenje.setWidgetResizable(True)
        
        # Create a widget to hold the checkboxes and set it as the scroll area's widget
        vsebina = QWidget()
        vsebina.setLayout(postavitev)
        drsenje.setWidget(vsebina)
        
        # Add the scroll area to the main layout
        glavni_postavitev = QVBoxLayout()
        glavni_postavitev.addWidget(drsenje)
        
        self.setLayout(glavni_postavitev)
    
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
                slovar_deklarantov[polje.text()] = self.slovar[polje.text()].split(sep='-') + [[]]
        
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
        
        # slika
        self.image_label = QLabel(self)
        pixmap = QPixmap('fedex.png')
        self.image_label.setPixmap(pixmap)
        self.image_label.setScaledContents(True)  # Enable scaled contents
        self.image_label.resize(pixmap.width(), pixmap.height())
        
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
        
        # dodamo sliko
        postavitev.addWidget(self.image_label)
        # gumbi
        self.izracunaj = QPushButton("Naredi distribucijo")
        self.izracunaj.clicked.connect(self.distribucija)
        postavitev.addWidget(self.izracunaj)
        
        self.razdeli_deklarantom = QPushButton("Razdeli med izbrane deklarante")
        self.razdeli_deklarantom.clicked.connect(self.razdeli_med_deklarante)
        postavitev.addWidget(self.razdeli_deklarantom)
        
        # menu ukazi
        izhod.triggered.connect(self._izhod)
        deklaranti.triggered.connect(self.izberi_deklarante)
        
        uvozi_csv.triggered.connect(self.izberi_podatke)
        
        shrani_podatke.triggered.connect(self.shrani_CSV_podatke)
        
        shrani_kot_excel.triggered.connect(self.shrani_kot_XSLX)
        
        # Create a central widget and set the layout
        central_widget = QWidget()
        central_widget.setLayout(postavitev)
        self.setCentralWidget(central_widget)
               
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
    
    def razdeli_med_deklarante(self):
        """Metoda nam razdeli podatke globalno."""
        # Poiscemo samo AWB stevilke
        # za pozneje lahko das pogoj: len(st) == 11 ali 12
        for k,v in slovar_deklarantov.items():
            v[2].clear()
        
        mno_odstopi0 = {[st for st in nab if st.isdigit()][0] for nab in mno_odstopi}
        mno_fiz0 = {[st for st in nab if st.isdigit()][0] for nab in mno_fiz}
        mno_fiz_formal = {[niz for niz in nab if niz.isdigit()][0] for nab in mno_fiz if 'FORMAL' in nab}
        mno_pod0 = {[st for st in nab if st.isdigit()][0] for nab in mno_pod}
        # ODSTOPI
        sez_odstopi = list(mno_odstopi0)
        slovar_deklarantov_odstopi['Odstopi'] = ['any', 100, []]
        for awb in sez_odstopi:
            slovar_deklarantov_odstopi['Odstopi'][2].append(awb)
        if 'odstop' in [deklarant[0] for deklarant in slovar_deklarantov.values()]:
            for k,v in slovar_deklarantov.items():
                if 'odstop' in v:
                    while mno_odstopi0 and len(slovar_deklarantov[k][2]) != int(slovar_deklarantov[k][1]):
                        slovar_deklarantov[k][2].append(mno_odstopi0.pop())
        
        while mno_fiz_formal:
            if 'formal' in [deklarant[0] for deklarant in slovar_deklarantov.values()]:
                for k,v in slovar_deklarantov.items():
                    if 'formal' in v:
                        while mno_fiz_formal and len(slovar_deklarantov[k][2]) <= int(slovar_deklarantov[k][1]):
                            slovar_deklarantov[k][2].append(mno_fiz_formal.pop())
                # potrebno je dodati break stavek da se ne zaciklamo
                break
        
        kljuc = itertools.cycle(slovar_deklarantov.keys())
        # po logiki najmanjse tabele v tabeli vrednosti kljuca
        najmanjsa_st = len(slovar_deklarantov[next(kljuc)][2])
        while mno_fiz_formal:
            naslednji_kljuc = next(kljuc)
            if len(slovar_deklarantov[naslednji_kljuc][2]) < int(slovar_deklarantov[naslednji_kljuc][1]) and slovar_deklarantov[naslednji_kljuc][0] != 'formal' and len(slovar_deklarantov[naslednji_kljuc][2]) == najmanjsa_st:
                slovar_deklarantov[naslednji_kljuc][2].append(mno_fiz_formal.pop())
                najmanjsa_st = min([len(tab[2]) for tab in slovar_deklarantov.values() if 'any' in tab])
        while mno_fiz0:
            naslednji_kljuc = next(kljuc)
            if len(slovar_deklarantov[naslednji_kljuc][2]) < int(slovar_deklarantov[naslednji_kljuc][1]) and slovar_deklarantov[naslednji_kljuc][0] != 'formal' and len(slovar_deklarantov[naslednji_kljuc][2]) == najmanjsa_st:
                slovar_deklarantov[naslednji_kljuc][2].append(mno_fiz0.pop())
                najmanjsa_st = min([len(tab[2]) for tab in slovar_deklarantov.values() if 'any' in tab])
        while mno_pod0:
            naslednji_kljuc = next(kljuc)
            if len(slovar_deklarantov[naslednji_kljuc][2]) < int(slovar_deklarantov[naslednji_kljuc][1]) and slovar_deklarantov[naslednji_kljuc][0] != 'formal' and len(slovar_deklarantov[naslednji_kljuc][2]) == najmanjsa_st:
                slovar_deklarantov[naslednji_kljuc][2].append(mno_pod0.pop())
                najmanjsa_st = min([len(tab[2]) for tab in slovar_deklarantov.values() if 'any' in tab])
        # preveri veckrat ce je pogoj pravi!!! za najmanjsa_st
        self.dist_okno = prikaziDistribucijo(slovar_deklarantov)
        self.dist_okno.show()
    
    def shrani_kot_XSLX(self):
        datoteka = QFileDialog.getSaveFileName(self, 'Shrani kot', os.getcwd(), 'Vrsta datoteke (*.xlsx)')
        if datoteka[0] != '':
            workbook = xlsxwriter.Workbook(datoteka[0])
            izmena = 'DOP' if time.gmtime()[3] < 12 else 'POP'
            worksheet = workbook.add_worksheet(str(time.gmtime()[2])+'.'+str(time.gmtime()[1])+' '+izmena)
            
            format_text = workbook.add_format({
                'border': 1,
                'font_size': 11,
                'align': 'center',
                'valign': 'vcenter',
                'text_wrap': True
            })
            format_celo_st = workbook.add_format({
                'border': 1,
                'font_size': 11,
                'align': 'center',
                'valign': 'vcenter',
                'num_format': '0',
                'text_wrap': True
            })
            format_st = workbook.add_format({
                'font_size': 11,
                'align': 'center',
                'valign': 'vcenter',
                'num_format': '0',
                'text_wrap': True
            })
            st = 0
            slovar_deklarantov2 = {k: v[2] for k, v in slovar_deklarantov.items()} | {k: v[2] for k, v in slovar_deklarantov_odstopi.items()}
            for k,v in slovar_deklarantov2.items():
                worksheet.write(0, st, k, format_text)
                worksheet.write(1, st, len(v), format_celo_st)
                for vrstica, awb in enumerate(v, 3):
                    worksheet.write(vrstica, st, int(awb), format_st)
                st += 1
            
            worksheet.autofit()
            # Zapremo zvezek
            workbook.close()
        
distribucija = QApplication([])

glavno_okno = mojaAplikacija()

glavno_okno.show()

distribucija.exec()