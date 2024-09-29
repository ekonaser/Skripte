# razred za dodajanje deklarantov
import os
from funkcije import preberi_deklarante

from PyQt6.QtWidgets import QWidget, QPushButton, QVBoxLayout, QCheckBox, QMessageBox, QLineEdit, QComboBox, QSpinBox

class dodajDeklaranta(QWidget):
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Dodaj deklarante")
        self.setGeometry(200, 200, 200, 200)
        
        self.tab = []
        
        postavitev = QVBoxLayout()
        
        tab_deklarantov = preberi_deklarante()
        
        self.vhodno_niz_ime = QLineEdit()
        self.vhodno_niz_ime.setPlaceholderText("Vnesite ime deklaranta")
        
        self.jezicni_menu = QComboBox()
        self.jezicni_menu.addItems(["karkoli", "odstopi", "formal"])
        
        self.vhodna_st = QLineEdit()
        self.vhodna_st.setPlaceholderText("Največje število deklaracij")
        
        self.dodaj_gumb = QPushButton("Dodaj")
        self.dodaj_gumb.clicked.connect(self.dodaj_deklaranta)
        
        self.zapri_gumb = QPushButton("Zapri")
        self.zapri_gumb.clicked.connect(self._zapri)
        
        postavitev.addWidget(self.vhodno_niz_ime)
        postavitev.addWidget(self.jezicni_menu)
        postavitev.addWidget(self.vhodna_st)
        postavitev.addWidget(self.dodaj_gumb)
        postavitev.addWidget(self.zapri_gumb)
        
        self.setLayout(postavitev)
        
    def dodaj_deklaranta(self):
        with open(os.getcwd() + '\\viri\\deklaranti.txt', 'a', encoding='utf-8') as dat:
            niz = self.vhodno_niz_ime.text() + ';' + self.jezicni_menu.currentText() + '-' + self.vhodna_st.text()
            dat.write(niz + '\n')
            self.tab.append(niz)
        self.vhodno_niz_ime.clear()
        self.jezicni_menu.setCurrentIndex(0)
        self.vhodna_st.clear()
    
    def _zapri(self):
        self.close()