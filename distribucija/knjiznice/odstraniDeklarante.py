import os
from funkcije import preberi_deklarante
from PyQt6.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, QMainWindow,
    QLabel, QFileDialog, QMessageBox, QTableWidget, QTableWidgetItem, QScrollArea, QHBoxLayout, QCheckBox
)

class odstraniDeklarante(QWidget):
    
    def __init__(self, tab):
        super().__init__()
        self.setWindowTitle("Trenutni deklaranti")
        self.setGeometry(300, 300, 300, 300)
        
        postavitev = QVBoxLayout()
        
        self.izberi_vse = QCheckBox("Vsi")
        self.izberi_vse.stateChanged.connect(self._izberi_vse)
        
        postavitev.addWidget(self.izberi_vse)
        
        self.slovar = dict(tab)
        self.potr_polja = []
        for par in tab:
            potrditveno_polje = QCheckBox(par[0])
            self.potr_polja.append(potrditveno_polje)
            postavitev.addWidget(potrditveno_polje)
        
        self.odstrani = QPushButton("Odstrani")
        self.odstrani.clicked.connect(self._odstrani)
        
        self.zapri = QPushButton("Zapri")
        self.zapri.clicked.connect(self._zapri)
              
        postavitev.addWidget(self.odstrani)
        postavitev.addWidget(self.zapri)
        
        self.setLayout(postavitev)
    
    def _odstrani(self):
        with open(os.getcwd() + '\\viri\\deklaranti.txt', 'w',encoding='utf-8') as dat:
            for polje in self.potr_polja:
                if polje.isChecked():
                    self.slovar.pop(polje.text())
            for k,v in self.slovar.items():
                dat.write(k + ';' + v + '\n')
        
        self.izbirno_okno = odstraniDeklarante(preberi_deklarante())
        self.izbirno_okno.show()
        self._zapri()
    
    def _izberi_vse(self):
        """Metoda nam vsa polja označi za izbrana"""
        vse = self.izberi_vse.isChecked()
        if vse:
            for polje in self.potr_polja:
                polje.setChecked(True)
        else:
            for polje in self.potr_polja:
                polje.setChecked(False)
                
    def _zapri(self):
        self.close()