# razred za izbiranje CSV podatkov s pomocjo gumbov dveh gumbov
# ter delitev teh podatkov s pomocjo metode razdeli
from globalne_spremenljivke import mno_fiz, mno_pod, mno_odstopov, mno_odstopi

from PyQt6.QtWidgets import (
    QWidget, QPushButton, QVBoxLayout, QMainWindow, QTableWidget, QTableWidgetItem, QScrollArea, QHBoxLayout
)
from PyQt6.QtCore import Qt

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