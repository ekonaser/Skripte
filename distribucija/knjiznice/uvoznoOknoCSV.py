# uvozno okno za csv podatke

from globalne_spremenljivke import slovar_CSV_podatkov

from PyQt6.QtWidgets import (
    QWidget, QPushButton, QVBoxLayout, QScrollArea, QCheckBox, QMessageBox
)

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