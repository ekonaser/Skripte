from globalne_spremenljivke import slovar_deklarantov, slovar_deklarantov_odstopi

from PyQt6.QtWidgets import QWidget, QPushButton, QVBoxLayout, QCheckBox, QMessageBox

class izberiDeklarante(QWidget):
    
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