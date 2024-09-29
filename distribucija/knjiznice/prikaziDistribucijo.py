# prikazno okno razdeljene distribucije

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QTableWidget, QTableWidgetItem
)

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