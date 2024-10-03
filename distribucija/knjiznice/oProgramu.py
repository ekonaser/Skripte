import os

from PyQt6.QtWidgets import QWidget, QVBoxLayout, QLabel, QGridLayout, QGroupBox
from PyQt6.QtCore import Qt

class oProgramu(QWidget):
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("O programu")
        
        postavitev = QGridLayout()
        
        besedilo1 = QLabel("FedEx Distribucija")
        besedilo2 = QLabel("Naredil: ")
        besedilo3 = QLabel("Naser Ogrešević")
        besedilo4 = QLabel("e-mail: ")
        email = QLabel('<a href="mailto:naser.ogresevic.OSV@fedex.com">naser.ogresevic.OSV@fedex.com</a>')
        email.setOpenExternalLinks(True)
        podokno = QGroupBox("Namen")
        podokno.setFixedSize(230, 120)
        postavitev_podokno = QVBoxLayout()
        besedilo5 = QLabel("Program je namenjen razdeljevanju\ndeklaracij deklarantom. Program\nuporablja transponiranje za avtomatsko\ndetekcijo stolpičev v csv/txt vhodni\ndatoteki.")
        besedilo5.setAlignment(Qt.AlignmentFlag.AlignJustify)
        postavitev.addWidget(besedilo1, 0, 0, 1, 2)
        postavitev.addWidget(besedilo2, 1, 0)
        postavitev.addWidget(besedilo3, 1, 1)
        postavitev.addWidget(besedilo4, 2, 0)
        postavitev.addWidget(email, 2, 1)
        
        postavitev_podokno.addWidget(besedilo5)
        podokno.setLayout(postavitev_podokno)
        postavitev.addWidget(podokno, 3, 0, 4, 2)
        
        self.setLayout(postavitev)