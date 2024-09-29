# razred ki dokument shrani kot txt
import os
import pprint

from PyQt6.QtWidgets import QFileDialog

class shraniTXT:
    
    def shrani_kot_TXT(self, slovar1, slovar2):
        datoteka = QFileDialog.getSaveFileName(self, 'Shrani kot', os.getcwd(), 'Vrsta datoteke (*.txt)')
        if datoteka[0] != '':
            with open(datoteka[0], 'w') as dat:
                pprint.pprint({k: v[2] for k, v in slovar1.items()} | {k: v[2] for k, v in slovar2.items()}, stream=dat)