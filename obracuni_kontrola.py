import sys
import os
import pandas as pd
import xml.etree.ElementTree as ET
from PyQt6.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, QMainWindow,
    QLabel, QFileDialog, QMessageBox, QTableWidget, QTableWidgetItem, QScrollArea, QHBoxLayout
)
from PyQt6.QtGui import QAction, QPixmap, QGuiApplication

slovar_FDX = {}
slovar_TNT = {}

class GlavnoOkno(QMainWindow):
    
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle('Obračuni')
        self.setGeometry(500, 500, 500, 500)
        
        postavitev = QVBoxLayout()
        
        datoteka = self.menuBar()
        meni1 = datoteka.addMenu("&Datoteka")
        
        uvozi_xls = QAction("Uvozi xls", self)
        izhod = QAction("Izhod", self)
        
        meni1.addAction(uvozi_xls)
        meni1.addSeparator()
        meni1.addAction(izhod)
        
        # menu ukazi
        uvozi_xls.triggered.connect(self._uvozi_xls)
        izhod.triggered.connect(self._izhod)
    
    def _uvozi_xls(self):
        datoteka = QFileDialog.getOpenFileName(self, 'Odpri xls datoteko', os.getcwd(), 'Vrsta datoteke (*.xls)')
        if datoteka[0] != '':
            self.pretvori_v_csv(datoteka[0])
    
    def _izhod(self):
        self.close()
    
    def pretvori_v_csv(self, pot):
        """Metoda nam specificno pretvori 'xml' datoteko zamaskirano kot xls v csv"""
        tree = ET.parse(pot)
        root = tree.getroot()

        # Namespace dictionary
        ns = {
            'ss': 'urn:schemas-microsoft-com:office:spreadsheet',
            'x': 'urn:schemas-microsoft-com:office:excel'
        }

        # Extract data
        data = []
        for row in root.findall('.//ss:Row', ns):
            row_data = []
            for cell in row.findall('ss:Cell', ns):
                # Check for merged cells
                merge_across = int(cell.attrib.get('ss:MergeAcross', 0))
                merge_down = int(cell.attrib.get('ss:MergeDown', 0))
                
                data_element = cell.find('ss:Data', ns)
                cell_text = data_element.text.replace('\n', ' ') if data_element is not None and data_element.text is not None else ''
                
                row_data.append(cell_text)
                
                # Handle merge across
                for _ in range(merge_across):
                    row_data.append('')
            
            data.append(row_data)
        
        # Convert to DataFrame
        df = pd.DataFrame(data)

        # Save to CSV with UTF-8 encoding
        df.to_csv('viri\\datoteka.csv', index=False, header=None, encoding='utf-8', sep=';')
    
    def razvrsti(self):
        return 0

aplikacija = QApplication([])
glavno_okno = GlavnoOkno()
glavno_okno.show()
aplikacija.exec()