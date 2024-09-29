# xslx modul

import os
import xlsxwriter

from PyQt6.QtWidgets import QFileDialog

class shraniXLSX:
    
    def shrani_kot_XLSX(self, slovar1, slovar2):
        datoteka = QFileDialog.getSaveFileName(self, 'Shrani kot', os.getcwd(), 'Vrsta datoteke (*.xlsx)')
        if datoteka[0] != '':
            workbook = xlsxwriter.Workbook(datoteka[0])
            worksheet = workbook.add_worksheet()
            
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
            slovar = {k: v[2] for k, v in slovar1.items()} | {k: v[2] for k, v in slovar2.items()}
            for k,v in slovar.items():
                worksheet.write(0, st, k, format_text)
                worksheet.write(1, st, len(v), format_celo_st)
                for vrstica, awb in enumerate(v, 3):
                    worksheet.write(vrstica, st, int(awb), format_st)
                st += 1
            
            worksheet.autofit()
            # Zapremo zvezek
            workbook.close()
    