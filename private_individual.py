import time
import xlsxwriter
import os

def private_individual() -> None:
    """Funkcija nam ustvari datoteko VAT AND DUTIES [danasnji datum] PRIVATE INDIVIDUAL.xlsx
    """
    # Ustvarimo excelov zvezek in prvi zavihek
    workbook = xlsxwriter.Workbook('VAT AND DUTIES ' + _danasnji_datum2[0] + '.' + _danasnji_datum2[1] + '.' + _danasnji_datum2[2] +  ' PRIVATE INDIVIDUAL.xlsx')
    worksheet = workbook.add_worksheet('TNT')

    # Razlicni formati za celice
    merge_format = workbook.add_format({
        'bold': 1,
        'border': 1,
        'font_size': 8,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#FFA500',
        'text_wrap': True
    })

    prva_format = workbook.add_format({
        'border': 1,
        'font_size': 8,
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': True
    })

    druga_format = workbook.add_format({
        'border': 1,
        'font_size': 8,
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': True
    })

    celostevilo_format = workbook.add_format({
        'align': 'right',
        'num_format': '0'
    })

    decst_format = workbook.add_format({
        'align': 'right',
        'num_format': '0.00'
    })
    
    prva_vrstica = ['Description',
                    'VAT',
                    'Duties',
                    'Other Government\nAgency',
                    'Additional\nLine Items',
                    'Clearance transfer',
                    'Disbursement Fee\nReturned Goods\nPre-Payment\n(Direct Payment Processing)\nReturned Goods\nTemporary Import\nPost Entry Adjustment',
                    'In-Bond Transit',
                    'Storage\nCustomized Service',
                    'Disbursement Fee\nReturned Goods\nPre-Payment (Direct\nPayment Processing)\nReturned Goods\nTemporary Import\nPost Entry\nAdjustment',
                    'Clearance\ntransfer',
                    'Additional Line\nItems',
                    'Storage\nCustomized Service',
                    'In-Bond Transit',
                    'In-Bond Transit',
                    'Storage\nCustomized Service',
                    'Disbursement Fee\nReturned Goods\nPre-Payment (Direct\nPayment Processing)\nReturned Goods\nTemporary Import\nPost Entry Adjustment',
                    'Clearance transfer']

    druga_vrstica = ['Con Note',
                     'Customer reference',
                     'MRN',
                     'Date',
                     'address',
                     'postcode',
                     'town',
                     'Accoun no.',
                     'Customer name',
                     'VAT ID',
                     'Customer name',
                     'VT',
                     'DT',
                     'CC1',
                     'CL0',
                     'CL3',
                     'CL5',
                     'CL6',
                     'HF',
                     'CG1',
                     'CL4',
                     'OCH',
                     'ST1',
                     'TR1',
                     'CL9',
                     'CLA',
                     'CLB',
                     'CLC']

    # zdruzimo celice od A3 do J3
    worksheet.merge_range('A3:J3', '\nTNT\n', merge_format)

    for y in range(10,28):
        worksheet.write(2,y,prva_vrstica.pop(0),prva_format)

    for y in range(0,28):
        worksheet.write(3,y,druga_vrstica[y],druga_format)

    worksheet.autofilter('A4:AB4')

    # vrstice steje tako kot python se pravi 0 = 1
    worksheet.set_row(2, 100)
    
    worksheet.autofit()

    # zapremo zvezek
    workbook.close()

if __name__ == '__main__':
    _danasnji_datum2 = [f"{time.gmtime()[2]:02d}", f"{time.gmtime()[1]:02d}", f"{time.gmtime()[0]}"]
    private_individual()