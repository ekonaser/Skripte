import xlsxwriter
import csv
import time
import os

#########################################################################################
def isfloat(st):
    """Funkcija preveri ali je dana stevilka float"""
    try:
        float(st.replace(',','.'))
        return True
    except ValueError:
        return False

def zamenjaj_z_none(row):
    return ['None' if not value else value for value in row]
#########################################################################################
def fedex_razvrscanje() -> None:
    """Funkcija nam ustvari fedex.csv datoteko na katero so zapisani razvrsceni podatki za fedex obracun.
        Funkcija to naredi s pomocjo posredne datoteke obracun_precisceno.csv katero na koncu izbrisemo.
    """
    fedex = open('fedex.csv','w',encoding='utf-8')

    sez = os.listdir(os.getcwd())
    for dat in sez:
        if 'OBRAČUN.csv' in dat:
            with open(dat,'r',newline='',encoding='utf-8') as infile:
                reader = csv.reader(infile, delimiter=';')
                rows = [zamenjaj_z_none(row) for row in reader]
            
            with open('obracun_precisceno.csv','w',newline='',encoding='utf-8') as outfile:
                writer = csv.writer(outfile, delimiter=';')
                writer.writerows(rows)
            
            with open('obracun_precisceno.csv','r',encoding='utf-8') as dt:
                for vr in dt:
                    break
                for vr in dt:
                    sez = vr.replace('\n','').split(sep=';')
                    if len(sez[0]) == 12:
                        fedex.write(sez[0]+';'+sez[1]+';'+sez[2]+';'+sez[3]+';'+sez[14]+';'+sez[15]+';'+sez[16]+';'+sez[4]+';'+sez[5]+';'+sez[6]+';'+sez[7]+';'+sez[8]+';'+sez[9]+';'+sez[10]+';'+sez[11]+';'+sez[12]+';'+sez[13]+'\n')
    # izbrisemo posredno datoteko, katera vsebuje vse podatke prvotne datoteke
    # le da je na praznih celicah zdaj vrednost 'None'
    os.remove(os.getcwd() + '\\' + 'obracun_precisceno.csv')
    fedex.close()
#########################################################################################
def fedex() -> None:
    """Funkcija nam fedex.csv razvrsti v datoteko VAT AND DUTIES [danasnji datum] FEDEX.xlsx
    """
    # Ustvarimo excelov zvezek in prvi zavihek
    workbook = xlsxwriter.Workbook('VAT AND DUTIES ' + _danasnji_datum2[0] + '.' + _danasnji_datum2[1] + '.' + _danasnji_datum2[2] +  ' FEDEX.xlsx')
    worksheet = workbook.add_worksheet('FDX')

    # Razlicni formati za celice
    merge_format = workbook.add_format({
        'bold': 1,
        'border': 1,
        'font_size': 8,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#BF00BF',
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
    
    special_format = workbook.add_format({
        'border': 1,
        'font_size': 8,
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': True,
        'num_format': '000'
    })

    celostevilo_format = workbook.add_format({
        'align': 'right',
        'num_format': '0'
    })

    decst_format = workbook.add_format({
        'align': 'right',
        'num_format': '0.00'
    })

    celostevilo_format_rod = workbook.add_format({
        'align': 'right',
        'num_format': '0',
        'bg_color': '#FFFF00'
    })

    decst_format_rod = workbook.add_format({
        'align': 'right',
        'num_format': '0.00',
        'bg_color': '#FFFF00'
    })

    oranzna = workbook.add_format({'bg_color': '#FFA500', 'border': 1,'text_wrap': True,'font_size': 8,'align': 'center','valign': 'vcenter',})
    oranzna_spec = workbook.add_format({'bg_color': '#FFA500', 'border': 1,'text_wrap': True,'font_size': 8,'align': 'center','valign': 'vcenter','num_format': '000'})
    zelena = workbook.add_format({'bg_color': '#00CD00', 'border': 1,'text_wrap': True,'font_size': 8,'align': 'center','valign': 'vcenter',})
    rod = workbook.add_format({'bg_color': '#FFFF00'})

    prva_vrstica = ['Description',
                    'VAT',
                    'Duties',
                    'Other\nGovernment\nAgency',
                    'Additional\nLine\nItems',
                    'Clearance\ntransfer',
                    'Disbursement\nFee',
                    'Pre-Payment\n(Direct\nPayment\nProcessing)',
                    'Returned\nGoods',
                    'Temporary\nImport',
                    'Post Entry\nAdjustment',
                    'In-Bond\nTransit',
                    'Storage',
                    'Customized Service',
                    'Brokerage Fee',
                    'VAT on\nAncillary\nFees (@\n22%)',
                    'Total\nCollected\nfrom\nCustomer']

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
                    '059',
                    '052',
                    '425',
                    '350',
                    '422',
                    '074',
                    '429',
                    '414',
                    '415',
                    '411',
                    '424',
                    '421',
                    '404',
                    '432',
                    '940',
                    '---']

    # zdruzimo celice od A3 do J3
    worksheet.merge_range('A3:J3', '\nFEDEX\n', merge_format)

    for y in range(10,27):
        worksheet.write(2,y,prva_vrstica.pop(0),prva_format)

    for y in range(0,11):
        worksheet.write(3,y,druga_vrstica[y],druga_format)
    
    for y in range(11,26):
        worksheet.write(3,y,int(druga_vrstica[y]),special_format)
    
    worksheet.write(3,26,druga_vrstica[26],druga_format)

    # oranzne celice
    worksheet.write(2,25,'VAT on\nAncillary\nFees (@\n22%)',oranzna)
    worksheet.write(3,25,int('940'),oranzna_spec)

    # zelene celice
    worksheet.write(2,26,'Total\nCollected\nfrom\nCustomer',zelena)
    worksheet.write(3,26,'---',zelena)

    worksheet.autofilter('A4:AA4')

    # vrstice steje tako kot python se pravi 0 = 1
    worksheet.set_row(2, 45)

    with open('fedex.csv', 'r',encoding='utf-8') as csvfile:
        csvreader = csv.reader(csvfile, delimiter=';')
        for row_num, row_data in enumerate(csvreader,4):
            row_data = [niz if niz != 'None' else '' for niz in row_data]
            if 'ROD' in row_data[2]:
                for col_num, col_data in enumerate(row_data):
                    if col_data.isnumeric():
                        worksheet.write(row_num, col_num, int(col_data), celostevilo_format_rod)
                    elif isfloat(col_data):
                        worksheet.write(row_num, col_num, float(col_data.replace(',','.')), decst_format_rod)
                    else:
                        worksheet.write(row_num, col_num, col_data, rod)
                for col_num in range(17,27):
                    if col_num == 25:
                        worksheet.write(row_num, col_num, float(3.30), decst_format_rod)
                    else:
                        worksheet.write(row_num, col_num, '', rod)
                    
            else:
                for col_num, col_data in enumerate(row_data):
                    if col_data.isnumeric():
                        worksheet.write(row_num, col_num, int(col_data), celostevilo_format)
                    elif isfloat(col_data):
                        worksheet.write(row_num, col_num, float(col_data.replace(',','.')), decst_format)
                    else:
                        worksheet.write(row_num, col_num, col_data)

    worksheet.autofit()

    # Close the workbook
    workbook.close()

if __name__ == '__main__':
    _danasnji_datum2 = [f"{time.gmtime()[2]:02d}", f"{time.gmtime()[1]:02d}", f"{time.gmtime()[0]}"]
    fedex_razvrscanje()
    fedex()
    # na koncu izbrisemo fedex.csv datoteko ker smo jo ze razvrstili
    os.remove(os.getcwd() + '\\' + 'fedex.csv')