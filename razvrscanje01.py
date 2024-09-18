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

def formati(workbook) -> dict:
    """Funkcija nam doda različne formate v excelov zvezek, ter vrne slovar teh formatov"""
    # Razlicni formati za celice
    slo_formatov ={}
    slo_formatov['merge_format'] = workbook.add_format({
        'bold': 1,
        'border': 1,
        'font_size': 8,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#BF00BF',
        'text_wrap': True
    })
    
    slo_formatov['merge_format2'] = workbook.add_format({
        'bold': 1,
        'border': 1,
        'font_size': 8,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#FFA500',
        'text_wrap': True
    })

    slo_formatov['prva_format'] = workbook.add_format({
        'border': 1,
        'font_size': 8,
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': True
    })

    slo_formatov['druga_format'] = workbook.add_format({
        'border': 1,
        'font_size': 8,
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': True
    })
    
    slo_formatov['special_format'] = workbook.add_format({
        'border': 1,
        'font_size': 8,
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': True,
        'num_format': '000'
    })

    slo_formatov['celostevilo_format'] = workbook.add_format({
        'align': 'right',
        'num_format': '0'
    })

    slo_formatov['decst_format'] = workbook.add_format({
        'align': 'right',
        'num_format': '0.00'
    })

    slo_formatov['celostevilo_format_rod'] = workbook.add_format({
        'align': 'right',
        'num_format': '0',
        'bg_color': '#FFFF00'
    })

    slo_formatov['decst_format_rod'] = workbook.add_format({
        'align': 'right',
        'num_format': '0.00',
        'bg_color': '#FFFF00'
    })

    slo_formatov['oranzna'] = workbook.add_format({'bg_color': '#FFA500', 'border': 1,'text_wrap': True,'font_size': 8,'align': 'center','valign': 'vcenter',})
    slo_formatov['oranzna_spec'] = workbook.add_format({'bg_color': '#FFA500', 'border': 1,'text_wrap': True,'font_size': 8,'align': 'center','valign': 'vcenter','num_format': '000'})
    slo_formatov['zelena'] = workbook.add_format({'bg_color': '#00CD00', 'border': 1,'text_wrap': True,'font_size': 8,'align': 'center','valign': 'vcenter',})
    slo_formatov['rod'] = workbook.add_format({'bg_color': '#FFFF00'})
    
    return slo_formatov
#########################################################################################
def razvrscanje_isb_out() -> None:
    isb_out = open('isb_out.csv','w',encoding='utf-8')
    with open('obracun_precisceno.csv','r',encoding='utf-8') as dat:
        for vr in dat:
            break
        for vr in dat:
            sez = vr.replace('\n','').split(sep=';')
            if len(sez[0]) == 11 and sez[4] == 'DDP':
                isb_out.write(sez[0]+';'+sez[1]+';'+sez[2]+';'+sez[3]+';'+sez[4]+';'+'None'+';'+'None'+';'+'None'+';'+'None'+';'+sez[8]+';'+sez[9]+';'+sez[10]+';'+sez[11]+';'+sez[12]+';'+sez[13]+'\n')
    isb_out.close()

def razvrscanje_private_individuals() -> None:
    private_individuals = open('private_individuals.csv','w',encoding='utf-8')
    with open('obracun_precisceno.csv','r',encoding='utf-8') as dat:
        for vr in dat:
            break
        for vr in dat:
            sez = vr.replace('\n','').split(sep=';')
            if len(sez[0]) == 11 and sez[6] == 'SIA5555555':
                private_individuals.write(sez[0]+';'+sez[1]+';'+sez[2]+';'+sez[3]+';'+sez[14]+';'+sez[15]+';'+sez[16]+';'+'24914'+';'+sez[4]+';'+sez[6]+';'+sez[7]+';'+sez[8]+';'+sez[9]+';'+sez[10]+';'+sez[11]+';'+sez[12]+';'+sez[13]+'\n')
    private_individuals.close()

def razvrscanje_navadna() -> None:
    slo_strank = {}
    navadna = open('navadna.csv','w',encoding='utf-8')
    with open('navadna_acc.csv', 'r', encoding='utf-8') as dat:
        for vr in dat:
            sez = vr.replace('\n','').split(sep=';')
            slo_strank[sez[1]] = sez[0]
    with open('obracun_precisceno.csv','r',encoding='utf-8') as dat:
        for vr in dat:
            break
        for vr in dat:
            sez = vr.replace('\n','').split(sep=';')
            acc_n = 'None'
            if sez[6] in slo_strank:
                acc_n = slo_strank[sez[6]]
            if len(sez[0]) == 11 and sez[4] in {'CPT', 'ZAČASNI UVOZ 42', 'VRAČILA 42', 'SANITARC/DRUG VLADNI ORGAN', 'TRANZIT'} and sez[6] != 'SIA5555555':
                navadna.write(sez[0]+';'+sez[1]+';'+sez[2]+';'+sez[3]+';'+acc_n+';'+sez[4]+';'+sez[6]+';'+sez[7]+';'+sez[8]+';'+sez[9]+';'+sez[10]+';'+sez[11]+';'+sez[12]+';'+sez[13]+'\n')
    navadna.close()


def razvrscanje_fedex() -> None:
    fedex = open('fedex.csv','w',encoding='utf-8')
    with open('obracun_precisceno.csv','r',encoding='utf-8') as dat:
        for vr in dat:
            break
        for vr in dat:
            sez = vr.replace('\n','').split(sep=';')
            if len(sez[0]) == 12:
                fedex.write(sez[0]+';'+sez[1]+';'+sez[2]+';'+sez[3]+';'+sez[14]+';'+sez[15]+';'+sez[16]+';'+sez[4]+';'+sez[5]+';'+sez[6]+';'+sez[7]+';'+sez[8]+';'+sez[9]+';'+sez[10]+';'+sez[11]+';'+sez[12]+';'+sez[13]+'\n')
    fedex.close()

def razvrscanje() -> None:
    """Funkcija nam ustvari fedex.csv datoteko na katero so zapisani razvrsceni podatki za fedex obracun.
        Funkcija to naredi s pomocjo posredne datoteke obracun_precisceno.csv katero na koncu izbrisemo.
    """
    
    sez = os.listdir(os.getcwd())
    for dat in sez:
        if 'OBRAČUN.csv' in dat:
            with open(dat,'r',newline='',encoding='utf-8') as infile:
                reader = csv.reader(infile, delimiter=';')
                rows = [zamenjaj_z_none(row) for row in reader]
            
            with open('obracun_precisceno.csv','w',newline='',encoding='utf-8') as outfile:
                writer = csv.writer(outfile, delimiter=';')
                writer.writerows(rows)
    # izbrisemo posredno datoteko, katera vsebuje vse podatke prvotne datoteke
    # le da je na praznih celicah zdaj vrednost 'None'
    # os.remove(os.getcwd() + '\\' + 'obracun_precisceno.csv')
#########################################################################################
def fedex() -> None:
    """Funkcija nam fedex.csv razvrsti v datoteko VAT AND DUTIES [danasnji datum] FEDEX.xlsx
    """
    # Ustvarimo excelov zvezek in prvi zavihek
    workbook = xlsxwriter.Workbook('VAT AND DUTIES ' + _danasnji_datum2[0] + '.' + _danasnji_datum2[1] + '.' + _danasnji_datum2[2] +  ' FEDEX.xlsx')
    worksheet = workbook.add_worksheet('FDX')

    slo_format = formati(workbook)

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
    worksheet.merge_range('A3:J3', '\nFEDEX\n', slo_format['merge_format'])

    for y in range(10,27):
        worksheet.write(2,y,prva_vrstica.pop(0),slo_format['prva_format'])

    for y in range(0,11):
        worksheet.write(3,y,druga_vrstica[y],slo_format['druga_format'])
    
    for y in range(11,26):
        worksheet.write(3,y,int(druga_vrstica[y]),slo_format['special_format'])
    
    worksheet.write(3,26,druga_vrstica[26],slo_format['druga_format'])

    # oranzne celice
    worksheet.write(2,25,'VAT on\nAncillary\nFees (@\n22%)',slo_format['oranzna'])
    worksheet.write(3,25,int('940'),slo_format['oranzna_spec'])

    # zelene celice
    worksheet.write(2,26,'Total\nCollected\nfrom\nCustomer',slo_format['zelena'])
    worksheet.write(3,26,'---',slo_format['zelena'])

    worksheet.autofilter('A4:AA4')

    # vrstice steje tako kot python se pravi 0 = 1
    worksheet.set_row(2, 45)

    with open('fedex.csv', 'r',encoding='utf-8') as csvfile:
        csvreader = csv.reader(csvfile, delimiter=';')
        for row_num, row_data in enumerate(csvreader,4):
            # niz 'None' zamenjamo za prazen niz
            row_data = [niz if niz != 'None' else '' for niz in row_data]
            if 'ROD' in row_data[2]:
                for col_num, col_data in enumerate(row_data):
                    # DATUM je na 5 indeksu
                    if col_data.isnumeric() and col_num == 5 or col_num == 0:
                        worksheet.write(row_num, col_num, int(col_data), slo_format['celostevilo_format_rod'])
                        continue
                    if col_data.isnumeric():
                        worksheet.write(row_num, col_num, int(col_data), slo_format['decst_format_rod'])
                    elif isfloat(col_data):
                        worksheet.write(row_num, col_num, float(col_data.replace(',','.')), slo_format['decst_format_rod'])
                    else:
                        worksheet.write(row_num, col_num, col_data, slo_format['rod'])
                for col_num in range(17,27):
                    if col_num == 25:
                        worksheet.write(row_num, col_num, float(3.30), slo_format['decst_format_rod'])
                    else:
                        worksheet.write(row_num, col_num, '', slo_format['rod'])
                    
            else:
                for col_num, col_data in enumerate(row_data):
                    # DATUM je na 5 indeksu
                    if col_data.isnumeric() and col_num == 5 or col_num == 0:
                        worksheet.write(row_num, col_num, int(col_data), slo_format['celostevilo_format'])
                        continue
                    if col_data.isnumeric():
                        worksheet.write(row_num, col_num, int(col_data), slo_format['decst_format'])
                    elif isfloat(col_data):
                        worksheet.write(row_num, col_num, float(col_data.replace(',','.')), slo_format['decst_format'])
                    else:
                        worksheet.write(row_num, col_num, col_data)
    worksheet.autofit()

    # Zapremo zvezek
    workbook.close()

#########################################################################################
def isb_out() -> None:
    """Funkcija nam ustvari datoteko VAT AND DUTIES [danasnji datum] ISB OUT.xlsx
    """
    # Ustvarimo excelov zvezek in prvi zavihek
    workbook = xlsxwriter.Workbook('VAT AND DUTIES ' + _danasnji_datum2[0] + '.' + _danasnji_datum2[1] + '.' + _danasnji_datum2[2] +  ' ISB OUT.xlsx')
    worksheet = workbook.add_worksheet('TNT')

    slo_format = formati(workbook)
    
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
                     'Accoun no.',
                     'Customer name',
                     'product',
                     'division',
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
    worksheet.merge_range('A3:H3', '\nTNT\n', slo_format['merge_format2'])

    for y in range(8,26):
        worksheet.write(2,y,prva_vrstica.pop(0),slo_format['prva_format'])

    for y in range(0,26):
        worksheet.write(3,y,druga_vrstica[y],slo_format['druga_format'])

    worksheet.autofilter('A4:Z4')

    # vrstice steje tako kot python se pravi 0 = 1
    worksheet.set_row(2, 100)
    
    with open('isb_out.csv', 'r',encoding='utf-8') as csvfile:
        csvreader = csv.reader(csvfile, delimiter=';')
        for row_num, row_data in enumerate(csvreader,4):
            # niz 'None' zamenjamo za prazen niz
            row_data = [niz if niz != 'None' else '' for niz in row_data]
            if 'ROD' in row_data[2]:
                for col_num, col_data in enumerate(row_data):
                    if col_data.isnumeric() and col_num == 0:
                        worksheet.write(row_num, col_num, int(col_data), slo_format['celostevilo_format_rod'])
                        continue
                    if col_data.isnumeric():
                        worksheet.write(row_num, col_num, int(col_data), slo_format['decst_format_rod'])
                    elif isfloat(col_data):
                        worksheet.write(row_num, col_num, float(col_data.replace(',','.')), slo_format['decst_format_rod'])
                    else:
                        worksheet.write(row_num, col_num, col_data, slo_format['rod'])
                for col_num in range(15,26):
                    if col_num == 24:
                        worksheet.write(row_num, col_num, float(3.30), slo_format['decst_format_rod'])
                    else:
                        worksheet.write(row_num, col_num, '', slo_format['rod'])
            else:
                for col_num, col_data in enumerate(row_data):
                    if col_data.isnumeric() and col_num == 0:
                        worksheet.write(row_num, col_num, int(col_data), slo_format['celostevilo_format'])
                        continue
                    if col_data.isnumeric():
                        worksheet.write(row_num, col_num, int(col_data), slo_format['decst_format'])
                    elif isfloat(col_data):
                        worksheet.write(row_num, col_num, float(col_data.replace(',','.')), slo_format['decst_format'])
                    else:
                        worksheet.write(row_num, col_num, col_data)
    
    worksheet.autofit()

    # zapremo zvezek
    workbook.close()
#########################################################################################
def navadna() -> None:
    """Funkcija nam ustvari datoteko VAT AND DUTIES [danasnji datum] NAVADNA.xlsx
    """
    # Ustvarimo excelov zvezek in prvi zavihek
    workbook = xlsxwriter.Workbook('VAT AND DUTIES ' + _danasnji_datum2[0] + '.' + _danasnji_datum2[1] + '.' + _danasnji_datum2[2] +  ' NAVADNA.xlsx')
    worksheet = workbook.add_worksheet('TNT')
    
    slo_format = formati(workbook)
    
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
    worksheet.merge_range('A3:G3', '\nTNT\n', slo_format['merge_format2'])

    for y in range(7,25):
        worksheet.write(2,y,prva_vrstica.pop(0),slo_format['prva_format'])

    for y in range(0,25):
        worksheet.write(3,y,druga_vrstica[y],slo_format['druga_format'])

    worksheet.autofilter('A4:Y4')
    
    # vrstice steje tako kot python se pravi 0 = 1
    worksheet.set_row(2, 100)
    
    with open('navadna.csv', 'r',encoding='utf-8') as csvfile:
        csvreader = csv.reader(csvfile, delimiter=';')
        for row_num, row_data in enumerate(csvreader,4):
            # niz 'None' zamenjamo za prazen niz
            row_data = [niz if niz != 'None' else '' for niz in row_data]
            if 'ROD' in row_data[2]:
                for col_num, col_data in enumerate(row_data):
                    if col_data.isnumeric() and col_num == 0:
                        worksheet.write(row_num, col_num, int(col_data), slo_format['celostevilo_format_rod'])
                        continue
                    if col_data.isnumeric():
                        worksheet.write(row_num, col_num, int(col_data), slo_format['decst_format_rod'])
                    elif isfloat(col_data):
                        worksheet.write(row_num, col_num, float(col_data.replace(',','.')), slo_format['decst_format_rod'])
                    else:
                        worksheet.write(row_num, col_num, col_data, slo_format['rod'])
                for col_num in range(14,25):
                    if col_num == 23:
                        worksheet.write(row_num, col_num, float(3.30), slo_format['decst_format_rod'])
                    else:
                        worksheet.write(row_num, col_num, '', slo_format['rod'])
                
            else:
                for col_num, col_data in enumerate(row_data):
                    if col_data.isnumeric() and col_num == 0 or col_num == 4:
                        worksheet.write(row_num, col_num, int(col_data), slo_format['celostevilo_format'])
                        continue
                    if col_data.isnumeric():
                        worksheet.write(row_num, col_num, int(col_data), slo_format['decst_format'])
                    elif isfloat(col_data):
                        worksheet.write(row_num, col_num, float(col_data.replace(',','.')), slo_format['decst_format'])
                    else:
                        worksheet.write(row_num, col_num, col_data)
    
    worksheet.autofit()

    # zapremo zvezek
    workbook.close()
#########################################################################################
def private_individual() -> None:
    """Funkcija nam ustvari datoteko VAT AND DUTIES [danasnji datum] PRIVATE INDIVIDUAL.xlsx
    """
    # Ustvarimo excelov zvezek in prvi zavihek
    workbook = xlsxwriter.Workbook('VAT AND DUTIES ' + _danasnji_datum2[0] + '.' + _danasnji_datum2[1] + '.' + _danasnji_datum2[2] +  ' PRIVATE INDIVIDUAL.xlsx')
    worksheet = workbook.add_worksheet('TNT')
    
    slo_format = formati(workbook)
    
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
    worksheet.merge_range('A3:J3', '\nTNT\n', slo_format['merge_format2'])

    for y in range(10,28):
        worksheet.write(2,y,prva_vrstica.pop(0),slo_format['prva_format'])

    for y in range(0,28):
        worksheet.write(3,y,druga_vrstica[y],slo_format['druga_format'])

    worksheet.autofilter('A4:AB4')

    # vrstice steje tako kot python se pravi 0 = 1
    worksheet.set_row(2, 100)
    
    with open('private_individuals.csv', 'r',encoding='utf-8') as csvfile:
        csvreader = csv.reader(csvfile, delimiter=';')
        for row_num, row_data in enumerate(csvreader,4):
            # niz 'None' zamenjamo za prazen niz
            row_data = [niz if niz != 'None' else '' for niz in row_data]
            if 'ROD' in row_data[2]:
                for col_num, col_data in enumerate(row_data):
                    # DATUM je na 5 indeksu
                    if col_data.isnumeric() and col_num == 7 or col_num == 5 or col_num == 0:
                        worksheet.write(row_num, col_num, int(col_data), slo_format['celostevilo_format_rod'])
                        continue
                    if col_data.isnumeric():
                        worksheet.write(row_num, col_num, int(col_data), slo_format['decst_format_rod'])
                    elif isfloat(col_data):
                        worksheet.write(row_num, col_num, float(col_data.replace(',','.')), slo_format['decst_format_rod'])
                    else:
                        worksheet.write(row_num, col_num, col_data, slo_format['rod'])
                for col_num in range(17,28):
                    if col_num == 26:
                        worksheet.write(row_num, col_num, float(3.30), slo_format['decst_format_rod'])
                    else:
                        worksheet.write(row_num, col_num, '', slo_format['rod'])
                    
            else:
                for col_num, col_data in enumerate(row_data):
                    # DATUM je na 5 indeksu
                    if col_data.isnumeric() and col_num == 7 or col_num == 5 or col_num == 0:
                        worksheet.write(row_num, col_num, int(col_data), slo_format['celostevilo_format'])
                        continue
                    if col_data.isnumeric():
                        worksheet.write(row_num, col_num, int(col_data), slo_format['decst_format'])
                    elif isfloat(col_data):
                        worksheet.write(row_num, col_num, float(col_data.replace(',','.')), slo_format['decst_format'])
                    else:
                        worksheet.write(row_num, col_num, col_data)
    
    worksheet.autofit()

    # zapremo zvezek
    workbook.close()
#########################################################################################

if __name__ == '__main__':
    _danasnji_datum2 = [f"{time.gmtime()[2]:02d}", f"{time.gmtime()[1]:02d}", f"{time.gmtime()[0]}"]
    razvrscanje()
    razvrscanje_fedex()
    razvrscanje_navadna()
    razvrscanje_private_individuals()
    razvrscanje_isb_out()
    fedex()
    isb_out()
    navadna()
    private_individual()
    # na koncu izbrisemo fedex.csv datoteko ker smo jo ze razvrstili
    # os.remove(os.getcwd() + '\\' + 'fedex.csv')