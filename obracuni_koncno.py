import sys
import os
import xml.etree.ElementTree as ET

def pretvori_v_csv(pot):
    """Metoda nam specificno pretvori 'xml' datoteko zamaskirano kot xls v csv"""
    tree = ET.parse(pot)
    root = tree.getroot()

    # Namespace dictionary
    ns = {
        'ss': 'urn:schemas-microsoft-com:office:spreadsheet',
        'x': 'urn:schemas-microsoft-com:office:excel'
        }
    
    datoteka_manjkajoci = open('manjkajoci.txt','w',encoding='utf-8')
    
    # CarM Scan podatke dodamo v množico
    mno = set()
    with open('carm_scan.csv','r',encoding='utf-8-sig') as dat:
        for vr in dat:
            awb, mrn = vr.strip().split(';')
            mno.add(awb)
            mno.add(mrn)
    
    slo_acc = {}
    with open('navadna_acc.csv','r',encoding='utf-8-sig') as dat:
        for vr in dat:
            st, davcna, naziv = vr.strip().split(';')
            slo_acc[davcna] = st
            
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
                
        if row_data[2] in mno or row_data[0] in mno:
            if len(row_data[0]) == 11 and row_data[0].isdigit():
                row_data[8] = slo_acc.get(row_data[9], row_data[9])
            data.append(row_data)
            mno.discard(row_data[2])
            mno.discard(row_data[0])
            
        elif len(row_data[0]) >= 18 and row_data[0][-18:] in mno:
            awb, mrn = row_data[0].replace(' ', '').split('-')
            row_data[0] = awb
            row_data[2] = mrn
            if len(awb) == 11 and awb.isdigit():
                row_data[8] = slo_acc.get(row_data[9], row_data[9])
            data.append(row_data)
            mno.discard(mrn)
            mno.discard(awb)
    
    with open('obracun.csv','w',encoding='utf-8') as dat:
        for tab in data:
            dat.write(';'.join(tab) + '\n')
    
    print('\n'.join(mno), file=datoteka_manjkajoci)
    datoteka_manjkajoci.close()

if __name__ == '__main__':
    pretvori_v_csv(os.getcwd() + '\\' + 'tlm_obracun.xls')
