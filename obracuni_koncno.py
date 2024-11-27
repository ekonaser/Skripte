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
    
    # CarM Scan podatke dodamo v množico
    mno = set()
    with open('carm_scan.csv','r',encoding='utf-8') as dat:
        for vr in dat:
            break
        for vr in dat:
            mno.add(vr[:-1])
    
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
            
        if row_data[2] in mno:
            data.append(row_data)
            mno.discard(row_data[2])
    
    with open('obracun.csv','w',encoding='utf-8') as dat:
        for tab in data:
            dat.write(';'.join(tab) + '\n')
    
    print(mno)

if __name__ == '__main__':
    pretvori_v_csv(os.getcwd() + '\\' + '.xls')
