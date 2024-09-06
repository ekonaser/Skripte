# Razredi
from datetime import datetime, timedelta


def zadnji_delovni_dan(date):
    # Pretvori datum v numpy datetime64 objekt
    date = np.datetime64(date)
    
    # Odštej en dan
    vceraj = date - np.timedelta64(1, 'D')
    
    # Če je prejšnji dan vikend, odštej dodatne dni
    if np.is_busday(vceraj) == False:
        vceraj = np.busday_offset(vceraj, -1, roll='backward')
    
    return str(vceraj)[-5:] + '-' + str(vceraj)[:4]

class strankaTrinet:
    
    def __init__(self, tab):
        """Konstruktor nam podatke iz podane tabele razvrsti kot naslednje atribute"""
        self.opravilna_stevilka = tab[0]
        self.lrn = tab[1]
        self.mrn = tab[3]
        self.mrn_datum = tab[4]
        self.st = tab[10]
        self.po = tab[13]
        self.op = tab[19]
        if tab[21] != 'None' and int(tab[21]) >= 6:
            self.st_po = str((int(tab[21]) - 5) * 8)
        else:
            self.st_po = '0'
        self.dajatev = tab[33]
        self.davek = tab[34]
        self.vrsta3 = tab[42]
        if tab[73] == 'None' or tab[73][-1:] == 'O' or 'SIA' in tab[73]:
            self.eori3 = 'SIA5555555'
        else:
            self.eori3 = tab[73]
        self.naziv1 = tab[82]
        self.ulica = tab[77]
        self.posta = tab[78]
        self.kraj = tab[79]
        self.naziv3 = tab[74]

class strankaObracun:
    
    def __init__(self, tab):
        """Konstruktor nam podatke iz podane tabele razvrsti kot naslednje atribute"""
        self.awb = tab[0]
        self.mrn = tab[1]
        self.customer_reference = ''
        self.opravilna_stevilka = tab[2]
        self.datum = tab[3].replace('/','-')
        self.customer_name1 = tab[4]
        self.davcna = tab[5]
        self.customer_name2 = tab[6]
        if tab[7] == 'None':
            self.vtb = ''
        else:
            self.vtb = tab[7]
        if tab[8] == 'None':
            self.dta = ''
        else:
            self.dta = tab[8]
        # samoobdavcitev
        self.hf = tab[9]
        self.cl0 = tab[10]
        self.cl3 = tab[11]
        self.cl5 = tab[12]
        self.naslov = tab[13]
        self.posta = tab[14]
        self.kraj = tab[15]
        self.cg1 = tab[16]
        self.cl4 = tab[17]
        self.cl6 = tab[18]
        self.cl8 = tab[19]
    
    def spremeni_atribute(self):
        for k,v in list(self.__dict__.items()):
            self.__dict__[k] = ''
            
    def cl5_funkcija(self,st1,st2) -> str:
        """Funkcija nam poračuna vrednosti dveh parametrov in vrne eno vrednost kot niz"""
        if st1 == 'None' or st1 == '':
            st1 = '0'
        if st2 == 'None' or st2 == '':
            st2 = '0'
        st1 = float(st1.replace(',',''))
        st2 = float(st2.replace(',',''))
        st = st1 + st2
        if st < 50:
            poracunanast = st * 0.3
            if st1 == 0 and st2 == 0:
                return '0'
            elif poracunanast <= 5:
                return '5'
            return str(poracunanast)
        elif st >= 50 and st <= 600:
            return '15'
        return str(st * 0.025)
    
    def glavna_metoda(self):
        """Metoda nam spremeni atribute objekta. Mozno je tudi da metoda klice druge metode znotraj
            objekta.
        """
        if self.datum == 'None':
            self.datum = '9-3-2024'
        
        if self.cl4 == 'H3' or self.cl4 == 'H4':
            self.opravilna_stevilka = 'ZAČASNI UVOZ 42'
        if self.cl6 == '61':
            self.opravilna_stevilka = 'VRAČILA 42'

        if self.hf == 'Da':
            self.vtb = ''
        
        if self.cl8 == 'CPT':
            self.opravilna_stevilka = 'CPT'
        
        elif self.cl8 == 'DDP':
            self.opravilna_stevilka = 'DDP'
        
        if 'ROD' in self.opravilna_stevilka.upper():
            self.opravilna_stevilka = 'CPT'
        
        elif 'DDP' in self.opravilna_stevilka.upper():
            self.opravilna_stevilka = 'DDP'
        
        elif 'SANI' in self.opravilna_stevilka.upper():
            self.opravilna_stevilka = 'SANITARC/DRUG VLADNI ORGAN'
        
        elif 'VRAČ' not in self.opravilna_stevilka and 'ZAČA' not in self.opravilna_stevilka:
            self.opravilna_stevilka = 'CPT'
        
        self.customer_reference = 'MRN ' + self.mrn.replace(' ROD','')[-6:] + ' DNE ' + self.datum
        
        self.cl5 = self.cl5_funkcija(self.vtb,self.dta)
        
        if self.mrn[:4] == '4611' or self.mrn[:4] == '4211':
            self.spremeni_atribute()
        if self.cg1 != 'None' and self.cg1 == self.customer_name2 and self.opravilna_stevilka == 'CPT':
            self.spremeni_atribute()
        if self.cg1 != 'None' and self.cg1 == self.customer_name2 and self.opravilna_stevilka == 'DDP':
            self.spremeni_atribute()
            
        if self.cg1 != 'None' and self.cl5 == '0' and self.opravilna_stevilka == 'CPT':
            self.spremeni_atribute()
        if self.cg1 != 'None' and self.cl5 == '0' and self.opravilna_stevilka == 'DDP':
            self.spremeni_atribute()
        
        self.hf = ''
        self.cl0 = ''
        self.cl3 = ''
        # vcasih nizi niso isti zato jih ne mores brisat!!!
        # self.cg1 = ''
        self.cl4 = ''
        self.cl6 = ''
        self.cl8 = '\n'