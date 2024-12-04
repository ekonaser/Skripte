import csv
import time
import os

# funkcije
######################################################################################################################
######################################################################################################################

def zamenjaj_z_none(row) -> list:
    return ['None' if not value else value for value in row]

# razred trinet
######################################################################################################################
######################################################################################################################
class strankaTrinet:
    
    def __init__(self, tab: list) -> None:
        """Konstruktor nam podatke iz podane tabele razvrsti kot
            naslednje atribute.
        """
        self.opravilna_stevilka = tab[0]
        self.mrn = tab[1]
        self.mrn_datum = tab[2]
        self.st = tab[3]
        self.po = tab[4]
        self.op = tab[5]
        if tab[6] != 'None' and int(tab[6]) >= 6:
            self.st_po = str((int(tab[6]) - 5) * 8) + '.00'
        else:
            self.st_po = '0'
        self.dajatev = tab[7]
        self.davek = tab[8]
        self.vrsta3 = tab[9]
        if tab[10] == 'None' or 'SIA' in tab[10]:
            self.eori3 = 'SIA5555555'
        else:
            self.eori3 = tab[10]
        self.naziv1 = tab[11]
        self.ulica = tab[12]
        self.posta = tab[13]
        self.kraj = tab[14]
        self.naziv3 = tab[15]

######################################################################################################################
######################################################################################################################

# razred obracun
######################################################################################################################
######################################################################################################################
class strankaObracun:
    
    def __init__(self, tab: list) -> None:
        """Konstruktor nam podatke iz podane tabele razvrsti kot
            naslednje atribute.
        """
        self.awb = tab[0]
        self.mrn = tab[1]
        self.customer_reference = ''
        self.opravilna_stevilka = tab[2].upper()
        self.datum = tab[3].replace('/','-')
        self.customer_name1 = tab[4]
        self.davcna = tab[5]
        self.customer_name2 = tab[6]
        if tab[7] == 'None':
            self.vtb = ''
        else:
            self.vtb = tab[7].replace('.','').replace(',','.')
        if tab[8] == 'None':
            self.dta = ''
        else:
            self.dta = tab[8].replace('.','').replace(',','.')
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
    
    def spremeni_atribute(self) -> None:
        """Metoda nam vse atribute objekta nastavi na prazen niz."""
        for k,v in list(self.__dict__.items()):
            self.__dict__[k] = ''
    
    def cl5_funkcija(self,st1: str,st2: str) -> str:
        """Funkcija nam poračuna vrednosti dveh parametrov in vrne eno vrednost kot niz"""
        if st1 == 'None' or st1 == '':
            st1 = '0'
        if st2 == 'None' or st2 == '':
            st2 = '0'
        st1 = float(st1)
        st2 = float(st2)
        st = st1 + st2
        if st < 50:
            poracunanast = st * 0.3
            if st1 == 0 and st2 == 0:
                return '0'
            elif poracunanast <= 5:
                return '5'
            return f"{poracunanast:.2f}"
        elif st >= 50 and st <= 600:
            return '15'
        return f"{(st * 0.025):.2f}"
    
    def glavna_metoda(self) -> None:
        """Metoda nam spremeni atribute objekta glede na podana pravila.
            Metoda ne vrne ničesar.
        """
        if self.mrn[:4] == '4611' or self.mrn[:4] == '4211':
            self.spremeni_atribute()
            # predčasno zaključimo metodo
            return None
        
        if self.datum == 'None':
            self.datum = _datum
        
        if self.cl4 == 'H3' or self.cl4 == 'H4' and 'STOR' not in self.opravilna_stevilka and 'SANI' not in self.opravilna_stevilka:
            self.opravilna_stevilka = 'ZAČASNI UVOZ 42'
        if self.cl6 == '61' and 'STOR' not in self.opravilna_stevilka and 'SANI' not in self.opravilna_stevilka:
            self.opravilna_stevilka = 'VRAČILA 42'

        if self.hf == 'Da':
            self.vtb = ''
        
        if self.cl8 == 'CPT':
            self.opravilna_stevilka = 'CPT'
        
        elif self.cl8 == 'DDP':
            self.opravilna_stevilka = 'DDP'
        
        if 'STOR' in self.opravilna_stevilka:
            self.opravilna_stevilka = "VRAČILO/STORAGE 42"

        elif 'SANI' in self.opravilna_stevilka:
            self.opravilna_stevilka = 'SANITARC/DRUG VLADNI ORGAN'

        elif 'ROD' in self.opravilna_stevilka:
            self.opravilna_stevilka = 'CPT'
        
        elif 'DDP' in self.opravilna_stevilka:
            self.opravilna_stevilka = 'DDP'
        
        elif 'VRAČ' not in self.opravilna_stevilka and 'ZAČA' not in self.opravilna_stevilka:
            self.opravilna_stevilka = 'CPT'
        
        self.customer_reference = 'MRN ' + self.mrn.replace(' ROD','')[-6:] + ' DNE ' + self.datum
        
        self.cl5 = self.cl5_funkcija(self.vtb,self.dta)
        
        # Naslednji pogoji morajo biti na koncu saj objektu
        # še vedno spreminjamo atribute in nekateri pogoji so
        # odvisni od slednjih
        if self.cg1 != 'None' and self.cg1 == self.customer_name2 and self.opravilna_stevilka == 'CPT' and self.cl0 == '0':
            self.spremeni_atribute()
            return None
        if self.cg1 != 'None' and self.cg1 == self.customer_name2 and self.opravilna_stevilka == 'DDP' and self.cl0 == '0':
            self.spremeni_atribute()
            return None
            
        if self.cg1 != 'None' and self.cl5 == '0' and self.opravilna_stevilka == 'CPT' and self.cl0 == '0':
            self.spremeni_atribute()
            return None
        if self.cg1 != 'None' and self.cl5 == '0' and self.opravilna_stevilka == 'DDP' and self.cl0 == '0':
            self.spremeni_atribute()
            return None
        
        #self.hf = ''
        # self.cl0 je stevilka ki eskalira za 8 krat ko je vec kot 5 zato je ne smes spremeniti
        # poracuna se ze v razredu [razredTrinet]
        self.cl0 = '' if self.cl0 == '0' or self.cl0 == 'None' else self.cl0
        self.cl3 = ''
        # včasih nizi niso isti zato jih ne moreš brisat!!!
        # self.cg1 = ''
        #self.cl4 = ''
        #self.cl6 = ''
        #self.cl8 = '\n'

######################################################################################################################
######################################################################################################################

def main() -> None:
    """"Glavna funkcija skripte"""
    with open('trinet.csv', 'r', newline='', encoding="UTF-8") as infile:
        reader = csv.reader(infile, delimiter=';')
        rows = [zamenjaj_z_none(row) for row in reader]

    with open('trinetPrecisceno.csv', 'w', newline='',encoding="UTF-8") as outfile:
        writer = csv.writer(outfile, delimiter=';')
        writer.writerows(rows)
    
    slo_trinet = {}
    with open('trinetPrecisceno.csv','r',encoding='UTF-8') as dat:
        for vr in dat:
            break
        for vr in dat:
            objekt = strankaTrinet(vr.replace('"','').replace('\n','').split(sep=';'))
            slo_trinet |= {objekt.mrn: [objekt.opravilna_stevilka,
                                        objekt.mrn_datum,
                                        '=IF(E2="DDP","SHIPPER","RECEIVER")',
                                        objekt.eori3,
                                        objekt.naziv1,
                                        objekt.davek,
                                        objekt.dajatev,
                                        objekt.op,
                                        objekt.st_po,
                                        "None",
                                        "None",
                                        objekt.ulica,
                                        objekt.posta,
                                        objekt.kraj,
                                        objekt.naziv3,
                                        objekt.st,
                                        objekt.po,
                                        objekt.vrsta3]}

    obracun = open('obracun.csv','w',encoding='UTF-8')
    # v carm_scan sta lahko samo dva stolpca
    # AWB, MRN
    with open('carm_scan.csv','r',encoding='UTF-8') as dat:
        for vr in dat:
            break
        for vr in dat:
            vr = vr.replace('\n','').split(sep=';')
            obracun.write(vr[0]+';'+vr[1])
            if vr[1] in slo_trinet:
                if slo_trinet[vr[1]][0][-3:].upper() == 'ROD':
                    obracun.write(" ROD;" + ";".join(slo_trinet[vr[1]]))
                else:
                    obracun.write(";" + ";".join(slo_trinet[vr[1]]))
            else:
                obracun.write(';None;None;None;None;None;None;None;None;None;None;None;None;None;None;None;None;None;None')
            obracun.write('\n')
    obracun.close()

    datoteka = _danasnji_datum[0]+' '+_danasnji_datum[1]+' '+_danasnji_datum[2]+' OBRAČUN.csv'
    obracun_koncno = open(datoteka,'w',encoding='UTF-8')
    obracun_koncno.write('Con Note;Customer reference;MRN;Date;Opravilna stevilka;Customer name;VAT ID;customer_name;VT-B00;DT-A00;HF;CL0;CL3;CL5;AD0;AF2;CC1;CG1;CL4;CL6;CL8\n')
    with open('obracun.csv','r',encoding='UTF-8') as dat:
        for vr in dat:
            objekt = strankaObracun(vr.split(sep=';'))
            objekt.glavna_metoda()
            sez = [ objekt.awb,
                    objekt.customer_reference,
                    objekt.mrn,
                    objekt.datum,
                    objekt.opravilna_stevilka,
                    objekt.customer_name1,
                    objekt.davcna,
                    objekt.customer_name2,
                    objekt.vtb,
                    objekt.dta,
                    objekt.hf,
                    objekt.cl0,
                    objekt.cl3,
                    objekt.cl5,
                    objekt.naslov,
                    objekt.posta,
                    objekt.kraj,
                    objekt.cg1,
                    objekt.cl4,
                    objekt.cl6,
                    objekt.cl8 ]
            if sez[0] != '':
                obracun_koncno.write(';'.join(sez))
    obracun_koncno.close()
    # Naslednjo vrstico zakomentiraj v kolikor želiš gledati
    # posredno datoteko 'obracun.csv'
    os.remove(os.getcwd()+'\\'+'obracun.csv')

if __name__ == '__main__':
    _datum = input("Vnesi datum zadnjega delovnega dne v obliki [dd.mm.yyyy]: ")
    _danasnji_datum = [f"{time.gmtime()[2]:02d}", f"{time.gmtime()[1]:02d}", f"{time.gmtime()[0]}"]
    zacetni_cas = time.time()
    main()
    koncni_cas = time.time()
    print(f"{koncni_cas - zacetni_cas:.2f} sek")
