from razredi import *
import csv
import time

def zamenjaj_z_none(row):
    return ['None' if not value else value for value in row]

def main():
    with open('trinet.csv', 'r', newline='', encoding="utf-8") as infile:
        reader = csv.reader(infile, delimiter=';')
        rows = [zamenjaj_z_none(row) for row in reader]

    with open('trinetPrecisceno.csv', 'w', newline='',encoding="utf-8") as outfile:
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
                                        objekt.naziv3,
                                        objekt.davek,
                                        objekt.dajatev,
                                        objekt.op,
                                        objekt.st_po,
                                        '22.50',
                                        "None",
                                        objekt.ulica,
                                        objekt.posta,
                                        objekt.kraj,
                                        objekt.naziv1,
                                        objekt.st,
                                        objekt.po,
                                        objekt.vrsta3]}

    obracun = open('obracun2.csv','w',encoding='UTF-8')
    with open('camr_scan.csv','r',encoding='UTF-8') as dat:
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
    
    obracun_koncno = open('obracun3.csv','w',encoding='UTF-8')
    obracun_koncno.write('Con Note;Customer reference;MRN;Date;Opravilna stevilka;Customer name;VAT ID;customer_name;VT-B00;DT-A00;HF;CL0;CL3;CL5;AD0;AF2;CC1;CG1;CL4;CL6;CL8\n')
    with open('obracun2.csv','r',encoding='UTF-8') as dat:
        for vr in dat:
            objekt = strankaObracun(vr.split(sep=';'))
            objekt.glavna_metoda()
            sez =[  objekt.awb,
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
                    objekt.cl8  ]
            if sez[0] != '':
                obracun_koncno.write(';'.join(sez))
    obracun_koncno.close()

if __name__ == '__main__':
    zacetni_cas = time.time()
    main()
    koncni_cas = time.time()
    print(f"{koncni_cas - zacetni_cas:.2f} sek")