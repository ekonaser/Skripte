import sqlite3

def zamenjaj_z_none(row):
    return [None if not value else value for value in row]

# Poveži se na obstoječo SQLite bazo
conn = sqlite3.connect('databaseFDX.db')
cursor = conn.cursor()

# Odpri CSV datoteko
with open('podatki.csv', newline='', encoding='utf-8-sig') as dat:

    # Vstavi vsako vrstico v bazo
    for vr in dat:
        vr = zamenjaj_z_none(vr.replace('\n','').split(sep=';'))
        cursor.execute('''
            INSERT INTO obracuni_fdx (
                AWB, MRN, Datum, Naslov, Pošta, Mesto, Pogoji, PrejemnikPosiljatelj, Davcna, Ime_stranke, Davek, Dajatve, Ostali_drzavni_organi, Dodatna_imenovanja, Odstopi, Garancija, Predplacilo, Zacasni_izvoz, Zacasni_uvoz, Prilagoditev, Tranzit, Skladiscenje, Dodatne_storitve, Provizija_za_posredovanje, Davek_od_garancije, Skupaj_stranka_placala
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', vr)

# Potrdi spremembe in zapri povezavo
conn.commit()
conn.close()

print("Podatki iz CSV datoteke so bili uspešno vstavljeni v bazo.")
