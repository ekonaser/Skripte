import sqlite3

# Poveži se na SQLite bazo (ali jo ustvari, če še ne obstaja)
conn = sqlite3.connect('databaseFDX.db')
cursor = conn.cursor()

# Ustvari tabelo z navedenimi stolpci
cursor.execute('''
    CREATE TABLE IF NOT EXISTS obracuni_fdx (
        AWB TEXT,
        MRN TEXT,
        Datum TEXT,
        Naslov TEXT,
        Pošta TEXT,
        Mesto TEXT,
        Pogoji TEXT,
        PrejemnikPosiljatelj TEXT,
        Davcna TEXT,
        Ime_stranke TEXT,
        Davek REAL,
        Dajatve REAL,
        Ostali_drzavni_organi REAL,
        Dodatna_imenovanja REAL,
        Odstopi REAL,
        Garancija REAL,
        Predplacilo REAL,
        Zacasni_izvoz REAL,
        Zacasni_uvoz REAL,
        Prilagoditev REAL,
        Tranzit REAL,
        Skladiscenje REAL,
        Dodatne_storitve REAL,
        Provizija_za_posredovanje REAL,
        Davek_od_garancije REAL,
        Skupaj_stranka_placala REAL
    )
''')

# Potrdi spremembe in zapri povezavo
conn.commit()
conn.close()

print("Baza in tabela uspešno ustvarjeni.")
