# funkcije
import os

def zamenjaj_z_none(row):
    return ['None' if not value else value for value in row]

def filter_odstopov(x: tuple, y: set):
    """Funkcija nam filtrira odstope s pomocjo mnozic"""
    if set(x).intersection(y):
        return True
    return False

def preberi_deklarante():
    """Funkcija nam prebere deklarante ter te vrne v tabeli"""
    with open(os.getcwd() + '\\viri\\deklaranti.txt', 'r', encoding='utf-8-sig') as dat:
        tab = []
        for vr in dat:
            tab.append(vr.replace('\n','').split(sep=';'))
    return tab
