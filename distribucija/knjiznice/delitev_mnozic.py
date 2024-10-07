# razred delitev
import itertools
import random
from prikaziDistribucijo import *

class delitevMnozic:
    
    def razdeli_med_deklarante(self, slovar1, slovar2, mno1, mno2, mno3):
        """Metoda nam razdeli podatke globalno.
        Args:
            slovar1 (dict): slovar deklarantov
            slovar2 (dict): slovar izkljucno za odstope in deklarante kateri so izloceni zaradi pogojev
            mno1     (set): mnozica fizicnih oseb
            mno2     (set): mnozica podjetij
            mno3     (set): mnozica odstopov
        """
        # Poiscemo samo AWB stevilke
        # za pozneje lahko das pogoj: len(st) == 11 ali 12
        for k,v in slovar1.items():
            v[2].clear()
        # razdeliti moramo na dve glavni skupini: FIZICNE, OSEBE
        # nato pa te na H7, H7IOSS, FORMAL
        mno_fiz_H7 = {[st for st in nab if st.isdigit()][0] for nab in mno1 if 'H7' in nab}
        mno_fiz_H7IOSS = {[st for st in nab if st.isdigit()][0] for nab in mno1 if 'H7IOSS' in nab}
        mno_fiz_FORMAL = {[niz for niz in nab if niz.isdigit()][0] for nab in mno1 if 'FORMAL' in nab}
        mno_odstopi = {[st for st in nab if st.isdigit()][0] for nab in mno3}
        #################################################################################################
        mno_pod_H7 = {[st for st in nab if st.isdigit()][0] for nab in mno2 if 'H7' in nab}
        mno_pod_H7IOSS = {[st for st in nab if st.isdigit()][0] for nab in mno2 if 'H7IOSS' in nab}
        mno_pod_FORMAL = {[st for st in nab if st.isdigit()][0] for nab in mno2 if 'FORMAL' in nab}
        # ODSTOPI SLOVAR
        ############################################################
        sez_kljucev = list(slovar1.keys())
        random.shuffle(sez_kljucev)
        kljuc = itertools.cycle(sez_kljucev)
        # po logiki najmanjse tabele v tabeli vrednosti kljuca
        najmanjsa_st = len(slovar1[next(kljuc)][2])
        mno_kljucev = set()
        pogoj_formal = False
        pogoj_odstop = False
        # [ODSTOP]
        ############################################################
        while mno_odstopi:
            naslednji_kljuc = next(kljuc)
            if len(slovar1[naslednji_kljuc][2]) < int(slovar1[naslednji_kljuc][1]) and slovar1[naslednji_kljuc][0] == 'odstop':
                slovar1[naslednji_kljuc][2].append(mno_odstopi.pop())
                if slovar1[naslednji_kljuc][0] == 'odstop':
                    pogoj_odstop = True
            if naslednji_kljuc in mno_kljucev and pogoj_odstop != True or len(slovar1[naslednji_kljuc][2]) == int(slovar1[naslednji_kljuc][1]):
                break
            mno_kljucev.add(naslednji_kljuc)
        # [FORMAL]
        ############################################################
        mno_kljucev.clear()
        while mno_fiz_FORMAL:
            naslednji_kljuc = next(kljuc)
            if len(slovar1[naslednji_kljuc][2]) < int(slovar1[naslednji_kljuc][1]) and slovar1[naslednji_kljuc][0] == 'formal':
                slovar1[naslednji_kljuc][2].append(mno_fiz_FORMAL.pop())
                if slovar1[naslednji_kljuc][0] == 'formal':
                    pogoj_formal = True
            # naslednji pogoj se izvede v kolikor smo obsli vse kljuce
            # in nihce od njih ne dela FORMAL posiljk
            if naslednji_kljuc in mno_kljucev and pogoj_formal != True or len(slovar1[naslednji_kljuc][2]) == int(slovar1[naslednji_kljuc][1]):
                break
            mno_kljucev.add(naslednji_kljuc)
        while mno_fiz_FORMAL:
            naslednji_kljuc = next(kljuc)
            if len(slovar1[naslednji_kljuc][2]) < int(slovar1[naslednji_kljuc][1]) and slovar1[naslednji_kljuc][0] != 'formal' and len(slovar1[naslednji_kljuc][2]) == najmanjsa_st:
                slovar1[naslednji_kljuc][2].append(mno_fiz_FORMAL.pop())
                najmanjsa_st = min([len(tab[2]) for tab in slovar1.values() if 'karkoli' in tab or 'odstop' in tab])
            if len(slovar1[naslednji_kljuc][2]) == int(slovar1[naslednji_kljuc][1]):
                slovar2 |= {naslednji_kljuc: slovar1[naslednji_kljuc]}
                slovar1.pop(naslednji_kljuc)
                sez_kljucev = list(slovar1.keys())
                kljuc = itertools.cycle(sez_kljucev)
        # H7
        ############################################################
        while mno_fiz_H7:
            naslednji_kljuc = next(kljuc)
            if len(slovar1[naslednji_kljuc][2]) < int(slovar1[naslednji_kljuc][1]) and slovar1[naslednji_kljuc][0] != 'formal' and len(slovar1[naslednji_kljuc][2]) == najmanjsa_st:
                slovar1[naslednji_kljuc][2].append(mno_fiz_H7.pop())
                najmanjsa_st = min([len(tab[2]) for tab in slovar1.values() if 'karkoli' in tab or 'odstop' in tab])
            if len(slovar1[naslednji_kljuc][2]) == int(slovar1[naslednji_kljuc][1]):
                slovar2 |= {naslednji_kljuc: slovar1[naslednji_kljuc]}
                slovar1.pop(naslednji_kljuc)
                sez_kljucev = list(slovar1.keys())
                kljuc = itertools.cycle(sez_kljucev)
        # H7IOSS
        ############################################################
        while mno_fiz_H7IOSS:
            naslednji_kljuc = next(kljuc)
            if len(slovar1[naslednji_kljuc][2]) < int(slovar1[naslednji_kljuc][1]) and slovar1[naslednji_kljuc][0] != 'formal' and len(slovar1[naslednji_kljuc][2]) == najmanjsa_st:
                slovar1[naslednji_kljuc][2].append(mno_fiz_H7IOSS.pop())
                najmanjsa_st = min([len(tab[2]) for tab in slovar1.values() if 'karkoli' in tab or 'odstop' in tab])
            if len(slovar1[naslednji_kljuc][2]) == int(slovar1[naslednji_kljuc][1]):
                slovar2 |= {naslednji_kljuc: slovar1[naslednji_kljuc]}
                slovar1.pop(naslednji_kljuc)
                sez_kljucev = list(slovar1.keys())
                kljuc = itertools.cycle(sez_kljucev)
        # H7
        ############################################################
        while mno_pod_H7:
            naslednji_kljuc = next(kljuc)
            if len(slovar1[naslednji_kljuc][2]) < int(slovar1[naslednji_kljuc][1]) and slovar1[naslednji_kljuc][0] != 'formal' and len(slovar1[naslednji_kljuc][2]) == najmanjsa_st:
                slovar1[naslednji_kljuc][2].append(mno_pod_H7.pop())
                najmanjsa_st = min([len(tab[2]) for tab in slovar1.values() if 'karkoli' in tab or 'odstop' in tab])
            if len(slovar1[naslednji_kljuc][2]) == int(slovar1[naslednji_kljuc][1]):
                slovar2 |= {naslednji_kljuc: slovar1[naslednji_kljuc]}
                slovar1.pop(naslednji_kljuc)
                sez_kljucev = list(slovar1.keys())
                kljuc = itertools.cycle(sez_kljucev)
        # FORMAL
        ############################################################
        while mno_pod_FORMAL:
            naslednji_kljuc = next(kljuc)
            if len(slovar1[naslednji_kljuc][2]) < int(slovar1[naslednji_kljuc][1]) and slovar1[naslednji_kljuc][0] != 'formal' and len(slovar1[naslednji_kljuc][2]) == najmanjsa_st:
                slovar1[naslednji_kljuc][2].append(mno_pod_FORMAL.pop())
                najmanjsa_st = min([len(tab[2]) for tab in slovar1.values() if tab[0] in {'karkoli', 'odstop'}])
            if len(slovar1[naslednji_kljuc][2]) == int(slovar1[naslednji_kljuc][1]):
                slovar2 |= {naslednji_kljuc: slovar1[naslednji_kljuc]}
                slovar1.pop(naslednji_kljuc)
                sez_kljucev = list(slovar1.keys())
                kljuc = itertools.cycle(sez_kljucev)
        # H7IOSS
        ############################################################
        while mno_pod_H7IOSS:
            naslednji_kljuc = next(kljuc)
            if len(slovar1[naslednji_kljuc][2]) < int(slovar1[naslednji_kljuc][1]) and slovar1[naslednji_kljuc][0] != 'formal' and len(slovar1[naslednji_kljuc][2]) == najmanjsa_st:
                slovar1[naslednji_kljuc][2].append(mno_pod_H7IOSS.pop())
                najmanjsa_st = min([len(tab[2]) for tab in slovar1.values() if 'karkoli' in tab or 'odstop' in tab])
            if len(slovar1[naslednji_kljuc][2]) == int(slovar1[naslednji_kljuc][1]):
                slovar2 |= {naslednji_kljuc: slovar1[naslednji_kljuc]}
                slovar1.pop(naslednji_kljuc)
                sez_kljucev = list(slovar1.keys())
                kljuc = itertools.cycle(sez_kljucev)
        # preveri veckrat ce je pogoj pravi!!! za najmanjsa_st
        # 07.10.2024 - dodan pogoj da jih mece ven iz slovarja
        
        # v naslednjih 4 vrsticah kode dodamo se odstope v slovar2
        sez_odstopi = list(mno_odstopi)
        slovar2['Odstopi'] = ['karkoli', 100, []]
        for awb in sez_odstopi:
            slovar2['Odstopi'][2].append(awb)
        
        self.dist_okno = prikaziDistribucijo(slovar1)
        self.dist_okno.show()