#include <iostream>
#include <sstream>
#include <fstream>
#include <vector>
#include <unordered_map>
#include <set>

using namespace std;

//###################################FUNKCIJE###################################//
wstring ocisti_niz(wstring stavek, wstring vzorec, wstring zamenjava = L" ") {
    int pos = 0;
    while ((pos = stavek.find(vzorec, pos)) != wstring::npos) {
        stavek.replace(pos, vzorec.length(), zamenjava);
        pos += zamenjava.length();
    }
    return stavek;
}

vector<wstring> razdeli(wstring stavek, wchar_t del = L';') {
    /*Funkcija nam vhodni niz razdeli s pomocjo podanega locila*/
    wstringstream tok_niza(stavek);
    vector<wstring> sez_bes;
    wstring beseda;

    while (getline(tok_niza, beseda, del)) {
        sez_bes.push_back(beseda);
    }
    
    return sez_bes;
}

unordered_map<wstring, wstring> ustvariSlovar(const string datoteka) {
    /*Funkcija nam iz vhodne datoteke ustvari slovar*/
    unordered_map<wstring, wstring> slo;
    wifstream dat(datoteka);
    vector<wstring> vek;
    wstring vr;
    while (getline(dat, vr)) {
        vek = razdeli(vr);
        slo.insert({vek[0], vek[1]});
    }
    return slo;
}

//####################################RAZRED####################################//
class Slovar{
    public:
        Slovar(const string datoteka) {
            slo = ustvariSlovar(datoteka);
        }
    // globalne spremenljivke
    unordered_map<wstring, wstring> slo;

    // metode
        set<wstring> vrednosti() {
            /*Metoda nam vrne mnozico vrednosti*/
            set<wstring> sez_vre;
            for (const auto par: slo) {
                sez_vre.insert(par.second);
            }
            return sez_vre;
        }

        set<wstring> kljuci() {
            /*Metoda nam vrne mnozico kljucev*/
            set<wstring> sez_vre;
            for (const auto par: slo) {
                sez_vre.insert(par.first);
            }
            return sez_vre;
        }
};
//##############################################################################//

vector<vector<wstring>> izlusci(const string datoteka) {
    /*  Funkcija nam iz .xml datoteke izlusci vrstice in stolpce
        ter podatke vrne kot vektorje v vektorju - matrika. Ta
        pristop je veliko bolj 'razumljiv' od .findall() metode
        v Pythonu.
    */
    vector<vector<wstring>> podatki;
    wifstream dat(datoteka);
    wstring vr;
    while (getline(dat, vr)) {
        // poiscemo vrstico pri kateri se podatki zacnejo
        if (vr.find(L"<Row") != wstring::npos) {
            vector<wstring> vrstica;
            while (getline(dat,vr)) {
                // Dodamo samo tisto vrstico v kateri so podatki
                if (vr.find(L"<Cell ss:Index=") != wstring::npos) {
                    vrstica.push_back(L"");
                } else if (vr.find(L"<Data ss:Type=\"String\">") != wstring::npos) {
                    // -30 zaradi tega ker je 'prazna' celica dolzine 30
                    // se prej pa ocistimo niz od '\n' oz. '&#10'
                    // in od & znaka oz. &amp;
                    if (vr.find(L"&#10;") != wstring::npos) {
                        vr = ocisti_niz(vr, L"&#10;");
                    } else if (vr.find(L"&amp;") != wstring::npos) {
                        vr = ocisti_niz(vr, L"&amp;", L"&");
                    }
                    vrstica.back() = vr.substr(23,size(vr)-30);
                } else if (vr.find(L"</Row>") != wstring::npos) {
                    podatki.push_back(vrstica);
                    break;
                }
            }
        }
    }
    dat.close();
    return podatki;
}

int main() {
    wofstream dat("obracun.csv");
    // slovar parov stevilk AWB:MRN
    unordered_map<wstring, wstring> slo_parov;
    // matrika podatkov
    vector<vector<wstring>> mat_pod;
    mat_pod = izlusci("obracun_tlm.xls");
    Slovar slovar_parov("carm_scan.csv");
    for (vector vek : mat_pod) {
        if (slovar_parov.kljuci().count(vek[0]) > 0 || 
            slovar_parov.vrednosti().count(vek[2]) > 0 ||
            vek[0] == L"FEDEX" || vek[0] == L"TNT") {
                for (wstring celica : vek) {
                    dat << celica << ";";
                }
                dat << "\n";
        }
    }
    dat.close();
    system("pause");
    return 0;
}
