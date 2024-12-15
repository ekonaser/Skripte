#include "funkcije.hpp"

//###################################FUNKCIJE###################################//

vector<string> razdeli(string stavek, char del = L';') {
    /*Funkcija nam vhodni niz razdeli s pomocjo podanega locila*/
    stringstream tok_niza(stavek);
    vector<string> sez_bes;
    string beseda;

    while (getline(tok_niza, beseda, del)) {
        sez_bes.push_back(beseda);
    }
    
    return sez_bes;
}

unordered_map<string, string> ustvariSlovar(const string datoteka, short int a = 0, short int b = 1) {
    /*Funkcija nam iz vhodne datoteke ustvari slovar*/
    unordered_map<string, string> slo;
    ifstream dat(datoteka);
    vector<string> vek;
    string vr;
    while (getline(dat, vr)) {
        vek = razdeli(vr);
        slo.insert({vek[a], vek[b]});
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
    unordered_map<string, string> slo;

    // metode
        set<string> vrednosti() {
            /*Metoda nam vrne mnozico vrednosti*/
            set<string> sez_vre;
            for (const auto par: slo) {
                sez_vre.insert(par.second);
            }
            return sez_vre;
        }

        set<string> kljuci() {
            /*Metoda nam vrne mnozico kljucev*/
            set<string> sez_vre;
            for (const auto par: slo) {
                sez_vre.insert(par.first);
            }
            return sez_vre;
        }
};
//##############################################################################//

vector<vector<string>> izlusci(const string datoteka) {
    /*  Funkcija nam iz .xml datoteke izlusci vrstice in stolpce
        ter podatke vrne kot vektorje v vektorju -> matrika. Ta
        pristop je veliko bolj 'razumljiv' od .findall() metode
        v Pythonu.
    */
    vector<vector<string>> podatki;
    ifstream dat(datoteka);
    string vr;
    while (getline(dat, vr)) {
        // poiscemo vrstico pri kateri se podatki zacnejo
        if (vr.find("<Row") != string::npos) {
            vector<string> vrstica;
            while (getline(dat,vr)) {
                // Dodamo samo tisto vrstico v kateri so podatki
                if (vr.find("<Cell ss:Index=") != string::npos) {
                    vrstica.push_back("");
                } else if (vr.find("<Data ss:Type=\"String\">") != string::npos) {
                    // -30 zaradi tega ker je 'prazna' celica dolzine 30
                    // se prej pa ocistimo niz od '\n' oz. '&#10'
                    // in od & znaka oz. &amp;
                    if (vr.find("&#10;") != string::npos) {
                        vr = ocisti_niz(vr, "&#10;");
                    } else if (vr.find("&amp;") != string::npos) {
                        vr = ocisti_niz(vr, "&amp;", "&");
                    }
                    vrstica.back() = vr.substr(23,size(vr)-30);
                } else if (vr.find("</Row>") != string::npos) {
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
    ofstream dat("obracun.csv");
    // slovar parov stevilk AWB:MRN
    Slovar slovar_parov("carm_scan.csv");
    // slovar acc numbers
    unordered_map<string, string> slo_acc_num = ustvariSlovar("navadna_acc.csv", 1, 0);
    // matrika podatkov
    vector<vector<string>> mat_pod;
    vector<vector<string>> mat_fdx;
    vector<vector<string>> mat_tnt;
    mat_pod = izlusci("obracun_tlm.xls");
    
    for (vector vek : mat_pod) {
        if (size(vek[0]) > 12) {
            vek[0].erase(remove(vek[0].begin(), vek[0].end(), ' '), vek[0].end());
            vek[0] = razdeli(vek[0],'-')[1];
        }
        if (slovar_parov.kljuci().count(vek[0]) > 0 || 
            slovar_parov.vrednosti().count(vek[2]) > 0 ||
            vek[0] == "FEDEX" || vek[0] == "TNT") {
                if (slo_acc_num.count(vek[9]) > 0 && size(vek[0]) == 11) {
                    vek[8] = slo_acc_num[vek[9]];
                }
                if (vek[7] == "") {
                    vek[7] = "CPT";
                }
                if (size(vek[0]) == 12) {
                    mat_fdx.push_back(vek);
                }
                if (size(vek[0]) == 11) {
                    mat_tnt.push_back(vek);
                }
                for (string celica : vek) {
                    dat << celica << ";";
                }
                dat << "\n";
        }
    }
    dat.close();
    FDX(mat_fdx);
    system("pause");
    return 0;
}
