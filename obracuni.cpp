#include <iostream>
#include <fstream>
#include <vector>

using namespace std;

wstring ocisti_niz(wstring stavek, wstring vzorec, wstring zamenjava = L" ") {
    int pos = 0;
    while ((pos = stavek.find(vzorec, pos)) != wstring::npos) {
        stavek.replace(pos, vzorec.length(), zamenjava);
        pos += zamenjava.length();
    }
    return stavek;
}

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
    // matrika podatkov
    vector<vector<wstring>> mat_pod;
    mat_pod = izlusci("obracun_tlm.xls");
    for (vector vec : mat_pod) {
        wstring vr;
        for (wstring celica : vec) {
            vr += celica + L";";
        }
        dat << vr << '\n';
    }
    dat.close();
    system("pause");
    return 0;
}
